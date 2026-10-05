module Rbt.Spreadsheet.Model

open System

/// Cell address (0-indexed)
type CellAddress = {
    Row: int
    Col: int
}

/// Cell value types
type CellValue =
    | Empty
    | Number of float
    | String of string
    | Boolean of bool
    | Error of string
    | Formula of formula: string * cachedValue: CellValue

let unwrapFormulaValue (value: CellValue) : CellValue =
    let mutable current = value
    let mutable doneUnwrapping = false
    while not doneUnwrapping do
        match current with
        | Formula (_, cached) -> current <- cached
        | _ -> doneUnwrapping <- true
    current

/// Border style for cell edges
type BorderStyle = {
    Style: string  // "thin", "medium", "thick", "dashed", "dotted", "double"
    Color: string option
}

/// How an Excel border style is drawn with CSS: (width in px, CSS line
/// style), or None for no border ("none", ""). Excel's styles have no exact
/// CSS twins; each gets the nearest one of the same weight (thin 1px,
/// medium 2px, thick 3px; `double` needs 3px for its two rules and the gap).
let borderCssParts (style: string) : (int * string) option =
    match (if isNull style then "" else style).ToLowerInvariant() with
    | "" | "none" -> None
    | "thin" -> Some (1, "solid")
    | "medium" -> Some (2, "solid")
    | "thick" -> Some (3, "solid")
    | "double" -> Some (3, "double")
    | "hair" | "dotted" -> Some (1, "dotted")
    | "dashed" | "dashdot" | "dashdotdot" -> Some (1, "dashed")
    | "mediumdotted" -> Some (2, "dotted")
    | "mediumdashed" | "mediumdashdot" | "mediumdashdotdot" | "slantdashdot" -> Some (2, "dashed")
    | "solid" -> Some (1, "solid")   // CSS names, as older callers passed them
    | _ -> Some (1, "solid")

/// Excel's order of border styles where two cells state a border for the edge
/// they share. NOT the stroke width: `double` outranks `thick` although
/// `thick` paints fatter. None < hair < dotted < the dashed family < thin <
/// medium < thick < double.
let borderPriority (style: string) : int =
    match (if isNull style then "" else style).ToLowerInvariant() with
    | "" | "none" -> 0
    | "hair" -> 1
    | "dotted" -> 2
    | "dashed" -> 3
    | "dashdot" -> 4
    | "dashdotdot" -> 5
    | "mediumdashed" -> 6
    | "mediumdashdot" -> 7
    | "mediumdashdotdot" -> 8
    | "slantdashdot" -> 9
    | "thin" -> 10
    | "medium" -> 11
    | "thick" -> 12
    | "double" -> 13
    | _ -> 10  // unknown: as thin

/// The one border drawn on the edge two cells share. `first` is what the
/// cell above / to the left states for it (its bottom / right border),
/// `second` what the cell below / to the right states (its top / left
/// border). The stronger style wins (`borderPriority`); of two equally strong
/// ones the later cell's - Excel stores a border applied to an edge on both
/// cells, and the cell written last (below / right) is the one it shows.
let resolveSharedBorder (first: BorderStyle option) (second: BorderStyle option) : BorderStyle option =
    let drawn (border: BorderStyle option) =
        match border with
        | Some b when borderPriority b.Style > 0 -> border
        | _ -> None
    match drawn first, drawn second with
    | None, other -> other
    | other, None -> other
    | Some a, Some b -> if borderPriority a.Style > borderPriority b.Style then Some a else Some b

/// The value of a CSS `border-<side>` declaration for a cell border:
/// "1px solid #000000" - always with a width and a line style, without
/// which CSS draws nothing.
let borderCss (border: BorderStyle) : string =
    match borderCssParts border.Style with
    | Some (width, line) -> string width + "px " + line + " " + (border.Color |> Option.defaultValue "#000000")
    | None -> "none"

/// Cell formatting
type CellFormat = {
    Bold: bool
    Italic: bool
    Underline: bool
    /// Distinguishes `<u val="double"/>` (two parallel strokes) from the
    /// default single underline. Renderers draw two strokes when true.
    UnderlineDouble: bool
    Strikethrough: bool
    FontName: string option
    TextColor: string option
    BgColor: string option
    FontSize: int option
    HorizontalAlign: string option  // "left", "center", "right", "justify"
    VerticalAlign: string option    // "top", "middle", "bottom"
    WrapText: bool
    /// Excel `<alignment indent="N"/>`. Each unit ≈ 3 character widths and
    /// pads inward from the cell's leading edge: from the left for left /
    /// distributed alignment, from the right for right alignment.
    Indent: int
    /// `<alignment shrinkToFit="1"/>` — when set, the cell's text is
    /// rendered at a reduced font size if its natural size would overflow
    /// the cell width. Excel never increases the size; only shrinks.
    ShrinkToFit: bool
    /// `<alignment textRotation="N"/>` degrees. 0 = horizontal (default),
    /// 1..90 = CCW (text rises left-to-right), 91..180 = CW (encoded as
    /// 90 + |angle|), 255 = stacked.
    TextRotation: int
    NumberFormat: string option
    BorderTop: BorderStyle option
    BorderBottom: BorderStyle option
    BorderLeft: BorderStyle option
    BorderRight: BorderStyle option
    /// Optional `<diagonal>` border style + color. Combined with
    /// DiagonalUp / DiagonalDown to choose which one or two diagonal
    /// strokes Excel paints across the cell rect.
    BorderDiagonal: BorderStyle option
    /// Stroke from bottom-left to top-right of the cell rectangle.
    DiagonalUp: bool
    /// Stroke from top-left to bottom-right of the cell rectangle.
    DiagonalDown: bool
    /// Excel auto-filter dropdown button on the right edge of the cell.
    /// Set for cells in the header row of a table with an <autoFilter> child.
    FilterButton: bool
    /// Small red triangle in the top-right corner of the cell. Excel
    /// renders this whenever the cell has a comment / note attached.
    /// Interactive viewers show the comment text as a tooltip on hover.
    HasComment: bool
    /// Small downward arrow on the right edge of the cell. Excel renders
    /// this for cells inside a `<dataValidation type="list">` range when
    /// the cell is selected; for static snapshot rendering we always show
    /// it so the visual cue carries through to PDF / PNG exports.
    HasListDropdown: bool
    /// Conditional-formatting data bar: (fillFraction 0..1, RRGGBB
    /// color). Renderers draw a horizontal bar from the cell's left
    /// edge spanning fraction × width, behind the text.
    DataBar: (float * string) option
    /// Conditional-formatting icon set: (iconIndex, iconCount). The
    /// renderer draws a small low/mid/high indicator at the cell's left.
    IconSet: (int * int) option
}

/// Default format
let defaultFormat = {
    Bold = false
    Italic = false
    Underline = false
    UnderlineDouble = false
    Strikethrough = false
    FontName = None
    TextColor = None
    BgColor = None
    FontSize = None
    HorizontalAlign = None
    VerticalAlign = None
    WrapText = false
    Indent = 0
    ShrinkToFit = false
    TextRotation = 0
    NumberFormat = None
    BorderTop = None
    BorderBottom = None
    BorderLeft = None
    BorderRight = None
    BorderDiagonal = None
    DiagonalUp = false
    DiagonalDown = false
    FilterButton = false
    HasComment = false
    HasListDropdown = false
    DataBar = None
    IconSet = None
}

/// Cell comment/note
type CellComment = {
    Text: string
    Author: string option
}

/// Vertical glyph alignment within a run / cell. Renderers shift the
/// baseline and scale the glyph size for super/subscript. Maps from
/// Excel's `<vertAlign val="superscript"/>` element.
type VertAlign =
    | VertAlignBaseline
    | VertAlignSuperscript
    | VertAlignSubscript

/// Per-run formatting override for rich-text cells. Each field is optional;
/// when None the run inherits from the cell's base CellFormat. Originates
/// from `<rPr>` children of `<r>` elements inside `<si>` shared strings.
type RunOverride = {
    Bold: bool option
    Italic: bool option
    Underline: bool option
    UnderlineDouble: bool option
    Strikethrough: bool option
    Color: string option       // hex with leading '#'
    FontName: string option
    FontSize: float option     // points
    VertAlign: VertAlign option
}

let emptyRunOverride = {
    Bold = None
    Italic = None
    Underline = None
    UnderlineDouble = None
    Strikethrough = None
    Color = None
    FontName = None
    FontSize = None
    VertAlign = None
}

/// Cell data
type CellData = {
    Value: CellValue
    Format: CellFormat
    Comment: CellComment option
}

let private cellDataIsFormula (cell: CellData) =
    match cell.Value with
    | Formula _ -> true
    | _ -> false

let cellsHaveFormulaCells (cells: Map<CellAddress, CellData>) =
    cells |> Map.exists (fun _ cell -> cellDataIsFormula cell)

/// Selection types
type Selection =
    | SingleCell of row: int * col: int
    | Range of startRow: int * startCol: int * endRow: int * endCol: int

/// Clipboard data for copy/paste operations
type ClipboardData = {
    Cells: (int * int * CellData) list  // (relRow, relCol, data)
    IsCut: bool
    SourceStartRow: int
    SourceStartCol: int
    /// Bottom-right corner of the copied/cut rectangle (inclusive). Together
    /// with the start it gives the block's size — needed to tile a paste and
    /// to re-point references after a cut is pasted.
    SourceEndRow: int
    SourceEndCol: int
}

/// Sort state for a column
type SortState = {
    Column: int
    Ascending: bool
}

/// The comparison of a `cellIs` conditional format or a data validation.
type CompareOperator =
    | OpGreaterThan
    | OpLessThan
    | OpGreaterOrEqual
    | OpLessOrEqual
    | OpEqual
    | OpNotEqual
    | OpBetween
    | OpNotBetween

/// Where a colour-scale stop, a data-bar end or an icon-set step sits on the
/// value axis of the rule's range (a `<cfvo>`).
type CfThresholdKind =
    /// Lowest / highest value in the range.
    | CfvoMin
    | CfvoMax
    /// A literal number.
    | CfvoNumber
    /// A percentage of the way from the lowest to the highest value.
    | CfvoPercent
    /// The value at this percentile of the range's numbers.
    | CfvoPercentile
    /// A formula, evaluated once for the rule.
    | CfvoFormula

type CfThreshold = {
    Kind: CfThresholdKind
    /// The number / percentage / formula text; None for min and max.
    Value: string option
}

type TextRuleKind =
    | TextContains
    | TextNotContains
    | TextBeginsWith
    | TextEndsWith

/// Conditional formatting rule. The first five cases are the literal rules
/// the editor's dialog creates; the rest mirror the OOXML rule types. A
/// formula in a rule is written for the top-left cell of the rule's first
/// range and is re-pointed for every other cell (relative references move,
/// `$`-locked ones stay).
type ConditionalFormatRule =
    | GreaterThan of float
    | LessThan of float
    | Between of float * float
    | Equal of float
    | ContainsText of string
    /// `cellIs`: the cell's value compared with one operand (two for
    /// between / not between), each a formula.
    | CellIs of CompareOperator * string * string option
    /// `expression`: formatted where the formula is true.
    | Expression of string
    /// Case-insensitive text test.
    | TextRule of TextRuleKind * string
    /// `top10`: the `rank` largest (or smallest) values; `percent` makes the
    /// rank a percentage of the range's count.
    | TopBottom of rank: int * percent: bool * bottom: bool
    /// `aboveAverage`: above (or below) the mean, optionally including it,
    /// or beyond `stdDev` standard deviations from it when stdDev > 0.
    | AboveAverage of above: bool * orEqual: bool * stdDev: int
    | DuplicateValues
    | UniqueValues
    /// Blank cells (true) / non-blank cells (false).
    | Blanks of bool
    /// Cells holding an error (true) / no error (false).
    | Errors of bool
    /// `timePeriod` ("a date occurring"): the cell holds a date in the
    /// period, counted from today - "today", "yesterday", "tomorrow",
    /// "last7Days", "thisWeek", "lastWeek", "nextWeek" (weeks start on
    /// Sunday), "thisMonth", "lastMonth", "nextMonth".
    | TimePeriod of string
    /// 2- or 3-colour scale: stops in ascending order with their colours.
    | ColorScale of (CfThreshold * string) list
    /// Data bar from `low` to `high`, drawn in `color`.
    | DataBar of low: CfThreshold * high: CfThreshold * color: string
    /// Icon set: `family` is the OOXML name ("3TrafficLights1", "5Arrows");
    /// one threshold per icon (the first is the floor).
    | IconSet of family: string * thresholds: CfThreshold list * reverse: bool

/// Conditional format (rule + style to apply)
type ConditionalFormat = {
    Rule: ConditionalFormatRule
    Style: CellFormat
}

/// What a data validation of the comparison kind measures.
type ValidationKind =
    | WholeKind
    | DecimalKind
    | DateKind
    | TimeKind
    | TextLengthKind

/// Data-validation rule (what a cell is allowed to contain).
type ValidationRule =
    /// Allowed values (a dropdown). Case-insensitive match.
    | ListValues of string list
    /// Numeric range [min, max] (inclusive); `None` = unbounded.
    | WholeNumber of float option * float option
    | DecimalRange of float option * float option
    /// String length range.
    | TextLength of float option * float option
    /// Date (Excel serial) range.
    | DateRange of float option * float option
    /// Any comparison, including not-between and not-equal, of a whole
    /// number / decimal / date / time / text length with operands that are
    /// numbers or formulas (`=TODAY()`, `=$B$1`).
    | Comparison of ValidationKind * CompareOperator * string * string option
    /// The entry is valid when this formula is true.
    | CustomFormula of string
    /// Allowed values read from cells when they are needed: a range
    /// (`$A$1:$A$9`, `Lists!$B$2:$B$20`) or a defined name, as a formula
    /// without the leading `=`. A relative range is written for the rule's
    /// anchor cell, like any formula of a rule.
    | ListSource of string
    /// Every value is allowed (Excel's "any value"): the validation only
    /// carries an input message.
    | AnyValue

/// What happens when an entry fails validation. Excel's "stop" refuses the
/// entry; "warning" and "information" only tell the user and accept it.
type ValidationStyle =
    | StopStyle
    | WarningStyle
    | InformationStyle

type DataValidation = {
    Rule: ValidationRule
    /// Shown when an entered value fails the rule.
    ErrorMessage: string option
    /// Shown as a tooltip while the cell is selected.
    InputMessage: string option
    /// Captions of the error alert and of the input message.
    ErrorTitle: string option
    InputTitle: string option
    Style: ValidationStyle
    /// Excel's "show error alert after invalid data is entered". When off
    /// the rule is advisory only: nothing is refused and nothing is shown.
    ShowError: bool
    /// Excel's "ignore blank". Carried for round trips: an emptied cell is
    /// accepted either way.
    AllowBlank: bool
    /// Excel's "show input message when cell is selected". When off the
    /// input title and message are kept but not shown.
    ShowInput: bool
    /// A list rule's "in-cell dropdown": whether the cell offers its values
    /// in a dropdown (the rule holds either way).
    ShowDropdown: bool
    /// The cell the rule's formulas are written for (the top-left of the
    /// validated range); None = the validated cell itself.
    Anchor: CellAddress option
}

/// A validation with the given rule and Excel's defaults for the rest.
let validationOf (rule: ValidationRule) : DataValidation =
    { Rule = rule; ErrorMessage = None; InputMessage = None; ErrorTitle = None; InputTitle = None
      Style = StopStyle; ShowError = true; AllowBlank = true; ShowInput = true; ShowDropdown = true; Anchor = None }

/// Merged cell range (primary cell is top-left)
type MergedCell = {
    StartRow: int
    StartCol: int
    EndRow: int
    EndCol: int
}

// ---------------------------------------------------------------------------
// Print / page setup
// ---------------------------------------------------------------------------

type PrintOrientation =
    | PortraitOrientation
    | LandscapeOrientation

/// Walk order for 2D pagination. `DownThenOver` (Excel default) iterates
/// every row-slice within the first column-slice before moving on to the
/// next column-slice. `OverThenDown` does the inverse.
type PageOrder =
    | DownThenOver
    | OverThenDown

/// Print behavior for cell comments. Excel default is `CellCommentsNone`
/// (don't print). `CellCommentsAsDisplayed` shows the red corner marker,
/// `CellCommentsAtEnd` would list comments after the printed pages.
type CellCommentsPrint =
    | CellCommentsNone
    | CellCommentsAsDisplayed
    | CellCommentsAtEnd

/// Page margins, all in *inches* (matching how OOXML encodes them and how
/// Excel exposes them to users; the PDF backend converts to points).
type PageMargins = {
    Top: float
    Bottom: float
    Left: float
    Right: float
    Header: float
    Footer: float
}

/// Boolean print options from `<printOptions>`.
type PrintOptions = {
    /// Print gridlines under cell content.
    GridLines: bool
    /// Print row numbers and column letters as headings.
    Headings: bool
    /// Center the printed range horizontally on the page.
    HorizontalCentered: bool
    /// Center it vertically.
    VerticalCentered: bool
}

/// Header / footer strings carry Excel's `&L &C &R` section markers plus
/// field codes — `&P` = page number, `&N` = total pages, `&A` = sheet
/// name, `&F` = file name, `&D` = date, `&T` = time, `&&` = literal `&`.
/// The PDF backend substitutes these per page.
type HeaderFooter = {
    OddHeader: string option
    OddFooter: string option
    EvenHeader: string option
    EvenFooter: string option
    FirstHeader: string option
    FirstFooter: string option
    /// When true, EvenHeader/EvenFooter apply to even-numbered pages.
    DifferentOddEven: bool
    /// When true, FirstHeader/FirstFooter apply to the first page.
    DifferentFirst: bool
    /// When true (Excel default), H/F text scales with the page's print
    /// scale. When false, H/F always renders at its raw point size.
    ScaleWithDoc: bool
}

/// Inclusive 0-indexed cell range. Used for print area and print titles.
/// Distinct from `MergedCell` (same shape, different meaning) to keep the
/// type system honest and avoid record-literal ambiguity.
type CellRange = {
    From: CellAddress
    To: CellAddress
}

/// A conditional format as it is defined: for one or more ranges, not per
/// cell. A whole-column rule is the range from row 0 to the last sheet row —
/// it is never clipped to the used range. Rules apply in ascending
/// `Priority`; a matching rule with `StopIfTrue` hides the rules after it.
/// `Id` names the rule on its sheet for as long as the workbook is open: it
/// stays with the rule through edits of it and of the other rules, through
/// undo and redo and through inserted and deleted rows. A file has no ids:
/// an imported sheet's rules are numbered 1, 2, 3 ... in order of precedence.
type ConditionalFormatRange = {
    Id: int
    Ranges: CellRange list
    Format: ConditionalFormat
    Priority: int
    StopIfTrue: bool
}

/// A data validation as it is defined: for one or more ranges, not per cell.
/// A whole-column validation is the range from row 0 to the last sheet row -
/// it is never clipped to the used range, and no cell is made for it. A cell
/// is covered by at most one validation. The formulas of the rule are written
/// for `Validation.Anchor`, or for the top-left cell of the first range when
/// that is None.
type ValidationRange = {
    /// Names the validation among those of its sheet (the host API addresses
    /// validations by it). Unique within the sheet; never reused while the
    /// sheet lives. Validations read from a file are numbered 1, 2, ... in
    /// file order.
    Id: int
    Ranges: CellRange list
    Validation: DataValidation
}

/// Excel table style reference (e.g. `TableStyleMedium7` + show-stripes
/// toggles). Mirrors `Xlsx.TableStyleInfo` so the model carries this
/// without forcing a dependency on the Office.Formats reader.
type TableStyle = {
    Name: string
    ShowFirstColumn: bool
    ShowLastColumn: bool
    ShowRowStripes: bool
    ShowColumnStripes: bool
    /// Pre-resolved header fill / banded-row fill (hex). Set for tables whose
    /// styling is applied at render time from `model.Tables` (e.g. tables made
    /// by `createTable`); left None for imported tables whose runner bakes the
    /// theme-accent styling straight into each cell's format. `applyTableStyles`
    /// only acts when `HeaderFill` is set, so it never double-styles the corpus.
    HeaderFill: string option
    StripeFill: string option
}

/// One table region on a sheet (i.e., `xl/tables/tableN.xml`). Coordinates
/// are 1-indexed to match the XLSX schema and the table-overlay helpers
/// that consume this. `HeaderRowCount` is 0 when the table has no header
/// row (rare); otherwise 1 (effectively always 1 for the corpus we
/// handle). `AutoFilter` is the 1-indexed (startCol, endCol) range that
/// gets dropdown buttons on its header cells. `Name` is the structured-ref
/// identifier; `HasTotals` marks a totals row occupying the last row.
type Table = {
    Name: string
    StartRow: int
    StartCol: int
    EndRow: int
    EndCol: int
    HeaderRowCount: int
    Style: TableStyle option
    AutoFilter: (int * int) option
    HasTotals: bool
}

/// Per-sheet page setup. All fields have sensible defaults so a model
/// without explicit setup still has a usable PageSetup record. Page
/// dimensions are in points (1 pt = 1/72 inch).
type PageSetup = {
    PageWidth: float
    PageHeight: float
    /// The OOXML paper-size code the dimensions stand for (1 = Letter,
    /// 9 = A4 ...; see `paperSizePoints`), kept so that a size the editor
    /// has no dimensions for goes back to the file as it came. None: the
    /// file names no paper (Excel then takes Letter).
    PaperSize: int option
    Orientation: PrintOrientation
    Margins: PageMargins
    /// When `FitToPage` is true, the printed range is scaled to fit at
    /// most `FitToWidth` × `FitToHeight` pages.
    FitToPage: bool
    FitToWidth: int
    FitToHeight: int
    /// Manual scale percent (10..400). Used when `FitToPage` is false.
    Scale: int
    Options: PrintOptions
    HeaderFooter: HeaderFooter
    /// Manual page breaks: 1-based row / column numbers; the break sits
    /// *after* the given row / column. Empty when there are no manuals.
    ManualRowBreaks: int list
    ManualColBreaks: int list
    /// 0-indexed inclusive range to print; None = print the whole used
    /// range. From `_xlnm.Print_Area`.
    PrintArea: CellRange option
    /// Rows to repeat at the top of every page (inclusive, 0-indexed).
    /// From the row part of `_xlnm.Print_Titles`. None = no repeat.
    RepeatRows: (int * int) option
    /// Columns to repeat on the left of every page (inclusive, 0-indexed).
    /// From the col part of `_xlnm.Print_Titles`.
    RepeatCols: (int * int) option
    /// Page number to display for the first printed page; `&P` substitutes
    /// this value on page 1 and increments from there. Excel-default is
    /// 1; only honored when `UseFirstPageNumber` is true.
    FirstPageNumber: int
    /// `useFirstPageNumber="1"` opts in to FirstPageNumber; when false,
    /// Excel ignores the value and numbers from 1.
    UseFirstPageNumber: bool
    /// Direction of 2D pagination walk. Default DownThenOver.
    PageOrder: PageOrder
    /// When true, every fill renders as white and every text / border /
    /// stroke renders as black. Matches Excel's `blackAndWhite="1"` print
    /// option.
    BlackAndWhite: bool
    /// How cell comments appear in print output. Default `CellCommentsNone`
    /// matches Excel's default of suppressing comments in print.
    CellComments: CellCommentsPrint
}

let defaultPageMargins = {
    Top = 0.75; Bottom = 0.75; Left = 0.7; Right = 0.7
    Header = 0.3; Footer = 0.3
}

let defaultPrintOptions = {
    GridLines = false; Headings = false
    HorizontalCentered = false; VerticalCentered = false
}

let defaultHeaderFooter = {
    OddHeader = None; OddFooter = None
    EvenHeader = None; EvenFooter = None
    FirstHeader = None; FirstFooter = None
    DifferentOddEven = false; DifferentFirst = false
    ScaleWithDoc = true
}

let defaultPageSetup = {
    // US Letter at 72 DPI = 612 × 792 pt.
    PageWidth = 612.0
    PageHeight = 792.0
    PaperSize = None
    Orientation = PortraitOrientation
    Margins = defaultPageMargins
    FitToPage = false
    FitToWidth = 1
    FitToHeight = 1
    Scale = 100
    Options = defaultPrintOptions
    HeaderFooter = defaultHeaderFooter
    ManualRowBreaks = []
    ManualColBreaks = []
    PrintArea = None
    FirstPageNumber = 1
    UseFirstPageNumber = false
    PageOrder = DownThenOver
    BlackAndWhite = false
    CellComments = CellCommentsNone
    RepeatRows = None
    RepeatCols = None
}

/// Standard OOXML paper-size codes → (width, height) in points (portrait).
/// Defaults to Letter when the code is unknown.
let paperSizePoints (code: int) : float * float =
    match code with
    | 1 -> 612.0, 792.0      // Letter   (8.5 × 11 in)
    | 3 -> 792.0, 1224.0     // Tabloid  (11 × 17 in)
    | 7 -> 522.0, 756.0      // Executive (7.25 × 10.5 in)
    | 5 -> 612.0, 1008.0     // Legal    (8.5 × 14 in)
    | 8 -> 841.68, 1190.55   // A3       (297 × 420 mm)
    | 9 -> 595.27, 841.89    // A4       (210 × 297 mm)
    | 11 -> 419.53, 595.27   // A5       (148 × 210 mm)
    | 13 -> 728.5, 1031.81   // B5 ISO   (176 × 250 mm)
    | 70 -> 595.27, 841.89   // A4 alt
    | _ -> 612.0, 792.0      // unknown → Letter

/// Spreadsheet model (forward declaration for Command type)
/// Where a drawing object sits on the sheet (as xdr:twoCellAnchor /
/// oneCellAnchor / absoluteAnchor). Rows and columns are 0-based; offsets,
/// positions and sizes are pixels at 100% (9525 EMU each).
type DrawingAnchor = {
    /// "twoCell" | "oneCell" | "absolute"
    Kind: string
    FromRow: int
    FromCol: int
    /// Offset of the top-left corner inside its cell.
    FromOffsetX: float
    FromOffsetY: float
    /// The bottom-right corner of a two-cell anchor.
    ToRow: int option
    ToCol: int option
    ToOffsetX: float
    ToOffsetY: float
    /// Position of an absolute anchor.
    X: float
    Y: float
    /// Size given by the anchor (one-cell, absolute) or by the object itself.
    Width: float option
    Height: float option
    /// How a two-cell anchor follows its cells: "twoCell" (move and size,
    /// the default), "oneCell" (move only) or "absolute" (stay put).
    EditAs: string option
}

type DrawingContent =
    /// A picture: content type ("image/png"), file extension, data.
    | Picture of contentType: string * extension: string * data: byte[]
    /// An object the editor does not model (chart, shape, group, ...), kept
    /// as the XML of its anchor element: kind ("chart" | "shape" | "other").
    | KeptObject of kind: string * xml: string

/// A picture or other drawing object on a sheet.
type SheetDrawing = {
    /// What the object is known by on its sheet: it keeps the number while
    /// it is moved, sized or restacked. A file's objects are numbered 1..n
    /// back to front when it is opened; 0 = not numbered.
    Id: int
    Anchor: DrawingAnchor
    Name: string option
    Description: string option
    Content: DrawingContent
}

/// A threaded comment: a comment with an author, a date and replies.
type CommentThread = {
    Id: string
    Author: string option
    /// Id of the workbook's person entry behind the author.
    PersonId: string option
    Date: string option
    Text: string
    Replies: CommentThread list
    /// The thread is marked resolved (`done="1"` on its first comment).
    Done: bool
}

/// A hyperlink on a range of cells: an address, or a place in the workbook
/// (`IsInternal`, e.g. "Sheet2!A1").
type SheetHyperlink = {
    Range: CellRange
    Target: string
    IsInternal: bool
    /// The screen tip shown while the pointer is over the link (`tooltip`).
    Tooltip: string option
    /// The text the file says the link shows (`display`); the cells hold
    /// their own text, this is carried beside them.
    Display: string option
}

/// The default cell style of a run of columns (`<col min max style>`): what
/// a cell of those columns looks like when it has no format of its own.
type ColumnStyle = {
    /// First and last column, 0-based, inclusive.
    First: int
    Last: int
    Format: CellFormat
    /// The style in the file the sheet came from (an index into what the
    /// file bridge kept), or -1 for a style set in the editor.
    Source: int
}

/// A what-if data table: Excel fills `Range` by putting each row / column
/// header value into the input cells and recalculating. It is not an
/// expression; the editor does not evaluate it. The cells of the range hold
/// the results the file had, and a recalculation leaves them alone.
type DataTable = {
    Range: CellRange
    TwoDimensional: bool
    /// A one-input table whose input runs along the row.
    RowInput: bool
    /// The input cells, as A1 references.
    Input1: string
    Input2: string option
    Input1Deleted: bool
    Input2Deleted: bool
    CalculateAlways: bool
}

/// A range that something the editor does not model sits on (a pivot table
/// of the file): kept so that it follows inserted and deleted rows / columns.
type KeptRange = {
    /// What the range belongs to (for a pivot table, the path of its part).
    Key: string
    Range: CellRange
}

/// What one column of an AutoFilter lets through.
type FilterCriteria =
    /// The cells whose displayed text is one of `values` (compared without
    /// regard to case), and the blank cells when `blanks`.
    | FilterValues of values: Set<string> * blanks: bool
    /// One or two comparisons (never between / not between), joined by AND
    /// (`matchAll`) or OR. An operand is a number or text; with "equal" and
    /// "not equal" text may hold the wildcards `*` and `?` (`~` escapes one),
    /// which is how "contains", "begins with" and "ends with" are written.
    | FilterCustom of conditions: (CompareOperator * string) list * matchAll: bool
    /// The `count` largest (`bottom`: smallest) numbers of the column, or
    /// that percentage of them.
    | FilterTop of count: int * percent: bool * bottom: bool
    /// The cells filled with `color` ("#RRGGBB"; "" = no fill), or whose
    /// text has that colour when `font` ("" = the automatic colour).
    | FilterColor of color: string * font: bool
    /// "aboveAverage" | "belowAverage": numbers above / below the column's mean.
    | FilterDynamic of kind: string
    /// A filter of a file that the editor does not evaluate (date groups,
    /// the other dynamic filters), as the XML of its `<filterColumn>`
    /// children: it hides no row by itself, the rows the file had hidden
    /// stay hidden until the filter is applied again, and it is written back.
    | FilterKept of xml: string

/// The filter of one column of the sheet's AutoFilter.
type ColumnFilter = {
    /// 0-based sheet column.
    Column: int
    Criteria: FilterCriteria
}

/// The sheet's AutoFilter: the range it covers (its first row is the header
/// row and is never hidden). `Implicit` marks one that was made by filtering
/// a column without setting a range first (the grid's own filter buttons):
/// it goes away with its last column filter.
type SheetFilter = {
    Range: CellRange
    Implicit: bool
}

/// Which sheet of which opened file a model is: `Book` identifies the import
/// (`XlsxBridge.XlsxSource.Id`), `Sheet` is the sheet's position in that file.
type SheetOrigin = {
    Book: int
    Sheet: int
}

/// What a sheet keeps about its rows and columns besides their sizes. A
/// hidden row / column is one of size 0 in `RowHeights` / `ColumnWidths`
/// (every layout reads those); the size it goes back to when it is shown
/// again is kept here. Outline levels (Excel's Group) are 1..7 per row /
/// column; `CollapsedRows` / `CollapsedCols` are the summary rows / columns
/// (the one after a group) whose group is collapsed.
type SheetAxes = {
    HiddenRowSizes: Map<int, float>
    HiddenColSizes: Map<int, float>
    RowLevels: Map<int, int>
    ColLevels: Map<int, int>
    CollapsedRows: Set<int>
    CollapsedCols: Set<int>
}

let emptySheetAxes : SheetAxes =
    { HiddenRowSizes = Map.empty; HiddenColSizes = Map.empty; RowLevels = Map.empty; ColLevels = Map.empty
      CollapsedRows = Set.empty; CollapsedCols = Set.empty }

type SpreadsheetModel = {
    /// Sparse cell storage - only non-empty cells
    Cells: Map<CellAddress, CellData>

    /// Current selection
    Selection: Selection

    /// Currently editing cell
    EditingCell: CellAddress option

    /// Edit buffer (what user is typing)
    EditBuffer: string

    /// Grid size
    RowCount: int
    ColCount: int

    /// Clipboard data for copy/paste
    Clipboard: ClipboardData option

    /// History for undo/redo
    History: History

    /// Current sort state
    SortState: SortState option

    /// The sheet's AutoFilter (its range), when it has one.
    AutoFilter: SheetFilter option

    /// The column filters of the AutoFilter, by 0-based sheet column. Empty
    /// when no column is filtered (and always when there is no AutoFilter).
    Filters: Map<int, ColumnFilter>

    /// The rows that pass the column filters among the rows that hold cells,
    /// as last asked for (`getVisibleRows`); None when cells changed since.
    FilterVisibleRows: Set<int> option

    /// The rows the AutoFilter hides, ascending. They are worked out when a
    /// filter is set or applied again (`reapplyFilter`), not when a cell
    /// changes - as in Excel, an edited row stays where it is until the
    /// filter is applied again - and they are part of the document: a file
    /// stores them as hidden rows. Empty when no column is filtered.
    FilterHiddenRows: int array

    /// Custom row heights (default 25.0)
    RowHeights: Map<int, float>

    /// Custom column widths (default 100.0)
    ColumnWidths: Map<int, float>

    /// The sizes hidden rows / columns go back to, and the outline (groups).
    Axes: SheetAxes

    /// Current search query
    SearchQuery: string option

    /// Case-sensitivity used for the current search results.
    SearchCaseSensitive: bool option

    /// Search results (cell addresses)
    SearchResults: CellAddress array

    /// Current search result index
    SearchIndex: int

    /// Conditional formats, by range. See `ConditionalFormatting` for their
    /// evaluation.
    ConditionalFormats: ConditionalFormatRange list

    /// Merged cell ranges
    MergedCells: MergedCell list

    /// Rich-text runs keyed by cell address. Set for cells whose value comes
    /// from a shared-string entry with multiple `<r>` runs (per-run color,
    /// bold, font). Each entry is the list of (text, override) pairs in
    /// source order; the runs concatenate to the plain cell value.
    RichRuns: Map<CellAddress, (string * RunOverride) list>

    /// Number of rows that stay pinned at the top while the body scrolls
    /// vertically. Set from `<pane ySplit="N" state="frozen"/>` in the
    /// sheet's sheetView. Zero = no freeze.
    FrozenRows: int

    /// Number of columns that stay pinned on the left while the body
    /// scrolls horizontally. Set from `<pane xSplit="N" state="frozen"/>`.
    /// Zero = no freeze.
    FrozenCols: int

    /// Print / page setup for this sheet. Always present (defaults when
    /// the source file has no explicit `<pageSetup>` etc.).
    PageSetup: PageSetup

    /// Table regions defined on the sheet (`xl/tables/tableN.xml`). The
    /// runner translates `Xlsx.TableDefinition` here so renderers and the
    /// layout harness can detect table-style overlays from the model
    /// alone, without needing the original workbook handle.
    Tables: Table list

    /// Cells where the Bold flag came from a conditional-formatting rule
    /// (dxf) rather than the cell's own font. The runner sets these when
    /// it bakes a CF rule's `<font><b/></font>` into the cell's format —
    /// the layout harness reads them so it can report the underlying
    /// font's bold (which is what Excel COM `cell.Font.Bold` returns,
    /// regardless of any CF overlay).
    CfBoldCells: Set<CellAddress>

    /// Whether the sheet shows the light grey screen gridlines. From the
    /// sheetView `showGridLines` attribute (default true). Many polished
    /// workbooks turn this off (`showGridLines="0"`) and rely on explicit
    /// cell borders, so the renderer must suppress the always-on grid.
    ShowGridLines: bool

    /// Whether the row numbers and column letters are shown (sheetView
    /// `showRowColHeaders`, default true).
    ShowHeaders: bool

    /// Cells show their formulas instead of the results (sheetView
    /// `showFormulas`, Excel's Ctrl+`).
    ShowFormulas: bool

    /// False: a cell holding the number 0 shows nothing (sheetView `showZeros`).
    ShowZeros: bool

    /// The sheet's zoom: 1.0 = 100 % (sheetView `zoomScale`). View state that
    /// is saved with the sheet; changing it is not an edit of the document.
    Zoom: float

    /// Dynamic-array spill ownership: each spilled (ghost) cell mapped to its
    /// anchor formula cell. Maintained by the recalc+spill pass so stale
    /// spilled cells are cleared before a re-spill. Empty for non-spilling
    /// workbooks.
    Spill: Map<CellAddress, CellAddress>

    /// Data validations, by range (see `ValidationRange`). `getValidation`
    /// gives the one on a cell.
    Validations: ValidationRange list

    /// VBA/VBScript module source attached to the workbook (the "VBA project").
    /// Empty when there are no macros. Event subs (Worksheet_Change /
    /// Workbook_Open) and user-run macros live here; `Vba.runOnModel` executes
    /// it. Persisted/round-tripped separately from cell data.
    Macros: string

    /// Pictures and other drawing objects, back to front. They move with
    /// their anchor cells when rows or columns are inserted or deleted.
    Drawings: SheetDrawing list

    /// Threaded comments by cell (shown where a cell's comment is shown).
    Threads: Map<CellAddress, CommentThread>

    /// Hyperlinks, by range. They move with their cells.
    Hyperlinks: SheetHyperlink list

    /// Default cell styles of columns. They move with their columns.
    ColumnStyles: ColumnStyle list

    /// What-if data tables. Their ranges and input cells follow structure edits.
    DataTables: DataTable list

    /// Ranges of things the editor does not model (pivot tables).
    KeptRanges: KeptRange list

    /// The colour of the sheet's tab: "#RRGGBB", or a colour of the file's
    /// theme / palette as the file names it ("theme:4", "theme:4:-0.25",
    /// "indexed:10"). None = no colour.
    TabColor: string option

    /// A sheet without a tab (the workbook says which sheets have one) that
    /// is "very hidden": Excel does not offer to unhide it.
    VeryHidden: bool

    /// Which sheet of the opened file this model is, when it came from one.
    /// It stays with the model through renames, reordering and added or
    /// removed sheets, so a save can carry over what the editor does not
    /// model. None for a sheet made in the editor.
    SourceSheet: SheetOrigin option

    /// Counts the row / column insertions and deletions the sheet has seen
    /// (an undo brings the earlier count back). What is kept of a file by
    /// cell reference without being modelled (sparklines, x14 conditional
    /// formats) is only written back while this is what it was at import.
    StructureVersion: int

    /// Set on the stand-in for a sheet of an opened file that has not been
    /// converted to a model yet (a sheet is converted when it is first shown,
    /// or when a formula reads it). The stand-in has no cells; `resolveSheet`
    /// gives the sheet itself. None for every other model.
    Deferred: DeferredSheet option
}

/// A sheet of an opened file, converted on demand (see `SpreadsheetModel.Deferred`).
and DeferredSheet = {
    /// The converted sheet, once something asked for it.
    mutable Loaded: SpreadsheetModel option
    /// Convert the sheet (it is converted once; this then gives `Loaded`).
    Load: unit -> SpreadsheetModel
    /// The sheet has formulas, conditional formats or validations: what it
    /// shows can depend on other sheets, and their edits on it.
    HasFormulas: bool
    /// The sheet's formula cells (address, formula text as the converted
    /// sheet holds it), read off the file without converting the sheet: what
    /// the dependency graph needs of a sheet nothing has asked for yet.
    Formulas: unit -> (CellAddress * string) array
}

/// Command for undo/redo (must come after SpreadsheetModel)
and Command = {
    Undo: SpreadsheetModel -> SpreadsheetModel
    Redo: SpreadsheetModel -> SpreadsheetModel
    Description: string
}

/// History for undo/redo
and History = {
    Past: Command list      // Undo stack
    Future: Command list    // Redo stack
    MaxSize: int           // Limit memory (default 50)
}

/// Empty history
let emptyHistory = {
    Past = []
    Future = []
    MaxSize = 50
}

/// Create empty spreadsheet
let createEmpty rowCount colCount = {
    Cells = Map.empty
    Selection = SingleCell(0, 0)
    EditingCell = None
    EditBuffer = ""
    RowCount = rowCount
    ColCount = colCount
    Clipboard = None
    History = emptyHistory
    SortState = None
    AutoFilter = None
    Filters = Map.empty
    FilterVisibleRows = None
    FilterHiddenRows = [||]
    RowHeights = Map.empty
    ColumnWidths = Map.empty
    Axes = emptySheetAxes
    SearchQuery = None
    SearchCaseSensitive = None
    SearchResults = [||]
    SearchIndex = 0
    ConditionalFormats = []
    MergedCells = []
    RichRuns = Map.empty
    FrozenRows = 0
    FrozenCols = 0
    PageSetup = defaultPageSetup
    Tables = []
    CfBoldCells = Set.empty
    ShowGridLines = true
    ShowHeaders = true
    ShowFormulas = false
    ShowZeros = true
    Zoom = 1.0
    Spill = Map.empty
    Validations = []
    Macros = ""
    Drawings = []
    Threads = Map.empty
    Hyperlinks = []
    ColumnStyles = []
    DataTables = []
    KeptRanges = []
    TabColor = None
    VeryHidden = false
    SourceSheet = None
    StructureVersion = 0
    Deferred = None
}

/// The model of a sheet: for the stand-in of a sheet that has not been
/// converted yet, the sheet itself (converting it now), else the model.
let resolveSheet (model: SpreadsheetModel) : SpreadsheetModel =
    match model.Deferred with
    | None -> model
    | Some deferred ->
        match deferred.Loaded with
        | Some loaded -> loaded
        | None -> deferred.Load ()

/// Whether the model is the stand-in of a sheet that is not converted yet.
let isDeferredSheet (model: SpreadsheetModel) : bool =
    match model.Deferred with
    | Some deferred -> deferred.Loaded.IsNone
    | None -> false

/// Convert column index to letter (0="A", 25="Z", 26="AA")
let rec columnToLetter (col: int): string =
    if col < 0 then ""
    elif col < 26 then
        string (char (int 'A' + col))
    else
        let quotient = (col / 26) - 1
        let remainder = col % 26
        columnToLetter quotient + columnToLetter remainder

/// Convert column letter to index ("A"=0, "Z"=25, "AA"=26)
let letterToColumn (letter: string): int =
    let mutable result = 0
    for c in letter.ToUpper() do
        result <- result * 26 + (int c - int 'A' + 1)
    result - 1

let tryParseNumber (text: string): float option =
    match Double.TryParse(text, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
    | true, value -> Some value
    | false, _ -> None

/// Parse cell reference like "A1" to (row, col)
let parseCellRef (ref: string): (int * int) option =
    try
        let mutable colPart = ""
        let mutable rowPart = ""
        let mutable inRow = false
        let mutable valid = true

        for c in ref do
            if Char.IsDigit(c) then
                inRow <- true
                rowPart <- rowPart + string c
            elif Char.IsLetter(c) && not inRow then
                colPart <- colPart + string c
            else
                valid <- false

        if not valid || String.IsNullOrEmpty(colPart) || String.IsNullOrEmpty(rowPart) then
            None
        else
            let col = letterToColumn colPart
            let row = Int32.Parse(rowPart) - 1  // 1-indexed to 0-indexed
            if row < 0 || col < 0 then None else Some (row, col)
    with
    | _ -> None

/// Format cell reference like (0, 0) to "A1"
let formatCellRef (row: int) (col: int): string =
    sprintf "%s%d" (columnToLetter col) (row + 1)

/// Get cell data at address
let getCell (addr: CellAddress) (model: SpreadsheetModel): CellData option =
    Map.tryFind addr model.Cells

/// Called after the cells of a sheet changed. The rows an AutoFilter hides
/// stay as they are (see `SpreadsheetModel.FilterHiddenRows`); only the set
/// of rows that pass, kept for `getVisibleRows`, is forgotten.
let invalidateFilterProjection (model: SpreadsheetModel): SpreadsheetModel =
    if Map.isEmpty model.Filters then
        if model.FilterVisibleRows.IsNone && model.FilterHiddenRows.Length = 0 then model
        else { model with FilterVisibleRows = None; FilterHiddenRows = [||] }
    elif model.FilterVisibleRows.IsNone then model
    else { model with FilterVisibleRows = None }

let private binarySearchInt (values: int array) target =
    let mutable lo = 0
    let mutable hi = values.Length - 1
    let mutable found = -1
    while found < 0 && lo <= hi do
        let mid = lo + (hi - lo) / 2
        let value = values.[mid]
        if value = target then found <- mid
        elif value < target then lo <- mid + 1
        else hi <- mid - 1
    found

/// Whether the AutoFilter hides a row.
let isFilterHiddenRow (row: int) (model: SpreadsheetModel): bool =
    model.FilterHiddenRows.Length > 0 && not (Map.isEmpty model.Filters) && binarySearchInt model.FilterHiddenRows row >= 0

/// Set cell data at address
let setCell (addr: CellAddress) (data: CellData) (model: SpreadsheetModel): SpreadsheetModel =
    match Map.tryFind addr model.Cells with
    | Some existing when existing = data -> model
    | existing ->
        // Writing a different value into a cell a dynamic array spilled into
        // makes it the user's cell: it stops being a spill ghost, so the next
        // spill pass sees it as an obstruction (#SPILL!) instead of
        // overwriting it. Restyling a ghost keeps it a ghost.
        let spill =
            if Map.isEmpty model.Spill || not (Map.containsKey addr model.Spill) then model.Spill
            else
                match existing with
                | Some ghost when ghost.Value = data.Value -> model.Spill
                | _ -> Map.remove addr model.Spill
        { model with Cells = Map.add addr data model.Cells; Spill = spill } |> invalidateFilterProjection

let setCellFormat (addr: CellAddress) (update: CellFormat -> CellFormat) (model: SpreadsheetModel): SpreadsheetModel =
    let existing = getCell addr model |> Option.defaultValue { Value = Empty; Format = defaultFormat; Comment = None }
    let updatedFormat = update existing.Format
    if updatedFormat = existing.Format then
        model
    else
        setCell addr { existing with Format = updatedFormat } model

let setCellTextColor (addr: CellAddress) color (model: SpreadsheetModel): SpreadsheetModel =
    setCellFormat addr (fun f -> { f with TextColor = Some color }) model

let setCellBgColor (addr: CellAddress) color (model: SpreadsheetModel): SpreadsheetModel =
    setCellFormat addr (fun f -> { f with BgColor = Some color }) model

let setCellNumberFormat (addr: CellAddress) numberFormat (model: SpreadsheetModel): SpreadsheetModel =
    setCellFormat addr (fun f -> { f with NumberFormat = Some numberFormat }) model

let setCellHorizontalAlign (addr: CellAddress) align (model: SpreadsheetModel): SpreadsheetModel =
    setCellFormat addr (fun f -> { f with HorizontalAlign = Some align }) model

let private borderMatches style color (border: BorderStyle option) =
    match border with
    | Some existing -> existing.Style = style && existing.Color = Some color
    | None -> false

let private hasRequestedBorder side style color (format: CellFormat) =
    match side with
    | "top" -> borderMatches style color format.BorderTop
    | "bottom" -> borderMatches style color format.BorderBottom
    | "left" -> borderMatches style color format.BorderLeft
    | "right" -> borderMatches style color format.BorderRight
    | "all" ->
        borderMatches style color format.BorderTop
        && borderMatches style color format.BorderBottom
        && borderMatches style color format.BorderLeft
        && borderMatches style color format.BorderRight
    | _ -> true

let private applyBorderToFormat side border (format: CellFormat) =
    match side with
    | "top" -> { format with BorderTop = Some border }
    | "bottom" -> { format with BorderBottom = Some border }
    | "left" -> { format with BorderLeft = Some border }
    | "right" -> { format with BorderRight = Some border }
    | "all" ->
        { format with
            BorderTop = Some border
            BorderBottom = Some border
            BorderLeft = Some border
            BorderRight = Some border }
    | _ -> format

/// What `ApplyBorder(side, style, color)` does to the format of the cell
/// (row, col) of the selected block (top, left, bottom, right):
/// "top" | "bottom" | "left" | "right" | "all" set that side of every cell,
/// "outside" the outline of the block only, "inside" the edges between its
/// cells, "none" clears all four sides. A style of "none" (or "") clears
/// the sides it is applied to. The format is returned as is when nothing
/// changes for this cell.
let borderFormatInBlock (top: int, left: int, bottom: int, right: int) (side: string) (style: string) (color: string)
                        (row: int) (col: int) (format: CellFormat) : CellFormat =
    let border =
        match borderCssParts style with
        | Some _ when side <> "none" -> Some { Style = style; Color = Some color }
        | _ -> None
    let set (t: bool) (b: bool) (l: bool) (r: bool) =
        if not (t || b || l || r) then format
        else
            let pick wanted current = if wanted then border else current
            let updated =
                { format with
                    BorderTop = pick t format.BorderTop
                    BorderBottom = pick b format.BorderBottom
                    BorderLeft = pick l format.BorderLeft
                    BorderRight = pick r format.BorderRight }
            if updated = format then format else updated
    match side with
    | "top" -> set true false false false
    | "bottom" -> set false true false false
    | "left" -> set false false true false
    | "right" -> set false false false true
    | "all" | "none" -> set true true true true
    | "outside" -> set (row = top) (row = bottom) (col = left) (col = right)
    | "inside" -> set (row > top) (row < bottom) (col > left) (col < right)
    | _ -> format

let setCellBorder (addr: CellAddress) (side: string) (style: string) (color: string) (model: SpreadsheetModel): SpreadsheetModel =
    let existing = getCell addr model |> Option.defaultValue { Value = Empty; Format = defaultFormat; Comment = None }
    if hasRequestedBorder side style color existing.Format then
        model
    else
        let border = { Style = style; Color = Some color }
        setCell addr { existing with Format = applyBorderToFormat side border existing.Format } model

/// Clear cell at address
let clearCell (addr: CellAddress) (model: SpreadsheetModel): SpreadsheetModel =
    if Map.containsKey addr model.Cells then
        { model with Cells = Map.remove addr model.Cells } |> invalidateFilterProjection
    else
        model

/// Clear all cells in a selection, preserving identity when nothing changes.
let clearSelectionCells (selection: Selection) (model: SpreadsheetModel): SpreadsheetModel =
    match selection with
    | SingleCell (row, col) ->
        clearCell { Row = row; Col = col } model
    | Range (startRow, startCol, endRow, endCol) when float (endRow - startRow + 1) * float (endCol - startCol + 1) <= 50000.0 ->
        // A block of moderate size: its cells are taken out one by one,
        // which leaves the rest of a large sheet's map as it is (filtering
        // the whole map costs seconds on a sheet of several hundred
        // thousand cells).
        // (Rows an AutoFilter hides keep their cells: as in Excel, an edit
        // of a filtered selection acts on the rows that are shown.)
        let mutable cells = model.Cells
        for row in startRow .. endRow do
            if not (isFilterHiddenRow row model) then
                for col in startCol .. endCol do
                    let addr = { Row = row; Col = col }
                    if Map.containsKey addr cells then cells <- Map.remove addr cells
        if obj.ReferenceEquals(cells, model.Cells) then model
        else { model with Cells = cells } |> invalidateFilterProjection
    | Range (startRow, startCol, endRow, endCol) ->
        let mutable changed = false
        let cells =
            model.Cells
            |> Map.filter (fun addr _ ->
                let remove =
                    addr.Row >= startRow && addr.Row <= endRow &&
                    addr.Col >= startCol && addr.Col <= endCol &&
                    not (isFilterHiddenRow addr.Row model)
                if remove then changed <- true
                not remove)
        if changed then { model with Cells = cells } |> invalidateFilterProjection else model

/// Round to the nearest integer, halves away from zero (Excel's rule). Written
/// out because Fable ignores MidpointRounding and reads it as a digit count.
let roundHalfAway (x: float) : float =
    let whole = Math.Truncate x
    if abs (x - whole) >= 0.5 then whole + float (sign x) else whole

/// Excel's day zero. UTC-kind so that day arithmetic and calendar fields do
/// not shift across DST changes under Fable, where other DateTimes are local.
let excelEpoch = DateTime(1899, 12, 30, 0, 0, 0, DateTimeKind.Utc)

/// Format a number according to format code
let formatNumber (format: string option) (value: float): string =
    match format with
    | Some "currency" ->
        // Currency format: $1,234.56
        sprintf "$%s" (NumberText.numberToString "N2" value)
    | Some "percentage" | Some "percent" ->
        // Percentage format: 45.67%
        sprintf "%s%%" (NumberText.numberToString "N2" (value * 100.0))
    | Some "date" ->
        try
            let date = excelEpoch.AddDays(value)
            sprintf "%02d/%02d/%04d" date.Month date.Day date.Year
        with
        | _ -> NumberText.numberToString "F2" value
    | Some "decimal2" ->
        // 2 decimal places: 123.46
        NumberText.numberToString "F2" value
    | Some "number" ->
        // General number with comma separators
        NumberText.numberToString "N2" value
    | _ ->
        // Default: just show the number
        string value

/// Display value for cell (handles formulas) - needs to be before getCellValue
let rec displayValue (value: CellValue): string =
    match value with
    | Empty -> ""
    | Number n -> NumberText.numberToString "G" n
    | String s -> s
    | Boolean b -> if b then "TRUE" else "FALSE"
    | Error e -> sprintf "#%s" e
    | Formula (_, cached) -> displayValue cached

/// Formula text as the user edits it. Cells typed in the web editor store the
/// leading "="; imported cells store the bare expression.
let formulaText (text: string) : string =
    if text.StartsWith "=" then text else "=" + text

/// Display value with the legacy named formats applied. Real Excel format
/// codes need CellText.cellDisplayText, which compiles after this module.
let displayValueWithFormat (value: CellValue) (format: CellFormat): string =
    match value with
    | Number n -> formatNumber format.NumberFormat n
    | _ -> displayValue value

/// Get cell value (returns Empty if not found)
let getCellValue (addr: CellAddress) (model: SpreadsheetModel): CellValue =
    match getCell addr model with
    | Some cell -> cell.Value
    | None -> Empty

/// The value a cell shows under the sheet's view options: its formula as
/// text when formulas are shown, nothing for the number 0 when zeros are
/// hidden, else the value itself.
let viewedValue (model: SpreadsheetModel) (value: CellValue) : CellValue =
    if model.ShowZeros && not model.ShowFormulas then value
    else
        match value with
        | Formula (text, _) when model.ShowFormulas -> String (if text.StartsWith "=" then text else "=" + text)
        | _ ->
            if model.ShowZeros then value
            else
                match unwrapFormulaValue value with
                | Number 0.0 -> Empty
                | _ -> value

let private dataRows (model: SpreadsheetModel): Set<int> =
    model.Cells
    |> Map.toSeq
    |> Seq.map (fun (addr, _) -> addr.Row)
    |> Set.ofSeq

/// The block of rows a column-level operation (sort, AutoFilter) acts on, as
/// (first row, last row). Rows with content form unbroken runs separated by
/// empty rows (Excel's "current region"); the block is the run holding most of
/// the column's entries, so a stray title or note elsewhere on the sheet is
/// not mistaken for the data. A table covering that run wins (its totals row
/// excluded). The block's first row is where a header would sit. None for an
/// empty sheet.
let dataBlockForColumn (col: int) (model: SpreadsheetModel) : (int * int) option =
    let mutable rows = Set.empty
    let mutable columnRows = Set.empty
    for KeyValue(addr, cellData) in model.Cells do
        if cellData.Value <> Empty then
            rows <- Set.add addr.Row rows
            if addr.Col = col then columnRows <- Set.add addr.Row columnRows
    if Set.isEmpty rows then None
    else
        let mutable best = None
        let mutable bestCount = -1
        let mutable runStart = -1
        let mutable previous = -2
        let mutable count = 0
        let close () =
            if runStart >= 0 && count > bestCount then
                best <- Some (runStart, previous)
                bestCount <- count
        for row in rows do
            if row <> previous + 1 then
                close ()
                runStart <- row
                count <- 0
            if Set.contains row columnRows then count <- count + 1
            previous <- row
        close ()
        match best with
        | None -> None
        | Some (first, last) ->
            let table =
                model.Tables
                |> List.tryFind (fun t ->
                    col >= t.StartCol - 1 && col <= t.EndCol - 1
                    && t.StartRow - 1 <= last && t.EndRow - 1 >= first)
            match table with
            | Some t -> Some (t.StartRow - 1, t.EndRow - 1 - (if t.HasTotals then 1 else 0))
            | None -> Some (first, last)

/// A row/column insertion or deletion. `at` is the 0-indexed first row/column
/// affected; `count` rows/columns are inserted before it or deleted from it.
type StructureEdit =
    | InsertRows of at: int * count: int
    | DeleteRows of at: int * count: int
    | InsertCols of at: int * count: int
    | DeleteCols of at: int * count: int

/// A block of cells moved by cut-and-paste: the rectangle it occupied and how
/// far it travelled. The predicates classify a reference's sheet qualifier
/// (None = unqualified) as seen from the formula being rewritten.
type CellMove = {
    Top: int
    Left: int
    Bottom: int
    Right: int
    DeltaRow: int
    DeltaCol: int
    /// Does this qualifier name the sheet the block was cut from?
    OnSourceSheet: string option -> bool
    /// Does this qualifier name the sheet the block was pasted on?
    OnTargetSheet: string option -> bool
    /// The qualifier a reference to a moved cell carries afterwards.
    Requalify: string option -> string option
    /// The qualifier an unqualified reference to a cell that did NOT move gets
    /// (a formula carried to another sheet keeps pointing at its old sheet).
    QualifyStationary: string option
}

/// Formula-text rewrites. They need the formula parser and printer, which
/// compile after this module (Formula.fs), so Formula.fs installs them here.
/// A change of the workbook's sheets or tables that the text of formulas has to follow.
type SheetRewrite =
    /// References to the sheet named `oldName` name it `newName`.
    | RenameSheetRefs of oldName: string * newName: string
    /// References to the sheet named `name` (it is deleted) become #REF!.
    /// `sheets` are the names of every sheet in tab order, the deleted one
    /// among them: a 3-D reference that ends at it ends at its neighbour
    /// inside the span instead.
    | DeleteSheetRefs of name: string * sheets: string list
    /// Structured references to the table `oldName` name it `newName`.
    | RenameTableRefs of oldName: string * newName: string

type FormulaHooks = {
    /// Shift every relative reference by (rows, cols) — copy/paste and fill.
    Shift: int -> int -> string -> string
    /// Re-point references after a structure edit. The predicate says whether
    /// a reference's sheet qualifier (None = unqualified) names the edited sheet.
    Adjust: StructureEdit -> (string option -> bool) -> string -> string
    /// Re-point references to a block of cells that was moved.
    Move: CellMove -> string -> string
    /// Evaluate a formula written for one cell as if it stood `rows`/`cols`
    /// away, at `cell` of `model` — how validation and conditional-format
    /// formulas are applied to each cell of their range.
    Evaluate: string -> int -> int -> SpreadsheetModel -> CellAddress -> CellValue
    /// Re-point references after a sheet was renamed or deleted, or a table renamed.
    Sheets: SheetRewrite -> string -> string
    /// The text a cell shows for a value under a format (CellText).
    Display: CellFormat -> CellValue -> string
}

let mutable private formulaHooks : FormulaHooks option = None

/// Called once by Formula.fs when it initialises.
let installFormulaHooks (hooks: FormulaHooks) : unit =
    formulaHooks <- Some hooks

let private currentFormulaHooks () : FormulaHooks option =
    match formulaHooks with
    | Some _ -> formulaHooks
    | None ->
#if !FABLE_COMPILER
        // A .NET library initialises a file only when one of its values is
        // first read, so a host that has not evaluated a formula yet would
        // find no hooks. Reading Formula.formulaHooksInstalled runs that
        // file's initialiser, which installs them.
        try
            let formulaModule = typeof<CellAddress>.Assembly.GetType "Rbt.Spreadsheet.Formula"
            if not (isNull formulaModule) then
                let installed = formulaModule.GetProperty "formulaHooksInstalled"
                if not (isNull installed) then
                    installed.GetValue(null, null) |> ignore
        with _ -> ()
#endif
        formulaHooks

/// Shift the relative references of a formula by (deltaRow, deltaCol), the way
/// copying it to another cell does. `$`-locked parts stay put; a reference
/// pushed off the sheet becomes #REF!. Text that does not parse is returned
/// unchanged.
let shiftFormulaReferences (deltaRow: int) (deltaCol: int) (formula: string) : string =
    if deltaRow = 0 && deltaCol = 0 then formula
    else
        match currentFormulaHooks () with
        | Some hooks -> hooks.Shift deltaRow deltaCol formula
        | None -> formula

/// Re-point a formula's references after rows/columns were inserted or
/// deleted on the sheet selected by `targetsEditedSheet`.
let adjustFormulaForEdit (edit: StructureEdit) (targetsEditedSheet: string option -> bool) (formula: string) : string =
    match currentFormulaHooks () with
    | Some hooks -> hooks.Adjust edit targetsEditedSheet formula
    | None -> formula

/// Re-point a formula's references to cells that a cut-and-paste moved.
let moveFormulaReferences (move: CellMove) (formula: string) : string =
    match currentFormulaHooks () with
    | Some hooks -> hooks.Move move formula
    | None -> formula

/// Re-point a formula's references after a sheet was renamed or deleted, or
/// a table renamed. Text that does not parse, or that does not mention the
/// sheet / table, is returned as it was.
let rewriteFormulaForSheets (change: SheetRewrite) (formula: string) : string =
    match currentFormulaHooks () with
    | Some hooks -> hooks.Sheets change formula
    | None -> formula

/// Evaluate `formula` (written for a cell `rows`/`cols` before `cell`) at
/// `cell`. Error "NAME?" when no formula engine is loaded.
let evaluateFormulaAt (formula: string) (rows: int) (cols: int) (model: SpreadsheetModel) (cell: CellAddress) : CellValue =
    match currentFormulaHooks () with
    | Some hooks -> hooks.Evaluate formula rows cols model cell
    | None -> Error "NAME?"

// ---------------------------------------------------------------------------
// AutoFilter
// ---------------------------------------------------------------------------
// A sheet has at most one AutoFilter: a range whose first row is the header
// row, and a filter for some of its columns. The rows the filters hide are
// worked out when a filter is set or applied again and are kept in the model
// (`FilterHiddenRows`); editing a cell does not hide or show a row.

/// The text a cell shows, with its number format applied: what a filter's
/// list of values is made of and matched against.
let filterDisplayText (cell: CellData option) : string =
    match cell with
    | None -> ""
    | Some data ->
        match currentFormulaHooks () with
        | Some hooks -> hooks.Display data.Format data.Value
        | None -> displayValueWithFormat (unwrapFormulaValue data.Value) data.Format

type private WildToken =
    | WildChar of char
    | WildOne
    | WildMany

/// Whether `text` matches a filter pattern: `*` is any run of characters,
/// `?` any one character, `~*` / `~?` / `~~` the character itself. Letters
/// match without regard to case; the whole text has to match.
let wildcardMatch (pattern: string) (text: string) : bool =
    let pattern = pattern.ToUpperInvariant()
    let text = text.ToUpperInvariant()
    let tokens = ResizeArray<WildToken>()
    let mutable i = 0
    while i < pattern.Length do
        let c = pattern.[i]
        if c = '~' && i + 1 < pattern.Length && (pattern.[i + 1] = '*' || pattern.[i + 1] = '?' || pattern.[i + 1] = '~') then
            tokens.Add (WildChar pattern.[i + 1])
            i <- i + 2
        else
            tokens.Add (if c = '*' then WildMany elif c = '?' then WildOne else WildChar c)
            i <- i + 1
    let isMany index = index < tokens.Count && (match tokens.[index] with WildMany -> true | _ -> false)
    let mutable t = 0
    let mutable p = 0
    let mutable star = -1
    let mutable mark = 0
    let mutable failed = false
    while not failed && t < text.Length do
        let direct =
            p < tokens.Count
            && (match tokens.[p] with
                | WildOne -> true
                | WildChar c -> c = text.[t]
                | WildMany -> false)
        if direct then
            t <- t + 1
            p <- p + 1
        elif isMany p then
            star <- p
            mark <- t
            p <- p + 1
        elif star >= 0 then
            p <- star + 1
            mark <- mark + 1
            t <- mark
        else failed <- true
    if failed then false
    else
        while isMany p do p <- p + 1
        p = tokens.Count

/// Write text so that a filter pattern matches it literally.
let escapeWildcards (text: string) : string =
    text.Replace("~", "~~").Replace("*", "~*").Replace("?", "~?")

let private filterNumberOf (cell: CellData option) : float option =
    match cell with
    | Some data ->
        (match unwrapFormulaValue data.Value with
         | Number n -> Some n
         | _ -> None)
    | None -> None

let private filterBlank (cell: CellData option) : bool =
    match cell with
    | None -> true
    | Some data ->
        match unwrapFormulaValue data.Value with
        | Empty | String "" -> true
        | _ -> false

/// The test one comparison of a custom filter makes of a cell.
let private filterConditionTest (op: CompareOperator) (operand: string) : CellData option -> bool =
    let number = tryParseNumber operand
    fun cell ->
        let value = match cell with Some data -> unwrapFormulaValue data.Value | None -> Empty
        match op with
        | OpEqual | OpNotEqual ->
            let equal =
                if operand = "" then filterBlank cell
                else
                    match number, value with
                    | Some n, Number x when x = n -> true
                    | _ -> not (filterBlank cell) && wildcardMatch operand (filterDisplayText cell)
            if op = OpEqual then equal else not equal
        | _ ->
            let holds (order: int) =
                match op with
                | OpGreaterThan -> order > 0
                | OpGreaterOrEqual -> order >= 0
                | OpLessThan -> order < 0
                | OpLessOrEqual -> order <= 0
                | _ -> false
            match number, value with
            | Some n, Number x -> holds (compare x n)
            | None, String text when text <> "" -> holds (System.String.Compare(text, operand, System.StringComparison.OrdinalIgnoreCase))
            | _ -> false

/// The test a column filter makes of a cell of its column; `firstRow` to
/// `lastRow` are the rows the filter covers (what "top 10" and "above
/// average" are measured over).
let private filterCriteriaTest (model: SpreadsheetModel) (col: int) (firstRow: int) (lastRow: int) (criteria: FilterCriteria) : CellData option -> bool =
    let numbers () =
        [| for row in firstRow .. lastRow do
            match filterNumberOf (Map.tryFind { Row = row; Col = col } model.Cells) with
            | Some n -> yield n
            | None -> () |]
    match criteria with
    | FilterValues (values, blanks) ->
        let wanted = System.Collections.Generic.HashSet<string>()
        for value in values do wanted.Add (value.ToUpperInvariant()) |> ignore
        let blanks = blanks || wanted.Contains ""
        fun cell ->
            if filterBlank cell then blanks
            else
                wanted.Contains ((filterDisplayText cell).ToUpperInvariant())
                || (match cell with
                    | Some data -> wanted.Contains ((displayValue data.Value).ToUpperInvariant())
                    | None -> false)
    | FilterCustom (conditions, matchAll) ->
        let tests = conditions |> List.map (fun (op, operand) -> filterConditionTest op operand)
        if tests.IsEmpty then (fun _ -> true)
        elif matchAll then (fun cell -> tests |> List.forall (fun test -> test cell))
        else (fun cell -> tests |> List.exists (fun test -> test cell))
    | FilterTop (count, percent, bottom) ->
        let sorted = numbers () |> Array.sort
        if sorted.Length = 0 then (fun _ -> false)
        else
            let wanted = if percent then int (float sorted.Length * float count / 100.0) else count
            let wanted = max 1 (min sorted.Length wanted)
            // (Ties with the last one that makes it are let through too, as in Excel.)
            let threshold = if bottom then sorted.[wanted - 1] else sorted.[sorted.Length - wanted]
            fun cell ->
                match filterNumberOf cell with
                | Some x -> if bottom then x <= threshold else x >= threshold
                | None -> false
    | FilterDynamic kind ->
        let all = numbers ()
        if kind <> "aboveAverage" && kind <> "belowAverage" then (fun _ -> true)
        elif all.Length = 0 then (fun _ -> false)
        else
            let mean = Array.sum all / float all.Length
            fun cell ->
                match filterNumberOf cell with
                | Some x -> if kind = "aboveAverage" then x > mean else x < mean
                | None -> false
    | FilterColor (color, font) ->
        let wanted = color.ToUpperInvariant()
        fun cell ->
            let own =
                match cell with
                | Some data -> (if font then data.Format.TextColor else data.Format.BgColor) |> Option.defaultValue ""
                | None -> ""
            own.ToUpperInvariant() = wanted
    | FilterKept _ -> fun _ -> true

/// The range a filter made without one covers: the block of data around a
/// column (see `dataBlockForColumn`), as wide as the block's filled cells.
let private implicitFilterRange (col: int) (model: SpreadsheetModel) : CellRange =
    match dataBlockForColumn col model with
    | None -> { From = { Row = 0; Col = col }; To = { Row = 0; Col = col } }
    | Some (first, last) ->
        let mutable lo = col
        let mutable hi = col
        for KeyValue(addr, data) in model.Cells do
            if addr.Row >= first && addr.Row <= last && data.Value <> Empty then
                if addr.Col < lo then lo <- addr.Col
                if addr.Col > hi then hi <- addr.Col
        { From = { Row = first; Col = lo }; To = { Row = last; Col = hi } }

/// The rows the AutoFilter filters, first to last (the rows below its header
/// row, down to the end of its range or of the sheet's cells), or None.
let filterDataRows (model: SpreadsheetModel) : (int * int) option =
    match model.AutoFilter with
    | None -> None
    | Some filter ->
        let mutable lastUsed = -1
        for KeyValue(addr, _) in model.Cells do
            if addr.Row > lastUsed then lastUsed <- addr.Row
        let first = filter.Range.From.Row + 1
        let last = min filter.Range.To.Row lastUsed
        if last < first then None else Some (first, last)

/// An implicit filter follows its block of data: rows added to the block
/// since are filtered too.
let private refreshImplicitRange (model: SpreadsheetModel) : SpreadsheetModel =
    match model.AutoFilter with
    | Some filter when filter.Implicit && not (Map.isEmpty model.Filters) ->
        let firstCol = model.Filters |> Map.toSeq |> Seq.map fst |> Seq.min
        let block = implicitFilterRange firstCol model
        let lastCol = model.Filters |> Map.toSeq |> Seq.map fst |> Seq.max
        let range =
            { From = { Row = block.From.Row; Col = min block.From.Col (min filter.Range.From.Col firstCol) }
              To = { Row = block.To.Row; Col = max block.To.Col (max filter.Range.To.Col lastCol) } }
        if range = filter.Range then model else { model with AutoFilter = Some { filter with Range = range } }
    | _ -> model

/// Whether each row passes every column filter, for the rows the filter
/// covers: (first data row, last data row, test).
let private filterRowTest (model: SpreadsheetModel) : (int * int * (int -> bool)) option =
    if Map.isEmpty model.Filters then None
    else
        match filterDataRows model with
        | None -> None
        | Some (first, last) ->
            let tests =
                model.Filters
                |> Map.toArray
                |> Array.map (fun (col, filter) -> col, filterCriteriaTest model col first last filter.Criteria)
            let passes (row: int) =
                let mutable pass = true
                let mutable i = 0
                while pass && i < tests.Length do
                    let col, test = tests.[i]
                    pass <- test (Map.tryFind { Row = row; Col = col } model.Cells)
                    i <- i + 1
                pass
            Some (first, last, passes)

/// Apply the column filters to the sheet as it is now: the rows that fail
/// one are hidden, every other row of the filter's range is shown.
let reapplyFilter (model: SpreadsheetModel) : SpreadsheetModel =
    if Map.isEmpty model.Filters then
        if model.FilterVisibleRows.IsNone && model.FilterHiddenRows.Length = 0 then model
        else { model with FilterVisibleRows = None; FilterHiddenRows = [||] }
    else
        let model = refreshImplicitRange model
        let hidden =
            match filterRowTest model with
            | None -> [||]
            | Some (first, last, passes) ->
                // A kept filter is not evaluated: the rows the file had hidden stay hidden.
                let kept = model.Filters |> Map.exists (fun _ filter -> match filter.Criteria with FilterKept _ -> true | _ -> false)
                [| for row in first .. last do
                    if not (passes row) || (kept && isFilterHiddenRow row model) then yield row |]
        if hidden = model.FilterHiddenRows && model.FilterVisibleRows.IsNone then model
        else { model with FilterVisibleRows = None; FilterHiddenRows = (if hidden = model.FilterHiddenRows then model.FilterHiddenRows else hidden) }

/// Give the sheet an AutoFilter over `range` (its first row is the header
/// row), or take it away with None. The filters of columns outside the new
/// range are dropped; rows are hidden and shown accordingly.
let setAutoFilter (range: CellRange option) (model: SpreadsheetModel) : SpreadsheetModel =
    match range with
    | None -> { model with AutoFilter = None; Filters = Map.empty } |> reapplyFilter
    | Some range ->
        let filters = model.Filters |> Map.filter (fun col _ -> col >= range.From.Col && col <= range.To.Col)
        { model with
            AutoFilter = Some { Range = range; Implicit = false }
            Filters = (if filters.Count = model.Filters.Count then model.Filters else filters) }
        |> reapplyFilter

/// Filter a column. Without an AutoFilter the sheet gets an implicit one
/// over the block of data around the column.
let setColumnFilter (col: int) (criteria: FilterCriteria) (model: SpreadsheetModel) : SpreadsheetModel =
    let filter =
        match model.AutoFilter with
        | Some filter when filter.Implicit && (col < filter.Range.From.Col || col > filter.Range.To.Col) ->
            { filter with Range = { From = { filter.Range.From with Col = min col filter.Range.From.Col }; To = { filter.Range.To with Col = max col filter.Range.To.Col } } }
        | Some filter -> filter
        | None -> { Range = implicitFilterRange col model; Implicit = true }
    { model with
        AutoFilter = Some filter
        Filters = Map.add col { Column = col; Criteria = criteria } model.Filters }
    |> reapplyFilter

/// Take the filter off a column (None: off every column). An implicit
/// AutoFilter goes with its last column filter.
let clearColumnFilter (col: int option) (model: SpreadsheetModel) : SpreadsheetModel =
    let filters =
        match col with
        | Some col -> Map.remove col model.Filters
        | None -> Map.empty
    let autoFilter =
        match model.AutoFilter with
        | Some filter when filter.Implicit && Map.isEmpty filters -> None
        | other -> other
    if Map.isEmpty model.Filters && autoFilter = model.AutoFilter then model
    else { model with Filters = filters; AutoFilter = autoFilter } |> reapplyFilter

/// The distinct texts the cells of a column show within the AutoFilter's
/// rows (the block of data around the column when there is no AutoFilter),
/// sorted - numbers first, by value - and whether the column has blank cells
/// there: what a filter's checklist offers.
let filterColumnValues (col: int) (model: SpreadsheetModel) : string list * bool =
    let rows =
        match model.AutoFilter with
        | Some _ -> filterDataRows model
        | None ->
            match dataBlockForColumn col model with
            | Some (first, last) when last > first -> Some (first + 1, last)
            | _ -> None
    match rows with
    | None -> [], false
    | Some (first, last) ->
        let seen = System.Collections.Generic.HashSet<string>()
        let found = ResizeArray<CellValue * string>()
        let mutable blanks = false
        for row in first .. last do
            let cell = Map.tryFind { Row = row; Col = col } model.Cells
            if filterBlank cell then blanks <- true
            else
                let text = filterDisplayText cell
                if seen.Add (text.ToUpperInvariant()) then
                    found.Add ((match cell with Some data -> unwrapFormulaValue data.Value | None -> Empty), text)
        let rank (value: CellValue) = match value with Number _ -> 0 | String _ -> 1 | Boolean _ -> 2 | _ -> 3
        let ordered =
            found
            |> Seq.sortWith (fun (a, textA) (b, textB) ->
                match a, b with
                | Number x, Number y -> compare x y
                | _ ->
                    if rank a <> rank b then compare (rank a) (rank b)
                    else System.String.Compare(textA, textB, System.StringComparison.OrdinalIgnoreCase))
            |> Seq.map snd
            |> List.ofSeq
        ordered, blanks

/// The rows of `top` to `bottom` that the AutoFilter does not hide.
let filterShownRows (top: int) (bottom: int) (model: SpreadsheetModel) : int list =
    if model.FilterHiddenRows.Length = 0 || Map.isEmpty model.Filters then [ top .. bottom ]
    else [ for row in top .. bottom do if not (isFilterHiddenRow row model) then yield row ]

/// Serialize selection to clipboard data
let serializeSelection (selection: Selection) (model: SpreadsheetModel) (isCut: bool): ClipboardData =
    match selection with
    | SingleCell (row, col) ->
        let addr = { Row = row; Col = col }
        let cellData = getCell addr model
        match cellData with
        | Some data ->
            { Cells = [(0, 0, data)]; IsCut = isCut; SourceStartRow = row; SourceStartCol = col; SourceEndRow = row; SourceEndCol = col }
        | None ->
            { Cells = []; IsCut = isCut; SourceStartRow = row; SourceStartCol = col; SourceEndRow = row; SourceEndCol = col }

    | Range (startRow, startCol, endRow, endCol) ->
        // Rows an AutoFilter hides are not copied: a copy holds the shown
        // rows one after the other (as Excel pastes them); a cut keeps the
        // places, since its paste is a move of the cells.
        let rows = filterShownRows startRow endRow model
        let compact = not isCut && rows.Length < endRow - startRow + 1
        let cells =
            rows
            |> List.mapi (fun index row -> index, row)
            |> List.collect (fun (index, row) ->
                [startCol .. endCol]
                |> List.choose (fun col ->
                    let addr = { Row = row; Col = col }
                    let relRow = if compact then index else row - startRow
                    let relCol = col - startCol
                    match getCell addr model with
                    | Some data -> Some (relRow, relCol, data)
                    | None -> None))
        let endRow = if compact then startRow + max 0 (rows.Length - 1) else endRow
        { Cells = cells; IsCut = isCut; SourceStartRow = startRow; SourceStartCol = startCol; SourceEndRow = endRow; SourceEndCol = endCol }

/// Paste clipboard data at target address
let pasteClipboard (targetAddr: CellAddress) (clipboardData: ClipboardData) (model: SpreadsheetModel): SpreadsheetModel =
    let deltaRow = targetAddr.Row - clipboardData.SourceStartRow
    let deltaCol = targetAddr.Col - clipboardData.SourceStartCol

    let mutable updatedModel = model

    for (relRow, relCol, cellData) in clipboardData.Cells do
        let pasteRow = targetAddr.Row + relRow
        let pasteCol = targetAddr.Col + relCol

        if pasteRow >= 0 && pasteRow < model.RowCount && pasteCol >= 0 && pasteCol < model.ColCount then
            let pasteAddr = { Row = pasteRow; Col = pasteCol }

            // A copied formula follows its new position; a cut formula is moved
            // as written (Excel keeps a moved formula pointing at the same cells).
            let adjustedValue =
                match cellData.Value with
                | Formula (formula, _) when not clipboardData.IsCut ->
                    Formula (shiftFormulaReferences deltaRow deltaCol formula, Empty)
                | Formula (formula, _) ->
                    Formula (formula, Empty)
                | other -> other

            let adjustedCellData = { cellData with Value = adjustedValue }
            updatedModel <- setCell pasteAddr adjustedCellData updatedModel

    updatedModel

/// Where a paste of `clipboard` into `selection` puts the block's top-left
/// corner. One place — the selection's top-left — unless the selection is a
/// whole multiple of the copied block's size, in which case the block is
/// tiled across it (Excel's rule). A cut is a move and never tiles.
let pasteAnchors (selection: Selection) (clipboard: ClipboardData) : CellAddress list =
    let height = max 1 (clipboard.SourceEndRow - clipboard.SourceStartRow + 1)
    let width = max 1 (clipboard.SourceEndCol - clipboard.SourceStartCol + 1)
    match selection with
    | SingleCell (row, col) -> [ { Row = row; Col = col } ]
    | Range (r0, c0, r1, c1) ->
        let top, left = min r0 r1, min c0 c1
        let rows, cols = abs (r1 - r0) + 1, abs (c1 - c0) + 1
        if not clipboard.IsCut && rows % height = 0 && cols % width = 0 && (rows > height || cols > width) then
            [ for i in 0 .. rows / height - 1 do
                for j in 0 .. cols / width - 1 do
                    yield { Row = top + i * height; Col = left + j * width } ]
        else [ { Row = top; Col = left } ]

/// The selection after a paste: the whole area the pasted blocks cover.
let pasteSelection (anchors: CellAddress list) (clipboard: ClipboardData) (fallback: Selection) : Selection =
    match anchors with
    | [] -> fallback
    | _ ->
        let height = max 1 (clipboard.SourceEndRow - clipboard.SourceStartRow + 1)
        let width = max 1 (clipboard.SourceEndCol - clipboard.SourceStartCol + 1)
        let top = anchors |> List.map (fun a -> a.Row) |> List.min
        let left = anchors |> List.map (fun a -> a.Col) |> List.min
        let bottom = (anchors |> List.map (fun a -> a.Row) |> List.max) + height - 1
        let right = (anchors |> List.map (fun a -> a.Col) |> List.max) + width - 1
        if top = bottom && left = right then SingleCell (top, left) else Range (top, left, bottom, right)

/// Paste the clipboard into a selection (tiling it when the selection is a
/// multiple of the block's size).
let pasteClipboardIntoSelection (selection: Selection) (clipboard: ClipboardData) (model: SpreadsheetModel) : SpreadsheetModel =
    pasteAnchors selection clipboard
    |> List.fold (fun acc anchor -> pasteClipboard anchor clipboard acc) model

/// Push command to history (clears future stack)
let pushCommand (command: Command) (model: SpreadsheetModel): SpreadsheetModel =
    let newPast = command :: model.History.Past
    // Limit stack size
    let trimmedPast =
        if List.length newPast > model.History.MaxSize then
            List.take model.History.MaxSize newPast
        else
            newPast

    let newHistory = {
        model.History with
            Past = trimmedPast
            Future = []  // Clear future on new command
    }
    { model with History = newHistory }

/// Undo last command
let undo (model: SpreadsheetModel): SpreadsheetModel option =
    match model.History.Past with
    | [] -> None  // Nothing to undo
    | command :: restPast ->
        // Execute undo
        let undoneModel = command.Undo model

        // Move command to future stack
        let newHistory = {
            undoneModel.History with
                Past = restPast
                Future = command :: undoneModel.History.Future
        }

        Some { undoneModel with History = newHistory }

/// Redo last undone command
let redo (model: SpreadsheetModel): SpreadsheetModel option =
    match model.History.Future with
    | [] -> None  // Nothing to redo
    | command :: restFuture ->
        // Execute redo
        let redoneModel = command.Redo model

        // Move command back to past stack
        let newHistory = {
            redoneModel.History with
                Past = command :: redoneModel.History.Past
                Future = restFuture
        }

        Some { redoneModel with History = newHistory }

/// Compare cell values for sorting (numbers < strings < booleans < empty)
let rec compareCellValues (v1: CellValue) (v2: CellValue): int =
    match v1, v2 with
    | Formula (_, cached1), Formula (_, cached2) -> compareCellValues cached1 cached2
    | Formula (_, cached), other -> compareCellValues cached other
    | other, Formula (_, cached) -> compareCellValues other cached
    | Number n1, Number n2 -> compare n1 n2
    | String s1, String s2 -> System.String.Compare(s1, s2, System.StringComparison.OrdinalIgnoreCase)
    | Boolean b1, Boolean b2 -> compare b1 b2
    | Error _, Error _ -> 0
    | Empty, Empty -> 0
    | _ ->
        let rec rank = function
            | Number _ -> 0
            | String _ -> 1
            | Boolean _ -> 2
            | Error _ -> 3
            | Empty -> 4
            | Formula (_, cached) -> rank cached
        compare (rank v1) (rank v2)

/// Sort the data block around a column by that column, in place. The block's
/// rows keep their position on the sheet (nothing is compacted to the top) and
/// rows outside it are untouched. A first row made only of text is taken as
/// the header and stays first. Blanks sort last in either direction, ties keep
/// their order, and a formula that moves keeps pointing at its own row (its
/// relative references shift with it). Cached values of other formulas are
/// stale afterwards — recalculate.
let sortByColumn (col: int) (ascending: bool) (model: SpreadsheetModel): SpreadsheetModel =
    let sorted = { model with SortState = Some { Column = col; Ascending = ascending } }
    match dataBlockForColumn col model with
    | None -> reapplyFilter sorted
    | Some (firstRow, lastRow) ->
        let rowCells =
            model.Cells
            |> Map.toList
            |> List.filter (fun (addr, _) -> addr.Row >= firstRow && addr.Row <= lastRow)
        let isTable =
            model.Tables |> List.exists (fun t -> t.StartRow - 1 = firstRow && col >= t.StartCol - 1 && col <= t.EndCol - 1 && t.HeaderRowCount > 0)
        let headerCells = rowCells |> List.filter (fun (addr, cellData) -> addr.Row = firstRow && cellData.Value <> Empty)
        let hasHeader =
            lastRow > firstRow
            && (isTable
                || (not (List.isEmpty headerCells)
                    && headerCells |> List.forall (fun (_, cellData) -> match cellData.Value with String _ -> true | _ -> false)))
        let dataStart = if hasHeader then firstRow + 1 else firstRow
        let isBlank row =
            match getCellValue { Row = row; Col = col } model |> unwrapFormulaValue with
            | Empty | String "" -> true
            | _ -> false
        let compareRows a b =
            match isBlank a, isBlank b with
            | true, true -> 0
            | true, false -> 1
            | false, true -> -1
            | false, false ->
                let va = getCellValue { Row = a; Col = col } model
                let vb = getCellValue { Row = b; Col = col } model
                if ascending then compareCellValues va vb else compareCellValues vb va
        let order = [ dataStart .. lastRow ] |> List.sortWith compareRows
        // order.[i] is the source row that lands on row dataStart + i.
        let target = order |> List.mapi (fun i source -> source, dataStart + i) |> Map.ofList
        let moveMap (items: Map<CellAddress, 'T>) (onMove: int -> 'T -> 'T) : Map<CellAddress, 'T> =
            let moving, staying =
                items |> Map.partition (fun addr _ -> addr.Row >= dataStart && addr.Row <= lastRow)
            moving
            |> Map.fold (fun acc addr value ->
                let row = Map.find addr.Row target
                Map.add { addr with Row = row } (onMove (row - addr.Row) value) acc) staying
        let cells =
            moveMap model.Cells (fun delta cellData ->
                match cellData.Value with
                | Formula (formula, cached) when delta <> 0 ->
                    { cellData with Value = Formula (shiftFormulaReferences delta 0 formula, cached) }
                | _ -> cellData)
        { sorted with
            Cells = cells
            RichRuns = moveMap model.RichRuns (fun _ runs -> runs) }
        |> reapplyFilter

/// Get unique values in a column (as strings for display)
let getColumnUniqueValues (col: int) (model: SpreadsheetModel): string list =
    model.Cells
    |> Map.toList
    |> List.filter (fun (addr, _) -> addr.Col = col)
    |> List.map (fun (addr, cellData) -> displayValue cellData.Value)
    |> List.distinct
    |> List.sort

/// Filter a column to the cells that show one of `visibleValues` ("" lets
/// the blank cells through).
let applyFilter (col: int) (visibleValues: Set<string>) (model: SpreadsheetModel): SpreadsheetModel =
    setColumnFilter col (FilterValues (visibleValues, Set.contains "" visibleValues)) model

/// Remove filter from a column
let removeFilter (col: int) (model: SpreadsheetModel): SpreadsheetModel =
    clearColumnFilter (Some col) model

/// The rows holding cells that pass every column filter as the cells are
/// now (rows outside the filter's range pass; its header row is tested like
/// a data row here, though it is never hidden). With no filter: every row
/// that holds a cell.
let getVisibleRows (model: SpreadsheetModel): Set<int> =
    match model.FilterVisibleRows with
    | Some visible when not (Map.isEmpty model.Filters) -> visible
    | _ ->
        let rows = dataRows model
        match (if Map.isEmpty model.Filters then None else filterRowTest (refreshImplicitRange model)) with
        | None -> rows
        | Some (first, last, passes) -> rows |> Set.filter (fun row -> row < first - 1 || row > last || passes row)

let getVisibleRowsHashSet (model: SpreadsheetModel): System.Collections.Generic.HashSet<int> =
    System.Collections.Generic.HashSet<int>(getVisibleRows model |> Set.toArray)

/// The model a renderer draws: since the rows an AutoFilter hides are kept
/// in the model (and `Layout.metricsOf` gives them no height), this is the
/// model itself. Kept for the hosts that call it before painting.
let applyFilterRowHeights (model: SpreadsheetModel): SpreadsheetModel =
    if Map.isEmpty model.Filters && (model.FilterVisibleRows.IsSome || model.FilterHiddenRows.Length > 0) then
        { model with FilterVisibleRows = None; FilterHiddenRows = [||] }
    else model

/// Distinct, sorted display strings in a column — the AutoFilter checklist's
/// values. Blanks excluded; uses the same `displayValue` as the filter test.
let columnUniqueValues (col: int) (model: SpreadsheetModel) : string list =
    model.Cells
    |> Map.toList
    |> List.choose (fun (addr, _) ->
        if addr.Col = col then
            match displayValue (getCellValue addr model) with
            | "" -> None
            | s -> Some s
        else None)
    |> List.distinct
    |> List.sort

let applyFilterValues (col: int) (visibleValues: string array) (model: SpreadsheetModel): SpreadsheetModel =
    applyFilter col (Set.ofArray visibleValues) model

let columnUniqueValueCount (col: int) (model: SpreadsheetModel): int =
    columnUniqueValues col model |> List.length

let visibleRowCount (model: SpreadsheetModel): int =
    getVisibleRows model |> Set.count

let cellCount (model: SpreadsheetModel): int =
    model.Cells.Count

/// Create / register an Excel table over a 0-indexed region (header row r0,
/// data rows r0+1..r1). Blank header cells are filled with "Column N";
/// the table is registered in `model.Tables` (1-indexed) with a
/// TableStyleMedium2 style + an autofilter across every column. When
/// `withTotals`, a totals row is appended below the data: a "Total" label in
/// the first column and `=SUBTOTAL(109,<col>)` in the rest, and the table
/// range grows to include it. Returns the updated model. Structured refs over
/// the result work once it's bridged into the workbook (Workbook.makeTable…).
let createTable (name: string) (r0: int) (c0: int) (r1: int) (c1: int)
                (withTotals: bool) (model: SpreadsheetModel): SpreadsheetModel =
    // Fill any blank header cell with a default column name.
    let mutable m = model
    for c in c0 .. c1 do
        let addr = { Row = r0; Col = c }
        match getCellValue addr m with
        | Empty | String "" ->
            m <- setCell addr { Value = String (sprintf "Column %d" (c - c0 + 1)); Format = defaultFormat; Comment = None } m
        | _ -> ()
    // Optional totals row of SUBTOTAL(109,…) formulas, just below the data.
    let totalsRow = r1 + 1
    if withTotals then
        for c in c0 .. c1 do
            let addr = { Row = totalsRow; Col = c }
            if c = c0 then
                m <- setCell addr { Value = String "Total"; Format = defaultFormat; Comment = None } m
            else
                let col = columnToLetter c
                let formula = sprintf "SUBTOTAL(109,%s%d:%s%d)" col (r0 + 2) col (r1 + 1)
                m <- setCell addr { Value = Formula (formula, Number 0.0); Comment = None; Format = defaultFormat } m
    let endRow0 = if withTotals then totalsRow else r1
    let table : Table =
        { Name = name
          StartRow = r0 + 1
          StartCol = c0 + 1
          EndRow = endRow0 + 1
          EndCol = c1 + 1
          HeaderRowCount = 1
          Style = Some { Name = "TableStyleMedium2"; ShowFirstColumn = false; ShowLastColumn = false; ShowRowStripes = true; ShowColumnStripes = false
                         HeaderFill = Some "#4472C4"; StripeFill = Some "#D9E1F2" }
          AutoFilter = Some (c0 + 1, c1 + 1)
          HasTotals = withTotals }
    { m with Tables = m.Tables @ [ table ] }

/// The table whose region contains the 0-indexed cell, if any. Table
/// coordinates are 1-indexed, so they're shifted for the comparison.
let tableAt (row: int) (col: int) (model: SpreadsheetModel): Table option =
    model.Tables
    |> List.tryFind (fun t ->
        row >= t.StartRow - 1 && row <= t.EndRow - 1 && col >= t.StartCol - 1 && col <= t.EndCol - 1)

/// Distinct, sorted display values in one column of a table — data rows only
/// (header and totals excluded). The checklist for a table's filter button.
let tableColumnUniqueValues (table: Table) (col: int) (model: SpreadsheetModel): string list =
    let dataStart = (table.StartRow - 1) + table.HeaderRowCount
    let dataEnd = (table.EndRow - 1) - (if table.HasTotals then 1 else 0)
    [ for r in dataStart .. dataEnd -> displayValue (getCellValue { Row = r; Col = col } model) ]
    |> List.filter (fun s -> s <> "")
    |> List.distinct
    |> List.sort

/// Render-time: bake table styling (header accent fill + white bold text +
/// filter buttons, and banded data rows) into cell formats from `model.Tables`.
/// Acts ONLY on tables whose `Style.HeaderFill` is set — i.e. tables created in
/// the app (`createTable`); imported tables leave it None because the importer
/// already bakes theme-accent styling into each cell, so this never
/// double-styles or disturbs the parity corpus. Uses orElse/|| merges so a
/// cell's own explicit colors still win. Empty cells in the region are created
/// styled so the band/header spans the full width.
let applyTableStyles (model: SpreadsheetModel): SpreadsheetModel =
    let styleTable (m: SpreadsheetModel) (t: Table) : SpreadsheetModel =
        match t.Style with
        | Some s when s.HeaderFill.IsSome ->
            let r0, c0, r1, c1 = t.StartRow - 1, t.StartCol - 1, t.EndRow - 1, t.EndCol - 1
            let mutable acc = m
            for r in r0 .. r1 do
                let isHeader = r < r0 + t.HeaderRowCount
                let dataIdx = r - r0 - t.HeaderRowCount
                for c in c0 .. c1 do
                    let addr = { Row = r; Col = c }
                    let baseData = getCell addr acc |> Option.defaultValue { Value = Empty; Format = defaultFormat; Comment = None }
                    let fmt = baseData.Format
                    let hasFilter =
                        match t.AutoFilter with Some (fc1, fc2) -> (c + 1) >= fc1 && (c + 1) <= fc2 | None -> false
                    let newFmt =
                        if isHeader then
                            { fmt with
                                BgColor = fmt.BgColor |> Option.orElse s.HeaderFill
                                TextColor = fmt.TextColor |> Option.orElse (Some "#FFFFFF")
                                Bold = true
                                FilterButton = fmt.FilterButton || hasFilter }
                        elif s.ShowRowStripes && dataIdx >= 0 && dataIdx % 2 = 0 then
                            { fmt with BgColor = fmt.BgColor |> Option.orElse s.StripeFill }
                        else fmt
                    if newFmt <> fmt then acc <- setCell addr { baseData with Format = newFmt } acc
            acc
        | _ -> m
    model.Tables |> List.fold styleTable model

// --- Row / column insertion and deletion ------------------------------------
// One implementation serves all four operations: everything a sheet stores by
// position (cells, merges, sizes, conditional formats, validations, tables,
// the spill map, filters) is moved the same way, and formulas are re-pointed.

let private structureEditParts (edit: StructureEdit) =
    match edit with
    | InsertRows (at, count) -> true, true, at, max 0 count
    | DeleteRows (at, count) -> true, false, at, max 0 count
    | InsertCols (at, count) -> false, true, at, max 0 count
    | DeleteCols (at, count) -> false, false, at, max 0 count

/// Where an index on the edited axis ends up; None when it was deleted.
let private moveEditIndex isInsert at count index =
    if isInsert then Some (if index >= at then index + count else index)
    elif index >= at + count then Some (index - count)
    elif index >= at then None
    else Some index

/// Where an inclusive span on the edited axis ends up: an insertion inside it
/// grows it, a deletion overlapping it shrinks it, None when nothing is left.
let private moveEditSpan isInsert at count lo hi =
    if isInsert then
        Some ((if lo >= at then lo + count else lo), (if hi >= at then hi + count else hi))
    elif lo >= at && hi < at + count then
        None
    else
        let newLo = if lo >= at + count then lo - count elif lo >= at then at else lo
        let newHi = if hi >= at + count then hi - count elif hi >= at then at - 1 else hi
        Some (newLo, newHi)

/// Re-point the formulas of a sheet after `edit` happened on the sheet that
/// `targetsEditedSheet` selects. Used for the *other* sheets of a workbook;
/// returns the same model when no formula changed.
let rewriteFormulasForEdit (edit: StructureEdit) (targetsEditedSheet: string option -> bool) (model: SpreadsheetModel) : SpreadsheetModel =
    let mutable cells = model.Cells
    let mutable changed = false
    for KeyValue(addr, cellData) in model.Cells do
        match cellData.Value with
        | Formula (formula, cached) ->
            let adjusted = adjustFormulaForEdit edit targetsEditedSheet formula
            if adjusted <> formula then
                cells <- Map.add addr { cellData with Value = Formula (adjusted, cached) } cells
                changed <- true
        | _ -> ()
    if changed then { model with Cells = cells } |> invalidateFilterProjection else model

/// Apply `rewrite` to every formula a sheet holds: its formula cells, the
/// formulas of its conditional formats and of its validations. The model
/// itself when nothing changed.
let mapModelFormulas (rewrite: string -> string) (model: SpreadsheetModel) : SpreadsheetModel =
    let mutable cells = model.Cells
    for KeyValue(addr, cellData) in model.Cells do
        match cellData.Value with
        | Formula (formula, cached) ->
            let rewritten = rewrite formula
            if rewritten <> formula then cells <- Map.add addr { cellData with Value = Formula (rewritten, cached) } cells
        | _ -> ()
    let threshold (value: CfThreshold) =
        match value.Kind with
        | CfvoFormula -> { value with Value = value.Value |> Option.map rewrite }
        | _ -> value
    let conditionalFormats =
        model.ConditionalFormats
        |> List.map (fun entry ->
            let rule =
                match entry.Format.Rule with
                | CellIs (op, first, second) -> CellIs (op, rewrite first, second |> Option.map rewrite)
                | Expression formula -> Expression (rewrite formula)
                | ColorScale stops -> ColorScale (stops |> List.map (fun (stop, color) -> threshold stop, color))
                | DataBar (low, high, color) -> DataBar (threshold low, threshold high, color)
                | IconSet (family, thresholds, reverse) -> IconSet (family, thresholds |> List.map threshold, reverse)
                | other -> other
            if rule = entry.Format.Rule then entry else { entry with Format = { entry.Format with Rule = rule } })
    let validations =
        model.Validations
        |> List.map (fun entry ->
            let rule =
                match entry.Validation.Rule with
                | Comparison (kind, op, first, second) -> Comparison (kind, op, rewrite first, second |> Option.map rewrite)
                | CustomFormula formula -> CustomFormula (rewrite formula)
                | ListSource source ->
                    let rewritten = rewrite ("=" + source)
                    ListSource (if rewritten.StartsWith "=" then rewritten.Substring 1 else rewritten)
                | other -> other
            if rule = entry.Validation.Rule then entry else { entry with Validation = { entry.Validation with Rule = rule } })
    let sameCells = obj.ReferenceEquals(cells, model.Cells)
    let sameFormats = conditionalFormats = model.ConditionalFormats
    let sameValidations = validations = model.Validations
    if sameCells && sameFormats && sameValidations then model
    else
        let changed =
            { model with
                Cells = cells
                ConditionalFormats = (if sameFormats then model.ConditionalFormats else conditionalFormats)
                Validations = (if sameValidations then model.Validations else validations) }
        if sameCells then changed else invalidateFilterProjection changed

/// Insert or delete rows/columns on a sheet. `targetsEditedSheet` tells which
/// sheet qualifiers in this sheet's own formulas refer back to it (always the
/// unqualified ones; a workbook also passes the sheet's name).
let applyStructureEditWith (targetsEditedSheet: string option -> bool) (edit: StructureEdit) (model: SpreadsheetModel) : SpreadsheetModel =
    let onRows, isInsert, at, count = structureEditParts edit
    if count = 0 || at < 0 then model else
    let moveIndex = moveEditIndex isInsert at count
    let moveSpan = moveEditSpan isInsert at count
    let moveAddr (addr: CellAddress) =
        if onRows then moveIndex addr.Row |> Option.map (fun row -> { addr with Row = row })
        else moveIndex addr.Col |> Option.map (fun col -> { addr with Col = col })
    let moveAddressMap (items: Map<CellAddress, 'T>) : Map<CellAddress, 'T> =
        items
        |> Map.fold (fun acc addr value ->
            match moveAddr addr with
            | Some moved -> Map.add moved value acc
            | None -> acc) Map.empty
    let moveIndexMap (items: Map<int, 'T>) : Map<int, 'T> =
        items
        |> Map.fold (fun acc index value ->
            match moveIndex index with
            | Some moved -> Map.add moved value acc
            | None -> acc) Map.empty

    let cells =
        model.Cells
        |> Map.fold (fun acc addr cellData ->
            match moveAddr addr with
            | None -> acc
            | Some moved ->
                let cellData =
                    match cellData.Value with
                    | Formula (formula, cached) ->
                        let adjusted = adjustFormulaForEdit edit targetsEditedSheet formula
                        if adjusted = formula then cellData
                        else { cellData with Value = Formula (adjusted, cached) }
                    | _ -> cellData
                Map.add moved cellData acc) Map.empty

    // A deleted selection falls back onto the row/column that took its place.
    let moveSelected index =
        match moveIndex index with
        | Some moved -> moved
        | None -> max 0 (at - 1)
    let selection =
        match model.Selection with
        | SingleCell (row, col) ->
            if onRows then SingleCell (moveSelected row, col) else SingleCell (row, moveSelected col)
        | Range (sr, sc, er, ec) ->
            if onRows then Range (moveSelected sr, sc, moveSelected er, ec)
            else Range (sr, moveSelected sc, er, moveSelected ec)

    let merges =
        model.MergedCells
        |> List.choose (fun merge ->
            if onRows then
                moveSpan merge.StartRow merge.EndRow
                |> Option.map (fun (lo, hi) -> { merge with StartRow = lo; EndRow = hi })
            else
                moveSpan merge.StartCol merge.EndCol
                |> Option.map (fun (lo, hi) -> { merge with StartCol = lo; EndCol = hi }))
        // A merge squeezed down to one cell is no merge at all.
        |> List.filter (fun merge -> merge.StartRow <> merge.EndRow || merge.StartCol <> merge.EndCol)

    // Table coordinates are 1-indexed; spans are computed 0-indexed.
    let tables =
        model.Tables
        |> List.choose (fun table ->
            if onRows then
                moveSpan (table.StartRow - 1) (table.EndRow - 1)
                |> Option.map (fun (lo, hi) -> { table with StartRow = lo + 1; EndRow = hi + 1 })
            else
                moveSpan (table.StartCol - 1) (table.EndCol - 1)
                |> Option.map (fun (lo, hi) ->
                    let autoFilter =
                        table.AutoFilter
                        |> Option.bind (fun (first, last) -> moveSpan (first - 1) (last - 1))
                        |> Option.map (fun (first, last) -> first + 1, last + 1)
                    { table with StartCol = lo + 1; EndCol = hi + 1; AutoFilter = autoFilter }))

    let spill =
        model.Spill
        |> Map.fold (fun acc ghost anchor ->
            match moveAddr ghost, moveAddr anchor with
            | Some movedGhost, Some movedAnchor -> Map.add movedGhost movedAnchor acc
            | _ -> acc) Map.empty

    // Print area and print titles follow the rows/columns they name; a
    // deleted print area or title band is dropped. Manual breaks sit *after*
    // a 1-based row/column, i.e. on the 0-based index before it.
    let pageSetup =
        let setup = model.PageSetup
        let printArea =
            setup.PrintArea
            |> Option.bind (fun area ->
                if onRows then
                    moveSpan area.From.Row area.To.Row
                    |> Option.map (fun (lo, hi) -> { From = { area.From with Row = lo }; To = { area.To with Row = hi } })
                else
                    moveSpan area.From.Col area.To.Col
                    |> Option.map (fun (lo, hi) -> { From = { area.From with Col = lo }; To = { area.To with Col = hi } }))
        let moveBreaks (breaks: int list) =
            breaks
            |> List.choose (fun after -> moveIndex (after - 1) |> Option.map (fun index -> index + 1))
            |> List.distinct
        let updated =
            { setup with
                PrintArea = printArea
                RepeatRows = if onRows then setup.RepeatRows |> Option.bind (fun (lo, hi) -> moveSpan lo hi) else setup.RepeatRows
                RepeatCols = if onRows then setup.RepeatCols else setup.RepeatCols |> Option.bind (fun (lo, hi) -> moveSpan lo hi)
                ManualRowBreaks = if onRows then moveBreaks setup.ManualRowBreaks else setup.ManualRowBreaks
                ManualColBreaks = if onRows then setup.ManualColBreaks else moveBreaks setup.ManualColBreaks }
        if updated = setup then setup else updated

    // Conditional formats follow their ranges; a rule left with no range is
    // gone, and the formulas inside the rules are re-pointed like any other.
    let conditionalFormats =
        let moveRange (range: CellRange) =
            if onRows then
                // A whole-column rule stays a whole-column rule.
                if range.From.Row = 0 && range.To.Row >= 1048575 then Some range
                else
                    moveSpan range.From.Row range.To.Row
                    |> Option.map (fun (lo, hi) -> { From = { range.From with Row = lo }; To = { range.To with Row = hi } })
            else
                if range.From.Col = 0 && range.To.Col >= 16383 then Some range
                else
                    moveSpan range.From.Col range.To.Col
                    |> Option.map (fun (lo, hi) -> { From = { range.From with Col = lo }; To = { range.To with Col = hi } })
        let adjust (formula: string) = adjustFormulaForEdit edit targetsEditedSheet formula
        let adjustThreshold (threshold: CfThreshold) =
            match threshold.Kind with
            | CfvoFormula -> { threshold with Value = threshold.Value |> Option.map adjust }
            | _ -> threshold
        let adjustRule (rule: ConditionalFormatRule) =
            match rule with
            | CellIs (op, first, second) -> CellIs (op, adjust first, second |> Option.map adjust)
            | Expression formula -> Expression (adjust formula)
            | ColorScale stops -> ColorScale (stops |> List.map (fun (threshold, color) -> adjustThreshold threshold, color))
            | DataBar (low, high, color) -> DataBar (adjustThreshold low, adjustThreshold high, color)
            | IconSet (family, thresholds, reverse) -> IconSet (family, thresholds |> List.map adjustThreshold, reverse)
            | other -> other
        model.ConditionalFormats
        |> List.choose (fun entry ->
            match entry.Ranges |> List.choose moveRange with
            | [] -> None
            | ranges ->
                let rule = adjustRule entry.Format.Rule
                if ranges = entry.Ranges && rule = entry.Format.Rule then Some entry
                else Some { entry with Ranges = ranges; Format = { entry.Format with Rule = rule } })

    // Validations follow their ranges, as conditional formats do; one left
    // with no range is gone. The formulas in their rules are re-pointed and
    // their anchors follow.
    let validations =
        let adjust (formula: string) = adjustFormulaForEdit edit targetsEditedSheet formula
        let moveRange (range: CellRange) =
            if onRows then
                if range.From.Row = 0 && range.To.Row >= 1048575 then Some range
                else
                    moveSpan range.From.Row range.To.Row
                    |> Option.map (fun (lo, hi) -> { From = { range.From with Row = lo }; To = { range.To with Row = hi } })
            else
                if range.From.Col = 0 && range.To.Col >= 16383 then Some range
                else
                    moveSpan range.From.Col range.To.Col
                    |> Option.map (fun (lo, hi) -> { From = { range.From with Col = lo }; To = { range.To with Col = hi } })
        model.Validations
        |> List.choose (fun entry ->
            match entry.Ranges |> List.choose moveRange with
            | [] -> None
            | ranges ->
                let validation = entry.Validation
                let rule =
                    match validation.Rule with
                    | Comparison (kind, op, first, second) -> Comparison (kind, op, adjust first, second |> Option.map adjust)
                    | CustomFormula formula -> CustomFormula (adjust formula)
                    | ListSource source ->
                        // (A defined name is not a reference: it stays as it is.)
                        let adjusted = adjust ("=" + source)
                        ListSource (if adjusted.StartsWith "=" then adjusted.Substring 1 else adjusted)
                    | other -> other
                // The formulas stay written for the cell they were written
                // for; when that cell is deleted, for the first cell left.
                let oldAnchor = defaultArg validation.Anchor entry.Ranges.Head.From
                let newAnchor = moveAddr oldAnchor |> Option.defaultValue ranges.Head.From
                let anchor = if validation.Anchor.IsNone && newAnchor = ranges.Head.From then None else Some newAnchor
                if ranges = entry.Ranges && rule = validation.Rule && anchor = validation.Anchor then Some entry
                else Some { entry with Ranges = ranges; Validation = { validation with Rule = rule; Anchor = anchor } })

    // Drawing objects follow their anchor cells. An object whose anchor row
    // or column is deleted is not deleted with it: it moves to where the
    // deletion happened. A two-cell anchor grows or shrinks with what is
    // inserted or deleted between its corners, unless it says "oneCell"
    // (move, keep the size) or "absolute" (stay).
    let drawings =
        let moveDrawing (drawing: SheetDrawing) =
            let anchor = drawing.Anchor
            if anchor.Kind = "absolute" || anchor.EditAs = Some "absolute" then drawing
            else
                let fromIndex = if onRows then anchor.FromRow else anchor.FromCol
                let fromOffset = if onRows then anchor.FromOffsetY else anchor.FromOffsetX
                let newFrom, newFromOffset =
                    match moveIndex fromIndex with
                    | Some moved -> moved, fromOffset
                    | None -> at, 0.0
                let toIndex = if onRows then anchor.ToRow else anchor.ToCol
                let toOffset = if onRows then anchor.ToOffsetY else anchor.ToOffsetX
                let newTo, newToOffset =
                    match toIndex with
                    | None -> None, toOffset
                    | Some index when anchor.EditAs = Some "oneCell" -> Some (max newFrom (index + (newFrom - fromIndex))), toOffset
                    | Some index ->
                        match moveIndex index with
                        | Some moved -> Some (max newFrom moved), toOffset
                        | None -> Some (max newFrom at), 0.0
                let moved =
                    if onRows then { anchor with FromRow = newFrom; FromOffsetY = newFromOffset; ToRow = newTo; ToOffsetY = newToOffset }
                    else { anchor with FromCol = newFrom; FromOffsetX = newFromOffset; ToCol = newTo; ToOffsetX = newToOffset }
                if moved = anchor then drawing else { drawing with Anchor = moved }
        if model.Drawings.IsEmpty then model.Drawings else model.Drawings |> List.map moveDrawing

    // A range that is not a rule: it follows its cells, and is gone when
    // every row / column of it was deleted.
    let moveCellRange (range: CellRange) : CellRange option =
        if onRows then
            moveSpan range.From.Row range.To.Row
            |> Option.map (fun (lo, hi) -> { From = { range.From with Row = lo }; To = { range.To with Row = hi } })
        else
            moveSpan range.From.Col range.To.Col
            |> Option.map (fun (lo, hi) -> { From = { range.From with Col = lo }; To = { range.To with Col = hi } })
    let moveRanged (range: 'T -> CellRange) (withRange: 'T -> CellRange -> 'T) (items: 'T list) : 'T list =
        if items.IsEmpty then items
        else
            let moved =
                items
                |> List.choose (fun item ->
                    moveCellRange (range item)
                    |> Option.map (fun next -> if next = range item then item else withRange item next))
            if moved.Length = items.Length && List.forall2 (fun (a: 'T) (b: 'T) -> obj.ReferenceEquals(a, b)) moved items then items else moved
    let hyperlinks = model.Hyperlinks |> moveRanged (fun link -> link.Range) (fun link range -> { link with Range = range })
    let keptRanges = model.KeptRanges |> moveRanged (fun kept -> kept.Range) (fun kept range -> { kept with Range = range })
    // The input cells of a data table are references: they move like the
    // references of a formula, and a deleted one is marked as such.
    let dataTables =
        let moveInput (reference: string) (deleted: bool) =
            if deleted || reference = "" then reference, deleted
            else
                let adjusted = adjustFormulaForEdit edit targetsEditedSheet ("=" + reference)
                let text = if adjusted.StartsWith "=" then adjusted.Substring 1 else adjusted
                if text.Contains "#REF!" then reference, true else text, false
        model.DataTables
        |> moveRanged (fun table -> table.Range) (fun table range -> { table with Range = range })
        |> fun tables ->
            if tables.IsEmpty then tables
            else
                let adjusted =
                    tables
                    |> List.map (fun table ->
                        let input1, deleted1 = moveInput table.Input1 table.Input1Deleted
                        let input2, deleted2 =
                            match table.Input2 with
                            | Some reference ->
                                let text, deleted = moveInput reference table.Input2Deleted
                                Some text, deleted
                            | None -> None, table.Input2Deleted
                        if input1 = table.Input1 && input2 = table.Input2 && deleted1 = table.Input1Deleted && deleted2 = table.Input2Deleted then table
                        else { table with Input1 = input1; Input2 = input2; Input1Deleted = deleted1; Input2Deleted = deleted2 })
                if List.forall2 (fun (a: DataTable) (b: DataTable) -> obj.ReferenceEquals(a, b)) adjusted tables then tables else adjusted
    // Column default styles belong to their columns.
    let columnStyles =
        if onRows || model.ColumnStyles.IsEmpty then model.ColumnStyles
        else
            model.ColumnStyles
            |> List.choose (fun style ->
                // A whole-sheet style (to the last column) stays one.
                if style.First = 0 && style.Last >= 16383 then Some style
                else
                    moveSpan style.First style.Last
                    |> Option.map (fun (first, last) ->
                        let last = if style.Last >= 16383 then 16383 else min 16383 last
                        if first = style.First && last = style.Last then style else { style with First = first; Last = last }))

    let delta = if isInsert then count else -count
    { model with
        StructureVersion = model.StructureVersion + 1
        Hyperlinks = hyperlinks
        KeptRanges = keptRanges
        DataTables = dataTables
        ColumnStyles = columnStyles
        Drawings = drawings
        Threads = (if Map.isEmpty model.Threads then model.Threads else moveAddressMap model.Threads)
        Cells = cells
        PageSetup = pageSetup
        Selection = selection
        RowCount = if onRows then max 1 (model.RowCount + delta) else model.RowCount
        ColCount = if onRows then model.ColCount else max 1 (model.ColCount + delta)
        RowHeights = if onRows then moveIndexMap model.RowHeights else model.RowHeights
        ColumnWidths = if onRows then model.ColumnWidths else moveIndexMap model.ColumnWidths
        Axes =
            (let axes = model.Axes
             if obj.ReferenceEquals(axes, emptySheetAxes) then axes
             else
                let moveSet (items: Set<int>) = items |> Seq.choose moveIndex |> Set.ofSeq
                if onRows then { axes with HiddenRowSizes = moveIndexMap axes.HiddenRowSizes; RowLevels = moveIndexMap axes.RowLevels; CollapsedRows = moveSet axes.CollapsedRows }
                else { axes with HiddenColSizes = moveIndexMap axes.HiddenColSizes; ColLevels = moveIndexMap axes.ColLevels; CollapsedCols = moveSet axes.CollapsedCols })
        ConditionalFormats = conditionalFormats
        RichRuns = moveAddressMap model.RichRuns
        MergedCells = merges
        Tables = tables
        CfBoldCells = model.CfBoldCells |> Seq.choose moveAddr |> Set.ofSeq
        Spill = spill
        Validations = validations
        AutoFilter = model.AutoFilter |> Option.bind (fun filter -> moveCellRange filter.Range |> Option.map (fun range -> { filter with Range = range }))
        FilterHiddenRows = (if onRows && model.FilterHiddenRows.Length > 0 then model.FilterHiddenRows |> Array.choose moveIndex else model.FilterHiddenRows)
        FilterVisibleRows = None
        Filters =
            if onRows then model.Filters
            else
                model.Filters
                |> Map.fold (fun acc col filter ->
                    match moveIndex col with
                    | Some moved -> Map.add moved { filter with Column = moved } acc
                    | None -> acc) Map.empty
        SortState =
            if onRows then model.SortState
            else
                model.SortState
                |> Option.bind (fun sort -> moveIndex sort.Column |> Option.map (fun col -> { sort with Column = col }))
        SearchResults = model.SearchResults |> Array.choose moveAddr
        SearchIndex = 0
    }
    |> fun edited ->
        // The rows a filter hides moved with their rows; a filter whose range is gone is gone.
        match edited.AutoFilter with
        | None when not (Map.isEmpty edited.Filters) || edited.FilterHiddenRows.Length > 0 -> { edited with Filters = Map.empty; FilterHiddenRows = [||] }
        | Some filter ->
            let filters = edited.Filters |> Map.filter (fun col _ -> col >= filter.Range.From.Col && col <= filter.Range.To.Col)
            if filters.Count = model.Filters.Count then edited
            elif Map.isEmpty filters then { edited with Filters = filters; FilterHiddenRows = [||]; AutoFilter = (if filter.Implicit then None else edited.AutoFilter) }
            else reapplyFilter { edited with Filters = filters }
        | None -> edited

/// Insert or delete rows/columns on a stand-alone sheet. Unqualified
/// references in its formulas follow the edit; a workbook host should call
/// `Workbook.applyStructureEdit`, which also fixes the other sheets, defined
/// names and structured-reference tables.
let applyStructureEdit (edit: StructureEdit) (model: SpreadsheetModel) : SpreadsheetModel =
    applyStructureEditWith Option.isNone edit model

/// Insert or delete rows/columns on a sheet whose name the host knows but
/// that is edited outside a `Workbook` (a host keeping its own sheet list).
/// References that name the sheet itself (`Data!A1` on sheet Data) follow the
/// edit just like unqualified ones — the same predicate
/// `Workbook.applyStructureEdit` uses for the edited sheet.
let applyStructureEditNamed (sheetName: string) (edit: StructureEdit) (model: SpreadsheetModel) : SpreadsheetModel =
    let targetsEditedSheet (qualifier: string option) =
        match qualifier with
        | None -> true
        | Some name -> name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)
    applyStructureEditWith targetsEditedSheet edit model

/// Insert a row at the specified index (shifts rows down)
let insertRow (atRow: int) (model: SpreadsheetModel): SpreadsheetModel =
    applyStructureEdit (InsertRows (atRow, 1)) model

/// Delete a row at the specified index (shifts rows up)
let deleteRow (atRow: int) (model: SpreadsheetModel): SpreadsheetModel =
    applyStructureEdit (DeleteRows (atRow, 1)) model

/// Insert a column at the specified index (shifts columns right)
let insertColumn (atCol: int) (model: SpreadsheetModel): SpreadsheetModel =
    applyStructureEdit (InsertCols (atCol, 1)) model

/// Delete a column at the specified index (shifts columns left)
let deleteColumn (atCol: int) (model: SpreadsheetModel): SpreadsheetModel =
    applyStructureEdit (DeleteCols (atCol, 1)) model

/// Get row height (with default)
let getRowHeight (row: int) (model: SpreadsheetModel): float =
    if isFilterHiddenRow row model then
        0.0
    else
        Map.tryFind row model.RowHeights
        |> Option.defaultValue 25.0

/// Set row height
let setRowHeight (row: int) (height: float) (model: SpreadsheetModel): SpreadsheetModel =
    if height = 25.0 then
        if Map.containsKey row model.RowHeights then { model with RowHeights = Map.remove row model.RowHeights } else model
    else
        match Map.tryFind row model.RowHeights with
        | Some existing when existing = height -> model
        | _ -> { model with RowHeights = Map.add row height model.RowHeights }

/// Get column width (with default)
let getColumnWidth (col: int) (model: SpreadsheetModel): float =
    Map.tryFind col model.ColumnWidths
    |> Option.defaultValue 100.0

/// Set column width
let setColumnWidth (col: int) (width: float) (model: SpreadsheetModel): SpreadsheetModel =
    if width = 100.0 then
        if Map.containsKey col model.ColumnWidths then { model with ColumnWidths = Map.remove col model.ColumnWidths } else model
    else
        match Map.tryFind col model.ColumnWidths with
        | Some existing when existing = width -> model
        | _ -> { model with ColumnWidths = Map.add col width model.ColumnWidths }

let usedBounds (model: SpreadsheetModel) : (int * int) option =
    let mutable any = false
    let mutable maxRow = 0
    let mutable maxCol = 0
    for KeyValue(addr, _) in model.Cells do
        any <- true
        if addr.Row > maxRow then maxRow <- addr.Row
        if addr.Col > maxCol then maxCol <- addr.Col
    for merge in model.MergedCells do
        any <- true
        if merge.EndRow > maxRow then maxRow <- merge.EndRow
        if merge.EndCol > maxCol then maxCol <- merge.EndCol
    if any then Some (maxRow, maxCol) else None

let lastUsedColumnInRow (row: int) (model: SpreadsheetModel) : int option =
    let mutable found = false
    let mutable lastCol = 0
    for KeyValue(addr, _) in model.Cells do
        if addr.Row = row then
            found <- true
            if addr.Col > lastCol then lastCol <- addr.Col
    for merge in model.MergedCells do
        if row >= merge.StartRow && row <= merge.EndRow then
            found <- true
            if merge.EndCol > lastCol then lastCol <- merge.EndCol
    if found then Some lastCol else None

/// Search cells for a query string
let private containsSearchText (query: string) (caseSensitive: bool) (text: string) =
#if FABLE_COMPILER
    if caseSensitive then text.Contains(query)
    else text.ToLowerInvariant().Contains(query.ToLowerInvariant())
#else
    let comparison = if caseSensitive then StringComparison.Ordinal else StringComparison.OrdinalIgnoreCase
    text.IndexOf(query, comparison) >= 0
#endif

let searchCells (query: string) (caseSensitive: bool) (model: SpreadsheetModel): CellAddress array =
    model.Cells
    |> Seq.choose (fun (KeyValue(addr, cellData)) ->
        let cellText =
            match cellData.Value with
            | Formula (formula, _) -> formula  // Search in formula text
            | other -> displayValue other

        if containsSearchText query caseSensitive cellText then Some addr else None)
    |> Seq.sortBy (fun addr -> (addr.Row, addr.Col))  // Sort by row then column
    |> Seq.toArray

/// Initialize search with a query
let startSearch (query: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel =
    let results = searchCells query caseSensitive model
    { model with
        SearchQuery = Some query
        SearchCaseSensitive = Some caseSensitive
        SearchResults = results
        SearchIndex = 0
    }

let nextSearchResultValue (model: SpreadsheetModel): SpreadsheetModel voption =
    if model.SearchResults.Length = 0 then
        ValueNone
    else
        let newIndex = (model.SearchIndex + 1) % model.SearchResults.Length
        let nextAddr = model.SearchResults.[newIndex]
        let updatedModel =
            { model with
                SearchIndex = newIndex
                Selection = SingleCell (nextAddr.Row, nextAddr.Col)
            }
        ValueSome updatedModel

/// Move to next search result
let nextSearchResult (model: SpreadsheetModel): SpreadsheetModel option =
    match nextSearchResultValue model with
    | ValueSome next -> Some next
    | ValueNone -> None

/// Replace text at current selection
let replaceAtSelection (findText: string) (replaceText: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel option =
    match model.Selection with
    | SingleCell (row, col) ->
        let addr = { Row = row; Col = col }
        match getCell addr model with
        | Some cellData ->
            let currentText = displayValue cellData.Value
            let textToCheck = if caseSensitive then currentText else currentText.ToLower()
            let findToCheck = if caseSensitive then findText else findText.ToLower()

            if textToCheck.Contains(findToCheck) then
                // Perform replacement
                let newText =
                    if caseSensitive then
                        currentText.Replace(findText, replaceText)
                    else
                        // Case-insensitive replacement is more complex - use simple approach
                        let regex = System.Text.RegularExpressions.Regex(findText, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                        regex.Replace(currentText, replaceText)

                let newValue = String newText
                let newCellData = { cellData with Value = newValue }
                Some (setCell addr newCellData model)
            else
                None
        | None -> None
    | _ -> None

/// Replace all occurrences
let rec replaceAll (findText: string) (replaceText: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel =
    let updated, _ = replaceAllWithCount findText replaceText caseSensitive model
    updated

and replaceAllWithCount (findText: string) (replaceText: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel * int =
    let updated, count, _, _ = replaceAllWithChanges findText replaceText caseSensitive model
    updated, count

and replaceAllWithChanges (findText: string) (replaceText: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel * int * ResizeArray<CellAddress * CellData option> * ResizeArray<CellAddress * CellData option> =
    let oldCells = ResizeArray<CellAddress * CellData option>()
    let newCells = ResizeArray<CellAddress * CellData option>()
    if String.IsNullOrEmpty findText then
        model, 0, oldCells, newCells
    else
        let replacement = replaceText.Replace("$", "$$")
        let regex =
            if caseSensitive then None
            else Some (System.Text.RegularExpressions.Regex(System.Text.RegularExpressions.Regex.Escape findText, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
        let replaceIn (text: string) =
            if caseSensitive then text.Replace(findText, replaceText)
            else regex.Value.Replace(text, replacement)

        let mutable cells = model.Cells
        let mutable count = 0
        for KeyValue(addr, cellData) in model.Cells do
            let currentText = displayValue cellData.Value
            if containsSearchText findText caseSensitive currentText then
                let newText = replaceIn currentText
                if newText <> currentText then
                    let nextCell = { cellData with Value = String newText }
                    oldCells.Add(addr, Some cellData)
                    newCells.Add(addr, Some nextCell)
                    cells <- Map.add addr nextCell cells
                    count <- count + 1
        if count = 0 then model, 0, oldCells, newCells
        else ({ model with Cells = cells } |> invalidateFilterProjection), count, oldCells, newCells

/// Does a cell value match a *literal* conditional format rule (the five
/// value-only kinds)? Rules that need the range's other cells or the formula
/// engine are evaluated by `ConditionalFormatting`.
let evaluateRule (rule: ConditionalFormatRule) (value: CellValue): bool =
    match rule, value with
    | GreaterThan threshold, Number n -> n > threshold
    | LessThan threshold, Number n -> n < threshold
    | Between (min, max), Number n -> n >= min && n <= max
    | Equal target, Number n -> n = target
    | ContainsText text, String s -> s.Contains(text)
    | ContainsText text, Formula (formula, _) -> formula.Contains(text)
    | _ -> false

let cellRangeContains (addr: CellAddress) (range: CellRange) =
    addr.Row >= range.From.Row && addr.Row <= range.To.Row && addr.Col >= range.From.Col && addr.Col <= range.To.Col

/// The conditional formats whose ranges cover a cell, in priority order.
let conditionalFormatsAt (addr: CellAddress) (model: SpreadsheetModel) : ConditionalFormatRange list =
    model.ConditionalFormats
    |> List.filter (fun entry -> entry.Ranges |> List.exists (cellRangeContains addr))
    |> List.sortBy (fun entry -> entry.Priority)

/// The style of the first literal rule on a cell that the value matches.
/// Covers only the value-only rule kinds; see `ConditionalFormatting.formatAt`
/// for the full evaluation.
let getConditionalFormat (addr: CellAddress) (value: CellValue) (model: SpreadsheetModel): CellFormat option =
    conditionalFormatsAt addr model
    |> List.tryFind (fun entry -> evaluateRule entry.Format.Rule value)
    |> Option.map (fun entry -> entry.Format.Style)

/// Add a conditional format for some ranges. Like a rule made in Excel's
/// dialog, it takes precedence over the rules already there.
let addConditionalFormatRange (ranges: CellRange list) (format: ConditionalFormat) (model: SpreadsheetModel): SpreadsheetModel =
    let priority =
        match model.ConditionalFormats with
        | [] -> 1
        | existing -> (existing |> List.map (fun entry -> entry.Priority) |> List.min) - 1
    let id = (model.ConditionalFormats |> List.fold (fun highest entry -> max highest entry.Id) 0) + 1
    { model with
        ConditionalFormats = { Id = id; Ranges = ranges; Format = format; Priority = priority; StopIfTrue = false } :: model.ConditionalFormats }

/// Add a conditional format rule to a cell
let addConditionalFormat (addr: CellAddress) (format: ConditionalFormat) (model: SpreadsheetModel): SpreadsheetModel =
    addConditionalFormatRange [ { From = addr; To = addr } ] format model

/// Remove all conditional formats from a cell. A rule covering more than the
/// cell keeps the rest of its range.
let removeConditionalFormats (addr: CellAddress) (model: SpreadsheetModel): SpreadsheetModel =
    let without (range: CellRange) : CellRange list =
        if not (cellRangeContains addr range) then [ range ]
        else
            [ if range.From.Row < addr.Row then
                  yield { From = range.From; To = { Row = addr.Row - 1; Col = range.To.Col } }
              if range.From.Col < addr.Col then
                  yield { From = { Row = addr.Row; Col = range.From.Col }; To = { Row = addr.Row; Col = addr.Col - 1 } }
              if range.To.Col > addr.Col then
                  yield { From = { Row = addr.Row; Col = addr.Col + 1 }; To = { Row = addr.Row; Col = range.To.Col } }
              if range.To.Row > addr.Row then
                  yield { From = { Row = addr.Row + 1; Col = range.From.Col }; To = range.To } ]
    let remaining =
        model.ConditionalFormats
        |> List.choose (fun entry ->
            match entry.Ranges |> List.collect without with
            | [] -> None
            | ranges -> Some (if ranges = entry.Ranges then entry else { entry with Ranges = ranges }))
    if remaining = model.ConditionalFormats then model else { model with ConditionalFormats = remaining }

// --- Conditional-format VISUALS: data bars / color scales / icon sets -------
// These compute over a range and bake the result into each cell's CellFormat
// (DataBar / IconSet / BgColor), which the renderers already paint. Applying
// them is a range operation (from the CF dialog or programmatically).

let rec private unwrapVal (v: CellValue) : CellValue =
    match v with Formula (_, cached) -> unwrapVal cached | x -> x

let private cellNumber (model: SpreadsheetModel) (addr: CellAddress) : float option =
    match getCellValue addr model |> unwrapVal with
    | Number n -> Some n
    | Boolean b -> Some(if b then 1.0 else 0.0)
    | _ -> None

let private rangeNumbers sr sc er ec (model: SpreadsheetModel) : (CellAddress * float) list =
    [ for r in sr .. er do
        for c in sc .. ec do
            match cellNumber model { Row = r; Col = c } with
            | Some n -> yield { Row = r; Col = c }, n
            | None -> () ]

let private updateFormat (addr: CellAddress) (f: CellFormat -> CellFormat) (model: SpreadsheetModel) : SpreadsheetModel =
    match getCell addr model with
    | Some cd -> setCell addr { cd with Format = f cd.Format } model
    | None -> setCell addr { Value = Empty; Format = f defaultFormat; Comment = None } model

let private hexToRgb (s: string) =
    let h = (if s.StartsWith "#" then s.Substring 1 else s)
    if h.Length = 6 then
        try (Convert.ToInt32(h.Substring(0, 2), 16), Convert.ToInt32(h.Substring(2, 2), 16), Convert.ToInt32(h.Substring(4, 2), 16))
        with _ -> (255, 255, 255)
    else (255, 255, 255)

let private lerpHex (a: string) (b: string) (t: float) : string =
    let (ar, ag, ab), (br, bg, bb) = hexToRgb a, hexToRgb b
    let m x y = int (float x + (float y - float x) * t)
    sprintf "#%02X%02X%02X" (m ar br) (m ag bg) (m ab bb)

/// Data bars over a range, baselined at 0 when all values are ≥ 0 (Excel's
/// automatic behaviour), else min→max.
let applyDataBars (sr: int) (sc: int) (er: int) (ec: int) (color: string) (model: SpreadsheetModel) : SpreadsheetModel =
    match rangeNumbers sr sc er ec model with
    | [] -> model
    | cells ->
        let vals = cells |> List.map snd
        let lo = min 0.0 (List.min vals)
        let hi = max 0.0 (List.max vals)
        let span = if hi - lo < 1e-12 then 1.0 else hi - lo
        cells
        |> List.fold (fun m (addr, v) ->
            updateFormat addr (fun f -> { f with DataBar = Some(max 0.0 (min 1.0 ((v - lo) / span)), color) }) m) model

/// 2- or 3-color scale (pass `mid = None` for 2-color) → per-cell BgColor.
let applyColorScale (sr: int) (sc: int) (er: int) (ec: int) (loColor: string) (mid: string option) (hiColor: string) (model: SpreadsheetModel) : SpreadsheetModel =
    match rangeNumbers sr sc er ec model with
    | [] -> model
    | cells ->
        let vals = cells |> List.map snd
        let lo, hi = List.min vals, List.max vals
        let span = if hi - lo < 1e-12 then 1.0 else hi - lo
        cells
        |> List.fold (fun m (addr, v) ->
            let t = (v - lo) / span
            let color =
                match mid with
                | None -> lerpHex loColor hiColor t
                | Some mc -> if t <= 0.5 then lerpHex loColor mc (t * 2.0) else lerpHex mc hiColor ((t - 0.5) * 2.0)
            updateFormat addr (fun f -> { f with BgColor = Some color }) m) model

/// Icon set over a range: `count` buckets by value (Excel's percent
/// thresholds), each cell tagged with its bucket index for the renderer.
let applyIconSet (sr: int) (sc: int) (er: int) (ec: int) (count: int) (model: SpreadsheetModel) : SpreadsheetModel =
    match rangeNumbers sr sc er ec model with
    | [] -> model
    | cells when count > 1 ->
        let vals = cells |> List.map snd
        let lo, hi = List.min vals, List.max vals
        let span = if hi - lo < 1e-12 then 1.0 else hi - lo
        cells
        |> List.fold (fun m (addr, v) ->
            let bucket = min (count - 1) (int ((v - lo) / span * float count))
            updateFormat addr (fun f -> { f with IconSet = Some(bucket, count) }) m) model
    | _ -> model

// --- Data validation --------------------------------------------------------

let private displayText (v: CellValue) : string =
    match unwrapVal v with
    | String s -> s
    | Number n -> if n = floor n && abs n < 1e15 then string (int64 n) else string n
    | Boolean b -> if b then "TRUE" else "FALSE"
    | _ -> ""

/// Does `x` satisfy the comparison against its operand(s)?
let compareHolds (op: CompareOperator) (x: float) (first: float) (second: float option) : bool =
    let lo, hi =
        match second with
        | Some other -> min first other, max first other
        | None -> first, first
    let eps = 1e-9
    match op with
    | OpGreaterThan -> x > first + eps
    | OpLessThan -> x < first - eps
    | OpGreaterOrEqual -> x >= first - eps
    | OpLessOrEqual -> x <= first + eps
    | OpEqual -> abs (x - first) <= eps
    | OpNotEqual -> abs (x - first) > eps
    | OpBetween -> x >= lo - eps && x <= hi + eps
    | OpNotBetween -> x < lo - eps || x > hi + eps

/// Check a value against a validation rule. `evaluate` gives the value of a
/// formula of the rule at the validated cell; `listValues` the current values
/// of a list rule's source range (None when it cannot be read: the rule then
/// refuses nothing). Blank is always allowed (clearing a cell is never
/// refused). Returns Ok or a reason.
let validateValueIn (listValues: string -> string list option) (evaluate: string -> CellValue) (rule: ValidationRule) (value: CellValue) : Result<unit, string> =
    let v = unwrapVal value
    let isBlank = match v with Empty -> true | String "" -> true | _ -> false
    let within lo hi x =
        (match lo with Some l -> x >= l - 1e-9 | None -> true) && (match hi with Some h -> x <= h + 1e-9 | None -> true)
    /// An operand: a plain number, or a formula to evaluate.
    let operand (text: string) : float option =
        let trimmed = text.Trim()
        match Double.TryParse(trimmed, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture) with
        | true, number -> Some number
        | _ ->
            match unwrapVal (evaluate trimmed) with
            | Number number -> Some number
            | Boolean flag -> Some (if flag then 1.0 else 0.0)
            | _ -> None
    let inList (vs: string list) =
        let t = (displayText v).Trim().ToLowerInvariant()
        if vs |> List.exists (fun a -> a.Trim().ToLowerInvariant() = t) then Result.Ok()
        else Result.Error(sprintf "Value must be one of: %s" (String.concat ", " vs))
    if isBlank then Result.Ok()
    else
        match rule with
        | AnyValue -> Result.Ok()
        | ListValues vs -> inList vs
        | ListSource source ->
            (match listValues source with
             | Some vs -> inList vs
             // A source that cannot be read does not lock the cell.
             | None -> Result.Ok())
        | WholeNumber (lo, hi) ->
            (match v with
             | Number n when n = floor n -> if within lo hi n then Result.Ok() else Result.Error "Enter a whole number within the allowed range."
             | Number _ -> Result.Error "Enter a whole number (no decimals)."
             | _ -> Result.Error "Enter a whole number.")
        | DecimalRange (lo, hi) ->
            (match v with
             | Number n -> if within lo hi n then Result.Ok() else Result.Error "Enter a number within the allowed range."
             | _ -> Result.Error "Enter a number.")
        | TextLength (lo, hi) ->
            let len = match v with String s -> float s.Length | _ -> float (displayText v).Length
            if within lo hi len then Result.Ok() else Result.Error "Text length is outside the allowed range."
        | DateRange (lo, hi) ->
            (match v with
             | Number n -> if within lo hi n then Result.Ok() else Result.Error "Enter a date within the allowed range."
             | _ -> Result.Error "Enter a date.")
        | Comparison (kind, op, first, second) ->
            let noun =
                match kind with
                | WholeKind -> "whole number"
                | DecimalKind -> "number"
                | DateKind -> "date"
                | TimeKind -> "time"
                | TextLengthKind -> "text of the allowed length"
            let measured =
                match kind, v with
                | TextLengthKind, String text -> Some (float text.Length)
                | TextLengthKind, other -> Some (float (displayText other).Length)
                | WholeKind, Number n when n <> floor n -> None
                // A time of day is the fractional part of a serial.
                | TimeKind, Number n when n < 0.0 -> None
                | _, Number n -> Some n
                | _ -> None
            match measured with
            | None -> Result.Error (sprintf "Enter a %s." noun)
            | Some x ->
                match operand first, second |> Option.map operand with
                // An operand that has no value yet cannot refuse anything.
                | None, _ | _, Some None -> Result.Ok()
                | Some firstValue, secondValue ->
                    let secondValue = secondValue |> Option.bind id
                    if compareHolds op x firstValue secondValue then Result.Ok()
                    else Result.Error (sprintf "Enter a %s that meets the rule set for this cell." noun)
        | CustomFormula formula ->
            match unwrapVal (evaluate formula) with
            | Boolean true -> Result.Ok()
            | Number n when n <> 0.0 -> Result.Ok()
            // A formula that cannot be evaluated does not lock the cell.
            | Error _ -> Result.Ok()
            | _ -> Result.Error "This value does not meet the validation rule set for this cell."

/// `validateValueIn` for a rule whose list (if it is one) is written out.
let validateValueWith (evaluate: string -> CellValue) (rule: ValidationRule) (value: CellValue) : Result<unit, string> =
    validateValueIn (fun _ -> None) evaluate rule value

/// Check a value against a validation rule without a sheet to evaluate
/// formulas on: operands must be literal numbers, and a custom-formula rule
/// accepts everything. Use `checkEntry` to validate an entry in a cell.
let validateValue (rule: ValidationRule) (value: CellValue) : Result<unit, string> =
    validateValueWith (fun _ -> Error "NAME?") rule value

// --- The validations of a sheet, by range ------------------------------------

/// What is left of a range when the cells of `hole` are taken out of it: up
/// to four ranges (above, left, right, below), the range itself when the two
/// do not meet, nothing when the hole covers it.
let subtractCellRange (hole: CellRange) (range: CellRange) : CellRange list =
    let meets =
        hole.From.Row <= range.To.Row && hole.To.Row >= range.From.Row
        && hole.From.Col <= range.To.Col && hole.To.Col >= range.From.Col
    if not meets then [ range ]
    else
        let top = max hole.From.Row range.From.Row
        let bottom = min hole.To.Row range.To.Row
        [ if range.From.Row < top then
              yield { From = range.From; To = { Row = top - 1; Col = range.To.Col } }
          if range.From.Col < hole.From.Col then
              yield { From = { Row = top; Col = range.From.Col }; To = { Row = bottom; Col = hole.From.Col - 1 } }
          if range.To.Col > hole.To.Col then
              yield { From = { Row = top; Col = hole.To.Col + 1 }; To = { Row = bottom; Col = range.To.Col } }
          if range.To.Row > bottom then
              yield { From = { Row = bottom + 1; Col = range.From.Col }; To = range.To } ]

/// The validations of a sheet laid out for looking one up by cell: every
/// range with the position of its validation, and the box around them all.
type private ValidationIndex = {
    Source: ValidationRange list
    Entries: ValidationRange array
    /// r0, c0, r1, c1, entry - five numbers per range.
    Boxes: int array
    Top: int
    Left: int
    Bottom: int
    Right: int
}

/// The index of the validation list looked up last (a sheet's list is a
/// value that is replaced when it changes, so it is known by identity; a few
/// are kept because renders of several sheets interleave).
let mutable private validationIndexCache : ValidationIndex list = []

let private buildValidationIndex (validations: ValidationRange list) : ValidationIndex =
    let entries = List.toArray validations
    let boxes = ResizeArray<int>()
    let mutable top = System.Int32.MaxValue
    let mutable left = System.Int32.MaxValue
    let mutable bottom = -1
    let mutable right = -1
    entries |> Array.iteri (fun index entry ->
        for range in entry.Ranges do
            boxes.Add range.From.Row; boxes.Add range.From.Col; boxes.Add range.To.Row; boxes.Add range.To.Col; boxes.Add index
            top <- min top range.From.Row
            left <- min left range.From.Col
            bottom <- max bottom range.To.Row
            right <- max right range.To.Col)
    { Source = validations; Entries = entries; Boxes = boxes.ToArray(); Top = top; Left = left; Bottom = bottom; Right = right }

let private validationIndexOf (validations: ValidationRange list) : ValidationIndex =
    match validationIndexCache |> List.tryFind (fun index -> obj.ReferenceEquals(index.Source, validations)) with
    | Some index -> index
    | None ->
        let index = buildValidationIndex validations
        validationIndexCache <- index :: (validationIndexCache |> List.truncate 3)
        index

/// The validation whose ranges cover a cell, as it is stored.
let validationEntryAt (addr: CellAddress) (model: SpreadsheetModel) : ValidationRange option =
    if model.Validations.IsEmpty then None
    else
        let index = validationIndexOf model.Validations
        if addr.Row < index.Top || addr.Row > index.Bottom || addr.Col < index.Left || addr.Col > index.Right then None
        else
            let boxes = index.Boxes
            let mutable found = -1
            let mutable i = 0
            while found < 0 && i < boxes.Length do
                if addr.Row >= boxes.[i] && addr.Row <= boxes.[i + 2] && addr.Col >= boxes.[i + 1] && addr.Col <= boxes.[i + 3] then found <- boxes.[i + 4]
                i <- i + 5
            if found < 0 then None else Some index.Entries.[found]

/// The cell the formulas of a validation are written for.
let validationAnchor (entry: ValidationRange) : CellAddress =
    match entry.Validation.Anchor, entry.Ranges with
    | Some anchor, _ -> anchor
    | None, first :: _ -> first.From
    | None, [] -> { Row = 0; Col = 0 }

/// The validation on a cell, with `Anchor` saying which cell its formulas
/// are written for.
let getValidation (addr: CellAddress) (model: SpreadsheetModel) : DataValidation option =
    validationEntryAt addr model
    |> Option.map (fun entry ->
        match entry.Validation.Anchor with
        | Some _ -> entry.Validation
        | None -> { entry.Validation with Anchor = Some (validationAnchor entry) })

/// The id a new validation of the sheet gets.
let nextValidationId (model: SpreadsheetModel) : int =
    (model.Validations |> List.fold (fun highest entry -> max highest entry.Id) 0) + 1

/// Take ranges out of every validation of the sheet. A validation that
/// loses its first cell keeps the cell its formulas are written for; one
/// left with no range is gone.
let clearValidationRanges (holes: CellRange list) (model: SpreadsheetModel) : SpreadsheetModel =
    let remaining =
        model.Validations
        |> List.choose (fun entry ->
            let ranges = holes |> List.fold (fun ranges hole -> ranges |> List.collect (subtractCellRange hole)) entry.Ranges
            if ranges = entry.Ranges then Some entry
            elif ranges.IsEmpty then None
            else Some { entry with Ranges = ranges; Validation = { entry.Validation with Anchor = Some (validationAnchor entry) } })
    if remaining = model.Validations then model else { model with Validations = remaining }

/// Put a validation on ranges, in place of whatever validations their cells
/// had, under the id given. No cell is made or changed.
let setValidationRangesWithId (id: int) (ranges: CellRange list) (validation: DataValidation) (model: SpreadsheetModel) : SpreadsheetModel =
    let cleared = clearValidationRanges ranges model
    { cleared with Validations = cleared.Validations @ [ { Id = id; Ranges = ranges; Validation = validation } ] }

/// `setValidationRangesWithId` under the next free id.
let setValidationRanges (ranges: CellRange list) (validation: DataValidation) (model: SpreadsheetModel) : SpreadsheetModel =
    setValidationRangesWithId (nextValidationId model) ranges validation model

/// Attach a validation rule to one cell.
let setValidation (addr: CellAddress) (dv: DataValidation) (model: SpreadsheetModel) : SpreadsheetModel =
    setValidationRanges [ { From = addr; To = addr } ] dv model

/// The values a list rule's source gives when it is a range of the sheet
/// itself, shifted by (deltaRow, deltaCol) where it is relative: the texts
/// its filled cells show. None for anything else (another sheet, a name).
let sameSheetListValues (model: SpreadsheetModel) (deltaRow: int) (deltaCol: int) (source: string) : string list option =
    let shifted = shiftFormulaReferences deltaRow deltaCol ("=" + source.Trim().TrimStart('='))
    let text = shifted.TrimStart('=').Replace("$", "").Trim()
    let parts = text.Split(':')
    // A corner: its row and its column, either of which a whole-column /
    // whole-row reference leaves out.
    let corner (part: string) : (int option * int option) option =
        let part = part.Trim()
        let mutable i = 0
        while i < part.Length && Char.IsLetter part.[i] do i <- i + 1
        let letters = part.Substring(0, i)
        let digits = part.Substring i
        if part.Length = 0 || letters.Length > 3 || not (digits |> Seq.forall Char.IsDigit) then None
        else
            Some ((if digits.Length > 0 then Some (int digits - 1) else None),
                  (if letters.Length > 0 then Some (letterToColumn (letters.ToUpperInvariant())) else None))
    if text.Contains "!" || text.Contains "(" || parts.Length = 0 || parts.Length > 2 then None
    else
        match corner parts.[0], corner parts.[parts.Length - 1] with
        | Some (r0, c0), Some (r1, c1) ->
            let usedRow, usedCol = usedBounds model |> Option.defaultValue (0, 0)
            let rowLo, rowHi =
                match r0, r1 with
                | Some a, Some b -> min a b, min (max a b) usedRow
                | _ -> 0, usedRow
            let colLo, colHi =
                match c0, c1 with
                | Some a, Some b -> min a b, min (max a b) usedCol
                | _ -> 0, usedCol
            if rowLo < 0 || colLo < 0 || (r0.IsNone && c0.IsNone) then None
            else
                Some [ for row in rowLo .. rowHi do
                        for col in colLo .. colHi do
                            match Map.tryFind { Row = row; Col = col } model.Cells with
                            | Some cell ->
                                let shown = displayText (unwrapVal cell.Value)
                                if shown <> "" then yield shown
                            | None -> () ]
        | _ -> None

/// Whether a cell offers the values of a list validation in a dropdown.
let showsListDropdown (addr: CellAddress) (model: SpreadsheetModel) : bool =
    match validationEntryAt addr model with
    | Some entry ->
        entry.Validation.ShowDropdown
        && (match entry.Validation.Rule with ListValues _ | ListSource _ -> true | _ -> false)
    | None -> false

/// What validating an entry found.
type ValidationOutcome =
    /// No rule, or the entry satisfies it.
    | EntryValid
    /// The entry breaks the rule. `StopStyle` means it must be refused;
    /// the other styles are advisory — the entry is kept and the user told.
    | EntryInvalid of style: ValidationStyle * title: string option * message: string

/// Reads the values of a list rule's source for a cell of a sheet: the
/// sheet, how far the cell is from the rule's anchor (rows, columns), the
/// source text. None when the source cannot be read.
type ListSourceReader = SpreadsheetModel -> int -> int -> string -> string list option

/// Judge a value for a cell against the cell's rule. With `always` the rule
/// is applied even when its error alert is switched off (what Excel's
/// "circle invalid data" does).
let private judgeEntry (always: bool) (listValues: ListSourceReader)
                       (evaluate: string -> int -> int -> SpreadsheetModel -> CellAddress -> CellValue)
                       (addr: CellAddress) (value: CellValue) (model: SpreadsheetModel) : ValidationOutcome =
    match getValidation addr model with
    | None -> EntryValid
    | Some dv when not dv.ShowError && not always -> EntryValid
    | Some dv ->
        let existing = getCell addr model
        let trialCell (v: CellValue) =
            match existing with
            | Some cellData -> { cellData with Value = v }
            | None -> { Value = v; Format = defaultFormat; Comment = None }
        // A typed formula is judged by its result.
        let entered =
            match value with
            | Formula (formula, Empty) ->
                Formula (formula, evaluate formula 0 0 (setCell addr (trialCell value) model) addr |> unwrapVal)
            | other -> other
        let trial =
            match existing with
            | Some cellData when obj.ReferenceEquals(cellData.Value, entered) || cellData.Value = entered -> model
            | _ -> setCell addr (trialCell entered) model
        let anchor = defaultArg dv.Anchor addr
        let deltaRow, deltaCol = addr.Row - anchor.Row, addr.Col - anchor.Col
        let evaluateRuleFormula (formula: string) =
            evaluate formula deltaRow deltaCol trial addr
        match validateValueIn (listValues trial deltaRow deltaCol) evaluateRuleFormula dv.Rule entered with
        | Result.Ok () -> EntryValid
        | Result.Error fallback -> EntryInvalid (dv.Style, dv.ErrorTitle, dv.ErrorMessage |> Option.defaultValue fallback)

/// Validate a prospective entry against the cell's rule, the way Excel does
/// when the edit is committed. `evaluate` runs a formula written for one cell
/// at another (see `evaluateFormulaAt`); the rule's formulas see the sheet
/// with the entry already in place, so `=COUNTIF($A:$A,A1)=1` counts it.
/// `listValues` reads the source of a list rule.
let checkEntryIn (listValues: ListSourceReader)
                 (evaluate: string -> int -> int -> SpreadsheetModel -> CellAddress -> CellValue)
                 (addr: CellAddress) (value: CellValue) (model: SpreadsheetModel) : ValidationOutcome =
    judgeEntry false listValues evaluate addr value model

/// `checkEntryIn` with list sources read from the sheet itself.
let checkEntryWith (evaluate: string -> int -> int -> SpreadsheetModel -> CellAddress -> CellValue)
                   (addr: CellAddress) (value: CellValue) (model: SpreadsheetModel) : ValidationOutcome =
    checkEntryIn sameSheetListValues evaluate addr value model

/// `checkEntryWith` on a stand-alone sheet (formulas cannot reach other sheets).
let checkEntry (addr: CellAddress) (value: CellValue) (model: SpreadsheetModel) : ValidationOutcome =
    checkEntryWith evaluateFormulaAt addr value model

/// The filled cells whose value breaks the validation on them (Excel's
/// "circle invalid data"), in row order, each with the alert of its rule. A
/// rule whose alert is switched off counts too. Only the cells that exist
/// are looked at: a whole-column validation costs what the column holds.
let invalidCells (listValues: ListSourceReader)
                 (evaluate: string -> int -> int -> SpreadsheetModel -> CellAddress -> CellValue)
                 (model: SpreadsheetModel) : (CellAddress * ValidationOutcome) list =
    if model.Validations.IsEmpty then []
    else
        [ for KeyValue(addr, cell) in model.Cells do
            match unwrapVal cell.Value with
            | Empty | String "" -> ()
            | _ ->
                // (The other cells of a spilled array belong to its formula.)
                if not (Map.containsKey addr model.Spill) && (validationEntryAt addr model).IsSome then
                    match judgeEntry true listValues evaluate addr cell.Value model with
                    | EntryValid -> ()
                    | invalid -> yield addr, invalid ]

/// Validate a prospective edit against the cell's rule (if any). Error only
/// when the entry must be refused: a rule whose style is warning or
/// information accepts every entry (use `checkEntry` to learn about those).
let validateEdit (addr: CellAddress) (value: CellValue) (model: SpreadsheetModel) : Result<unit, string> =
    match checkEntry addr value model with
    | EntryInvalid (StopStyle, _, message) -> Result.Error message
    | _ -> Result.Ok()

/// Parse CSV string to list of rows
let parseCSV (csvText: string): string list list =
    let rows = ResizeArray<string list>()
    let mutable row = ResizeArray<string>()
    let cell = System.Text.StringBuilder()
    let mutable inQuotes = false
    let mutable i = 0

    let finishCell () =
        row.Add(cell.ToString())
        cell.Clear() |> ignore

    let finishRow () =
        finishCell()
        rows.Add(row |> Seq.toList)
        row <- ResizeArray<string>()

    while i < csvText.Length do
        let c = csvText.[i]
        if inQuotes then
            if c = '"' then
                if i + 1 < csvText.Length && csvText.[i + 1] = '"' then
                    cell.Append('"') |> ignore
                    i <- i + 1
                else
                    inQuotes <- false
            else
                cell.Append(c) |> ignore
        else
            match c with
            | '"' when cell.Length = 0 ->
                inQuotes <- true
            | ',' ->
                finishCell()
            | '\r' ->
                if i + 1 < csvText.Length && csvText.[i + 1] = '\n' then
                    i <- i + 1
                finishRow()
            | '\n' ->
                finishRow()
            | _ ->
                cell.Append(c) |> ignore
        i <- i + 1

    if cell.Length > 0 || row.Count > 0 then
        finishRow()

    rows |> Seq.toList

/// Import CSV data into spreadsheet
let rec importCSV (csvText: string) (startRow: int) (startCol: int) (model: SpreadsheetModel): SpreadsheetModel =
    let imported, _, _ = importCSVWithChanges csvText startRow startCol model
    imported

and importCSVWithChanges (csvText: string) (startRow: int) (startCol: int) (model: SpreadsheetModel): SpreadsheetModel * ResizeArray<CellAddress * CellData option> * ResizeArray<CellAddress * CellData option> =
    let rows = parseCSV csvText
    let mutable updatedModel = model
    let oldCells = ResizeArray<CellAddress * CellData option>()
    let newCells = ResizeArray<CellAddress * CellData option>()

    let record addr next =
        let old = getCell addr updatedModel
        if old <> next then
            oldCells.Add(addr, old)
            newCells.Add(addr, next)

    rows |> List.iteri (fun rowIdx rowData ->
        rowData |> List.iteri (fun colIdx cellText ->
            let addr = { Row = startRow + rowIdx; Col = startCol + colIdx }

            // Parse cell value
            let value =
                if String.IsNullOrEmpty(cellText) then
                    Empty
                elif cellText.StartsWith("=") then
                    Formula (cellText, Empty)
                elif cellText.ToLower() = "true" || cellText.ToLower() = "false" then
                    Boolean (cellText.ToLower() = "true")
                else
                    // Try to parse as number, otherwise string
                    match tryParseNumber cellText with
                    | Some n -> Number n
                    | None -> String cellText

            if value = Empty then
                record addr None
                updatedModel <- clearCell addr updatedModel
            else
                let cellData = {
                    Value = value
                    Format = defaultFormat
                    Comment = None
                }
                record addr (Some cellData)
                updatedModel <- setCell addr cellData updatedModel
        )
    )

    invalidateFilterProjection updatedModel, oldCells, newCells

/// Export spreadsheet to CSV string
let exportCSV (model: SpreadsheetModel): string =
    // Find the range of used cells
    let allAddrs = model.Cells |> Map.toList |> List.map fst

    if List.isEmpty allAddrs then
        ""
    else
        let maxRow = allAddrs |> List.map (fun addr -> addr.Row) |> List.max
        let maxCol = allAddrs |> List.map (fun addr -> addr.Col) |> List.max

        // Build CSV row by row
        let rows =
            [0 .. maxRow]
            |> List.map (fun row ->
                [0 .. maxCol]
                |> List.map (fun col ->
                    let addr = { Row = row; Col = col }
                    match getCell addr model with
                    | Some cellData ->
                        let valueText = displayValueWithFormat cellData.Value cellData.Format
                        // Escape commas and quotes
                        if valueText.Contains(",") || valueText.Contains("\"") || valueText.Contains("\n") || valueText.Contains("\r") then
                            sprintf "\"%s\"" (valueText.Replace("\"", "\"\""))
                        else
                            valueText
                    | None -> ""
                )
                |> String.concat ","
            )

        String.concat "\n" rows

/// Check if a cell address is within a merged range
let isInMergedRange (addr: CellAddress) (merge: MergedCell): bool =
    addr.Row >= merge.StartRow && addr.Row <= merge.EndRow &&
    addr.Col >= merge.StartCol && addr.Col <= merge.EndCol

/// Merges by the rows they cover, for one list of merges (the list a model
/// holds; a change of the merges gives the model another list).
type private MergeIndex = {
    Source: MergedCell list
    /// Row -> the merges covering it, in list order.
    ByRow: System.Collections.Generic.Dictionary<int, MergedCell list>
    /// Merges too tall to list row by row (a merged column): looked at one by one.
    Tall: MergedCell list
}

let mutable private mergeIndexCache : MergeIndex list = []

let private buildMergeIndex (merges: MergedCell list) : MergeIndex =
    let byRow = System.Collections.Generic.Dictionary<int, MergedCell list>()
    let tall = ResizeArray<MergedCell>()
    // Back to front, so that each row's list ends up in list order.
    for merge in List.rev merges do
        if merge.EndRow - merge.StartRow > 4096 then tall.Insert(0, merge)
        else
            for row in merge.StartRow .. merge.EndRow do
                match byRow.TryGetValue row with
                | true, existing -> byRow.[row] <- merge :: existing
                | false, _ -> byRow.[row] <- [ merge ]
    { Source = merges; ByRow = byRow; Tall = List.ofSeq tall }

let private mergeIndexOf (merges: MergedCell list) : MergeIndex =
    let cached = mergeIndexCache
    match cached |> List.tryFind (fun index -> obj.ReferenceEquals(index.Source, merges)) with
    | Some index -> index
    | None ->
        let index = buildMergeIndex merges
        mergeIndexCache <- index :: List.truncate 7 cached
        index

/// Find merged cell containing this address (if any). The first merge of the
/// model's list that covers the cell, found through an index by row.
let findMergedCell (addr: CellAddress) (model: SpreadsheetModel): MergedCell option =
    match model.MergedCells with
    | [] -> None
    | merges ->
        let index = mergeIndexOf merges
        let inRow =
            match index.ByRow.TryGetValue addr.Row with
            | true, candidates -> candidates |> List.tryFind (fun merge -> addr.Col >= merge.StartCol && addr.Col <= merge.EndCol)
            | false, _ -> None
        match inRow, index.Tall with
        | found, [] -> found
        | None, tall -> tall |> List.tryFind (fun merge -> isInMergedRange addr merge)
        | Some found, tall ->
            // Overlapping merges (a broken file): the one that comes first in the list.
            match tall |> List.tryFind (fun merge -> isInMergedRange addr merge) with
            | Some other when (merges |> List.findIndex (fun m -> obj.ReferenceEquals(m, other))) < (merges |> List.findIndex (fun m -> obj.ReferenceEquals(m, found))) -> Some other
            | _ -> Some found

/// Get primary cell address (top-left) of a merge
let getPrimaryCell (merge: MergedCell): CellAddress =
    { Row = merge.StartRow; Col = merge.StartCol }

/// Check if a cell is the primary cell of a merge
let isPrimaryCell (addr: CellAddress) (merge: MergedCell): bool =
    addr.Row = merge.StartRow && addr.Col = merge.StartCol

/// Add a merged cell range
let addMergedCell (startRow: int) (startCol: int) (endRow: int) (endCol: int) (model: SpreadsheetModel): SpreadsheetModel =
    if startRow > endRow || startCol > endCol then
        model  // Invalid range
    else
        let merge = {
            StartRow = startRow
            StartCol = startCol
            EndRow = endRow
            EndCol = endCol
        }
        { model with MergedCells = merge :: model.MergedCells }

/// Remove merged cell containing this address
let removeMergedCell (addr: CellAddress) (model: SpreadsheetModel): SpreadsheetModel =
    let newMerges =
        model.MergedCells
        |> List.filter (fun merge -> not (isInMergedRange addr merge))
    { model with MergedCells = newMerges }

let private mergeIntersectsRange startRow endRowExclusive startCol endColExclusive (merge: MergedCell) =
    merge.EndRow >= startRow &&
    merge.StartRow < endRowExclusive &&
    merge.EndCol >= startCol &&
    merge.StartCol < endColExclusive

let mergedCellsIntersectingRange startRow endRowExclusive startCol endColExclusive (model: SpreadsheetModel) =
    model.MergedCells
    |> List.filter (mergeIntersectsRange startRow endRowExclusive startCol endColExclusive)
