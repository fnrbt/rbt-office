module RedBat.Spreadsheet.Model

open System

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private clearArray (items: 'T array) =
    for i = 0 to items.Length - 1 do
        items.[i] <- Unchecked.defaultof<'T>

let private referenceHash (value: obj) =
    System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value)

let private referenceSlot (value: obj) slotCount =
    referenceHash value &&& (slotCount - 1)

let private prependResizeArray (items: ResizeArray<'T>) tail =
    let mutable result = tail
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private arrayToList (items: 'T array) =
    let mutable result = []
    for i = items.Length - 1 downto 0 do
        result <- items.[i] :: result
    result

let private binarySearchInt (values: int array) target =
    let mutable lo = 0
    let mutable hi = values.Length - 1
    let mutable found = -1
    while found < 0 && lo <= hi do
        let mid = lo + (hi - lo) / 2
        let value = values.[mid]
        if value = target then
            found <- mid
        elif value < target then
            lo <- mid + 1
        else
            hi <- mid - 1
    found

let private appendItem item items =
    let values = ResizeArray(List.length items + 1)
    for value in items do
        values.Add value
    values.Add item
    resizeArrayToList values

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

let private cellValueIsFormula value =
    match value with
    | Formula _ -> true
    | _ -> false

let private cellDataIsFormula (cell: CellData) =
    cellValueIsFormula cell.Value

type private FormulaPresenceCacheEntry = {
    Cells: Map<CellAddress, CellData>
    CellCount: int
    SourceTextCharCount: int
    FormulaCount: int option
    HasFormulas: bool
}

let mutable private formulaPresenceCache: FormulaPresenceCacheEntry list = []
let private formulaPresenceCellsComparer =
    { new System.Collections.Generic.IEqualityComparer<Map<CellAddress, CellData>> with
        member _.Equals(left, right) = obj.ReferenceEquals(left, right)
        member _.GetHashCode(value) = referenceHash value }
let private formulaPresenceCacheByCells =
    System.Collections.Generic.Dictionary<Map<CellAddress, CellData>, FormulaPresenceCacheEntry>(formulaPresenceCellsComparer)
let private formulaPresenceFastSlots = Array.zeroCreate<FormulaPresenceCacheEntry> 256
let private formulaPresenceCacheMaxEntries = 64
let private formulaPresenceCacheMaxTotalCells = 500000
let private formulaPresenceCacheMaxTotalSourceTextChars = 5000000

let private formulaPresenceCacheEntryMatches cells (entry: FormulaPresenceCacheEntry) =
    obj.ReferenceEquals(entry.Cells, cells)

let private setFormulaPresenceCache entries =
    formulaPresenceCache <- entries
    formulaPresenceCacheByCells.Clear()
    clearArray formulaPresenceFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<FormulaPresenceCacheEntry>
        if not (formulaPresenceCacheByCells.TryGetValue(entry.Cells, &existing)) then
            formulaPresenceCacheByCells.[entry.Cells] <- entry
            let slot = referenceSlot entry.Cells formulaPresenceFastSlots.Length
            formulaPresenceFastSlots.[slot] <- entry

let private tryGetFormulaPresenceCache cells =
    let slot = referenceSlot cells formulaPresenceFastSlots.Length
    let fast = formulaPresenceFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && formulaPresenceCacheEntryMatches cells fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<FormulaPresenceCacheEntry>
        if formulaPresenceCacheByCells.TryGetValue(cells, &cached) then
            formulaPresenceFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private formulaPresenceCacheWithoutCells cells =
    let mutable cached = Unchecked.defaultof<FormulaPresenceCacheEntry>
    if formulaPresenceCacheByCells.TryGetValue(cells, &cached) then
        formulaPresenceCache
        |> List.filter (formulaPresenceCacheEntryMatches cells >> not)
    else
        formulaPresenceCache

let private removeFormulaPresenceCacheForCells cells =
    let entries = formulaPresenceCacheWithoutCells cells
    if not (obj.ReferenceEquals(entries, formulaPresenceCache)) then
        setFormulaPresenceCache entries

let private formulaPresenceCacheEntryRetainable (entry: FormulaPresenceCacheEntry) =
    entry.CellCount <= formulaPresenceCacheMaxTotalCells
    && entry.SourceTextCharCount <= formulaPresenceCacheMaxTotalSourceTextChars

let private sourceTextAddCapped limit value total =
    if total > limit then total
    else min (limit + 1) (total + max 0 value)

let private optionTextCharCountFor limit total (value: string option) =
    match value with
    | Some value -> sourceTextAddCapped limit value.Length total
    | None -> total

let private borderTextCharCountFor limit total (border: BorderStyle option) =
    match border with
    | Some border ->
        let total = sourceTextAddCapped limit border.Style.Length total
        optionTextCharCountFor limit total border.Color
    | None -> total

let cellValueTextCharCountFor maxTotalTextChars value =
    let limit = max 0 maxTotalTextChars
    let mutable total = 0
    let mutable current = value
    let mutable doneCounting = false
    while total <= limit && not doneCounting do
        match current with
        | CellValue.String value
        | CellValue.Error value ->
            total <- sourceTextAddCapped limit value.Length total
            doneCounting <- true
        | CellValue.Formula (formula, cachedValue) ->
            total <- sourceTextAddCapped limit formula.Length total
            current <- cachedValue
        | CellValue.Empty
        | CellValue.Number _
        | CellValue.Boolean _ ->
            doneCounting <- true
    min total (limit + 1)

let cellFormatTextCharCountFor maxTotalTextChars (format: CellFormat) =
    let limit = max 0 maxTotalTextChars
    let mutable total = 0
    if total <= limit then total <- optionTextCharCountFor limit total format.FontName
    if total <= limit then total <- optionTextCharCountFor limit total format.TextColor
    if total <= limit then total <- optionTextCharCountFor limit total format.BgColor
    if total <= limit then total <- optionTextCharCountFor limit total format.HorizontalAlign
    if total <= limit then total <- optionTextCharCountFor limit total format.VerticalAlign
    if total <= limit then total <- optionTextCharCountFor limit total format.NumberFormat
    if total <= limit then total <- borderTextCharCountFor limit total format.BorderTop
    if total <= limit then total <- borderTextCharCountFor limit total format.BorderBottom
    if total <= limit then total <- borderTextCharCountFor limit total format.BorderLeft
    if total <= limit then total <- borderTextCharCountFor limit total format.BorderRight
    if total <= limit then total <- borderTextCharCountFor limit total format.BorderDiagonal
    if total <= limit then
        match format.DataBar with
        | Some (_, color) -> total <- sourceTextAddCapped limit color.Length total
        | None -> ()
    min total (limit + 1)

let cellCommentTextCharCountFor maxTotalTextChars (comment: CellComment option) =
    let limit = max 0 maxTotalTextChars
    match comment with
    | Some comment ->
        let total = sourceTextAddCapped limit comment.Text.Length 0
        optionTextCharCountFor limit total comment.Author
    | None -> 0

let cellDataTextCharCountFor maxTotalTextChars cell =
    let limit = max 0 maxTotalTextChars
    let mutable total = cellValueTextCharCountFor limit cell.Value
    if total <= limit then total <- sourceTextAddCapped limit (cellFormatTextCharCountFor limit cell.Format) total
    if total <= limit then total <- sourceTextAddCapped limit (cellCommentTextCharCountFor limit cell.Comment) total
    min total (limit + 1)

let private formulaPresenceCellDataStringCharCount cell =
    cellDataTextCharCountFor formulaPresenceCacheMaxTotalSourceTextChars cell

let private cellMapSourceTextCharCount maxTotalSourceTextChars (cells: Map<CellAddress, CellData>) =
    let limit = maxTotalSourceTextChars + 1
    let mutable total = 0
    cells
    |> Map.exists (fun _ cell ->
        total <- sourceTextAddCapped maxTotalSourceTextChars (cellDataTextCharCountFor maxTotalSourceTextChars cell) total
        total > maxTotalSourceTextChars)
    |> ignore
    min total limit

let private cellMapSourceTextAndFormulaMetrics maxTotalSourceTextChars (cells: Map<CellAddress, CellData>) =
    let limit = maxTotalSourceTextChars + 1
    let mutable total = 0
    let mutable formulaCount = 0
    cells
    |> Map.exists (fun _ cell ->
        if cellDataIsFormula cell then
            formulaCount <- formulaCount + 1
        total <- sourceTextAddCapped maxTotalSourceTextChars (cellDataTextCharCountFor maxTotalSourceTextChars cell) total
        total > maxTotalSourceTextChars)
    |> ignore
    let sourceTextCharCount = min total limit
    if total <= maxTotalSourceTextChars then
        sourceTextCharCount, Some formulaCount, formulaCount > 0
    else
        sourceTextCharCount, None, formulaCount > 0

let private cellMapSourceTextCharCountFor maxTotalCells maxTotalSourceTextChars (cells: Map<CellAddress, CellData>) =
    if cells.Count > maxTotalCells then
        maxTotalSourceTextChars + 1
    else
        cellMapSourceTextCharCount maxTotalSourceTextChars cells

let private formulaPresenceSourceTextCharCount cells =
    cellMapSourceTextCharCountFor
        formulaPresenceCacheMaxTotalCells
        formulaPresenceCacheMaxTotalSourceTextChars
        cells

let private rememberCellsFormulaPresenceWithSourceTextCharCount cells formulaCount hasFormulas sourceTextCharCount =
    let entry = {
        Cells = cells
        CellCount = max 1 (Map.count cells)
        SourceTextCharCount = sourceTextCharCount
        FormulaCount = formulaCount
        HasFormulas = hasFormulas
    }
    if not (formulaPresenceCacheEntryRetainable entry) then
        removeFormulaPresenceCacheForCells cells
    else
        let trim entries =
            let rec loop keptCount keptCells keptSourceTextChars acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateCells = max 1 candidate.CellCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    if formulaPresenceCacheEntryRetainable candidate
                       && keptCount < formulaPresenceCacheMaxEntries
                       && keptCells + candidateCells <= formulaPresenceCacheMaxTotalCells
                       && keptSourceTextChars + candidateSourceTextChars <= formulaPresenceCacheMaxTotalSourceTextChars then
                        loop
                            (keptCount + 1)
                            (keptCells + candidateCells)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptCells keptSourceTextChars acc tail
            loop 0 0 0 [] entries
        (entry :: formulaPresenceCacheWithoutCells cells)
        |> trim
        |> setFormulaPresenceCache

let private rememberCellsFormulaPresence cells formulaCount hasFormulas =
    rememberCellsFormulaPresenceWithSourceTextCharCount cells formulaCount hasFormulas (formulaPresenceSourceTextCharCount cells)

let private adjustedFormulaPresenceSourceTextCharCount cachedSourceTextCharCount oldCell data =
    let limit = formulaPresenceCacheMaxTotalSourceTextChars + 1
    if cachedSourceTextCharCount >= limit then
        limit
    else
        let oldChars =
            match oldCell with
            | Some cell -> formulaPresenceCellDataStringCharCount cell
            | None -> 0
        let nextCount =
            cachedSourceTextCharCount
            - oldChars
            + cellDataTextCharCountFor formulaPresenceCacheMaxTotalSourceTextChars data
        max 0 (min limit nextCount)

let private adjustedFormulaPresenceSourceTextCharCountAfterRemove cachedSourceTextCharCount removedSourceTextCharCount =
    let limit = formulaPresenceCacheMaxTotalSourceTextChars + 1
    if cachedSourceTextCharCount >= limit then
        limit
    else
        max 0 (cachedSourceTextCharCount - removedSourceTextCharCount)

let private cachedCellsHaveFormulaCells cells =
    match tryGetFormulaPresenceCache cells with
    | ValueSome cached ->
        Some cached.HasFormulas
    | ValueNone ->
        None

let private cachedFormulaPresence cells =
    match tryGetFormulaPresenceCache cells with
    | ValueSome cached -> Some cached
    | ValueNone -> None

let rememberCellsHaveFormulaCells cells hasFormulas =
    let count = if hasFormulas then None else Some 0
    rememberCellsFormulaPresence cells count hasFormulas

let rememberCellsFormulaCount cells formulaCount =
    rememberCellsFormulaPresence cells (Some formulaCount) (formulaCount > 0)

let cellsFormulaCountIfKnown cells =
    match cachedFormulaPresence cells with
    | Some cached -> cached.FormulaCount
    | None -> None

let private scanCellsHaveFormulaCells (cells: Map<CellAddress, CellData>) =
    cells |> Map.exists (fun _ cell -> cellDataIsFormula cell)

let cellsHaveFormulaCells cells =
    match cachedCellsHaveFormulaCells cells with
    | Some hasFormulas -> hasFormulas
    | None ->
        let hasFormulas = scanCellsHaveFormulaCells cells
        rememberCellsHaveFormulaCells cells hasFormulas
        hasFormulas

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
}

/// Sort state for a column
type SortState = {
    Column: int
    Ascending: bool
}

/// Column filter (values to show - unchecked values are hidden)
type ColumnFilter = {
    Column: int
    VisibleValues: Set<string>  // String representation of values to show
}

/// Conditional formatting rule
type ConditionalFormatRule =
    | GreaterThan of float
    | LessThan of float
    | Between of float * float
    | Equal of float
    | ContainsText of string

/// Conditional format (rule + style to apply)
type ConditionalFormat = {
    Rule: ConditionalFormatRule
    Style: CellFormat
}

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

type DataValidation = {
    Rule: ValidationRule
    /// Shown when an entered value fails the rule.
    ErrorMessage: string option
    /// Shown as a tooltip while the cell is selected.
    InputMessage: string option
}

/// Merged cell range (primary cell is top-left)
type MergedCell = {
    StartRow: int
    StartCol: int
    EndRow: int
    EndCol: int
}

[<Struct>]
type private CellMergedReferenceKey = {
    Cells: Map<CellAddress, CellData>
    MergedCells: MergedCell list
}

[<Struct>]
type private CellRowCountReferenceKey = {
    Cells: Map<CellAddress, CellData>
    RowCount: int
}

[<Struct>]
type private CellGridReferenceKey = {
    Cells: Map<CellAddress, CellData>
    RowCount: int
    ColCount: int
}

let private cellMapReferenceComparer =
    { new System.Collections.Generic.IEqualityComparer<Map<CellAddress, CellData>> with
        member _.Equals(left, right) = obj.ReferenceEquals(left, right)
        member _.GetHashCode(value) = referenceHash value }

let private cellMergedReferenceHash cells mergedCells =
    (referenceHash cells * 397) ^^^ referenceHash mergedCells

let private cellMergedReferenceKey cells mergedCells : CellMergedReferenceKey =
    { Cells = cells
      MergedCells = mergedCells }

let private cellRowCountReferenceHash cells rowCount =
    (referenceHash cells * 397) ^^^ rowCount

let private cellRowCountReferenceKey cells rowCount : CellRowCountReferenceKey =
    { Cells = cells
      RowCount = rowCount }

let private cellGridReferenceHash cells rowCount colCount =
    ((referenceHash cells * 397) ^^^ rowCount) * 397 ^^^ colCount

let private cellGridReferenceKey cells rowCount colCount : CellGridReferenceKey =
    { Cells = cells
      RowCount = rowCount
      ColCount = colCount }

let private cellMergedReferenceComparer =
    { new System.Collections.Generic.IEqualityComparer<CellMergedReferenceKey> with
        member _.Equals(left, right) =
            obj.ReferenceEquals(left.Cells, right.Cells)
            && obj.ReferenceEquals(left.MergedCells, right.MergedCells)

        member _.GetHashCode(key) =
            cellMergedReferenceHash key.Cells key.MergedCells }

let private cellRowCountReferenceComparer =
    { new System.Collections.Generic.IEqualityComparer<CellRowCountReferenceKey> with
        member _.Equals(left, right) =
            obj.ReferenceEquals(left.Cells, right.Cells)
            && left.RowCount = right.RowCount

        member _.GetHashCode(key) =
            cellRowCountReferenceHash key.Cells key.RowCount }

let private cellGridReferenceComparer =
    { new System.Collections.Generic.IEqualityComparer<CellGridReferenceKey> with
        member _.Equals(left, right) =
            obj.ReferenceEquals(left.Cells, right.Cells)
            && left.RowCount = right.RowCount
            && left.ColCount = right.ColCount

        member _.GetHashCode(key) =
            cellGridReferenceHash key.Cells key.RowCount key.ColCount }

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
    | 5 -> 612.0, 1008.0     // Legal    (8.5 × 14 in)
    | 8 -> 841.68, 1190.55   // A3       (297 × 420 mm)
    | 9 -> 595.27, 841.89    // A4       (210 × 297 mm)
    | 11 -> 419.53, 595.27   // A5       (148 × 210 mm)
    | 13 -> 728.5, 1031.81   // B5 ISO   (176 × 250 mm)
    | 70 -> 595.27, 841.89   // A4 alt
    | _ -> 612.0, 792.0      // unknown → Letter

/// Spreadsheet model (forward declaration for Command type)
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

    /// Active filters by column
    Filters: Map<int, ColumnFilter>

    /// Rows hidden by active AutoFilters, kept separate from user row heights.
    FilterHiddenRows: int array

    /// Custom row heights (default 25.0)
    RowHeights: Map<int, float>

    /// Custom column widths (default 100.0)
    ColumnWidths: Map<int, float>

    /// Current search query
    SearchQuery: string option

    /// Case-sensitivity used for the current search query
    SearchCaseSensitive: bool option

    /// Search results (cell addresses)
    SearchResults: CellAddress array

    /// Current search result index
    SearchIndex: int

    /// Conditional formats by cell address
    ConditionalFormats: Map<CellAddress, ConditionalFormat list>

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

    /// Dynamic-array spill ownership: each spilled (ghost) cell mapped to its
    /// anchor formula cell. Maintained by the recalc+spill pass so stale
    /// spilled cells are cleared before a re-spill. Empty for non-spilling
    /// workbooks.
    Spill: Map<CellAddress, CellAddress>

    /// Data-validation rules keyed by cell.
    Validations: Map<CellAddress, DataValidation>

    /// VBA/VBScript module source attached to the workbook (the "VBA project").
    /// Empty when there are no macros. Event subs (Worksheet_Change /
    /// Workbook_Open) and user-run macros live here; `Vba.runOnModel` executes
    /// it. Persisted/round-tripped separately from cell data.
    Macros: string
}

/// Command for undo/redo (must come after SpreadsheetModel)
and Command = {
    Undo: SpreadsheetModel -> SpreadsheetModel
    Redo: SpreadsheetModel -> SpreadsheetModel
    Description: string
    RetainedItemCount: int
    RetainedTextCharCount: int
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

let private historyMaxTotalRetainedItems = 500000
let private historyMaxTotalRetainedTextChars = 20000000

let private addCapped limit value total =
    if total > limit then total
    else min (limit + 1) (total + max 0 value)

let commandRetainedTextLength text =
    min (historyMaxTotalRetainedTextChars + 1) (max 0 (String.length text))

let commandRetainedCellOptionItemCount = function
    | Some _ -> 1
    | None -> 0

let commandRetainedCellDataTextCharCount (cell: CellData) =
    cellDataTextCharCountFor historyMaxTotalRetainedTextChars cell

let commandRetainedCellOptionTextCharCount = function
    | Some cell -> commandRetainedCellDataTextCharCount cell
    | None -> 0

let commandRetainedCellMapItemCount (cells: Map<CellAddress, CellData>) =
    max 1 cells.Count

let commandRetainedCellMapTextCharCount (cells: Map<CellAddress, CellData>) =
    if cells.Count > historyMaxTotalRetainedItems then
        historyMaxTotalRetainedTextChars + 1
    else
        match cachedFormulaPresence cells with
        | Some cached -> min (historyMaxTotalRetainedTextChars + 1) cached.SourceTextCharCount
        | None ->
            let sourceTextCharCount, formulaCount, hasFormulas =
                cellMapSourceTextAndFormulaMetrics historyMaxTotalRetainedTextChars cells
            rememberCellsFormulaPresenceWithSourceTextCharCount cells formulaCount hasFormulas sourceTextCharCount
            sourceTextCharCount

let commandRetainedClipboardItemCount (clipboard: ClipboardData) =
    let mutable total = 0
    let mutable remaining = clipboard.Cells
    while total <= historyMaxTotalRetainedItems && not remaining.IsEmpty do
        total <- addCapped historyMaxTotalRetainedItems 1 total
        remaining <- remaining.Tail
    max 1 total

let commandRetainedClipboardTextCharCount (clipboard: ClipboardData) =
    let mutable total = 0
    let mutable remaining = clipboard.Cells
    while total <= historyMaxTotalRetainedTextChars && not remaining.IsEmpty do
        let _, _, cell = remaining.Head
        total <- addCapped historyMaxTotalRetainedTextChars (commandRetainedCellDataTextCharCount cell) total
        remaining <- remaining.Tail
    total

let commandRetainedCellOptionPairsItemCount (items: seq<CellAddress * CellData option>) =
    let mutable total = 0
    let mutable seen = 0
    use enumerator = items.GetEnumerator()
    while total <= historyMaxTotalRetainedItems
          && seen <= historyMaxTotalRetainedItems
          && enumerator.MoveNext() do
        let _, cell = enumerator.Current
        total <- total + commandRetainedCellOptionItemCount cell
        seen <- seen + 1
    max 1 total

let commandRetainedCellOptionPairsTextCharCount (items: seq<CellAddress * CellData option>) =
    let mutable total = 0
    let mutable seen = 0
    use enumerator = items.GetEnumerator()
    while total <= historyMaxTotalRetainedTextChars
          && seen <= historyMaxTotalRetainedItems
          && enumerator.MoveNext() do
        let _, cell = enumerator.Current
        total <- addCapped historyMaxTotalRetainedTextChars (commandRetainedCellOptionTextCharCount cell) total
        seen <- seen + 1
    total

let commandRetainedSumTextCharCounts (counts: int list) =
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = counts
    while total <= historyMaxTotalRetainedTextChars
          && seen <= historyMaxTotalRetainedItems
          && not remaining.IsEmpty do
        total <- addCapped historyMaxTotalRetainedTextChars remaining.Head total
        seen <- seen + 1
        remaining <- remaining.Tail
    total

let commandRetainedTextCharCountForItemCount retainedItemCount getTextCharCount =
    if retainedItemCount > historyMaxTotalRetainedItems then
        historyMaxTotalRetainedTextChars + 1
    else
        getTextCharCount ()

let commandRetainedTextCharCountForItemCountWithInitial retainedItemCount initialTextCharCount getTextCharCount =
    if retainedItemCount > historyMaxTotalRetainedItems
       || initialTextCharCount > historyMaxTotalRetainedTextChars then
        historyMaxTotalRetainedTextChars + 1
    else
        addCapped historyMaxTotalRetainedTextChars (getTextCharCount ()) initialTextCharCount

let commandRetainedSumItemCounts (counts: int list) =
    let mutable total = 0
    let mutable remaining = counts
    while total <= historyMaxTotalRetainedItems && not remaining.IsEmpty do
        total <- addCapped historyMaxTotalRetainedItems remaining.Head total
        remaining <- remaining.Tail
    max 1 total

let private trimHistoryStack maxSize stack =
    let rec loop remaining kept count totalItems totalTextChars =
        match remaining with
        | [] -> List.rev kept
        | command :: rest ->
            let retainedItems = max 1 command.RetainedItemCount
            let retainedTextChars = max 1 command.RetainedTextCharCount
            let nextTotalItems = totalItems + retainedItems
            let nextTotalTextChars = totalTextChars + retainedTextChars
            if count < maxSize
               && retainedItems <= historyMaxTotalRetainedItems
               && retainedTextChars <= historyMaxTotalRetainedTextChars
               && nextTotalItems <= historyMaxTotalRetainedItems
               && nextTotalTextChars <= historyMaxTotalRetainedTextChars then
                loop rest (command :: kept) (count + 1) nextTotalItems nextTotalTextChars
            else
                loop rest kept count totalItems totalTextChars
    loop stack [] 0 0 0

/// Create empty spreadsheet
let createEmpty rowCount colCount =
    let model = {
        Cells = Map.empty
        Selection = SingleCell(0, 0)
        EditingCell = None
        EditBuffer = ""
        RowCount = rowCount
        ColCount = colCount
        Clipboard = None
        History = emptyHistory
        SortState = None
        Filters = Map.empty
        FilterHiddenRows = [||]
        RowHeights = Map.empty
        ColumnWidths = Map.empty
        SearchQuery = None
        SearchCaseSensitive = None
        SearchResults = [||]
        SearchIndex = 0
        ConditionalFormats = Map.empty
        MergedCells = []
        RichRuns = Map.empty
        FrozenRows = 0
        FrozenCols = 0
        PageSetup = defaultPageSetup
        Tables = []
        CfBoldCells = Set.empty
        ShowGridLines = true
        Spill = Map.empty
        Validations = Map.empty
        Macros = ""
    }
    rememberCellsHaveFormulaCells model.Cells false
    model

type private PasteIdentityCache = {
    Cells: Map<CellAddress, CellData>
    CellCount: int
    SourceTextCharCount: int
    RowCount: int
    ColCount: int
    Entries: System.Collections.Generic.Dictionary<PasteIdentityCacheKey, unit>
    Order: System.Collections.Generic.Queue<PasteIdentityCacheKey>
    mutable TotalClipboardCellCount: int
    mutable TotalClipboardTextCharCount: int
}

and private PasteIdentityCacheKey = {
    Target: CellAddress
    Clipboard: ClipboardData
    ClipboardCellCount: int
    ClipboardTextCharCount: int
}

let mutable private pasteIdentityCache: PasteIdentityCache list = []
let private pasteIdentityCacheBySource =
    System.Collections.Generic.Dictionary<CellGridReferenceKey, PasteIdentityCache>(cellGridReferenceComparer)
let private pasteIdentityFastSlots = Array.zeroCreate<PasteIdentityCache> 256
let private pasteIdentityCacheMaxEntries = 64
let private pasteIdentityCacheMaxSources = 64
let private pasteIdentityCacheMaxTotalCells = 200000
let private pasteIdentityCacheMaxTotalClipboardCells = 200000
let private pasteIdentityCacheMaxTotalSourceTextChars = 5000000
let private pasteIdentityCacheMaxTotalClipboardTextChars = 5000000

let private clipboardDataRetainedMetrics maxTotalCells maxTotalSourceTextChars (clipboardData: ClipboardData) =
    let textLimit = maxTotalSourceTextChars + 1
    let mutable cellCount = 0
    let mutable textCharCount = 0
    let mutable remaining = clipboardData.Cells
    while cellCount <= maxTotalCells
          && textCharCount <= maxTotalSourceTextChars
          && not remaining.IsEmpty do
        let _, _, cell = remaining.Head
        cellCount <- addCapped maxTotalCells 1 cellCount
        textCharCount <- addCapped maxTotalSourceTextChars (cellDataTextCharCountFor maxTotalSourceTextChars cell) textCharCount
        remaining <- remaining.Tail
    max 1 cellCount, min textCharCount textLimit

let private pasteIdentityCacheMatches cells rowCount colCount (cached: PasteIdentityCache) =
    obj.ReferenceEquals(cached.Cells, cells)
    && cached.RowCount = rowCount
    && cached.ColCount = colCount

let private setPasteIdentityCache entries =
    pasteIdentityCache <- entries
    pasteIdentityCacheBySource.Clear()
    clearArray pasteIdentityFastSlots
    for entry in entries do
        let key = cellGridReferenceKey entry.Cells entry.RowCount entry.ColCount
        let mutable existing = Unchecked.defaultof<PasteIdentityCache>
        if not (pasteIdentityCacheBySource.TryGetValue(key, &existing)) then
            pasteIdentityCacheBySource.[key] <- entry
            let slot = cellGridReferenceHash entry.Cells entry.RowCount entry.ColCount &&& (pasteIdentityFastSlots.Length - 1)
            pasteIdentityFastSlots.[slot] <- entry

let private tryGetPasteIdentityCache cells rowCount colCount =
    let slot = cellGridReferenceHash cells rowCount colCount &&& (pasteIdentityFastSlots.Length - 1)
    let fast = pasteIdentityFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && pasteIdentityCacheMatches cells rowCount colCount fast then
        ValueSome fast
    else
        let key = cellGridReferenceKey cells rowCount colCount
        let mutable cached = Unchecked.defaultof<PasteIdentityCache>
        if pasteIdentityCacheBySource.TryGetValue(key, &cached) then
            pasteIdentityFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private pasteIdentityCacheRetainable (cached: PasteIdentityCache) =
    cached.Entries.Count > 0
    && cached.Entries.Count <= pasteIdentityCacheMaxEntries
    && cached.CellCount <= pasteIdentityCacheMaxTotalCells
    && cached.SourceTextCharCount <= pasteIdentityCacheMaxTotalSourceTextChars
    && cached.TotalClipboardCellCount <= pasteIdentityCacheMaxTotalClipboardCells
    && cached.TotalClipboardTextCharCount <= pasteIdentityCacheMaxTotalClipboardTextChars

let private rememberPasteIdentityCache cached =
    let trim (entries: PasteIdentityCache list) =
        let rec loop keptCount keptCells keptSourceTextChars keptClipboardCells keptClipboardTextChars acc remaining =
            match remaining with
            | [] -> List.rev acc
            | candidate :: tail ->
                let candidateCells = max 1 candidate.CellCount
                let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                let candidateClipboardCells = max 1 candidate.TotalClipboardCellCount
                let candidateClipboardTextChars = max 1 candidate.TotalClipboardTextCharCount
                if pasteIdentityCacheRetainable candidate
                   && keptCount < pasteIdentityCacheMaxSources
                   && keptCells + candidateCells <= pasteIdentityCacheMaxTotalCells
                   && keptSourceTextChars + candidateSourceTextChars <= pasteIdentityCacheMaxTotalSourceTextChars
                   && keptClipboardCells + candidateClipboardCells <= pasteIdentityCacheMaxTotalClipboardCells
                   && keptClipboardTextChars + candidateClipboardTextChars <= pasteIdentityCacheMaxTotalClipboardTextChars then
                    loop
                        (keptCount + 1)
                        (keptCells + candidateCells)
                        (keptSourceTextChars + candidateSourceTextChars)
                        (keptClipboardCells + candidateClipboardCells)
                        (keptClipboardTextChars + candidateClipboardTextChars)
                        (candidate :: acc)
                        tail
                else
                    loop keptCount keptCells keptSourceTextChars keptClipboardCells keptClipboardTextChars acc tail
        loop 0 0 0 0 0 [] entries
    (cached
     :: (pasteIdentityCache
         |> List.filter (pasteIdentityCacheMatches cached.Cells cached.RowCount cached.ColCount >> not)))
    |> trim
    |> setPasteIdentityCache

let private pasteIdentityKeyComparer =
    { new System.Collections.Generic.IEqualityComparer<PasteIdentityCacheKey> with
        member _.Equals(left, right) =
            left.Target = right.Target
                && obj.ReferenceEquals(left.Clipboard, right.Clipboard)

        member _.GetHashCode(key) =
            (hash key.Target * 397)
                ^^^ System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.Clipboard) }

let private removePasteIdentityCacheEntry key (entries: System.Collections.Generic.Dictionary<PasteIdentityCacheKey, unit>) (cache: PasteIdentityCache) =
    if entries.Remove key then
        cache.TotalClipboardCellCount <- max 0 (cache.TotalClipboardCellCount - key.ClipboardCellCount)
        cache.TotalClipboardTextCharCount <-
            max 0 (cache.TotalClipboardTextCharCount - key.ClipboardTextCharCount)

let private addPasteIdentityCacheEntry key (cache: PasteIdentityCache) =
    if key.ClipboardCellCount > pasteIdentityCacheMaxTotalClipboardCells
       || key.ClipboardTextCharCount > pasteIdentityCacheMaxTotalClipboardTextChars then
        removePasteIdentityCacheEntry key cache.Entries cache
    else
        if not (cache.Entries.ContainsKey key) then
            cache.TotalClipboardCellCount <- cache.TotalClipboardCellCount + key.ClipboardCellCount
            cache.TotalClipboardTextCharCount <-
                cache.TotalClipboardTextCharCount + key.ClipboardTextCharCount
            cache.Order.Enqueue key
        cache.Entries.[key] <- ()
        while (cache.Entries.Count > pasteIdentityCacheMaxEntries
               || cache.TotalClipboardCellCount > pasteIdentityCacheMaxTotalClipboardCells
               || cache.TotalClipboardTextCharCount > pasteIdentityCacheMaxTotalClipboardTextChars)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if not (pasteIdentityKeyComparer.Equals(oldest, key)) then
                removePasteIdentityCacheEntry oldest cache.Entries cache
        if cache.Entries.Count > pasteIdentityCacheMaxEntries
           || cache.TotalClipboardCellCount > pasteIdentityCacheMaxTotalClipboardCells
           || cache.TotalClipboardTextCharCount > pasteIdentityCacheMaxTotalClipboardTextChars then
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue key
            cache.Entries.[key] <- ()
            cache.TotalClipboardCellCount <- key.ClipboardCellCount
            cache.TotalClipboardTextCharCount <- key.ClipboardTextCharCount

let private pasteIdentityKeyWithCounts
    (targetAddr: CellAddress)
    (clipboardData: ClipboardData)
    clipboardCellCount
    clipboardTextCharCount
    =
    {
        Target = targetAddr
        Clipboard = clipboardData
        ClipboardCellCount = clipboardCellCount
        ClipboardTextCharCount = clipboardTextCharCount
    }

let private pasteIdentityLookupKey (targetAddr: CellAddress) (clipboardData: ClipboardData) =
    pasteIdentityKeyWithCounts targetAddr clipboardData 0 0

let private tryPasteIdentityCache targetAddr clipboardData (model: SpreadsheetModel) =
    match tryGetPasteIdentityCache model.Cells model.RowCount model.ColCount with
    | ValueSome cached when pasteIdentityCacheRetainable cached ->
        let key = pasteIdentityLookupKey targetAddr clipboardData
        if cached.Entries.ContainsKey key then Some model else None
    | ValueSome cached ->
        rememberPasteIdentityCache cached
        None
    | ValueNone -> None

let private rememberPasteIdentityWithSourceTextCharCount
    sourceTextCharCount
    clipboardCellCount
    clipboardTextCharCount
    (targetAddr: CellAddress)
    (clipboardData: ClipboardData)
    (model: SpreadsheetModel)
    =
    let key = pasteIdentityKeyWithCounts targetAddr clipboardData clipboardCellCount clipboardTextCharCount
    let cache =
        match tryGetPasteIdentityCache model.Cells model.RowCount model.ColCount with
        | ValueSome cached ->
            cached
        | ValueNone ->
            {
                Cells = model.Cells
                CellCount = max 1 model.Cells.Count
                SourceTextCharCount = sourceTextCharCount
                RowCount = model.RowCount
                ColCount = model.ColCount
                Entries = System.Collections.Generic.Dictionary<PasteIdentityCacheKey, unit>(pasteIdentityKeyComparer)
                Order = System.Collections.Generic.Queue<PasteIdentityCacheKey>()
                TotalClipboardCellCount = 0
                TotalClipboardTextCharCount = 0
            }
    addPasteIdentityCacheEntry key cache
    rememberPasteIdentityCache cache

let private rememberPasteIdentityIfRetainable (targetAddr: CellAddress) (clipboardData: ClipboardData) (model: SpreadsheetModel) =
    let clipboardCellCount, clipboardTextCharCount =
        clipboardDataRetainedMetrics
            pasteIdentityCacheMaxTotalClipboardCells
            pasteIdentityCacheMaxTotalClipboardTextChars
            clipboardData
    if model.Cells.Count <= pasteIdentityCacheMaxTotalCells
       && clipboardCellCount <= pasteIdentityCacheMaxTotalClipboardCells then
        let sourceTextCharCount =
            cellMapSourceTextCharCountFor
                pasteIdentityCacheMaxTotalCells
                pasteIdentityCacheMaxTotalSourceTextChars
                model.Cells
        if sourceTextCharCount <= pasteIdentityCacheMaxTotalSourceTextChars then
            if clipboardTextCharCount <= pasteIdentityCacheMaxTotalClipboardTextChars then
                rememberPasteIdentityWithSourceTextCharCount
                    sourceTextCharCount
                    clipboardCellCount
                    clipboardTextCharCount
                    targetAddr
                    clipboardData
                    model

let private rememberPasteIdentity (targetAddr: CellAddress) (clipboardData: ClipboardData) (model: SpreadsheetModel) =
    rememberPasteIdentityIfRetainable targetAddr clipboardData model

let private rememberPasteIdentityForChangedResult
    (targetAddr: CellAddress)
    (clipboardData: ClipboardData)
    (model: SpreadsheetModel)
    =
    rememberPasteIdentityIfRetainable targetAddr clipboardData model

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
    for c in letter do
        let upper = Char.ToUpperInvariant c
        result <- result * 26 + (int upper - int 'A' + 1)
    result - 1

let tryParseNumber (text: string): float option =
    match Double.TryParse(text, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
    | true, value -> Some value
    | false, _ -> None

/// Parse cell reference like "A1" to (row, col)
let parseCellRef (ref: string): (int * int) option =
    let mutable col = 0
    let mutable row = 0
    let mutable hasCol = false
    let mutable hasRow = false
    let mutable inRow = false
    let mutable valid = true
    let mutable index = 0
    while valid && index < ref.Length do
        let c = ref.[index]
        if Char.IsDigit c then
            inRow <- true
            hasRow <- true
            let digit = int c - int '0'
            if row > (Int32.MaxValue - digit) / 10 then
                valid <- false
            else
                row <- row * 10 + digit
        elif Char.IsLetter c && not inRow then
            hasCol <- true
            let upper = Char.ToUpperInvariant c
            let digit = int upper - int 'A' + 1
            if col > (Int32.MaxValue - digit) / 26 then
                valid <- false
            else
                col <- col * 26 + digit
        else
            valid <- false
        index <- index + 1
    if valid && hasCol && hasRow && row > 0 && col > 0 then
        Some (row - 1, col - 1)
    else
        None

/// Format cell reference like (0, 0) to "A1"
let formatCellRef (row: int) (col: int): string =
    sprintf "%s%d" (columnToLetter col) (row + 1)

/// Get cell data at address
let getCell (addr: CellAddress) (model: SpreadsheetModel): CellData option =
    Map.tryFind addr model.Cells

/// Set cell data at address
let setCell (addr: CellAddress) (data: CellData) (model: SpreadsheetModel): SpreadsheetModel =
    let mutable oldCell = Unchecked.defaultof<CellData>
    let hasOldCell = model.Cells.TryGetValue(addr, &oldCell)
    if hasOldCell && oldCell = data then
        model
    else
        let cells = Map.add addr data model.Cells
        let dataIsFormula = cellDataIsFormula data
        let oldCellOption = if hasOldCell then Some oldCell else None
        let oldIsFormula =
            if hasOldCell then cellDataIsFormula oldCell
            else false
        match cachedFormulaPresence model.Cells with
        | Some cached ->
            let sourceTextCharCount = adjustedFormulaPresenceSourceTextCharCount cached.SourceTextCharCount oldCellOption data
            match cached.FormulaCount with
            | Some count ->
                let nextCount = max 0 (count + (if dataIsFormula && not oldIsFormula then 1 else 0) - (if oldIsFormula && not dataIsFormula then 1 else 0))
                rememberCellsFormulaPresenceWithSourceTextCharCount cells (Some nextCount) (nextCount > 0) sourceTextCharCount
            | None when not cached.HasFormulas ->
                let nextCount = if dataIsFormula then 1 else 0
                rememberCellsFormulaPresenceWithSourceTextCharCount cells (Some nextCount) (nextCount > 0) sourceTextCharCount
            | None ->
                let hasFormulas =
                    if oldIsFormula && not dataIsFormula then
                        scanCellsHaveFormulaCells cells
                    else
                        true
                if oldIsFormula && not dataIsFormula then
                    rememberCellsFormulaPresenceWithSourceTextCharCount cells (if hasFormulas then None else Some 0) hasFormulas sourceTextCharCount
                else
                    rememberCellsFormulaPresenceWithSourceTextCharCount cells None hasFormulas sourceTextCharCount
        | None when dataIsFormula ->
            rememberCellsHaveFormulaCells cells true
        | None ->
            ()
        { model with Cells = cells }

let setCellFormat (addr: CellAddress) (update: CellFormat -> CellFormat) (model: SpreadsheetModel): SpreadsheetModel =
    let mutable cellData = Unchecked.defaultof<CellData>
    if model.Cells.TryGetValue(addr, &cellData) then
        let newFormat = update cellData.Format
        if newFormat = cellData.Format then
            model
        else
            setCell addr { cellData with Format = newFormat } model
    else
        let newFormat = update defaultFormat
        if newFormat = defaultFormat then
            model
        else
            setCell addr { Value = Empty; Format = newFormat; Comment = None } model

let setCellTextColor (addr: CellAddress) color (model: SpreadsheetModel): SpreadsheetModel =
    let mutable cellData = Unchecked.defaultof<CellData>
    if model.Cells.TryGetValue(addr, &cellData) then
        match cellData.Format.TextColor with
        | Some existingColor when existingColor = color ->
            model
        | _ ->
            setCell addr { cellData with Format = { cellData.Format with TextColor = Some color } } model
    else
        setCell addr { Value = Empty; Format = { defaultFormat with TextColor = Some color }; Comment = None } model

let setCellBgColor (addr: CellAddress) color (model: SpreadsheetModel): SpreadsheetModel =
    let mutable cellData = Unchecked.defaultof<CellData>
    if model.Cells.TryGetValue(addr, &cellData) then
        match cellData.Format.BgColor with
        | Some existingColor when existingColor = color ->
            model
        | _ ->
            setCell addr { cellData with Format = { cellData.Format with BgColor = Some color } } model
    else
        setCell addr { Value = Empty; Format = { defaultFormat with BgColor = Some color }; Comment = None } model

let setCellNumberFormat (addr: CellAddress) numberFormat (model: SpreadsheetModel): SpreadsheetModel =
    let mutable cellData = Unchecked.defaultof<CellData>
    if model.Cells.TryGetValue(addr, &cellData) then
        match cellData.Format.NumberFormat with
        | Some existingFormat when existingFormat = numberFormat ->
            model
        | _ ->
            setCell addr { cellData with Format = { cellData.Format with NumberFormat = Some numberFormat } } model
    else
        setCell addr { Value = Empty; Format = { defaultFormat with NumberFormat = Some numberFormat }; Comment = None } model

let setCellHorizontalAlign (addr: CellAddress) align (model: SpreadsheetModel): SpreadsheetModel =
    let mutable cellData = Unchecked.defaultof<CellData>
    if model.Cells.TryGetValue(addr, &cellData) then
        match cellData.Format.HorizontalAlign with
        | Some existingAlign when existingAlign = align ->
            model
        | _ ->
            setCell addr { cellData with Format = { cellData.Format with HorizontalAlign = Some align } } model
    else
        setCell addr { Value = Empty; Format = { defaultFormat with HorizontalAlign = Some align }; Comment = None } model

let private borderMatches (style: string) (color: string) (border: BorderStyle option) =
    match border with
    | Some existing ->
        existing.Style = style
            && match existing.Color with
               | Some existingColor -> existingColor = color
               | None -> false
    | None -> false

let private hasRequestedBorder (side: string) (style: string) (color: string) (format: CellFormat) =
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

let private applyBorderToFormat (side: string) (border: BorderStyle) (format: CellFormat) =
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

let setCellBorder (addr: CellAddress) (side: string) (style: string) (color: string) (model: SpreadsheetModel): SpreadsheetModel =
    let mutable cellData = Unchecked.defaultof<CellData>
    if model.Cells.TryGetValue(addr, &cellData) then
        if hasRequestedBorder side style color cellData.Format then
            model
        else
            let border = { Style = style; Color = Some color }
            setCell addr { cellData with Format = applyBorderToFormat side border cellData.Format } model
    elif hasRequestedBorder side style color defaultFormat then
        model
    else
        let border = { Style = style; Color = Some color }
        setCell addr { Value = Empty; Format = applyBorderToFormat side border defaultFormat; Comment = None } model

/// Clear cell at address
let clearCell (addr: CellAddress) (model: SpreadsheetModel): SpreadsheetModel =
    let mutable oldCell = Unchecked.defaultof<CellData>
    if not (model.Cells.TryGetValue(addr, &oldCell)) then
        model
    else
        let cells = Map.remove addr model.Cells
        let oldIsFormula = cellDataIsFormula oldCell
        match cachedFormulaPresence model.Cells with
        | Some cached ->
            let sourceTextCharCount =
                adjustedFormulaPresenceSourceTextCharCountAfterRemove
                    cached.SourceTextCharCount
                    (formulaPresenceCellDataStringCharCount oldCell)
            match cached.FormulaCount with
            | Some count ->
                let nextCount = max 0 (count - (if oldIsFormula then 1 else 0))
                rememberCellsFormulaPresenceWithSourceTextCharCount cells (Some nextCount) (nextCount > 0) sourceTextCharCount
            | None when not cached.HasFormulas ->
                rememberCellsFormulaPresenceWithSourceTextCharCount cells (Some 0) false sourceTextCharCount
            | None ->
                let hasFormulas =
                    if oldIsFormula then
                        scanCellsHaveFormulaCells cells
                    else
                        true
                rememberCellsFormulaPresenceWithSourceTextCharCount cells (if hasFormulas then None else Some 0) hasFormulas sourceTextCharCount
        | None ->
            ()
        { model with Cells = cells }

let private rememberClearedCellsFormulaPresence oldCells newCells removedFormulaCount removedSourceTextCharCount =
    match cachedFormulaPresence oldCells with
    | Some cached ->
        let sourceTextCharCount =
            adjustedFormulaPresenceSourceTextCharCountAfterRemove cached.SourceTextCharCount removedSourceTextCharCount
        match cached.FormulaCount with
        | Some count ->
            let nextCount = max 0 (count - removedFormulaCount)
            rememberCellsFormulaPresenceWithSourceTextCharCount newCells (Some nextCount) (nextCount > 0) sourceTextCharCount
        | None when not cached.HasFormulas ->
            rememberCellsFormulaPresenceWithSourceTextCharCount newCells (Some 0) false sourceTextCharCount
        | None ->
            let hasFormulas =
                if removedFormulaCount > 0 then
                    scanCellsHaveFormulaCells newCells
                else
                    true
            rememberCellsFormulaPresenceWithSourceTextCharCount newCells (if hasFormulas then None else Some 0) hasFormulas sourceTextCharCount
    | None ->
        ()

/// Clear all cells in a selection, preserving identity when nothing changes.
let clearSelectionCells (selection: Selection) (model: SpreadsheetModel): SpreadsheetModel =
    match selection with
    | SingleCell (row, col) ->
        clearCell { Row = row; Col = col } model
    | Range (startRow, startCol, endRow, endCol) ->
        if Map.isEmpty model.Cells then
            model
        else
            let rowCount = max 0 (endRow - startRow + 1)
            let colCount = max 0 (endCol - startCol + 1)
            let area = int64 rowCount * int64 colCount
            let mutable cells = model.Cells
            let mutable changed = false
            let mutable removedFormulaCount = 0
            let mutable removedSourceTextCharCount = 0
            let mutable clearedAllCells = false
            let formulaPresence = cachedFormulaPresence model.Cells
            let trackRemovedFormulaPresence = formulaPresence.IsSome
            let cachedNoFormulas =
                match formulaPresence with
                | Some cached -> not cached.HasFormulas
                | None -> false
            let mutable keepNoFormulaPresence = false
            let inRange addr =
                addr.Row >= startRow && addr.Row <= endRow && addr.Col >= startCol && addr.Col <= endCol
            let recordRemoved cellData =
                if trackRemovedFormulaPresence then
                    removedSourceTextCharCount <-
                        sourceTextAddCapped
                            formulaPresenceCacheMaxTotalSourceTextChars
                            (formulaPresenceCellDataStringCharCount cellData)
                            removedSourceTextCharCount
                    if cellDataIsFormula cellData then
                        removedFormulaCount <- removedFormulaCount + 1
            let removeCell addr cellData =
                cells <- Map.remove addr cells
                changed <- true
                recordRemoved cellData
            let rebuildWithKeptCells (keptCount: int) =
                let kept = ResizeArray<CellAddress * CellData>(keptCount)
                model.Cells
                |> Map.iter (fun addr cellData ->
                    if not (inRange addr) then
                        kept.Add(addr, cellData))
                Map.ofSeq kept
            if (not trackRemovedFormulaPresence || cachedNoFormulas)
               && area > 4096L
               && area >= (int64 model.Cells.Count + 3L) / 4L then
                let mutable removedCount = 0
                let mutable keptCount = 0
                model.Cells
                |> Map.iter (fun addr _ ->
                    if inRange addr then
                        removedCount <- removedCount + 1
                    else
                        keptCount <- keptCount + 1)
                if removedCount > 0 then
                    changed <- true
                    keepNoFormulaPresence <- cachedNoFormulas
                    if keptCount = 0 then
                        cells <- Map.empty
                        clearedAllCells <- true
                    elif int64 removedCount * 3L >= int64 model.Cells.Count * 2L then
                        cells <- model.Cells |> Map.filter (fun addr _ -> not (inRange addr))
                    else
                        model.Cells
                        |> Map.iter (fun addr _ ->
                            if inRange addr then
                                cells <- Map.remove addr cells)
            elif area > int64 model.Cells.Count * 4L then
                let mutable removedCount = 0
                let mutable keptCount = 0
                model.Cells
                |> Map.iter (fun addr cellData ->
                    if inRange addr then
                        changed <- true
                        removedCount <- removedCount + 1
                        recordRemoved cellData
                    else
                        keptCount <- keptCount + 1)
                if changed then
                    if keptCount = 0 then
                        cells <- Map.empty
                        clearedAllCells <- true
                    elif int64 removedCount * 2L >= int64 model.Cells.Count then
                        cells <- rebuildWithKeptCells keptCount
                    else
                        model.Cells
                        |> Map.iter (fun addr _ ->
                            if inRange addr then
                                cells <- Map.remove addr cells)
            elif area > 4096L then
                let mutable removedCount = 0
                for row in startRow .. endRow do
                    for col in startCol .. endCol do
                        let addr = { Row = row; Col = col }
                        let mutable cellData = Unchecked.defaultof<CellData>
                        if model.Cells.TryGetValue(addr, &cellData) then
                            removedCount <- removedCount + 1
                            recordRemoved cellData
                if removedCount > 0 then
                    changed <- true
                    let keptCount = model.Cells.Count - removedCount
                    if keptCount = 0 then
                        cells <- Map.empty
                        clearedAllCells <- true
                    elif int64 removedCount * 2L >= int64 model.Cells.Count then
                        cells <- rebuildWithKeptCells keptCount
                    else
                        for row in startRow .. endRow do
                            for col in startCol .. endCol do
                                let addr = { Row = row; Col = col }
                                if model.Cells.ContainsKey addr then
                                    cells <- Map.remove addr cells
            else
                for row in startRow .. endRow do
                    for col in startCol .. endCol do
                        let addr = { Row = row; Col = col }
                        let mutable cellData = Unchecked.defaultof<CellData>
                        if cells.TryGetValue(addr, &cellData) then
                            removeCell addr cellData
            if changed then
                if clearedAllCells then
                    rememberCellsFormulaPresenceWithSourceTextCharCount cells (Some 0) false 0
                elif keepNoFormulaPresence then
                    rememberCellsFormulaPresence cells (Some 0) false
                elif trackRemovedFormulaPresence then
                    rememberClearedCellsFormulaPresence model.Cells cells removedFormulaCount removedSourceTextCharCount
                { model with Cells = cells }
            else
                model

/// Format a number according to format code
let formatNumber (format: string option) (value: float): string =
    match format with
    | Some "currency" ->
        // Currency format: $1,234.56
        sprintf "$%s" (value.ToString("N2"))
    | Some "percentage" ->
        // Percentage format: 45.67%
        sprintf "%s%%" ((value * 100.0).ToString("N2"))
    | Some "date" ->
        // Excel date format (days since 1900-01-01)
        try
            let baseDate = System.DateTime(1900, 1, 1)
            let date = baseDate.AddDays(value - 1.0)  // Excel dates are 1-indexed
            date.ToString("MM/dd/yyyy")
        with
        | _ -> value.ToString("F2")
    | Some "decimal2" ->
        // 2 decimal places: 123.46
        value.ToString("F2")
    | Some "number" ->
        // General number with comma separators
        value.ToString("N2")
    | _ ->
        // Default: just show the number
        string value

/// Display value for cell (handles formulas) - needs to be before getCellValue
let displayValue (value: CellValue): string =
    match unwrapFormulaValue value with
    | Empty -> ""
    | Number n -> n.ToString("G", Globalization.CultureInfo.InvariantCulture)
    | String s -> s
    | Boolean b -> if b then "TRUE" else "FALSE"
    | Error e -> sprintf "#%s" e
    | Formula _ -> ""

/// Display value with formatting applied
let displayValueWithFormat (value: CellValue) (format: CellFormat): string =
    match value with
    | Number n -> formatNumber format.NumberFormat n
    | _ -> displayValue value

/// Get cell value (returns Empty if not found)
let getCellValue (addr: CellAddress) (model: SpreadsheetModel): CellValue =
    match getCell addr model with
    | Some cell -> cell.Value
    | None -> Empty

/// Adjust formula references when pasting (relative reference adjustment)
let rec adjustFormulaReferences (formula: string) (deltaRow: int) (deltaCol: int): string =
    try
        // Simple adjustment: parse the formula and adjust cell references
        // For now, we'll use a simple string replacement approach
        // TODO: Implement proper AST-based formula adjustment for absolute refs ($A$1)
        // Find all cell references in the formula (simple regex-like approach)
        let mutable i = 0
        let mutable result = System.Text.StringBuilder()

        while i < formula.Length do
            let c = formula.[i]
            if Char.IsLetter(c) && (i = 0 || not (Char.IsLetterOrDigit(formula.[i-1]))) then
                // Potential cell reference
                let mutable col = 0
                let mutable j = i

                // Collect column letters
                while j < formula.Length && Char.IsLetter(formula.[j]) do
                    let upper = Char.ToUpperInvariant formula.[j]
                    let digit = int upper - int 'A' + 1
                    if col > (Int32.MaxValue - digit) / 26 then
                        failwith "Column reference overflow"
                    col <- col * 26 + digit
                    j <- j + 1

                // Collect row digits
                let rowStart = j
                let mutable row = 0
                while j < formula.Length && Char.IsDigit(formula.[j]) do
                    let digit = int formula.[j] - int '0'
                    if row > (Int32.MaxValue - digit) / 10 then
                        failwith "Row reference overflow"
                    row <- row * 10 + digit
                    j <- j + 1

                // Check if this is a valid cell reference
                if j > rowStart then
                    // Valid cell reference - adjust it
                    let col = col - 1
                    let row = row - 1

                    let newRow = row + deltaRow
                    let newCol = col + deltaCol

                    if newRow >= 0 && newCol >= 0 then
                        let newRef = formatCellRef newRow newCol
                        result.Append(newRef) |> ignore
                    else
                        // Reference out of bounds, keep original
                        result.Append(formula.Substring(i, j - i)) |> ignore

                    i <- j
                else
                    // Not a cell reference, just a word
                    result.Append(c) |> ignore
                    i <- i + 1
            else
                result.Append(c) |> ignore
                i <- i + 1

        result.ToString()
    with
    | _ -> formula  // On error, return original formula

/// Serialize selection to clipboard data
let serializeSelection (selection: Selection) (model: SpreadsheetModel) (isCut: bool): ClipboardData =
    match selection with
    | SingleCell (row, col) ->
        let addr = { Row = row; Col = col }
        let cellData = getCell addr model
        match cellData with
        | Some data ->
            { Cells = [(0, 0, data)]; IsCut = isCut; SourceStartRow = row; SourceStartCol = col }
        | None ->
            { Cells = []; IsCut = isCut; SourceStartRow = row; SourceStartCol = col }

    | Range (startRow, startCol, endRow, endCol) ->
        let cells = ResizeArray<int * int * CellData>()
        let rowCount = max 0 (endRow - startRow + 1)
        let colCount = max 0 (endCol - startCol + 1)
        let area = int64 rowCount * int64 colCount
        if area > int64 model.Cells.Count * 4L then
            model.Cells
            |> Map.iter (fun addr data ->
                if addr.Row >= startRow && addr.Row <= endRow && addr.Col >= startCol && addr.Col <= endCol then
                    cells.Add(addr.Row - startRow, addr.Col - startCol, data))
        else
            for row in startRow .. endRow do
                for col in startCol .. endCol do
                    let addr = { Row = row; Col = col }
                    match getCell addr model with
                    | Some data -> cells.Add(row - startRow, col - startCol, data)
                    | None -> ()

        { Cells = resizeArrayToList cells; IsCut = isCut; SourceStartRow = startRow; SourceStartCol = startCol }

/// Paste clipboard data at target address
let pasteClipboard (targetAddr: CellAddress) (clipboardData: ClipboardData) (model: SpreadsheetModel): SpreadsheetModel =
    match tryPasteIdentityCache targetAddr clipboardData model with
    | Some cached -> cached
    | None ->
        let deltaRow = targetAddr.Row - clipboardData.SourceStartRow
        let deltaCol = targetAddr.Col - clipboardData.SourceStartCol

        let mutable adjustedFormulaCache: System.Collections.Generic.Dictionary<string, string> option = None

        let adjustedCellData (cellData: CellData) =
            match cellData.Value with
            | Formula (formula, _) ->
                let cache =
                    match adjustedFormulaCache with
                    | Some cache -> cache
                    | None ->
                        let cache = System.Collections.Generic.Dictionary<string, string>()
                        adjustedFormulaCache <- Some cache
                        cache
                let mutable adjustedFormula = Unchecked.defaultof<string>
                if not (cache.TryGetValue(formula, &adjustedFormula)) then
                    adjustedFormula <-
                        if deltaRow = 0 && deltaCol = 0 then formula
                        else adjustFormulaReferences formula deltaRow deltaCol
                    cache.[formula] <- adjustedFormula
                // Will be re-evaluated when set
                { cellData with Value = Formula (adjustedFormula, Empty) }
            | _ -> cellData

        if Map.isEmpty model.Cells then
            let mutable changed = false
            let pastedCells =
                seq {
                    for (relRow, relCol, cellData) in clipboardData.Cells do
                        let pasteRow = targetAddr.Row + relRow
                        let pasteCol = targetAddr.Col + relCol

                        if pasteRow >= 0 && pasteRow < model.RowCount && pasteCol >= 0 && pasteCol < model.ColCount then
                            changed <- true
                            let pasteAddr = { Row = pasteRow; Col = pasteCol }
                            let pastedCellData =
                                match cellData.Value with
                                | Formula _ -> adjustedCellData cellData
                                | _ -> cellData
                            yield pasteAddr, pastedCellData
                }
            let cells = Map.ofSeq pastedCells
            if changed then
                let result = { model with Cells = cells }
                rememberPasteIdentityForChangedResult targetAddr clipboardData result
                result
            else
                rememberPasteIdentity targetAddr clipboardData model
                model
        else
            let mutable cells = model.Cells
            let mutable changed = false

            for (relRow, relCol, cellData) in clipboardData.Cells do
                let pasteRow = targetAddr.Row + relRow
                let pasteCol = targetAddr.Col + relCol

                if pasteRow >= 0 && pasteRow < model.RowCount && pasteCol >= 0 && pasteCol < model.ColCount then
                    let pasteAddr = { Row = pasteRow; Col = pasteCol }
                    let pastedCellData =
                        match cellData.Value with
                        | Formula _ -> adjustedCellData cellData
                        | _ -> cellData
                    let mutable oldCell = Unchecked.defaultof<CellData>
                    if not (cells.TryGetValue(pasteAddr, &oldCell)) || oldCell <> pastedCellData then
                        cells <- Map.add pasteAddr pastedCellData cells
                        changed <- true

            if changed then
                let result = { model with Cells = cells }
                rememberPasteIdentityForChangedResult targetAddr clipboardData result
                result
            else
                rememberPasteIdentity targetAddr clipboardData model
                model

/// Push command to history (clears future stack)
let pushCommand (command: Command) (model: SpreadsheetModel): SpreadsheetModel =
    if model.History.MaxSize <= 0 then
        if model.History.Past.IsEmpty && model.History.Future.IsEmpty then
            model
        else
            { model with History = { model.History with Past = []; Future = [] } }
    else
        let newPast = command :: model.History.Past
        // Limit stack size
        let trimmedPast = trimHistoryStack model.History.MaxSize newPast

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
                Future = (command :: undoneModel.History.Future) |> trimHistoryStack undoneModel.History.MaxSize
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
                Past = (command :: redoneModel.History.Past) |> trimHistoryStack redoneModel.History.MaxSize
                Future = restFuture
        }

        Some { redoneModel with History = newHistory }

/// Compare cell values for sorting (numbers < strings < booleans < empty)
let compareCellValues (v1: CellValue) (v2: CellValue): int =
    let v1 = unwrapFormulaValue v1
    let v2 = unwrapFormulaValue v2
    match v1, v2 with
    | Number n1, Number n2 -> compare n1 n2
    | String s1, String s2 -> System.String.Compare(s1, s2, System.StringComparison.OrdinalIgnoreCase)
    | Boolean b1, Boolean b2 -> compare b1 b2
    | Error _, Error _ -> 0
    | Empty, Empty -> 0
    | _ ->
        let rank = function
            | Number _ -> 0
            | String _ -> 1
            | Boolean _ -> 2
            | Error _ -> 3
            | Empty -> 4
            | Formula _ -> 5
        compare (rank v1) (rank v2)

let private sortedStringSetValues (values: System.Collections.Generic.HashSet<string>) : string list =
    let arr = Array.zeroCreate<string> values.Count
    let mutable index = 0
    for value in values do
        arr.[index] <- value
        index <- index + 1
    Array.sortInPlace arr
    arrayToList arr

let private sortedStringSetValuesWithTextMetrics limit (values: System.Collections.Generic.HashSet<string>) =
    let arr = Array.zeroCreate<string> values.Count
    let mutable index = 0
    for value in values do
        arr.[index] <- value
        index <- index + 1
    Array.sortInPlace arr
    let mutable result = []
    let mutable textCharCount = 0
    for i = arr.Length - 1 downto 0 do
        let value = arr.[i]
        if textCharCount <= limit then
            textCharCount <- addCapped limit value.Length textCharCount
        result <- value :: result
    result, arr.Length, textCharCount

type private ColumnUniqueValuesCache = {
    Cells: Map<CellAddress, CellData>
    CellCount: int
    SourceTextCharCount: int
    Entries: System.Collections.Generic.Dictionary<int, ColumnUniqueValuesCacheEntry>
    Order: System.Collections.Generic.Queue<int>
    mutable TotalValueCount: int
    mutable TotalValueTextCharCount: int
}

and private ColumnUniqueValuesCacheEntry = {
    ValuesWithBlanks: string list
    ValuesWithoutBlanks: string list
    ValueCount: int
    ValueTextCharCount: int
}

let mutable private columnUniqueValuesCache: ColumnUniqueValuesCache list = []
let private columnUniqueValuesCacheByCells =
    System.Collections.Generic.Dictionary<Map<CellAddress, CellData>, ColumnUniqueValuesCache>(cellMapReferenceComparer)
let private columnUniqueValuesFastSlots = Array.zeroCreate<ColumnUniqueValuesCache> 256
let private columnUniqueValuesCacheMaxEntries = 64
let private columnUniqueValuesCacheMaxSources = 64
let private columnUniqueValuesCacheMaxTotalValues = 200000
let private columnUniqueValuesCacheMaxTotalValueTextChars = 5000000
let private columnUniqueValuesCacheMaxTotalCells = 500000
let private columnUniqueValuesCacheMaxTotalSourceTextChars = 5000000

let private columnUniqueValuesCacheMatches cells (cached: ColumnUniqueValuesCache) =
    obj.ReferenceEquals(cached.Cells, cells)

let private setColumnUniqueValuesCache entries =
    columnUniqueValuesCache <- entries
    columnUniqueValuesCacheByCells.Clear()
    clearArray columnUniqueValuesFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<ColumnUniqueValuesCache>
        if not (columnUniqueValuesCacheByCells.TryGetValue(entry.Cells, &existing)) then
            columnUniqueValuesCacheByCells.[entry.Cells] <- entry
            let slot = referenceSlot entry.Cells columnUniqueValuesFastSlots.Length
            columnUniqueValuesFastSlots.[slot] <- entry

let private tryGetColumnUniqueValuesCache cells =
    let slot = referenceSlot cells columnUniqueValuesFastSlots.Length
    let fast = columnUniqueValuesFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && columnUniqueValuesCacheMatches cells fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<ColumnUniqueValuesCache>
        if columnUniqueValuesCacheByCells.TryGetValue(cells, &cached) then
            columnUniqueValuesFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private columnUniqueValuesCacheWithoutCells cells =
    let mutable cached = Unchecked.defaultof<ColumnUniqueValuesCache>
    if columnUniqueValuesCacheByCells.TryGetValue(cells, &cached) then
        columnUniqueValuesCache
        |> List.filter (columnUniqueValuesCacheMatches cells >> not)
    else
        columnUniqueValuesCache

let private removeColumnUniqueValuesCacheForCells cells =
    let entries = columnUniqueValuesCacheWithoutCells cells
    if not (obj.ReferenceEquals(entries, columnUniqueValuesCache)) then
        setColumnUniqueValuesCache entries

let private columnUniqueValuesCacheRetainable (cached: ColumnUniqueValuesCache) =
    cached.Entries.Count > 0
    && cached.TotalValueCount <= columnUniqueValuesCacheMaxTotalValues
    && cached.TotalValueTextCharCount <= columnUniqueValuesCacheMaxTotalValueTextChars
    && cached.CellCount <= columnUniqueValuesCacheMaxTotalCells
    && cached.SourceTextCharCount <= columnUniqueValuesCacheMaxTotalSourceTextChars

let private columnUniqueValuesSourceTextCharCountFor valueCount valueTextCharCount cells =
    if valueCount > columnUniqueValuesCacheMaxTotalValues
       || valueTextCharCount > columnUniqueValuesCacheMaxTotalValueTextChars then
        columnUniqueValuesCacheMaxTotalSourceTextChars + 1
    else
        cellMapSourceTextCharCountFor
            columnUniqueValuesCacheMaxTotalCells
            columnUniqueValuesCacheMaxTotalSourceTextChars
            cells

let private rememberColumnUniqueValuesCache cached =
    if not (columnUniqueValuesCacheRetainable cached) then
        removeColumnUniqueValuesCacheForCells cached.Cells
    else
        let mutable keptCount = 1
        let mutable keptValues = cached.TotalValueCount
        let mutable keptValueTextChars = cached.TotalValueTextCharCount
        let mutable keptCells = cached.CellCount
        let mutable keptSourceTextChars = max 1 cached.SourceTextCharCount
        let retained =
            columnUniqueValuesCacheWithoutCells cached.Cells
            |> List.choose (fun candidate ->
                let candidateValues = candidate.TotalValueCount
                let candidateValueTextChars = max 1 candidate.TotalValueTextCharCount
                let candidateCells = candidate.CellCount
                let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                if columnUniqueValuesCacheRetainable candidate
                   && keptCount < columnUniqueValuesCacheMaxSources
                   && keptValues + candidateValues <= columnUniqueValuesCacheMaxTotalValues
                   && keptValueTextChars + candidateValueTextChars <= columnUniqueValuesCacheMaxTotalValueTextChars
                   && keptCells + candidateCells <= columnUniqueValuesCacheMaxTotalCells
                   && keptSourceTextChars + candidateSourceTextChars <= columnUniqueValuesCacheMaxTotalSourceTextChars then
                    keptCount <- keptCount + 1
                    keptValues <- keptValues + candidateValues
                    keptValueTextChars <- keptValueTextChars + candidateValueTextChars
                    keptCells <- keptCells + candidateCells
                    keptSourceTextChars <- keptSourceTextChars + candidateSourceTextChars
                    Some candidate
                else
                    None)
        cached :: retained
        |> setColumnUniqueValuesCache

let private removeColumnUniqueValuesCacheEntry col (cache: ColumnUniqueValuesCache) =
    let mutable existing = Unchecked.defaultof<ColumnUniqueValuesCacheEntry>
    if cache.Entries.TryGetValue(col, &existing) then
        cache.TotalValueCount <- max 0 (cache.TotalValueCount - existing.ValueCount)
        cache.TotalValueTextCharCount <- max 0 (cache.TotalValueTextCharCount - existing.ValueTextCharCount)
        cache.Entries.Remove col |> ignore

let private addColumnUniqueValuesCacheEntry col entry (cache: ColumnUniqueValuesCache) =
    let mutable existing = Unchecked.defaultof<ColumnUniqueValuesCacheEntry>
    if entry.ValueCount > columnUniqueValuesCacheMaxTotalValues
       || entry.ValueTextCharCount > columnUniqueValuesCacheMaxTotalValueTextChars then
        removeColumnUniqueValuesCacheEntry col cache
    else
        if cache.Entries.TryGetValue(col, &existing) then
            cache.TotalValueCount <- max 0 (cache.TotalValueCount - existing.ValueCount)
            cache.TotalValueTextCharCount <- max 0 (cache.TotalValueTextCharCount - existing.ValueTextCharCount)
        else
            cache.Order.Enqueue col
        cache.Entries.[col] <- entry
        cache.TotalValueCount <- cache.TotalValueCount + entry.ValueCount
        cache.TotalValueTextCharCount <- cache.TotalValueTextCharCount + entry.ValueTextCharCount
        while (cache.Entries.Count > columnUniqueValuesCacheMaxEntries
               || cache.TotalValueCount > columnUniqueValuesCacheMaxTotalValues
               || cache.TotalValueTextCharCount > columnUniqueValuesCacheMaxTotalValueTextChars)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if oldest <> col then
                removeColumnUniqueValuesCacheEntry oldest cache
        if cache.Entries.Count > columnUniqueValuesCacheMaxEntries
           || cache.TotalValueCount > columnUniqueValuesCacheMaxTotalValues
           || cache.TotalValueTextCharCount > columnUniqueValuesCacheMaxTotalValueTextChars then
            let current = cache.Entries.[col]
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue col
            cache.Entries.[col] <- current
            cache.TotalValueCount <- current.ValueCount
            cache.TotalValueTextCharCount <- current.ValueTextCharCount

type private SortIdentityCache = {
    Cells: Map<CellAddress, CellData>
    CellCount: int
    SourceTextCharCount: int
    Entries: System.Collections.Generic.Dictionary<struct (int * bool), unit>
    Order: System.Collections.Generic.Queue<struct (int * bool)>
}

let mutable private sortIdentityCache: SortIdentityCache list = []
let private sortIdentityCacheByCells =
    System.Collections.Generic.Dictionary<Map<CellAddress, CellData>, SortIdentityCache>(cellMapReferenceComparer)
let private sortIdentityFastSlots = Array.zeroCreate<SortIdentityCache> 256
let private sortIdentityCacheMaxEntries = 64
let private sortIdentityCacheMaxSources = 64
let private sortIdentityCacheMaxTotalCells = 500000
let private sortIdentityCacheMaxTotalSourceTextChars = 5000000

let private sortIdentityCacheMatches cells (cached: SortIdentityCache) =
    obj.ReferenceEquals(cached.Cells, cells)

let private setSortIdentityCache entries =
    sortIdentityCache <- entries
    sortIdentityCacheByCells.Clear()
    clearArray sortIdentityFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<SortIdentityCache>
        if not (sortIdentityCacheByCells.TryGetValue(entry.Cells, &existing)) then
            sortIdentityCacheByCells.[entry.Cells] <- entry
            let slot = referenceSlot entry.Cells sortIdentityFastSlots.Length
            sortIdentityFastSlots.[slot] <- entry

let private tryGetSortIdentityCache cells =
    let slot = referenceSlot cells sortIdentityFastSlots.Length
    let fast = sortIdentityFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && sortIdentityCacheMatches cells fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<SortIdentityCache>
        if sortIdentityCacheByCells.TryGetValue(cells, &cached) then
            sortIdentityFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private sortIdentityCacheWithoutCells cells =
    let mutable cached = Unchecked.defaultof<SortIdentityCache>
    if sortIdentityCacheByCells.TryGetValue(cells, &cached) then
        sortIdentityCache
        |> List.filter (sortIdentityCacheMatches cells >> not)
    else
        sortIdentityCache

let private removeSortIdentityCacheForCells cells =
    let entries = sortIdentityCacheWithoutCells cells
    if not (obj.ReferenceEquals(entries, sortIdentityCache)) then
        setSortIdentityCache entries

let private sortIdentityCacheRetainable (cached: SortIdentityCache) =
    cached.CellCount <= sortIdentityCacheMaxTotalCells
    && cached.SourceTextCharCount <= sortIdentityCacheMaxTotalSourceTextChars

let private sortIdentitySourceTextCharCount cells =
    cellMapSourceTextCharCountFor
        sortIdentityCacheMaxTotalCells
        sortIdentityCacheMaxTotalSourceTextChars
        cells

let private rememberSortIdentityCache cached =
    if not (sortIdentityCacheRetainable cached) then
        removeSortIdentityCacheForCells cached.Cells
    else
        let trim (entries: SortIdentityCache list) =
            let rec loop keptCount keptCells keptSourceTextChars acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateCells = max 1 candidate.CellCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    if sortIdentityCacheRetainable candidate
                       && keptCount < sortIdentityCacheMaxSources
                       && keptCells + candidateCells <= sortIdentityCacheMaxTotalCells
                       && keptSourceTextChars + candidateSourceTextChars <= sortIdentityCacheMaxTotalSourceTextChars then
                        loop
                            (keptCount + 1)
                            (keptCells + candidateCells)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptCells keptSourceTextChars acc tail
            loop 0 0 0 [] entries
        (cached :: sortIdentityCacheWithoutCells cached.Cells)
        |> trim
        |> setSortIdentityCache

let private addSortIdentityCacheEntry key (entries: System.Collections.Generic.Dictionary<struct (int * bool), unit>) (order: System.Collections.Generic.Queue<struct (int * bool)>) =
    if not (entries.ContainsKey key) then
        order.Enqueue key
    entries.[key] <- ()
    while entries.Count > sortIdentityCacheMaxEntries && order.Count > 0 do
        let oldest = order.Dequeue()
        if oldest <> key then
            entries.Remove oldest |> ignore
    if entries.Count > sortIdentityCacheMaxEntries then
        entries.Clear()
        order.Clear()
        order.Enqueue key
        entries.[key] <- ()

let private rememberSortIdentityWithSourceTextCharCount sourceTextCharCount cells col ascending =
    let key = struct (col, ascending)
    let cache =
        match tryGetSortIdentityCache cells with
        | ValueSome cached -> cached
        | ValueNone ->
            {
                Cells = cells
                CellCount = max 1 cells.Count
                SourceTextCharCount = sourceTextCharCount
                Entries = System.Collections.Generic.Dictionary<struct (int * bool), unit>()
                Order = System.Collections.Generic.Queue<struct (int * bool)>()
            }
    addSortIdentityCacheEntry key cache.Entries cache.Order
    rememberSortIdentityCache cache

let private rememberSortIdentity cells col ascending =
    rememberSortIdentityWithSourceTextCharCount (sortIdentitySourceTextCharCount cells) cells col ascending

let private hasSortIdentity cells col ascending =
    let key = struct (col, ascending)
    match tryGetSortIdentityCache cells with
    | ValueSome cached ->
        cached.Entries.ContainsKey key
    | ValueNone -> false

let private sortStateMatches col ascending (sortState: SortState option) =
    match sortState with
    | Some state -> state.Column = col && state.Ascending = ascending
    | None -> false

let private makeSortState col ascending =
    Some { Column = col; Ascending = ascending }

let private collectSortRowsWithValues col (cells: Map<CellAddress, CellData>) =
    let rowsWithValues = ResizeArray<int * CellValue>()
    let mutable haveLastRow = false
    let mutable lastRow = 0
    cells
    |> Map.iter (fun addr cellData ->
        if not haveLastRow || addr.Row <> lastRow then
            rowsWithValues.Add(addr.Row, CellValue.Empty)
            haveLastRow <- true
            lastRow <- addr.Row
        if addr.Col = col then
            rowsWithValues.[rowsWithValues.Count - 1] <- (addr.Row, cellData.Value))
    rowsWithValues

let private sortRowsAreIdentitySorted col ascending (cells: Map<CellAddress, CellData>) =
    let sourceTextLimit = sortIdentityCacheMaxTotalSourceTextChars
    let mutable haveCurrentRow = false
    let mutable currentRow = 0
    let mutable currentValue = CellValue.Empty
    let mutable haveRows = false
    let mutable rowPosition = 0
    let mutable previousValue = CellValue.Empty
    let mutable havePreviousValue = false
    let mutable alreadyIdentitySorted = true
    let mutable sourceTextCharCount = 0

    let finishRow () =
        if haveCurrentRow then
            haveRows <- true
            if currentRow <> rowPosition then
                alreadyIdentitySorted <- false
            if havePreviousValue then
                let comparison = compareCellValues previousValue currentValue
                if (ascending && comparison > 0) || (not ascending && comparison < 0) then
                    alreadyIdentitySorted <- false
            previousValue <- currentValue
            havePreviousValue <- true
            rowPosition <- rowPosition + 1

    cells
    |> Map.iter (fun addr cellData ->
        if sourceTextCharCount <= sourceTextLimit then
            sourceTextCharCount <-
                addCapped
                    sourceTextLimit
                    (cellDataTextCharCountFor sourceTextLimit cellData)
                    sourceTextCharCount
        if not haveCurrentRow || addr.Row <> currentRow then
            finishRow()
            haveCurrentRow <- true
            currentRow <- addr.Row
            currentValue <- CellValue.Empty
        if addr.Col = col then
            currentValue <- cellData.Value)

    finishRow()
    haveRows, alreadyIdentitySorted, sourceTextCharCount

let private columnUniqueValuesCached includeBlanks col (model: SpreadsheetModel) =
    let tryCached () =
        match tryGetColumnUniqueValuesCache model.Cells with
        | ValueSome cached ->
            let mutable entry = Unchecked.defaultof<ColumnUniqueValuesCacheEntry>
            if cached.Entries.TryGetValue(col, &entry) then
                Some (if includeBlanks then entry.ValuesWithBlanks else entry.ValuesWithoutBlanks)
            else
                None
        | ValueNone -> None

    match tryCached () with
    | Some values -> values
    | _ ->
        let valuesWithBlanks = System.Collections.Generic.HashSet<string>()
        let valuesWithoutBlanks = System.Collections.Generic.HashSet<string>()
        let mutable errorDisplayCache = Unchecked.defaultof<System.Collections.Generic.Dictionary<string, string>>
        let displayCached value =
            match unwrapFormulaValue value with
            | Error code ->
                if obj.ReferenceEquals(errorDisplayCache, null) then
                    errorDisplayCache <- System.Collections.Generic.Dictionary<string, string>()
                let mutable cached = Unchecked.defaultof<string>
                if errorDisplayCache.TryGetValue(code, &cached) then
                    cached
                else
                    let text = "#" + code
                    errorDisplayCache.[code] <- text
                    text
            | other -> displayValue other
        model.Cells
        |> Map.iter (fun addr cellData ->
            if addr.Col = col then
                match displayCached cellData.Value with
                | "" ->
                    valuesWithBlanks.Add "" |> ignore
                | value ->
                    valuesWithBlanks.Add value |> ignore
                    valuesWithoutBlanks.Add value |> ignore)
        let withBlanks, withBlanksCount, valueTextCharCount =
            sortedStringSetValuesWithTextMetrics columnUniqueValuesCacheMaxTotalValueTextChars valuesWithBlanks
        let withoutBlanks = sortedStringSetValues valuesWithoutBlanks
        let valueCount = withBlanksCount + valuesWithoutBlanks.Count
        let cache =
            match tryGetColumnUniqueValuesCache model.Cells with
            | ValueSome cached -> cached
            | ValueNone ->
                {
                Cells = model.Cells
                CellCount = max 1 model.Cells.Count
                SourceTextCharCount =
                    columnUniqueValuesSourceTextCharCountFor
                        valueCount
                        valueTextCharCount
                        model.Cells
                Entries = System.Collections.Generic.Dictionary<int, ColumnUniqueValuesCacheEntry>()
                Order = System.Collections.Generic.Queue<int>()
                TotalValueCount = 0
                TotalValueTextCharCount = 0
            }
        addColumnUniqueValuesCacheEntry col {
            ValuesWithBlanks = withBlanks
            ValuesWithoutBlanks = withoutBlanks
            ValueCount = valueCount
            ValueTextCharCount = valueTextCharCount
        } cache
        rememberColumnUniqueValuesCache cache
        if includeBlanks then withBlanks else withoutBlanks

/// Sort spreadsheet by column
let sortByColumn (col: int) (ascending: bool) (model: SpreadsheetModel): SpreadsheetModel =
    let hasRequestedSortState = sortStateMatches col ascending model.SortState
    let modelWithSortState () =
        if hasRequestedSortState then model else { model with SortState = makeSortState col ascending }

    if hasSortIdentity model.Cells col ascending then
        modelWithSortState()
    else
        let hasRows, alreadyIdentitySorted, sourceTextCharCount = sortRowsAreIdentitySorted col ascending model.Cells
        if not hasRows then
            // No data to sort
            rememberSortIdentityWithSourceTextCharCount sourceTextCharCount model.Cells col ascending
            modelWithSortState()
        elif alreadyIdentitySorted then
            rememberSortIdentityWithSourceTextCharCount sourceTextCharCount model.Cells col ascending
            modelWithSortState()
        else
            // Collect all rows that have data in any column, plus the value in
            // the sort column when present. Map iteration is row-major, so this
            // avoids a second Map.tryFind pass over every populated row.
            let rowsWithValues = collectSortRowsWithValues col model.Cells

            // Sort rows by their values in the sort column
            rowsWithValues.Sort(System.Comparison<int * CellValue>(fun (row1, v1) (row2, v2) ->
                if ascending then
                    match compareCellValues v1 v2 with
                    | 0 -> compare row1 row2
                    | c -> c
                else
                    match compareCellValues v2 v1 with
                    | 0 -> compare row1 row2
                    | c -> c))

            // Create mapping from old row index to new row index
            let rowMapping = System.Collections.Generic.Dictionary<int, int>(rowsWithValues.Count)
            let mutable identityMapping = true
            for newIdx in 0 .. rowsWithValues.Count - 1 do
                let oldRow, _ = rowsWithValues.[newIdx]
                rowMapping.[oldRow] <- newIdx
                if oldRow <> newIdx then
                    identityMapping <- false

            if identityMapping then
                rememberSortIdentity model.Cells col ascending
                modelWithSortState()
            else
                // Rebuild cells with new row indices
                let reorderedCells = Array.zeroCreate<CellAddress * CellData> model.Cells.Count
                let mutable reorderedIndex = 0
                model.Cells
                |> Map.iter (fun addr cellData ->
                    let newAddr =
                        match rowMapping.TryGetValue addr.Row with
                        | true, newRow -> { addr with Row = newRow }
                        | false, _ -> addr
                    reorderedCells.[reorderedIndex] <- newAddr, cellData
                    reorderedIndex <- reorderedIndex + 1)
                let newCells = Map.ofArray reorderedCells

                let nextModel = {
                    model with
                        Cells = newCells
                        SortState = makeSortState col ascending
                }
                rememberSortIdentity nextModel.Cells col ascending
                nextModel

/// Get unique values in a column (as strings for display)
let getColumnUniqueValues (col: int) (model: SpreadsheetModel): string list =
    columnUniqueValuesCached true col model

let private collectRowsWithCellsArray (model: SpreadsheetModel) =
    let rows = ResizeArray<int>()
    let mutable lastRow = System.Int32.MinValue
    model.Cells
    |> Map.iter (fun addr _ ->
        if addr.Row <> lastRow then
            rows.Add addr.Row
            lastRow <- addr.Row)
    rows.ToArray()

let private collectRowsWithCells (model: SpreadsheetModel) =
    let rows = System.Collections.Generic.HashSet<int>()
    for row in collectRowsWithCellsArray model do
        rows.Add row |> ignore
    rows

let private rowPassesFilters (model: SpreadsheetModel) row =
    model.Filters
    |> Map.forall (fun col filter ->
        let addr = { Row = row; Col = col }
        let value = getCellValue addr model
        let displayStr = displayValue value
        Set.contains displayStr filter.VisibleValues)

let private rowValuesPassFilters (rowValues: CellValue array) (filterValues: Set<string> array) =
    let mutable passes = true
    let mutable i = 0
    while passes && i < rowValues.Length do
        passes <- Set.contains (displayValue rowValues.[i]) filterValues.[i]
        i <- i + 1
    passes

let private rowValuesPassFiltersVersioned
    (rowValues: CellValue array)
    (rowVersions: int array)
    rowVersion
    (filterValues: Set<string> array) =
    let mutable passes = true
    let mutable i = 0
    while passes && i < rowValues.Length do
        if rowVersions.[i] = rowVersion then
            passes <- Set.contains (displayValue rowValues.[i]) filterValues.[i]
        else
            passes <- Set.contains "" filterValues.[i]
        i <- i + 1
    passes

let private filterEmptyAllowances (filterValues: Set<string> array) =
    let allowsEmpty = Array.zeroCreate<bool> filterValues.Length
    let mutable rejectedCount = 0
    for i = 0 to filterValues.Length - 1 do
        let allowed = Set.contains "" filterValues.[i]
        allowsEmpty.[i] <- allowed
        if not allowed then
            rejectedCount <- rejectedCount + 1
    allowsEmpty, rejectedCount

let private touchedRowValuesPassFilters
    (rowValues: CellValue array)
    (filterValues: Set<string> array)
    (touchedIndexes: ResizeArray<int>)
    missingRejectedCount =
    let mutable passes = missingRejectedCount = 0
    let mutable i = 0
    while passes && i < touchedIndexes.Count do
        let filterIndex = touchedIndexes.[i]
        passes <- Set.contains (displayValue rowValues.[filterIndex]) filterValues.[filterIndex]
        i <- i + 1
    passes

let private filterColumnIndex (filterCols: int array) col =
    match filterCols.Length with
    | 1 ->
        if col = filterCols.[0] then 0 else -1
    | 2 ->
        if col = filterCols.[0] then 0
        elif col = filterCols.[1] then 1
        else -1
    | 3 ->
        if col = filterCols.[0] then 0
        elif col = filterCols.[1] then 1
        elif col = filterCols.[2] then 2
        else -1
    | _ ->
        binarySearchInt filterCols col

let private filterColumnsAreContiguous (filterCols: int array) =
    let mutable contiguous = filterCols.Length > 0
    let mutable i = 1
    while contiguous && i < filterCols.Length do
        contiguous <- filterCols.[i] = filterCols.[i - 1] + 1
        i <- i + 1
    contiguous

let private filterColumnIndexFor contiguous firstCol lastCol (filterCols: int array) col =
    if contiguous then
        if col >= firstCol && col <= lastCol then col - firstCol else -1
    else
        filterColumnIndex filterCols col

let private filterHiddenRowsByCellScan (filterCols: int array) (filterValues: Set<string> array) (model: SpreadsheetModel) =
    let hidden = ResizeArray<int>()
    if filterCols.Length <= 3 then
        let rowValues = Array.create filterCols.Length CellValue.Empty
        let mutable currentRow = System.Int32.MinValue
        let mutable headerRow = System.Int32.MinValue
        let mutable hasRow = false

        let finishRow () =
            if hasRow && currentRow <> headerRow && not (rowValuesPassFilters rowValues filterValues) then
                hidden.Add currentRow

        model.Cells
        |> Map.iter (fun addr cellData ->
            if addr.Row <> currentRow then
                finishRow()
                currentRow <- addr.Row
                if not hasRow then
                    headerRow <- currentRow
                hasRow <- true
                Array.fill rowValues 0 rowValues.Length CellValue.Empty
            let filterIndex = filterColumnIndex filterCols addr.Col
            if filterIndex >= 0 then
                rowValues.[filterIndex] <- cellData.Value)

        finishRow()
    else
        let rowValues = Array.create filterCols.Length CellValue.Empty
        let rowVersions = Array.zeroCreate<int> filterCols.Length
        let touchedIndexes = ResizeArray<int>()
        let emptyAllowed, emptyRejectedCount = filterEmptyAllowances filterValues
        let mutable currentRow = System.Int32.MinValue
        let mutable currentRowVersion = 0
        let mutable headerRow = System.Int32.MinValue
        let mutable hasRow = false
        let mutable missingRejectedCount = emptyRejectedCount
        let contiguousFilters = filterColumnsAreContiguous filterCols
        let firstFilterCol = filterCols.[0]
        let lastFilterCol = filterCols.[filterCols.Length - 1]

        let finishRow () =
            if hasRow
               && currentRow <> headerRow
               && not (touchedRowValuesPassFilters rowValues filterValues touchedIndexes missingRejectedCount) then
                hidden.Add currentRow

        model.Cells
        |> Map.iter (fun addr cellData ->
            if addr.Row <> currentRow then
                finishRow()
                currentRow <- addr.Row
                currentRowVersion <- currentRowVersion + 1
                touchedIndexes.Clear()
                missingRejectedCount <- emptyRejectedCount
                if not hasRow then
                    headerRow <- currentRow
                hasRow <- true
            let filterIndex = filterColumnIndexFor contiguousFilters firstFilterCol lastFilterCol filterCols addr.Col
            if filterIndex >= 0 then
                if rowVersions.[filterIndex] <> currentRowVersion then
                    rowVersions.[filterIndex] <- currentRowVersion
                    touchedIndexes.Add filterIndex
                    if not emptyAllowed.[filterIndex] then
                        missingRejectedCount <- missingRejectedCount - 1
                rowValues.[filterIndex] <- cellData.Value
            )

        finishRow()
    hidden.ToArray()

let private filterHiddenRowsByRowLookup (filterCols: int array) (filterValues: Set<string> array) (model: SpreadsheetModel) =
    let rowValues = Array.create filterCols.Length CellValue.Empty
    let hidden = ResizeArray<int>()
    let mutable headerRow = System.Int32.MinValue
    let mutable hasHeader = false
    for row in 0 .. model.RowCount - 1 do
        Array.fill rowValues 0 rowValues.Length CellValue.Empty
        let mutable hasFilterCell = false
        for i in 0 .. filterCols.Length - 1 do
            match Map.tryFind { Row = row; Col = filterCols.[i] } model.Cells with
            | Some cellData ->
                rowValues.[i] <- cellData.Value
                hasFilterCell <- true
            | None -> ()
        if hasFilterCell then
            if not hasHeader then
                headerRow <- row
                hasHeader <- true
            elif row <> headerRow && not (rowValuesPassFilters rowValues filterValues) then
                hidden.Add row
    hidden.ToArray()

let private filterHiddenRowsUncached (model: SpreadsheetModel) =
    if Map.isEmpty model.Filters then [||]
    else
        let filterEntries = model.Filters |> Map.toArray
        let filterCols = filterEntries |> Array.map fst
        let filterValues = filterEntries |> Array.map (fun (_, filter) -> filter.VisibleValues)
        let rowLookupProbeCount = int64 model.RowCount * int64 filterCols.Length
        if model.RowCount > 0 && rowLookupProbeCount <= int64 model.Cells.Count / 4L then
            filterHiddenRowsByRowLookup filterCols filterValues model
        else
            filterHiddenRowsByCellScan filterCols filterValues model

type private FilterHiddenRowsCache = {
    Cells: Map<CellAddress, CellData>
    CellCount: int
    SourceTextCharCount: int
    RowCount: int
    Entries: System.Collections.Generic.Dictionary<Map<int, ColumnFilter>, FilterHiddenRowsCacheEntry>
    Order: System.Collections.Generic.Queue<Map<int, ColumnFilter>>
    mutable TotalRowCount: int
    mutable TotalFilterValueCount: int
    mutable TotalFilterTextCharCount: int
}

and private FilterHiddenRowsCacheEntry = {
    Rows: int array
    FilterValueCount: int
    FilterTextCharCount: int
}

let mutable private filterHiddenRowsCache: FilterHiddenRowsCache list = []
let private filterHiddenRowsCacheBySource =
    System.Collections.Generic.Dictionary<CellRowCountReferenceKey, FilterHiddenRowsCache>(cellRowCountReferenceComparer)
let private filterHiddenRowsFastSlots = Array.zeroCreate<FilterHiddenRowsCache> 256
let private filterHiddenRowsCacheMaxEntries = 64
let private filterHiddenRowsCacheMaxSources = 64
let private filterHiddenRowsCacheMaxRowCount = 2000000
let private filterHiddenRowsCacheMaxTotalRowCount = 2000000
let private filterHiddenRowsCacheMaxTotalCells = 500000
let private filterHiddenRowsCacheMaxTotalSourceTextChars = 5000000
let private filterHiddenRowsCacheMaxFilterValues = 100000
let private filterHiddenRowsCacheMaxTotalFilterValues = 200000
let private filterHiddenRowsCacheMaxFilterTextChars = 5000000
let private filterHiddenRowsCacheMaxTotalFilterTextChars = 5000000

let private filterValueCount (filters: Map<int, ColumnFilter>) =
    let mutable count = 0
    use entries = (filters :> seq<_>).GetEnumerator()
    while count <= filterHiddenRowsCacheMaxFilterValues && entries.MoveNext() do
        count <- addCapped filterHiddenRowsCacheMaxFilterValues entries.Current.Value.VisibleValues.Count count
    count

let private filterTextCharCount maxTotalFilterTextChars (filters: Map<int, ColumnFilter>) =
    let limit = maxTotalFilterTextChars + 1
    let mutable total = 0
    use filters = (filters :> seq<_>).GetEnumerator()
    while total <= maxTotalFilterTextChars && filters.MoveNext() do
        let filter = filters.Current.Value
        use values = (filter.VisibleValues :> seq<string>).GetEnumerator()
        while total <= maxTotalFilterTextChars && values.MoveNext() do
            total <- addCapped maxTotalFilterTextChars values.Current.Length total
    min total limit

let private filterHiddenRowsCacheMatches cells rowCount (cached: FilterHiddenRowsCache) =
    obj.ReferenceEquals(cached.Cells, cells)
    && cached.RowCount = rowCount

let private setFilterHiddenRowsCache entries =
    filterHiddenRowsCache <- entries
    filterHiddenRowsCacheBySource.Clear()
    clearArray filterHiddenRowsFastSlots
    for entry in entries do
        let key = cellRowCountReferenceKey entry.Cells entry.RowCount
        let mutable existing = Unchecked.defaultof<FilterHiddenRowsCache>
        if not (filterHiddenRowsCacheBySource.TryGetValue(key, &existing)) then
            filterHiddenRowsCacheBySource.[key] <- entry
            let slot = cellRowCountReferenceHash entry.Cells entry.RowCount &&& (filterHiddenRowsFastSlots.Length - 1)
            filterHiddenRowsFastSlots.[slot] <- entry

let private tryGetFilterHiddenRowsCache cells rowCount =
    let slot = cellRowCountReferenceHash cells rowCount &&& (filterHiddenRowsFastSlots.Length - 1)
    let fast = filterHiddenRowsFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && filterHiddenRowsCacheMatches cells rowCount fast then
        ValueSome fast
    else
        let key = cellRowCountReferenceKey cells rowCount
        let mutable cached = Unchecked.defaultof<FilterHiddenRowsCache>
        if filterHiddenRowsCacheBySource.TryGetValue(key, &cached) then
            filterHiddenRowsFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private filterHiddenRowsCacheWithoutSource cells rowCount =
    let key = cellRowCountReferenceKey cells rowCount
    let mutable cached = Unchecked.defaultof<FilterHiddenRowsCache>
    if filterHiddenRowsCacheBySource.TryGetValue(key, &cached) then
        filterHiddenRowsCache
        |> List.filter (filterHiddenRowsCacheMatches cells rowCount >> not)
    else
        filterHiddenRowsCache

let private removeFilterHiddenRowsCacheForSource cells rowCount =
    let entries = filterHiddenRowsCacheWithoutSource cells rowCount
    if not (obj.ReferenceEquals(entries, filterHiddenRowsCache)) then
        setFilterHiddenRowsCache entries

let private filterHiddenRowsCacheRetainable (cached: FilterHiddenRowsCache) =
    cached.Entries.Count > 0
    && cached.TotalRowCount <= filterHiddenRowsCacheMaxTotalRowCount
    && cached.CellCount <= filterHiddenRowsCacheMaxTotalCells
    && cached.SourceTextCharCount <= filterHiddenRowsCacheMaxTotalSourceTextChars
    && cached.TotalFilterValueCount <= filterHiddenRowsCacheMaxTotalFilterValues
    && cached.TotalFilterTextCharCount <= filterHiddenRowsCacheMaxTotalFilterTextChars

let private filterHiddenRowsSourceTextCharCountFor rowCount filterValues filterTextChars cells =
    if rowCount > filterHiddenRowsCacheMaxRowCount
       || rowCount > filterHiddenRowsCacheMaxTotalRowCount
       || filterValues > filterHiddenRowsCacheMaxFilterValues
       || filterTextChars > filterHiddenRowsCacheMaxFilterTextChars then
        filterHiddenRowsCacheMaxTotalSourceTextChars + 1
    else
        cellMapSourceTextCharCountFor
            filterHiddenRowsCacheMaxTotalCells
            filterHiddenRowsCacheMaxTotalSourceTextChars
            cells

let private rememberFilterHiddenRowsCache cached =
    if not (filterHiddenRowsCacheRetainable cached) then
        removeFilterHiddenRowsCacheForSource cached.Cells cached.RowCount
    else
        let mutable keptCount = 1
        let mutable keptRows = cached.TotalRowCount
        let mutable keptCells = max 1 cached.CellCount
        let mutable keptSourceTextChars = max 1 cached.SourceTextCharCount
        let mutable keptFilterValues = cached.TotalFilterValueCount
        let mutable keptFilterTextChars = cached.TotalFilterTextCharCount
        let retained =
            filterHiddenRowsCacheWithoutSource cached.Cells cached.RowCount
            |> List.choose (fun candidate ->
                let candidateRows = candidate.TotalRowCount
                let candidateCells = max 1 candidate.CellCount
                let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                let candidateFilterValues = candidate.TotalFilterValueCount
                let candidateFilterTextChars = candidate.TotalFilterTextCharCount
                if filterHiddenRowsCacheRetainable candidate
                   && keptCount < filterHiddenRowsCacheMaxSources
                   && keptRows + candidateRows <= filterHiddenRowsCacheMaxTotalRowCount
                   && keptCells + candidateCells <= filterHiddenRowsCacheMaxTotalCells
                   && keptSourceTextChars + candidateSourceTextChars <= filterHiddenRowsCacheMaxTotalSourceTextChars
                   && keptFilterValues + candidateFilterValues <= filterHiddenRowsCacheMaxTotalFilterValues
                   && keptFilterTextChars + candidateFilterTextChars <= filterHiddenRowsCacheMaxTotalFilterTextChars then
                    keptCount <- keptCount + 1
                    keptRows <- keptRows + candidateRows
                    keptCells <- keptCells + candidateCells
                    keptSourceTextChars <- keptSourceTextChars + candidateSourceTextChars
                    keptFilterValues <- keptFilterValues + candidateFilterValues
                    keptFilterTextChars <- keptFilterTextChars + candidateFilterTextChars
                    Some candidate
                else
                    None)
        cached :: retained
        |> setFilterHiddenRowsCache

let private removeFilterHiddenRowsCacheEntry filters (cache: FilterHiddenRowsCache) =
    let mutable existing = Unchecked.defaultof<FilterHiddenRowsCacheEntry>
    if cache.Entries.TryGetValue(filters, &existing) then
        cache.TotalRowCount <- max 0 (cache.TotalRowCount - existing.Rows.Length)
        cache.TotalFilterValueCount <- max 0 (cache.TotalFilterValueCount - existing.FilterValueCount)
        cache.TotalFilterTextCharCount <- max 0 (cache.TotalFilterTextCharCount - existing.FilterTextCharCount)
        cache.Entries.Remove filters |> ignore

let private addFilterHiddenRowsCacheEntryWithMetrics filters (rows: int array) filterValues filterTextChars (cache: FilterHiddenRowsCache) =
    let mutable existing = Unchecked.defaultof<FilterHiddenRowsCacheEntry>
    if rows.Length > filterHiddenRowsCacheMaxRowCount
       || filterValues > filterHiddenRowsCacheMaxFilterValues then
        removeFilterHiddenRowsCacheEntry filters cache
    elif filterTextChars > filterHiddenRowsCacheMaxFilterTextChars then
        removeFilterHiddenRowsCacheEntry filters cache
    else
        let alreadyCached = cache.Entries.TryGetValue(filters, &existing)
        if alreadyCached then
            cache.TotalRowCount <- max 0 (cache.TotalRowCount - existing.Rows.Length)
            cache.TotalFilterValueCount <- max 0 (cache.TotalFilterValueCount - existing.FilterValueCount)
            cache.TotalFilterTextCharCount <- max 0 (cache.TotalFilterTextCharCount - existing.FilterTextCharCount)
        else
            cache.Order.Enqueue filters
        cache.Entries.[filters] <- { Rows = rows; FilterValueCount = filterValues; FilterTextCharCount = filterTextChars }
        cache.TotalRowCount <- cache.TotalRowCount + rows.Length
        cache.TotalFilterValueCount <- cache.TotalFilterValueCount + filterValues
        cache.TotalFilterTextCharCount <- cache.TotalFilterTextCharCount + filterTextChars
        while (cache.Entries.Count > filterHiddenRowsCacheMaxEntries
               || cache.TotalRowCount > filterHiddenRowsCacheMaxTotalRowCount
               || cache.TotalFilterValueCount > filterHiddenRowsCacheMaxTotalFilterValues
               || cache.TotalFilterTextCharCount > filterHiddenRowsCacheMaxTotalFilterTextChars)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if oldest <> filters then
                removeFilterHiddenRowsCacheEntry oldest cache
        if cache.Entries.Count > filterHiddenRowsCacheMaxEntries
           || cache.TotalRowCount > filterHiddenRowsCacheMaxTotalRowCount
           || cache.TotalFilterValueCount > filterHiddenRowsCacheMaxTotalFilterValues
           || cache.TotalFilterTextCharCount > filterHiddenRowsCacheMaxTotalFilterTextChars then
            let current = cache.Entries.[filters]
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue filters
            cache.Entries.[filters] <- current
            cache.TotalRowCount <- current.Rows.Length
            cache.TotalFilterValueCount <- current.FilterValueCount
            cache.TotalFilterTextCharCount <- current.FilterTextCharCount

let private addFilterHiddenRowsCacheEntry filters (rows: int array) (cache: FilterHiddenRowsCache) =
    let filterValues = filterValueCount filters
    let filterTextChars =
        if rows.Length > filterHiddenRowsCacheMaxRowCount
           || filterValues > filterHiddenRowsCacheMaxFilterValues then
            filterHiddenRowsCacheMaxFilterTextChars + 1
        else
            filterTextCharCount filterHiddenRowsCacheMaxFilterTextChars filters
    addFilterHiddenRowsCacheEntryWithMetrics filters rows filterValues filterTextChars cache

let private filterHiddenRows (model: SpreadsheetModel) =
    match tryGetFilterHiddenRowsCache model.Cells model.RowCount with
    | ValueSome cached ->
        let mutable entry = Unchecked.defaultof<FilterHiddenRowsCacheEntry>
        if cached.Entries.TryGetValue(model.Filters, &entry) then
            entry.Rows
        else
            let rows = filterHiddenRowsUncached model
            addFilterHiddenRowsCacheEntry model.Filters rows cached
            rememberFilterHiddenRowsCache cached
            rows
    | ValueNone ->
        let rows = filterHiddenRowsUncached model
        let filterValues = filterValueCount model.Filters
        let filterTextChars =
            if rows.Length > filterHiddenRowsCacheMaxRowCount
               || filterValues > filterHiddenRowsCacheMaxFilterValues then
                filterHiddenRowsCacheMaxFilterTextChars + 1
            else
                filterTextCharCount filterHiddenRowsCacheMaxFilterTextChars model.Filters
        let entries = System.Collections.Generic.Dictionary<Map<int, ColumnFilter>, FilterHiddenRowsCacheEntry>()
        let order = System.Collections.Generic.Queue<Map<int, ColumnFilter>>()
        let cache = {
            Cells = model.Cells
            CellCount = model.Cells.Count
            SourceTextCharCount =
                filterHiddenRowsSourceTextCharCountFor
                    rows.Length
                    filterValues
                    filterTextChars
                    model.Cells
            RowCount = model.RowCount
            Entries = entries
            Order = order
            TotalRowCount = 0
            TotalFilterValueCount = 0
            TotalFilterTextCharCount = 0
        }
        addFilterHiddenRowsCacheEntryWithMetrics model.Filters rows filterValues filterTextChars cache
        rememberFilterHiddenRowsCache
            cache
        rows

let private tryFilterHiddenRowsCachedMetrics cells rowCount filters =
    match tryGetFilterHiddenRowsCache cells rowCount with
    | ValueSome cached ->
        let mutable entry = Unchecked.defaultof<FilterHiddenRowsCacheEntry>
        if cached.Entries.TryGetValue(filters, &entry) then
            ValueSome (struct (cached.SourceTextCharCount, entry.FilterValueCount, entry.FilterTextCharCount))
        else
            ValueNone
    | ValueNone ->
        ValueNone

let private intArraysEqual (left: int array) (right: int array) =
    if left.Length <> right.Length then
        false
    else
        let mutable equal = true
        let mutable i = 0
        while equal && i < left.Length do
            if left.[i] <> right.[i] then
                equal <- false
            i <- i + 1
        equal

let private filterProjectionMatchesCurrent (model: SpreadsheetModel) =
    match tryGetFilterHiddenRowsCache model.Cells model.RowCount with
    | ValueSome cached ->
        let mutable entry = Unchecked.defaultof<FilterHiddenRowsCacheEntry>
        cached.Entries.TryGetValue(model.Filters, &entry)
            && (obj.ReferenceEquals(model.FilterHiddenRows, entry.Rows)
                || intArraysEqual model.FilterHiddenRows entry.Rows)
    | ValueNone -> false

/// Apply filter to a column
let applyFilter (col: int) (visibleValues: Set<string>) (model: SpreadsheetModel): SpreadsheetModel =
    let mutable existing = Unchecked.defaultof<ColumnFilter>
    if model.Filters.TryGetValue(col, &existing)
       && existing.Column = col
       && (obj.ReferenceEquals(existing.VisibleValues, visibleValues) || existing.VisibleValues = visibleValues)
       && (model.FilterHiddenRows.Length = 0 || filterProjectionMatchesCurrent model) then
        model
    else
        let filter = { Column = col; VisibleValues = visibleValues }
        { model with Filters = Map.add col filter model.Filters; FilterHiddenRows = [||] }

/// Remove filter from a column
let removeFilter (col: int) (model: SpreadsheetModel): SpreadsheetModel =
    if not (Map.containsKey col model.Filters) && model.FilterHiddenRows.Length = 0 then
        model
    else
        { model with Filters = Map.remove col model.Filters; FilterHiddenRows = [||] }

let private collectVisibleRowsArray (model: SpreadsheetModel) =
    if Map.isEmpty model.Filters then
        collectRowsWithCellsArray model
    else
        let filterEntries = model.Filters |> Map.toArray
        let filterCols = filterEntries |> Array.map fst
        let filterValues = filterEntries |> Array.map (fun (_, filter) -> filter.VisibleValues)
        let visible = ResizeArray<int>()
        if filterCols.Length <= 3 then
            let rowValues = Array.create filterCols.Length CellValue.Empty
            let mutable currentRow = System.Int32.MinValue
            let mutable hasRow = false

            let finishRow () =
                if hasRow && rowValuesPassFilters rowValues filterValues then
                    visible.Add currentRow

            model.Cells
            |> Map.iter (fun addr cellData ->
                if addr.Row <> currentRow then
                    finishRow()
                    currentRow <- addr.Row
                    hasRow <- true
                    Array.fill rowValues 0 rowValues.Length CellValue.Empty
                let filterIndex = filterColumnIndex filterCols addr.Col
                if filterIndex >= 0 then
                    rowValues.[filterIndex] <- cellData.Value)

            finishRow()
        else
            let rowValues = Array.create filterCols.Length CellValue.Empty
            let rowVersions = Array.zeroCreate<int> filterCols.Length
            let touchedIndexes = ResizeArray<int>()
            let emptyAllowed, emptyRejectedCount = filterEmptyAllowances filterValues
            let mutable currentRow = System.Int32.MinValue
            let mutable currentRowVersion = 0
            let mutable hasRow = false
            let mutable missingRejectedCount = emptyRejectedCount
            let contiguousFilters = filterColumnsAreContiguous filterCols
            let firstFilterCol = filterCols.[0]
            let lastFilterCol = filterCols.[filterCols.Length - 1]

            let finishRow () =
                if hasRow && touchedRowValuesPassFilters rowValues filterValues touchedIndexes missingRejectedCount then
                    visible.Add currentRow

            model.Cells
            |> Map.iter (fun addr cellData ->
                if addr.Row <> currentRow then
                    finishRow()
                    currentRow <- addr.Row
                    currentRowVersion <- currentRowVersion + 1
                    touchedIndexes.Clear()
                    missingRejectedCount <- emptyRejectedCount
                    hasRow <- true
                let filterIndex = filterColumnIndexFor contiguousFilters firstFilterCol lastFilterCol filterCols addr.Col
                if filterIndex >= 0 then
                    if rowVersions.[filterIndex] <> currentRowVersion then
                        rowVersions.[filterIndex] <- currentRowVersion
                        touchedIndexes.Add filterIndex
                        if not emptyAllowed.[filterIndex] then
                            missingRejectedCount <- missingRejectedCount - 1
                    rowValues.[filterIndex] <- cellData.Value
                )

            finishRow()
        visible.ToArray()

type private VisibleRowsCache = {
    Cells: Map<CellAddress, CellData>
    CellCount: int
    SourceTextCharCount: int
    Entries: System.Collections.Generic.Dictionary<Map<int, ColumnFilter>, VisibleRowsCacheEntry>
    Order: System.Collections.Generic.Queue<Map<int, ColumnFilter>>
    mutable TotalRowCount: int
    mutable TotalFilterValueCount: int
    mutable TotalFilterTextCharCount: int
}

and private VisibleRowsCacheEntry = {
    Rows: Set<int>
    FilterValueCount: int
    FilterTextCharCount: int
}

let mutable private visibleRowsCache: VisibleRowsCache list = []
let private visibleRowsCacheByCells =
    System.Collections.Generic.Dictionary<Map<CellAddress, CellData>, VisibleRowsCache>(cellMapReferenceComparer)
let private visibleRowsFastSlots = Array.zeroCreate<VisibleRowsCache> 256
let private visibleRowsCacheMaxEntries = 64
let private visibleRowsCacheMaxSources = 64
let private visibleRowsCacheMaxRowCount = 200000
let private visibleRowsCacheMaxTotalRowCount = 200000
let private visibleRowsCacheMaxTotalCells = 500000
let private visibleRowsCacheMaxTotalSourceTextChars = 5000000
let private visibleRowsCacheMaxFilterValues = 100000
let private visibleRowsCacheMaxTotalFilterValues = 200000
let private visibleRowsCacheMaxFilterTextChars = 5000000
let private visibleRowsCacheMaxTotalFilterTextChars = 5000000

let private visibleRowsCacheMatches cells (cached: VisibleRowsCache) =
    obj.ReferenceEquals(cached.Cells, cells)

let private setVisibleRowsCache entries =
    visibleRowsCache <- entries
    visibleRowsCacheByCells.Clear()
    clearArray visibleRowsFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<VisibleRowsCache>
        if not (visibleRowsCacheByCells.TryGetValue(entry.Cells, &existing)) then
            visibleRowsCacheByCells.[entry.Cells] <- entry
            let slot = referenceSlot entry.Cells visibleRowsFastSlots.Length
            visibleRowsFastSlots.[slot] <- entry

let private tryGetVisibleRowsCache cells =
    let slot = referenceSlot cells visibleRowsFastSlots.Length
    let fast = visibleRowsFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && visibleRowsCacheMatches cells fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<VisibleRowsCache>
        if visibleRowsCacheByCells.TryGetValue(cells, &cached) then
            visibleRowsFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private visibleRowsCacheWithoutCells cells =
    let mutable cached = Unchecked.defaultof<VisibleRowsCache>
    if visibleRowsCacheByCells.TryGetValue(cells, &cached) then
        visibleRowsCache
        |> List.filter (visibleRowsCacheMatches cells >> not)
    else
        visibleRowsCache

let private removeVisibleRowsCacheForCells cells =
    let entries = visibleRowsCacheWithoutCells cells
    if not (obj.ReferenceEquals(entries, visibleRowsCache)) then
        setVisibleRowsCache entries

let private visibleRowsCacheRetainable (cached: VisibleRowsCache) =
    cached.Entries.Count > 0
    && cached.TotalRowCount <= visibleRowsCacheMaxTotalRowCount
    && cached.CellCount <= visibleRowsCacheMaxTotalCells
    && cached.SourceTextCharCount <= visibleRowsCacheMaxTotalSourceTextChars
    && cached.TotalFilterValueCount <= visibleRowsCacheMaxTotalFilterValues
    && cached.TotalFilterTextCharCount <= visibleRowsCacheMaxTotalFilterTextChars

let private visibleRowsSourceTextCharCountFor rowCount filterValues filterTextChars cells =
    if rowCount > visibleRowsCacheMaxRowCount
       || rowCount > visibleRowsCacheMaxTotalRowCount
       || filterValues > visibleRowsCacheMaxFilterValues
       || filterTextChars > visibleRowsCacheMaxFilterTextChars then
        visibleRowsCacheMaxTotalSourceTextChars + 1
    else
        cellMapSourceTextCharCountFor
            visibleRowsCacheMaxTotalCells
            visibleRowsCacheMaxTotalSourceTextChars
            cells

let private visibleRowsCacheMetricsFor rowCount cells sourceRowCount filters =
    match tryFilterHiddenRowsCachedMetrics cells sourceRowCount filters with
    | ValueSome (struct (cachedSourceTextChars, filterValues, filterTextChars)) ->
        let sourceTextChars =
            if rowCount > visibleRowsCacheMaxRowCount
               || rowCount > visibleRowsCacheMaxTotalRowCount
               || filterValues > visibleRowsCacheMaxFilterValues
               || filterTextChars > visibleRowsCacheMaxFilterTextChars then
                visibleRowsCacheMaxTotalSourceTextChars + 1
            elif cachedSourceTextChars <= visibleRowsCacheMaxTotalSourceTextChars then
                cachedSourceTextChars
            else
                visibleRowsSourceTextCharCountFor rowCount filterValues filterTextChars cells
        struct (filterValues, filterTextChars, sourceTextChars)
    | ValueNone ->
        let filterValues = filterValueCount filters
        let filterTextChars =
            if rowCount > visibleRowsCacheMaxRowCount
               || filterValues > visibleRowsCacheMaxFilterValues then
                visibleRowsCacheMaxFilterTextChars + 1
            else
                filterTextCharCount visibleRowsCacheMaxFilterTextChars filters
        let sourceTextChars =
            visibleRowsSourceTextCharCountFor rowCount filterValues filterTextChars cells
        struct (filterValues, filterTextChars, sourceTextChars)

let private rememberVisibleRowsCache cached =
    if not (visibleRowsCacheRetainable cached) then
        removeVisibleRowsCacheForCells cached.Cells
    else
        let mutable keptCount = 1
        let mutable keptRows = cached.TotalRowCount
        let mutable keptCells = max 1 cached.CellCount
        let mutable keptSourceTextChars = max 1 cached.SourceTextCharCount
        let mutable keptFilterValues = cached.TotalFilterValueCount
        let mutable keptFilterTextChars = cached.TotalFilterTextCharCount
        let retained =
            visibleRowsCacheWithoutCells cached.Cells
            |> List.choose (fun candidate ->
                let candidateRows = candidate.TotalRowCount
                let candidateCells = max 1 candidate.CellCount
                let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                let candidateFilterValues = candidate.TotalFilterValueCount
                let candidateFilterTextChars = candidate.TotalFilterTextCharCount
                if visibleRowsCacheRetainable candidate
                   && keptCount < visibleRowsCacheMaxSources
                   && keptRows + candidateRows <= visibleRowsCacheMaxTotalRowCount
                   && keptCells + candidateCells <= visibleRowsCacheMaxTotalCells
                   && keptSourceTextChars + candidateSourceTextChars <= visibleRowsCacheMaxTotalSourceTextChars
                   && keptFilterValues + candidateFilterValues <= visibleRowsCacheMaxTotalFilterValues
                   && keptFilterTextChars + candidateFilterTextChars <= visibleRowsCacheMaxTotalFilterTextChars then
                    keptCount <- keptCount + 1
                    keptRows <- keptRows + candidateRows
                    keptCells <- keptCells + candidateCells
                    keptSourceTextChars <- keptSourceTextChars + candidateSourceTextChars
                    keptFilterValues <- keptFilterValues + candidateFilterValues
                    keptFilterTextChars <- keptFilterTextChars + candidateFilterTextChars
                    Some candidate
                else
                    None)
        cached :: retained
        |> setVisibleRowsCache

let private removeVisibleRowsCacheEntry filters (cache: VisibleRowsCache) =
    let mutable existing = Unchecked.defaultof<VisibleRowsCacheEntry>
    if cache.Entries.TryGetValue(filters, &existing) then
        cache.TotalRowCount <- max 0 (cache.TotalRowCount - existing.Rows.Count)
        cache.TotalFilterValueCount <- max 0 (cache.TotalFilterValueCount - existing.FilterValueCount)
        cache.TotalFilterTextCharCount <- max 0 (cache.TotalFilterTextCharCount - existing.FilterTextCharCount)
        cache.Entries.Remove filters |> ignore

let private addVisibleRowsCacheEntryWithMetrics filters (rows: Set<int>) filterValues filterTextChars (cache: VisibleRowsCache) =
    let mutable existing = Unchecked.defaultof<VisibleRowsCacheEntry>
    if rows.Count > visibleRowsCacheMaxRowCount
       || filterValues > visibleRowsCacheMaxFilterValues then
        removeVisibleRowsCacheEntry filters cache
    elif filterTextChars > visibleRowsCacheMaxFilterTextChars then
        removeVisibleRowsCacheEntry filters cache
    else
        let alreadyCached = cache.Entries.TryGetValue(filters, &existing)
        if alreadyCached then
            cache.TotalRowCount <- max 0 (cache.TotalRowCount - existing.Rows.Count)
            cache.TotalFilterValueCount <- max 0 (cache.TotalFilterValueCount - existing.FilterValueCount)
            cache.TotalFilterTextCharCount <- max 0 (cache.TotalFilterTextCharCount - existing.FilterTextCharCount)
        else
            cache.Order.Enqueue filters
        cache.Entries.[filters] <- { Rows = rows; FilterValueCount = filterValues; FilterTextCharCount = filterTextChars }
        cache.TotalRowCount <- cache.TotalRowCount + rows.Count
        cache.TotalFilterValueCount <- cache.TotalFilterValueCount + filterValues
        cache.TotalFilterTextCharCount <- cache.TotalFilterTextCharCount + filterTextChars
        while (cache.Entries.Count > visibleRowsCacheMaxEntries
               || cache.TotalRowCount > visibleRowsCacheMaxTotalRowCount
               || cache.TotalFilterValueCount > visibleRowsCacheMaxTotalFilterValues
               || cache.TotalFilterTextCharCount > visibleRowsCacheMaxTotalFilterTextChars)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if oldest <> filters then
                removeVisibleRowsCacheEntry oldest cache
        if cache.Entries.Count > visibleRowsCacheMaxEntries
           || cache.TotalRowCount > visibleRowsCacheMaxTotalRowCount
           || cache.TotalFilterValueCount > visibleRowsCacheMaxTotalFilterValues
           || cache.TotalFilterTextCharCount > visibleRowsCacheMaxTotalFilterTextChars then
            let current = cache.Entries.[filters]
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue filters
            cache.Entries.[filters] <- current
            cache.TotalRowCount <- current.Rows.Count
            cache.TotalFilterValueCount <- current.FilterValueCount
            cache.TotalFilterTextCharCount <- current.FilterTextCharCount

let private addVisibleRowsCacheEntry filters (rows: Set<int>) (cache: VisibleRowsCache) =
    let filterValues = filterValueCount filters
    let filterTextChars =
        if rows.Count > visibleRowsCacheMaxRowCount
           || filterValues > visibleRowsCacheMaxFilterValues then
            visibleRowsCacheMaxFilterTextChars + 1
        else
            filterTextCharCount visibleRowsCacheMaxFilterTextChars filters
    addVisibleRowsCacheEntryWithMetrics filters rows filterValues filterTextChars cache

let private tryVisibleRowsCached filters cells =
    match tryGetVisibleRowsCache cells with
    | ValueSome cached ->
        let mutable entry = Unchecked.defaultof<VisibleRowsCacheEntry>
        if cached.Entries.TryGetValue(filters, &entry) then ValueSome entry.Rows else ValueNone
    | ValueNone -> ValueNone

let private rememberVisibleRows (rows: Set<int>) (model: SpreadsheetModel) =
    match tryGetVisibleRowsCache model.Cells with
    | ValueSome cached ->
        let mutable entry = Unchecked.defaultof<VisibleRowsCacheEntry>
        if not (cached.Entries.TryGetValue(model.Filters, &entry)) then
            let struct (filterValues, filterTextChars, _) =
                visibleRowsCacheMetricsFor rows.Count model.Cells model.RowCount model.Filters
            addVisibleRowsCacheEntryWithMetrics model.Filters rows filterValues filterTextChars cached
            rememberVisibleRowsCache cached
    | ValueNone ->
        let struct (filterValues, filterTextChars, sourceTextCharCount) =
            visibleRowsCacheMetricsFor rows.Count model.Cells model.RowCount model.Filters
        let entries = System.Collections.Generic.Dictionary<Map<int, ColumnFilter>, VisibleRowsCacheEntry>()
        let order = System.Collections.Generic.Queue<Map<int, ColumnFilter>>()
        let cache = {
            Cells = model.Cells
            CellCount = model.Cells.Count
            SourceTextCharCount = sourceTextCharCount
            Entries = entries
            Order = order
            TotalRowCount = 0
            TotalFilterValueCount = 0
            TotalFilterTextCharCount = 0
        }
        addVisibleRowsCacheEntryWithMetrics model.Filters rows filterValues filterTextChars cache
        rememberVisibleRowsCache cache

let isFilterHiddenRow (row: int) (model: SpreadsheetModel): bool =
    not (Map.isEmpty model.Filters) && binarySearchInt model.FilterHiddenRows row >= 0

let private collectVisibleRowsHashSetUncached (model: SpreadsheetModel) =
    if Map.isEmpty model.Filters then
        System.Collections.Generic.HashSet<int>(collectRowsWithCellsArray model)
    else
        let visible = System.Collections.Generic.HashSet<int>()
        let filterEntries = model.Filters |> Map.toArray
        let filterCols = filterEntries |> Array.map fst
        let filterValues = filterEntries |> Array.map (fun (_, filter) -> filter.VisibleValues)
        if filterCols.Length <= 3 then
            let rowValues = Array.create filterCols.Length CellValue.Empty
            let mutable currentRow = System.Int32.MinValue
            let mutable hasRow = false

            let finishRow () =
                if hasRow && rowValuesPassFilters rowValues filterValues then
                    visible.Add currentRow |> ignore

            model.Cells
            |> Map.iter (fun addr cellData ->
                if addr.Row <> currentRow then
                    finishRow()
                    currentRow <- addr.Row
                    hasRow <- true
                    Array.fill rowValues 0 rowValues.Length CellValue.Empty
                let filterIndex = filterColumnIndex filterCols addr.Col
                if filterIndex >= 0 then
                    rowValues.[filterIndex] <- cellData.Value)

            finishRow()
        else
            let rowValues = Array.create filterCols.Length CellValue.Empty
            let rowVersions = Array.zeroCreate<int> filterCols.Length
            let touchedIndexes = ResizeArray<int>()
            let emptyAllowed, emptyRejectedCount = filterEmptyAllowances filterValues
            let mutable currentRow = System.Int32.MinValue
            let mutable currentRowVersion = 0
            let mutable hasRow = false
            let mutable missingRejectedCount = emptyRejectedCount
            let contiguousFilters = filterColumnsAreContiguous filterCols
            let firstFilterCol = filterCols.[0]
            let lastFilterCol = filterCols.[filterCols.Length - 1]

            let finishRow () =
                if hasRow && touchedRowValuesPassFilters rowValues filterValues touchedIndexes missingRejectedCount then
                    visible.Add currentRow |> ignore

            model.Cells
            |> Map.iter (fun addr cellData ->
                if addr.Row <> currentRow then
                    finishRow()
                    currentRow <- addr.Row
                    currentRowVersion <- currentRowVersion + 1
                    touchedIndexes.Clear()
                    missingRejectedCount <- emptyRejectedCount
                    hasRow <- true
                let filterIndex = filterColumnIndexFor contiguousFilters firstFilterCol lastFilterCol filterCols addr.Col
                if filterIndex >= 0 then
                    if rowVersions.[filterIndex] <> currentRowVersion then
                        rowVersions.[filterIndex] <- currentRowVersion
                        touchedIndexes.Add filterIndex
                        if not emptyAllowed.[filterIndex] then
                            missingRejectedCount <- missingRejectedCount - 1
                    rowValues.[filterIndex] <- cellData.Value
                )

            finishRow()
        visible

/// Get visible rows as a hash set for repeated membership tests.
let getVisibleRowsHashSet (model: SpreadsheetModel): System.Collections.Generic.HashSet<int> =
    if Map.isEmpty model.Filters then
        collectVisibleRowsHashSetUncached model
    else
        match tryVisibleRowsCached model.Filters model.Cells with
        | ValueSome rows ->
            System.Collections.Generic.HashSet<int>(Set.toArray rows)
        | ValueNone ->
            collectVisibleRowsHashSetUncached model

/// Get visible rows based on all active filters
let getVisibleRows (model: SpreadsheetModel): Set<int> =
    match tryGetVisibleRowsCache model.Cells with
    | ValueSome cached ->
        let mutable entry = Unchecked.defaultof<VisibleRowsCacheEntry>
        if cached.Entries.TryGetValue(model.Filters, &entry) then
            entry.Rows
        else
            let rows = collectVisibleRowsArray model |> Set.ofArray
            rememberVisibleRows rows model
            rows
    | ValueNone ->
        let rows = collectVisibleRowsArray model |> Set.ofArray
        rememberVisibleRows rows model
        rows

/// Ensure the filter-hidden row projection exists for callers that constructed
/// or modified Filters directly instead of going through applyFilter/removeFilter.
let applyFilterRowHeights (model: SpreadsheetModel): SpreadsheetModel =
    if Map.isEmpty model.Filters then
        if model.FilterHiddenRows.Length = 0 then model else { model with FilterHiddenRows = [||] }
    else
        let hiddenRows = filterHiddenRows model
        if obj.ReferenceEquals(model.FilterHiddenRows, hiddenRows) || intArraysEqual model.FilterHiddenRows hiddenRows then model
        else { model with FilterHiddenRows = hiddenRows }

/// Distinct, sorted display strings in a column — the AutoFilter checklist's
/// values. Blanks excluded; uses the same `displayValue` as the filter test.
let columnUniqueValues (col: int) (model: SpreadsheetModel) : string list =
    columnUniqueValuesCached false col model

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
    { m with Tables = appendItem table m.Tables }

/// The table whose region contains the 0-indexed cell, if any. Table
/// coordinates are 1-indexed, so they're shifted for the comparison.
let tableAt (row: int) (col: int) (model: SpreadsheetModel): Table option =
    let mutable found = None
    let mutable remaining = model.Tables
    while found.IsNone && not remaining.IsEmpty do
        match remaining with
        | table :: tail ->
            if row >= table.StartRow - 1
               && row <= table.EndRow - 1
               && col >= table.StartCol - 1
               && col <= table.EndCol - 1 then
                found <- Some table
            remaining <- tail
        | [] -> ()
    found

/// Distinct, sorted display values in one column of a table — data rows only
/// (header and totals excluded). The checklist for a table's filter button.
type private TableColumnUniqueValuesCache = {
    Cells: Map<CellAddress, CellData>
    CellCount: int
    SourceTextCharCount: int
    Entries: System.Collections.Generic.Dictionary<TableColumnUniqueValuesKey, TableColumnUniqueValuesCacheEntry>
    Order: System.Collections.Generic.Queue<TableColumnUniqueValuesKey>
    mutable TotalValueCount: int
    mutable TotalValueTextCharCount: int
}

and private TableColumnUniqueValuesCacheEntry = {
    Values: string list
    ValueCount: int
    ValueTextCharCount: int
}

and private TableColumnUniqueValuesKey = {
    StartRow: int
    EndRow: int
    HeaderRowCount: int
    HasTotals: bool
    Column: int
}

let mutable private tableColumnUniqueValuesCache: TableColumnUniqueValuesCache list = []
let private tableColumnUniqueValuesCacheByCells =
    System.Collections.Generic.Dictionary<Map<CellAddress, CellData>, TableColumnUniqueValuesCache>(cellMapReferenceComparer)
let private tableColumnUniqueValuesFastSlots = Array.zeroCreate<TableColumnUniqueValuesCache> 256
let private tableColumnUniqueValuesCacheMaxEntries = 64
let private tableColumnUniqueValuesCacheMaxSources = 64
let private tableColumnUniqueValuesCacheMaxTotalValues = 200000
let private tableColumnUniqueValuesCacheMaxTotalValueTextChars = 5000000
let private tableColumnUniqueValuesCacheMaxTotalCells = 500000
let private tableColumnUniqueValuesCacheMaxTotalSourceTextChars = 5000000

let private tableColumnUniqueValuesCacheMatches cells (cached: TableColumnUniqueValuesCache) =
    obj.ReferenceEquals(cached.Cells, cells)

let private setTableColumnUniqueValuesCache entries =
    tableColumnUniqueValuesCache <- entries
    tableColumnUniqueValuesCacheByCells.Clear()
    clearArray tableColumnUniqueValuesFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<TableColumnUniqueValuesCache>
        if not (tableColumnUniqueValuesCacheByCells.TryGetValue(entry.Cells, &existing)) then
            tableColumnUniqueValuesCacheByCells.[entry.Cells] <- entry
            let slot = referenceSlot entry.Cells tableColumnUniqueValuesFastSlots.Length
            tableColumnUniqueValuesFastSlots.[slot] <- entry

let private tryGetTableColumnUniqueValuesCache cells =
    let slot = referenceSlot cells tableColumnUniqueValuesFastSlots.Length
    let fast = tableColumnUniqueValuesFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && tableColumnUniqueValuesCacheMatches cells fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<TableColumnUniqueValuesCache>
        if tableColumnUniqueValuesCacheByCells.TryGetValue(cells, &cached) then
            tableColumnUniqueValuesFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private tableColumnUniqueValuesCacheWithoutCells cells =
    let mutable cached = Unchecked.defaultof<TableColumnUniqueValuesCache>
    if tableColumnUniqueValuesCacheByCells.TryGetValue(cells, &cached) then
        tableColumnUniqueValuesCache
        |> List.filter (tableColumnUniqueValuesCacheMatches cells >> not)
    else
        tableColumnUniqueValuesCache

let private removeTableColumnUniqueValuesCacheForCells cells =
    let entries = tableColumnUniqueValuesCacheWithoutCells cells
    if not (obj.ReferenceEquals(entries, tableColumnUniqueValuesCache)) then
        setTableColumnUniqueValuesCache entries

let private tableColumnUniqueValuesCacheRetainable (cached: TableColumnUniqueValuesCache) =
    cached.Entries.Count > 0
    && cached.TotalValueCount <= tableColumnUniqueValuesCacheMaxTotalValues
    && cached.TotalValueTextCharCount <= tableColumnUniqueValuesCacheMaxTotalValueTextChars
    && cached.CellCount <= tableColumnUniqueValuesCacheMaxTotalCells
    && cached.SourceTextCharCount <= tableColumnUniqueValuesCacheMaxTotalSourceTextChars

let private tableColumnUniqueValuesSourceTextCharCountFor valueCount valueTextCharCount cells =
    if valueCount > tableColumnUniqueValuesCacheMaxTotalValues
       || valueTextCharCount > tableColumnUniqueValuesCacheMaxTotalValueTextChars then
        tableColumnUniqueValuesCacheMaxTotalSourceTextChars + 1
    else
        cellMapSourceTextCharCountFor
            tableColumnUniqueValuesCacheMaxTotalCells
            tableColumnUniqueValuesCacheMaxTotalSourceTextChars
            cells

let private rememberTableColumnUniqueValuesCache cached =
    if not (tableColumnUniqueValuesCacheRetainable cached) then
        removeTableColumnUniqueValuesCacheForCells cached.Cells
    else
        let mutable keptCount = 1
        let mutable keptValues = cached.TotalValueCount
        let mutable keptValueTextChars = cached.TotalValueTextCharCount
        let mutable keptCells = cached.CellCount
        let mutable keptSourceTextChars = max 1 cached.SourceTextCharCount
        let retained =
            tableColumnUniqueValuesCacheWithoutCells cached.Cells
            |> List.choose (fun candidate ->
                let candidateValues = candidate.TotalValueCount
                let candidateValueTextChars = max 1 candidate.TotalValueTextCharCount
                let candidateCells = candidate.CellCount
                let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                if tableColumnUniqueValuesCacheRetainable candidate
                   && keptCount < tableColumnUniqueValuesCacheMaxSources
                   && keptValues + candidateValues <= tableColumnUniqueValuesCacheMaxTotalValues
                   && keptValueTextChars + candidateValueTextChars <= tableColumnUniqueValuesCacheMaxTotalValueTextChars
                   && keptCells + candidateCells <= tableColumnUniqueValuesCacheMaxTotalCells
                   && keptSourceTextChars + candidateSourceTextChars <= tableColumnUniqueValuesCacheMaxTotalSourceTextChars then
                    keptCount <- keptCount + 1
                    keptValues <- keptValues + candidateValues
                    keptValueTextChars <- keptValueTextChars + candidateValueTextChars
                    keptCells <- keptCells + candidateCells
                    keptSourceTextChars <- keptSourceTextChars + candidateSourceTextChars
                    Some candidate
                else
                    None)
        cached :: retained
        |> setTableColumnUniqueValuesCache

let private removeTableColumnUniqueValuesCacheEntry key (cache: TableColumnUniqueValuesCache) =
    let mutable existing = Unchecked.defaultof<TableColumnUniqueValuesCacheEntry>
    if cache.Entries.TryGetValue(key, &existing) then
        cache.TotalValueCount <- max 0 (cache.TotalValueCount - existing.ValueCount)
        cache.TotalValueTextCharCount <- max 0 (cache.TotalValueTextCharCount - existing.ValueTextCharCount)
        cache.Entries.Remove key |> ignore

let private addTableColumnUniqueValuesCacheEntry key entry (cache: TableColumnUniqueValuesCache) =
    let mutable existing = Unchecked.defaultof<TableColumnUniqueValuesCacheEntry>
    if entry.ValueCount > tableColumnUniqueValuesCacheMaxTotalValues
       || entry.ValueTextCharCount > tableColumnUniqueValuesCacheMaxTotalValueTextChars then
        removeTableColumnUniqueValuesCacheEntry key cache
    else
        if cache.Entries.TryGetValue(key, &existing) then
            cache.TotalValueCount <- max 0 (cache.TotalValueCount - existing.ValueCount)
            cache.TotalValueTextCharCount <- max 0 (cache.TotalValueTextCharCount - existing.ValueTextCharCount)
        else
            cache.Order.Enqueue key
        cache.Entries.[key] <- entry
        cache.TotalValueCount <- cache.TotalValueCount + entry.ValueCount
        cache.TotalValueTextCharCount <- cache.TotalValueTextCharCount + entry.ValueTextCharCount
        while (cache.Entries.Count > tableColumnUniqueValuesCacheMaxEntries
               || cache.TotalValueCount > tableColumnUniqueValuesCacheMaxTotalValues
               || cache.TotalValueTextCharCount > tableColumnUniqueValuesCacheMaxTotalValueTextChars)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if oldest <> key then
                removeTableColumnUniqueValuesCacheEntry oldest cache
        if cache.Entries.Count > tableColumnUniqueValuesCacheMaxEntries
           || cache.TotalValueCount > tableColumnUniqueValuesCacheMaxTotalValues
           || cache.TotalValueTextCharCount > tableColumnUniqueValuesCacheMaxTotalValueTextChars then
            let current = cache.Entries.[key]
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue key
            cache.Entries.[key] <- current
            cache.TotalValueCount <- current.ValueCount
            cache.TotalValueTextCharCount <- current.ValueTextCharCount

let tableColumnUniqueValues (table: Table) (col: int) (model: SpreadsheetModel): string list =
    let key = {
        StartRow = table.StartRow
        EndRow = table.EndRow
        HeaderRowCount = table.HeaderRowCount
        HasTotals = table.HasTotals
        Column = col
    }
    match tryGetTableColumnUniqueValuesCache model.Cells with
    | ValueSome cached ->
        let mutable entry = Unchecked.defaultof<TableColumnUniqueValuesCacheEntry>
        if cached.Entries.TryGetValue(key, &entry) then
            entry.Values
        else
            let dataStart = (table.StartRow - 1) + table.HeaderRowCount
            let dataEnd = (table.EndRow - 1) - (if table.HasTotals then 1 else 0)
            let dataRowCount = max 0 (dataEnd - dataStart + 1)
            let values = System.Collections.Generic.HashSet<string>()
            let mutable errorDisplayCache = Unchecked.defaultof<System.Collections.Generic.Dictionary<string, string>>
            let displayCached value =
                match unwrapFormulaValue value with
                | Error code ->
                    if obj.ReferenceEquals(errorDisplayCache, null) then
                        errorDisplayCache <- System.Collections.Generic.Dictionary<string, string>()
                    let mutable cached = Unchecked.defaultof<string>
                    if errorDisplayCache.TryGetValue(code, &cached) then
                        cached
                    else
                        let text = "#" + code
                        errorDisplayCache.[code] <- text
                        text
                | other -> displayValue other
            let addValue value =
                match displayCached value with
                | "" -> ()
                | value -> values.Add value |> ignore
            if dataRowCount > model.Cells.Count / 2 then
                model.Cells
                |> Map.iter (fun addr cellData ->
                    if addr.Col = col && addr.Row >= dataStart && addr.Row <= dataEnd then
                        addValue cellData.Value)
            else
                for r in dataStart .. dataEnd do
                    addValue (getCellValue { Row = r; Col = col } model)
            let result, valueCount, valueTextCharCount =
                sortedStringSetValuesWithTextMetrics tableColumnUniqueValuesCacheMaxTotalValueTextChars values
            addTableColumnUniqueValuesCacheEntry key {
                Values = result
                ValueCount = valueCount
                ValueTextCharCount = valueTextCharCount
            } cached
            rememberTableColumnUniqueValuesCache cached
            result
    | ValueNone ->
        let dataStart = (table.StartRow - 1) + table.HeaderRowCount
        let dataEnd = (table.EndRow - 1) - (if table.HasTotals then 1 else 0)
        let dataRowCount = max 0 (dataEnd - dataStart + 1)
        let values = System.Collections.Generic.HashSet<string>()
        let mutable errorDisplayCache = Unchecked.defaultof<System.Collections.Generic.Dictionary<string, string>>
        let displayCached value =
            match unwrapFormulaValue value with
            | Error code ->
                if obj.ReferenceEquals(errorDisplayCache, null) then
                    errorDisplayCache <- System.Collections.Generic.Dictionary<string, string>()
                let mutable cached = Unchecked.defaultof<string>
                if errorDisplayCache.TryGetValue(code, &cached) then
                    cached
                else
                    let text = "#" + code
                    errorDisplayCache.[code] <- text
                    text
            | other -> displayValue other
        let addValue value =
            match displayCached value with
            | "" -> ()
            | value -> values.Add value |> ignore
        if dataRowCount > model.Cells.Count / 2 then
            model.Cells
            |> Map.iter (fun addr cellData ->
                if addr.Col = col && addr.Row >= dataStart && addr.Row <= dataEnd then
                    addValue cellData.Value)
        else
            for r in dataStart .. dataEnd do
                addValue (getCellValue { Row = r; Col = col } model)
        let result, valueCount, valueTextCharCount =
            sortedStringSetValuesWithTextMetrics tableColumnUniqueValuesCacheMaxTotalValueTextChars values
        let entries = System.Collections.Generic.Dictionary<TableColumnUniqueValuesKey, TableColumnUniqueValuesCacheEntry>()
        let order = System.Collections.Generic.Queue<TableColumnUniqueValuesKey>()
        let cache = {
            Cells = model.Cells
            CellCount = max 1 model.Cells.Count
            SourceTextCharCount =
                tableColumnUniqueValuesSourceTextCharCountFor
                    valueCount
                    valueTextCharCount
                    model.Cells
            Entries = entries
            Order = order
            TotalValueCount = 0
            TotalValueTextCharCount = 0
        }
        addTableColumnUniqueValuesCacheEntry key {
            Values = result
            ValueCount = valueCount
            ValueTextCharCount = valueTextCharCount
        } cache
        rememberTableColumnUniqueValuesCache
            cache
        result

/// Render-time: bake table styling (header accent fill + white bold text +
/// filter buttons, and banded data rows) into cell formats from `model.Tables`.
/// Acts ONLY on tables whose `Style.HeaderFill` is set — i.e. tables created in
/// the app (`createTable`); imported tables leave it None because the importer
/// already bakes theme-accent styling into each cell, so this never
/// double-styles or disturbs the parity corpus. Uses orElse/|| merges so a
/// cell's own explicit colors still win. Empty cells in the region are created
/// styled so the band/header spans the full width.
type private TableStyleCellsCacheEntry = {
    SourceCells: Map<CellAddress, CellData>
    SourceTables: Table list
    StyledCells: Map<CellAddress, CellData>
    CellCount: int
    TableCount: int
    SourceTextCharCount: int
}

[<Struct>]
type private TableStyleCellsReferenceKey = {
    Cells: Map<CellAddress, CellData>
    Tables: Table list
}

let mutable private tableStyleCellsCache: TableStyleCellsCacheEntry list = []
let private tableStyleCellsCacheMaxEntries = 64
let private tableStyleCellsCacheMaxTotalCells = 200000
let private tableStyleCellsCacheMaxTotalTables = 100000
let private tableStyleCellsCacheMaxTotalSourceTextChars = 5000000

let private tableStyleCellsCacheEntryMatches cells tables (entry: TableStyleCellsCacheEntry) =
    obj.ReferenceEquals(entry.SourceCells, cells)
    && obj.ReferenceEquals(entry.SourceTables, tables)

let private tableStyleCellsReferenceHash cells tables =
    (referenceHash cells * 397) ^^^ referenceHash tables

let private tableStyleCellsReferenceKey cells tables : TableStyleCellsReferenceKey =
    { Cells = cells
      Tables = tables }

let private tableStyleCellsReferenceComparer =
    { new System.Collections.Generic.IEqualityComparer<TableStyleCellsReferenceKey> with
        member _.Equals(left, right) =
            obj.ReferenceEquals(left.Cells, right.Cells)
            && obj.ReferenceEquals(left.Tables, right.Tables)

        member _.GetHashCode(key) =
            tableStyleCellsReferenceHash key.Cells key.Tables }

let private tableStyleCellsCacheBySource =
    System.Collections.Generic.Dictionary<TableStyleCellsReferenceKey, TableStyleCellsCacheEntry>(tableStyleCellsReferenceComparer)
let private tableStyleCellsFastSlots = Array.zeroCreate<TableStyleCellsCacheEntry> 256

let private setTableStyleCellsCache entries =
    tableStyleCellsCache <- entries
    tableStyleCellsCacheBySource.Clear()
    clearArray tableStyleCellsFastSlots
    for entry in entries do
        let key = tableStyleCellsReferenceKey entry.SourceCells entry.SourceTables
        let mutable existing = Unchecked.defaultof<TableStyleCellsCacheEntry>
        if not (tableStyleCellsCacheBySource.TryGetValue(key, &existing)) then
            tableStyleCellsCacheBySource.[key] <- entry
            tableStyleCellsFastSlots.[tableStyleCellsReferenceHash entry.SourceCells entry.SourceTables &&& (tableStyleCellsFastSlots.Length - 1)] <- entry

let private tryGetTableStyleCellsEntry cells tables =
    let slot = tableStyleCellsReferenceHash cells tables &&& (tableStyleCellsFastSlots.Length - 1)
    let fast = tableStyleCellsFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && tableStyleCellsCacheEntryMatches cells tables fast then
        ValueSome fast
    else
        let key = tableStyleCellsReferenceKey cells tables
        let mutable cached = Unchecked.defaultof<TableStyleCellsCacheEntry>
        if tableStyleCellsCacheBySource.TryGetValue(key, &cached) then
            tableStyleCellsFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private tableStyleCellsCacheWithoutSource cells tables =
    let key = tableStyleCellsReferenceKey cells tables
    let mutable cached = Unchecked.defaultof<TableStyleCellsCacheEntry>
    if tableStyleCellsCacheBySource.TryGetValue(key, &cached)
       && tableStyleCellsCacheEntryMatches cells tables cached then
        tableStyleCellsCache
        |> List.filter (tableStyleCellsCacheEntryMatches cells tables >> not)
    else
        tableStyleCellsCache

let private retainedTableStyleCellCount (cells: Map<CellAddress, CellData>) (styledCells: Map<CellAddress, CellData>) =
    if obj.ReferenceEquals(cells, styledCells) then styledCells.Count
    else cells.Count + styledCells.Count

let private trimTableStyleCellsCache entries =
    let rec loop remaining kept count totalCells totalTables totalSourceTextChars =
        match remaining with
        | [] -> List.rev kept
        | entry :: rest ->
            let nextTotalCells = totalCells + entry.CellCount
            let nextTotalTables = totalTables + entry.TableCount
            let candidateSourceTextChars = max 1 entry.SourceTextCharCount
            let nextTotalSourceTextChars = totalSourceTextChars + candidateSourceTextChars
            if count < tableStyleCellsCacheMaxEntries
                && nextTotalCells <= tableStyleCellsCacheMaxTotalCells
                && nextTotalTables <= tableStyleCellsCacheMaxTotalTables
                && nextTotalSourceTextChars <= tableStyleCellsCacheMaxTotalSourceTextChars then
                loop rest (entry :: kept) (count + 1) nextTotalCells nextTotalTables nextTotalSourceTextChars
            else
                loop rest kept count totalCells totalTables totalSourceTextChars
    loop entries [] 0 0 0 0

let private tableStyleCellsCacheEntryRetainable entry =
    entry.CellCount <= tableStyleCellsCacheMaxTotalCells
    && entry.TableCount <= tableStyleCellsCacheMaxTotalTables
    && entry.SourceTextCharCount <= tableStyleCellsCacheMaxTotalSourceTextChars

let private tableStyleCellsTableCount (tables: Table list) =
    let mutable count = 0
    let mutable remaining = tables
    while count <= tableStyleCellsCacheMaxTotalTables && not remaining.IsEmpty do
        count <- count + 1
        remaining <- remaining.Tail
    count

let private tableStyleCellsOptionStringCharCount (value: string option) =
    match value with
    | Some text -> text.Length
    | None -> 0

let private tableStyleCellsStyleStringCharCount (style: TableStyle option) =
    match style with
    | Some style ->
        style.Name.Length
        + tableStyleCellsOptionStringCharCount style.HeaderFill
        + tableStyleCellsOptionStringCharCount style.StripeFill
    | None -> 0

let private tableStyleCellsSourceTextCharCount (cells: Map<CellAddress, CellData>) (tables: Table list) =
    let limit = tableStyleCellsCacheMaxTotalSourceTextChars
    let cappedLimit = limit + 1
    let mutable total = 0
    let mutable remaining = tables
    while total <= limit && not remaining.IsEmpty do
        let table = remaining.Head
        total <- addCapped limit table.Name.Length total
        if total <= limit then
            total <- addCapped limit (tableStyleCellsStyleStringCharCount table.Style) total
        remaining <- remaining.Tail
    if total <= limit then
        total <-
            addCapped
                limit
                (cellMapSourceTextCharCountFor
                    tableStyleCellsCacheMaxTotalCells
                    limit
                    cells)
                total
    min total cappedLimit

let private tableStyleCellsSourceTextCharCountFor cellCount tableCount cells tables =
    if cellCount > tableStyleCellsCacheMaxTotalCells
       || tableCount > tableStyleCellsCacheMaxTotalTables then
        tableStyleCellsCacheMaxTotalSourceTextChars + 1
    else
        tableStyleCellsSourceTextCharCount cells tables

let private rememberTableStyleCellsEntry entry =
    (entry :: tableStyleCellsCacheWithoutSource entry.SourceCells entry.SourceTables)
    |> trimTableStyleCellsCache
    |> setTableStyleCellsCache

let private rememberTableStyleCells cells tables styledCells =
    let cellCount = retainedTableStyleCellCount cells styledCells
    let tableCount =
        if cellCount > tableStyleCellsCacheMaxTotalCells then
            tableStyleCellsCacheMaxTotalTables + 1
        else
            tableStyleCellsTableCount tables
    let entry = {
        SourceCells = cells
        SourceTables = tables
        StyledCells = styledCells
        CellCount = cellCount
        TableCount = tableCount
        SourceTextCharCount = tableStyleCellsSourceTextCharCountFor cellCount tableCount cells tables
    }
    if tableStyleCellsCacheEntryRetainable entry then
        rememberTableStyleCellsEntry entry

let applyTableStyles (model: SpreadsheetModel): SpreadsheetModel =
    match tryGetTableStyleCellsEntry model.Cells model.Tables with
    | ValueSome cached ->
        if obj.ReferenceEquals(cached.StyledCells, model.Cells) then model
        else { model with Cells = cached.StyledCells }
    | ValueNone ->
        let pending = System.Collections.Generic.Dictionary<CellAddress, CellData>()
        let mutable changed = false

        let currentCellData addr =
            let mutable pendingCell = Unchecked.defaultof<CellData>
            if pending.TryGetValue(addr, &pendingCell) then
                pendingCell
            else
                match Map.tryFind addr model.Cells with
                | Some cellData -> cellData
                | None -> { Value = Empty; Format = defaultFormat; Comment = None }

        let applyPending () =
            if not changed then
                model.Cells
            elif int64 pending.Count * 2L < int64 (max 1 model.Cells.Count) then
                let mutable cells = model.Cells
                for KeyValue(addr, cellData) in pending do
                    cells <- Map.add addr cellData cells
                cells
            else
                let entries = ResizeArray<CellAddress * CellData>(model.Cells.Count + pending.Count)
                for KeyValue(addr, cellData) in model.Cells do
                    entries.Add(addr, cellData)
                for KeyValue(addr, cellData) in pending do
                    entries.Add(addr, cellData)
                Map.ofArray (entries.ToArray())

        let styleTable (t: Table) =
            match t.Style with
            | Some s when s.HeaderFill.IsSome ->
                let r0, c0, r1, c1 = t.StartRow - 1, t.StartCol - 1, t.EndRow - 1, t.EndCol - 1
                for r in r0 .. r1 do
                    let isHeader = r < r0 + t.HeaderRowCount
                    let dataIdx = r - r0 - t.HeaderRowCount
                    let isStripe = s.ShowRowStripes && dataIdx >= 0 && dataIdx % 2 = 0 && s.StripeFill.IsSome
                    if isHeader || isStripe then
                        for c in c0 .. c1 do
                            let addr = { Row = r; Col = c }
                            let baseData = currentCellData addr
                            let fmt = baseData.Format
                            let newFmt =
                                if isHeader then
                                    let hasFilter =
                                        match t.AutoFilter with Some (fc1, fc2) -> (c + 1) >= fc1 && (c + 1) <= fc2 | None -> false
                                    { fmt with
                                        BgColor = fmt.BgColor |> Option.orElse s.HeaderFill
                                        TextColor = fmt.TextColor |> Option.orElse (Some "#FFFFFF")
                                        Bold = true
                                        FilterButton = fmt.FilterButton || hasFilter }
                                else
                                    { fmt with BgColor = fmt.BgColor |> Option.orElse s.StripeFill }
                            if newFmt <> fmt then
                                pending.[addr] <- { baseData with Format = newFmt }
                                changed <- true
            | _ -> ()

        for table in model.Tables do
            styleTable table

        let cells = applyPending ()
        rememberTableStyleCells model.Cells model.Tables cells

        if changed then { model with Cells = cells } else model

let private countCellsWhere predicate (cells: Map<CellAddress, CellData>) =
    let mutable count = 0
    for KeyValue(addr, _) in cells do
        if predicate addr then
            count <- count + 1
    count

let private sparseRewriteWorth affectedCount (cells: Map<CellAddress, CellData>) =
    int64 affectedCount * 2L < int64 cells.Count

let private rewriteSparseCells (cells: Map<CellAddress, CellData>) (removals: ResizeArray<CellAddress>) (additions: ResizeArray<CellAddress * CellData>) =
    let mutable rewritten = cells
    for addr in removals do
        rewritten <- Map.remove addr rewritten
    for addr, cellData in additions do
        rewritten <- Map.add addr cellData rewritten
    rewritten

/// Insert a row at the specified index (shifts rows down)
let insertRow (atRow: int) (model: SpreadsheetModel): SpreadsheetModel =
    if atRow < 0 || atRow > model.RowCount then
        model
    else
        let maxUsedRow =
            if Map.isEmpty model.Cells then
                None
            else
                let maxAddr, _ = Map.maxKeyValue model.Cells
                Some maxAddr.Row

        let hasAffectedCells =
            match maxUsedRow with
            | Some maxRow -> maxRow >= atRow
            | None -> false

        let newCells =
            if hasAffectedCells then
                let nearUsedTail =
                    match maxUsedRow with
                    | Some maxRow -> atRow > maxRow / 2
                    | None -> false
                if nearUsedTail then
                    let affectedCount = countCellsWhere (fun addr -> addr.Row >= atRow) model.Cells
                    if sparseRewriteWorth affectedCount model.Cells then
                        let removals = ResizeArray<CellAddress>(affectedCount)
                        let additions = ResizeArray<CellAddress * CellData>(affectedCount)
                        for KeyValue(addr, cellData) in model.Cells do
                            if addr.Row >= atRow then
                                removals.Add addr
                                additions.Add({ addr with Row = addr.Row + 1 }, cellData)
                        rewriteSparseCells model.Cells removals additions
                    else
                        // Shift all rows >= atRow down by 1
                        let mutable shifted = Map.empty
                        model.Cells
                        |> Map.iter (fun addr cellData ->
                            let newAddr =
                                if addr.Row >= atRow then { addr with Row = addr.Row + 1 }
                                else addr
                            shifted <- Map.add newAddr cellData shifted)
                        shifted
                else
                    // Shift all rows >= atRow down by 1
                    let mutable shifted = Map.empty
                    model.Cells
                    |> Map.iter (fun addr cellData ->
                        let newAddr =
                            if addr.Row >= atRow then { addr with Row = addr.Row + 1 }
                            else addr
                        shifted <- Map.add newAddr cellData shifted)
                    shifted
            else
                model.Cells

        // Update selection if affected
        let newSelection =
            match model.Selection with
            | SingleCell (row, col) ->
                if row >= atRow then SingleCell (row + 1, col)
                else model.Selection
            | Range (sr, sc, er, ec) ->
                let newSr = if sr >= atRow then sr + 1 else sr
                let newEr = if er >= atRow then er + 1 else er
                Range (newSr, sc, newEr, ec)

        { model with
            Cells = newCells
            Selection = newSelection
            RowCount = model.RowCount + 1
        }

/// Delete a row at the specified index (shifts rows up)
let deleteRow (atRow: int) (model: SpreadsheetModel): SpreadsheetModel =
    if atRow < 0 || atRow >= model.RowCount || model.RowCount <= 1 then
        model
    else
        let maxUsedRow =
            if Map.isEmpty model.Cells then
                None
            else
                let maxAddr, _ = Map.maxKeyValue model.Cells
                Some maxAddr.Row

        let hasAffectedCells =
            match maxUsedRow with
            | Some maxRow -> maxRow >= atRow
            | None -> false

        let newCells =
            if hasAffectedCells then
                let nearUsedTail =
                    match maxUsedRow with
                    | Some maxRow -> atRow > maxRow / 2
                    | None -> false
                if nearUsedTail then
                    let affectedCount = countCellsWhere (fun addr -> addr.Row >= atRow) model.Cells
                    if sparseRewriteWorth affectedCount model.Cells then
                        let removals = ResizeArray<CellAddress>(affectedCount)
                        let additions = ResizeArray<CellAddress * CellData>(max 0 (affectedCount - 1))
                        for KeyValue(addr, cellData) in model.Cells do
                            if addr.Row >= atRow then
                                removals.Add addr
                                if addr.Row > atRow then
                                    additions.Add({ addr with Row = addr.Row - 1 }, cellData)
                        rewriteSparseCells model.Cells removals additions
                    else
                        // Remove cells in the deleted row and shift rows up
                        let mutable shifted = Map.empty
                        model.Cells
                        |> Map.iter (fun addr cellData ->
                            if addr.Row <> atRow then
                                let newAddr =
                                    if addr.Row > atRow then { addr with Row = addr.Row - 1 }
                                    else addr
                                shifted <- Map.add newAddr cellData shifted)
                        shifted
                else
                    // Remove cells in the deleted row and shift rows up
                    let mutable shifted = Map.empty
                    model.Cells
                    |> Map.iter (fun addr cellData ->
                        if addr.Row <> atRow then
                            let newAddr =
                                if addr.Row > atRow then { addr with Row = addr.Row - 1 }
                                else addr
                            shifted <- Map.add newAddr cellData shifted)
                    shifted
            else
                model.Cells

        // Update selection if affected
        let newSelection =
            match model.Selection with
            | SingleCell (row, col) ->
                if row = atRow then SingleCell (max 0 (atRow - 1), col)
                elif row > atRow then SingleCell (row - 1, col)
                else model.Selection
            | Range (sr, sc, er, ec) ->
                let newSr = if sr > atRow then sr - 1 elif sr = atRow then max 0 (atRow - 1) else sr
                let newEr = if er > atRow then er - 1 elif er = atRow then max 0 (atRow - 1) else er
                Range (newSr, sc, newEr, ec)

        { model with
            Cells = newCells
            Selection = newSelection
            RowCount = model.RowCount - 1
        }

type private MaxUsedColumnCache = {
    Cells: Map<CellAddress, CellData>
    CellCount: int
    SourceTextCharCount: int
    MaxColumn: int option
}

type private StructuralCellSourceTextCache = {
    Cells: Map<CellAddress, CellData>
    SourceTextCharCount: int
}

let mutable private structuralCellSourceTextCache: StructuralCellSourceTextCache option = None
let private structuralCellSourceTextCacheMaxTotalCells = 500000
let private structuralCellSourceTextCacheMaxTotalSourceTextChars = 5000000

let private structuralCellSourceTextCharCountFor (cells: Map<CellAddress, CellData>) =
    if Map.count cells > structuralCellSourceTextCacheMaxTotalCells then
        structuralCellSourceTextCacheMaxTotalSourceTextChars + 1
    else
        match structuralCellSourceTextCache with
        | Some cached when obj.ReferenceEquals(cached.Cells, cells) ->
            cached.SourceTextCharCount
        | _ ->
            let sourceTextCharCount =
                cellMapSourceTextCharCountFor
                    structuralCellSourceTextCacheMaxTotalCells
                    structuralCellSourceTextCacheMaxTotalSourceTextChars
                    cells
            if sourceTextCharCount <= structuralCellSourceTextCacheMaxTotalSourceTextChars then
                structuralCellSourceTextCache <-
                    Some ({
                        Cells = cells
                        SourceTextCharCount = sourceTextCharCount
                    }: StructuralCellSourceTextCache)
            sourceTextCharCount

let mutable private maxUsedColumnCache: MaxUsedColumnCache list = []
let private maxUsedColumnCacheByCells =
    System.Collections.Generic.Dictionary<Map<CellAddress, CellData>, MaxUsedColumnCache>(cellMapReferenceComparer)
let private maxUsedColumnFastSlots = Array.zeroCreate<MaxUsedColumnCache> 256
let private maxUsedColumnCacheMaxSources = 64
let private maxUsedColumnCacheMaxTotalCells = structuralCellSourceTextCacheMaxTotalCells
let private maxUsedColumnCacheMaxTotalSourceTextChars = structuralCellSourceTextCacheMaxTotalSourceTextChars

let private maxUsedColumnCacheMatches cells (cached: MaxUsedColumnCache) =
    obj.ReferenceEquals(cached.Cells, cells)

let private setMaxUsedColumnCache entries =
    maxUsedColumnCache <- entries
    maxUsedColumnCacheByCells.Clear()
    clearArray maxUsedColumnFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<MaxUsedColumnCache>
        if not (maxUsedColumnCacheByCells.TryGetValue(entry.Cells, &existing)) then
            maxUsedColumnCacheByCells.[entry.Cells] <- entry
            let slot = referenceSlot entry.Cells maxUsedColumnFastSlots.Length
            maxUsedColumnFastSlots.[slot] <- entry

let private tryGetMaxUsedColumnCache cells =
    let slot = referenceSlot cells maxUsedColumnFastSlots.Length
    let fast = maxUsedColumnFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && maxUsedColumnCacheMatches cells fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<MaxUsedColumnCache>
        if maxUsedColumnCacheByCells.TryGetValue(cells, &cached) then
            maxUsedColumnFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private maxUsedColumnCacheWithoutCells cells =
    let mutable cached = Unchecked.defaultof<MaxUsedColumnCache>
    if maxUsedColumnCacheByCells.TryGetValue(cells, &cached) then
        maxUsedColumnCache
        |> List.filter (maxUsedColumnCacheMatches cells >> not)
    else
        maxUsedColumnCache

let private removeMaxUsedColumnCacheForCells cells =
    let entries = maxUsedColumnCacheWithoutCells cells
    if not (obj.ReferenceEquals(entries, maxUsedColumnCache)) then
        setMaxUsedColumnCache entries

let private maxUsedColumnCacheRetainable (cached: MaxUsedColumnCache) =
    cached.CellCount <= maxUsedColumnCacheMaxTotalCells
    && cached.SourceTextCharCount <= maxUsedColumnCacheMaxTotalSourceTextChars

let private rememberMaxUsedColumn (cached: MaxUsedColumnCache) =
    if not (maxUsedColumnCacheRetainable cached) then
        removeMaxUsedColumnCacheForCells cached.Cells
    else
        let trim (entries: MaxUsedColumnCache list) =
            let rec loop keptCount keptCells keptSourceTextChars (acc: MaxUsedColumnCache list) (remaining: MaxUsedColumnCache list) =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateCells = max 1 candidate.CellCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    if maxUsedColumnCacheRetainable candidate
                       && keptCount < maxUsedColumnCacheMaxSources
                       && keptCells + candidateCells <= maxUsedColumnCacheMaxTotalCells
                       && keptSourceTextChars + candidateSourceTextChars <= maxUsedColumnCacheMaxTotalSourceTextChars then
                        loop
                            (keptCount + 1)
                            (keptCells + candidateCells)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptCells keptSourceTextChars acc tail
            loop 0 0 0 [] entries
        (cached :: maxUsedColumnCacheWithoutCells cached.Cells)
        |> trim
        |> setMaxUsedColumnCache

let private maxUsedColumn (model: SpreadsheetModel) =
    match tryGetMaxUsedColumnCache model.Cells with
    | ValueSome cached when maxUsedColumnCacheRetainable cached ->
        cached.MaxColumn
    | ValueSome cached ->
        rememberMaxUsedColumn cached
        cached.MaxColumn
    | ValueNone ->
        let mutable found = false
        let mutable maxCol = 0
        model.Cells
        |> Map.iter (fun addr _ ->
            found <- true
            if addr.Col > maxCol then maxCol <- addr.Col)
        let result = if found then Some maxCol else None
        rememberMaxUsedColumn {
            Cells = model.Cells
            CellCount = max 1 model.Cells.Count
            SourceTextCharCount = structuralCellSourceTextCharCountFor model.Cells
            MaxColumn = result
        }
        result

/// Insert a column at the specified index (shifts columns right)
let insertColumn (atCol: int) (model: SpreadsheetModel): SpreadsheetModel =
    if atCol < 0 || atCol > model.ColCount then
        model
    else
        let maxCol = maxUsedColumn model
        let hasAffectedCells =
            match maxCol with
            | Some maxCol -> maxCol >= atCol
            | None -> false

        let newCells =
            if hasAffectedCells then
                let nearUsedTail =
                    match maxCol with
                    | Some maxCol -> atCol > maxCol / 2
                    | None -> false
                if nearUsedTail then
                    let affectedCount = countCellsWhere (fun addr -> addr.Col >= atCol) model.Cells
                    if sparseRewriteWorth affectedCount model.Cells then
                        let removals = ResizeArray<CellAddress>(affectedCount)
                        let additions = ResizeArray<CellAddress * CellData>(affectedCount)
                        for KeyValue(addr, cellData) in model.Cells do
                            if addr.Col >= atCol then
                                removals.Add addr
                                additions.Add({ addr with Col = addr.Col + 1 }, cellData)
                        rewriteSparseCells model.Cells removals additions
                    else
                        // Shift all columns >= atCol right by 1
                        let mutable shifted = Map.empty
                        model.Cells
                        |> Map.iter (fun addr cellData ->
                            let newAddr =
                                if addr.Col >= atCol then { addr with Col = addr.Col + 1 }
                                else addr
                            shifted <- Map.add newAddr cellData shifted)
                        shifted
                else
                    // Shift all columns >= atCol right by 1
                    let mutable shifted = Map.empty
                    model.Cells
                    |> Map.iter (fun addr cellData ->
                        let newAddr =
                            if addr.Col >= atCol then { addr with Col = addr.Col + 1 }
                            else addr
                        shifted <- Map.add newAddr cellData shifted)
                    shifted
            else
                model.Cells

        // Update selection if affected
        let newSelection =
            match model.Selection with
            | SingleCell (row, col) ->
                if col >= atCol then SingleCell (row, col + 1)
                else model.Selection
            | Range (sr, sc, er, ec) ->
                let newSc = if sc >= atCol then sc + 1 else sc
                let newEc = if ec >= atCol then ec + 1 else ec
                Range (sr, newSc, er, newEc)

        { model with
            Cells = newCells
            Selection = newSelection
            ColCount = model.ColCount + 1
        }

/// Delete a column at the specified index (shifts columns left)
let deleteColumn (atCol: int) (model: SpreadsheetModel): SpreadsheetModel =
    if atCol < 0 || atCol >= model.ColCount || model.ColCount <= 1 then
        model
    else
        let maxCol = maxUsedColumn model
        let hasAffectedCells =
            match maxCol with
            | Some maxCol -> maxCol >= atCol
            | None -> false

        let newCells =
            if hasAffectedCells then
                let nearUsedTail =
                    match maxCol with
                    | Some maxCol -> atCol > maxCol / 2
                    | None -> false
                if nearUsedTail then
                    let affectedCount = countCellsWhere (fun addr -> addr.Col >= atCol) model.Cells
                    if sparseRewriteWorth affectedCount model.Cells then
                        let removals = ResizeArray<CellAddress>(affectedCount)
                        let additions = ResizeArray<CellAddress * CellData>(max 0 (affectedCount - 1))
                        for KeyValue(addr, cellData) in model.Cells do
                            if addr.Col >= atCol then
                                removals.Add addr
                                if addr.Col > atCol then
                                    additions.Add({ addr with Col = addr.Col - 1 }, cellData)
                        rewriteSparseCells model.Cells removals additions
                    else
                        // Remove cells in the deleted column and shift columns left
                        let mutable shifted = Map.empty
                        model.Cells
                        |> Map.iter (fun addr cellData ->
                            if addr.Col <> atCol then
                                let newAddr =
                                    if addr.Col > atCol then { addr with Col = addr.Col - 1 }
                                    else addr
                                shifted <- Map.add newAddr cellData shifted)
                        shifted
                else
                    // Remove cells in the deleted column and shift columns left
                    let mutable shifted = Map.empty
                    model.Cells
                    |> Map.iter (fun addr cellData ->
                        if addr.Col <> atCol then
                            let newAddr =
                                if addr.Col > atCol then { addr with Col = addr.Col - 1 }
                                else addr
                            shifted <- Map.add newAddr cellData shifted)
                    shifted
            else
                model.Cells

        // Update selection if affected
        let newSelection =
            match model.Selection with
            | SingleCell (row, col) ->
                if col = atCol then SingleCell (row, max 0 (atCol - 1))
                elif col > atCol then SingleCell (row, col - 1)
                else model.Selection
            | Range (sr, sc, er, ec) ->
                let newSc = if sc > atCol then sc - 1 elif sc = atCol then max 0 (atCol - 1) else sc
                let newEc = if ec > atCol then ec - 1 elif ec = atCol then max 0 (atCol - 1) else ec
                Range (sr, newSc, er, newEc)

        { model with
            Cells = newCells
            Selection = newSelection
            ColCount = model.ColCount - 1
        }

let private defaultRowHeight = 25.0
let private defaultColumnWidth = 100.0

/// Get row height (with default)
let getRowHeight (row: int) (model: SpreadsheetModel): float =
    if isFilterHiddenRow row model then
        0.0
    else
        Map.tryFind row model.RowHeights
        |> Option.defaultValue defaultRowHeight

/// Set row height
let setRowHeight (row: int) (height: float) (model: SpreadsheetModel): SpreadsheetModel =
    if row < 0 then
        model
    elif height = defaultRowHeight then
        if Map.containsKey row model.RowHeights then
            { model with RowHeights = Map.remove row model.RowHeights }
        else
            model
    else
        let mutable existing = 0.0
        if model.RowHeights.TryGetValue(row, &existing) then
            if existing = height then model
            else { model with RowHeights = Map.add row height model.RowHeights }
        else
            { model with RowHeights = Map.add row height model.RowHeights }

/// Get column width (with default)
let getColumnWidth (col: int) (model: SpreadsheetModel): float =
    Map.tryFind col model.ColumnWidths
    |> Option.defaultValue defaultColumnWidth

/// Set column width
let setColumnWidth (col: int) (width: float) (model: SpreadsheetModel): SpreadsheetModel =
    if col < 0 then
        model
    elif width = defaultColumnWidth then
        if Map.containsKey col model.ColumnWidths then
            { model with ColumnWidths = Map.remove col model.ColumnWidths }
        else
            model
    else
        let mutable existing = 0.0
        if model.ColumnWidths.TryGetValue(col, &existing) then
            if existing = width then model
            else { model with ColumnWidths = Map.add col width model.ColumnWidths }
        else
            { model with ColumnWidths = Map.add col width model.ColumnWidths }

/// Sparse used range of the sheet's real content. Scans the cell map once
/// and includes merged ranges so view sizing and Ctrl+End do not ignore
/// populated-looking merged regions.
type private UsedBoundsCache = {
    Cells: Map<CellAddress, CellData>
    CellCount: int
    SourceTextCharCount: int
    MergedCells: MergedCell list
    MergedCellCount: int
    Bounds: (int * int) option
}

let mutable private usedBoundsCache: UsedBoundsCache list = []
let private usedBoundsCacheBySource =
    System.Collections.Generic.Dictionary<CellMergedReferenceKey, UsedBoundsCache>(cellMergedReferenceComparer)
let private usedBoundsFastSlots = Array.zeroCreate<UsedBoundsCache> 256
let private usedBoundsCacheMaxSources = 64
let private usedBoundsCacheMaxTotalCells = structuralCellSourceTextCacheMaxTotalCells
let private usedBoundsCacheMaxTotalMergedCells = 100000
let private usedBoundsCacheMaxTotalSourceTextChars = structuralCellSourceTextCacheMaxTotalSourceTextChars

let private mergedCellCountCapped limit (mergedCells: MergedCell list) =
    let mutable count = 0
    let mutable remaining = mergedCells
    while count <= limit && not remaining.IsEmpty do
        count <- count + 1
        remaining <- remaining.Tail
    count

let private usedBoundsCacheMatches cells mergedCells (cached: UsedBoundsCache) =
    obj.ReferenceEquals(cached.Cells, cells)
    && obj.ReferenceEquals(cached.MergedCells, mergedCells)

let private setUsedBoundsCache entries =
    usedBoundsCache <- entries
    usedBoundsCacheBySource.Clear()
    clearArray usedBoundsFastSlots
    for entry in entries do
        let key = cellMergedReferenceKey entry.Cells entry.MergedCells
        let mutable existing = Unchecked.defaultof<UsedBoundsCache>
        if not (usedBoundsCacheBySource.TryGetValue(key, &existing)) then
            usedBoundsCacheBySource.[key] <- entry
            let slot = cellMergedReferenceHash entry.Cells entry.MergedCells &&& (usedBoundsFastSlots.Length - 1)
            usedBoundsFastSlots.[slot] <- entry

let private tryGetUsedBoundsCache cells mergedCells =
    let slot = cellMergedReferenceHash cells mergedCells &&& (usedBoundsFastSlots.Length - 1)
    let fast = usedBoundsFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && usedBoundsCacheMatches cells mergedCells fast then
        ValueSome fast
    else
        let key = cellMergedReferenceKey cells mergedCells
        let mutable cached = Unchecked.defaultof<UsedBoundsCache>
        if usedBoundsCacheBySource.TryGetValue(key, &cached) then
            usedBoundsFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private usedBoundsCacheWithoutSource cells mergedCells =
    let key = cellMergedReferenceKey cells mergedCells
    let mutable cached = Unchecked.defaultof<UsedBoundsCache>
    if usedBoundsCacheBySource.TryGetValue(key, &cached) then
        usedBoundsCache
        |> List.filter (usedBoundsCacheMatches cells mergedCells >> not)
    else
        usedBoundsCache

let private removeUsedBoundsCacheForSource cells mergedCells =
    let entries = usedBoundsCacheWithoutSource cells mergedCells
    if not (obj.ReferenceEquals(entries, usedBoundsCache)) then
        setUsedBoundsCache entries

let private usedBoundsCacheRetainable (cached: UsedBoundsCache) =
    cached.CellCount <= usedBoundsCacheMaxTotalCells
    && cached.MergedCellCount <= usedBoundsCacheMaxTotalMergedCells
    && cached.SourceTextCharCount <= usedBoundsCacheMaxTotalSourceTextChars

let private rememberUsedBounds cached =
    if not (usedBoundsCacheRetainable cached) then
        removeUsedBoundsCacheForSource cached.Cells cached.MergedCells
    else
        let trim (entries: UsedBoundsCache list) =
            let rec loop keptCount keptCells keptMergedCells keptSourceTextChars acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateCells = max 1 candidate.CellCount
                    let candidateMergedCells = max 1 candidate.MergedCellCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    if usedBoundsCacheRetainable candidate
                       && keptCount < usedBoundsCacheMaxSources
                       && keptCells + candidateCells <= usedBoundsCacheMaxTotalCells
                       && keptMergedCells + candidateMergedCells <= usedBoundsCacheMaxTotalMergedCells
                       && keptSourceTextChars + candidateSourceTextChars <= usedBoundsCacheMaxTotalSourceTextChars then
                        loop
                            (keptCount + 1)
                            (keptCells + candidateCells)
                            (keptMergedCells + candidateMergedCells)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptCells keptMergedCells keptSourceTextChars acc tail
            loop 0 0 0 0 [] entries
        (cached :: usedBoundsCacheWithoutSource cached.Cells cached.MergedCells)
        |> trim
        |> setUsedBoundsCache

type private ExportCsvCache = {
    Cells: Map<CellAddress, CellData>
    CellCount: int
    SourceTextCharCount: int
    MergedCells: MergedCell list
    MergedCellCount: int
    TextLength: int
    Text: string
}

let mutable private exportCsvCache: ExportCsvCache list = []
let private exportCsvCacheBySource =
    System.Collections.Generic.Dictionary<CellMergedReferenceKey, ExportCsvCache>(cellMergedReferenceComparer)
let private exportCsvFastSlots = Array.zeroCreate<ExportCsvCache> 256
let private exportCsvCacheMaxSources = 64
let private exportCsvCacheMaxTotalChars = 5000000
let private exportCsvCacheMaxTotalCells = 500000
let private exportCsvCacheMaxTotalMergedCells = 200000
let private exportCsvCacheMaxTotalSourceTextChars = 5000000

let private exportCsvCacheMatches cells mergedCells (cached: ExportCsvCache) =
    obj.ReferenceEquals(cached.Cells, cells)
    && obj.ReferenceEquals(cached.MergedCells, mergedCells)

let private setExportCsvCache entries =
    exportCsvCache <- entries
    exportCsvCacheBySource.Clear()
    clearArray exportCsvFastSlots
    for entry in entries do
        let key = cellMergedReferenceKey entry.Cells entry.MergedCells
        let mutable existing = Unchecked.defaultof<ExportCsvCache>
        if not (exportCsvCacheBySource.TryGetValue(key, &existing)) then
            exportCsvCacheBySource.[key] <- entry
            let slot = cellMergedReferenceHash entry.Cells entry.MergedCells &&& (exportCsvFastSlots.Length - 1)
            exportCsvFastSlots.[slot] <- entry

let private tryGetExportCsvCache cells mergedCells =
    let slot = cellMergedReferenceHash cells mergedCells &&& (exportCsvFastSlots.Length - 1)
    let fast = exportCsvFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && exportCsvCacheMatches cells mergedCells fast then
        ValueSome fast
    else
        let key = cellMergedReferenceKey cells mergedCells
        let mutable cached = Unchecked.defaultof<ExportCsvCache>
        if exportCsvCacheBySource.TryGetValue(key, &cached) then
            exportCsvFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private exportCsvCacheWithoutSource cells mergedCells =
    let key = cellMergedReferenceKey cells mergedCells
    let mutable cached = Unchecked.defaultof<ExportCsvCache>
    if exportCsvCacheBySource.TryGetValue(key, &cached) then
        exportCsvCache
        |> List.filter (exportCsvCacheMatches cells mergedCells >> not)
    else
        exportCsvCache

let private removeExportCsvCacheForSource cells mergedCells =
    let entries = exportCsvCacheWithoutSource cells mergedCells
    if not (obj.ReferenceEquals(entries, exportCsvCache)) then
        setExportCsvCache entries

let private exportCsvCacheRetainable (cached: ExportCsvCache) =
    cached.TextLength <= exportCsvCacheMaxTotalChars
    && cached.CellCount <= exportCsvCacheMaxTotalCells
    && cached.MergedCellCount <= exportCsvCacheMaxTotalMergedCells
    && cached.SourceTextCharCount <= exportCsvCacheMaxTotalSourceTextChars

let private exportCsvSourceTextCharCountFor textLength cellCount mergedCellCount cells =
    if textLength > exportCsvCacheMaxTotalChars
       || cellCount > exportCsvCacheMaxTotalCells
       || mergedCellCount > exportCsvCacheMaxTotalMergedCells then
        exportCsvCacheMaxTotalSourceTextChars + 1
    else
        cellMapSourceTextCharCountFor
            exportCsvCacheMaxTotalCells
            exportCsvCacheMaxTotalSourceTextChars
            cells

let private rememberExportCsv cached =
    if not (exportCsvCacheRetainable cached) then
        removeExportCsvCacheForSource cached.Cells cached.MergedCells
    else
        let trim entries =
            let rec loop keptCount keptChars keptCells keptMergedCells keptSourceTextChars acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateChars = max 1 candidate.TextLength
                    let candidateCells = max 1 candidate.CellCount
                    let candidateMergedCells = max 1 candidate.MergedCellCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    if exportCsvCacheRetainable candidate
                       && keptCount < exportCsvCacheMaxSources
                       && keptChars + candidateChars <= exportCsvCacheMaxTotalChars
                       && keptCells + candidateCells <= exportCsvCacheMaxTotalCells
                       && keptMergedCells + candidateMergedCells <= exportCsvCacheMaxTotalMergedCells
                       && keptSourceTextChars + candidateSourceTextChars <= exportCsvCacheMaxTotalSourceTextChars then
                        loop
                            (keptCount + 1)
                            (keptChars + candidateChars)
                            (keptCells + candidateCells)
                            (keptMergedCells + candidateMergedCells)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptChars keptCells keptMergedCells keptSourceTextChars acc tail
            loop 0 0 0 0 0 [] entries
        (cached :: exportCsvCacheWithoutSource cached.Cells cached.MergedCells)
        |> trim
        |> setExportCsvCache

let usedBounds (model: SpreadsheetModel) : (int * int) option =
    match tryGetUsedBoundsCache model.Cells model.MergedCells with
    | ValueSome cached when usedBoundsCacheRetainable cached ->
        cached.Bounds
    | ValueSome cached ->
        rememberUsedBounds cached
        cached.Bounds
    | ValueNone ->
        let mutable any = false
        let mutable maxRow = 0
        let mutable maxCol = 0

        if not (Map.isEmpty model.Cells) then
            any <- true
            let maxAddr, _ = Map.maxKeyValue model.Cells
            maxRow <- maxAddr.Row

            model.Cells
            |> Map.iter (fun addr _ ->
                if addr.Col > maxCol then maxCol <- addr.Col)

        for merge in model.MergedCells do
            any <- true
            if merge.EndRow > maxRow then maxRow <- merge.EndRow
            if merge.EndCol > maxCol then maxCol <- merge.EndCol

        let bounds = if any then Some (maxRow, maxCol) else None
        let mergedCellCount = mergedCellCountCapped usedBoundsCacheMaxTotalMergedCells model.MergedCells
        rememberUsedBounds {
            Cells = model.Cells
            CellCount = max 1 model.Cells.Count
            SourceTextCharCount =
                if mergedCellCount > usedBoundsCacheMaxTotalMergedCells then
                    usedBoundsCacheMaxTotalSourceTextChars + 1
                else
                    structuralCellSourceTextCharCountFor model.Cells
            MergedCells = model.MergedCells
            MergedCellCount = mergedCellCount
            Bounds = bounds
        }
        bounds

/// Last populated column in a row. Used by End-key navigation; keeps that
/// row-specific scan in one place and accounts for merged cells crossing
/// the row.
type private LastUsedColumnInRowCache = {
    Cells: Map<CellAddress, CellData>
    CellCount: int
    SourceTextCharCount: int
    MergedCells: MergedCell list
    MergedCellCount: int
    CellRowLastColumns: System.Collections.Generic.Dictionary<int, int>
    MergeRowLastColumns: System.Collections.Generic.Dictionary<int, int> option
    Entries: System.Collections.Generic.Dictionary<int, int option>
    Order: System.Collections.Generic.Queue<int>
}

let mutable private lastUsedColumnInRowCache: LastUsedColumnInRowCache list = []
let private lastUsedColumnInRowCacheBySource =
    System.Collections.Generic.Dictionary<CellMergedReferenceKey, LastUsedColumnInRowCache>(cellMergedReferenceComparer)
let private lastUsedColumnInRowFastSlots = Array.zeroCreate<LastUsedColumnInRowCache> 256
let private lastUsedColumnInRowCacheMaxEntries = 64
let private lastUsedColumnInRowCacheMaxSources = 64
let private lastUsedColumnInRowCacheMaxTotalCells = structuralCellSourceTextCacheMaxTotalCells
let private lastUsedColumnInRowCacheMaxTotalMergedCells = 100000
let private lastUsedColumnInRowCacheMaxTotalSourceTextChars = structuralCellSourceTextCacheMaxTotalSourceTextChars

let private lastUsedColumnInRowCacheMatches cells mergedCells (cached: LastUsedColumnInRowCache) =
    obj.ReferenceEquals(cached.Cells, cells)
    && obj.ReferenceEquals(cached.MergedCells, mergedCells)

let private setLastUsedColumnInRowCache entries =
    lastUsedColumnInRowCache <- entries
    lastUsedColumnInRowCacheBySource.Clear()
    clearArray lastUsedColumnInRowFastSlots
    for entry in entries do
        let key = cellMergedReferenceKey entry.Cells entry.MergedCells
        let mutable existing = Unchecked.defaultof<LastUsedColumnInRowCache>
        if not (lastUsedColumnInRowCacheBySource.TryGetValue(key, &existing)) then
            lastUsedColumnInRowCacheBySource.[key] <- entry
            let slot = cellMergedReferenceHash entry.Cells entry.MergedCells &&& (lastUsedColumnInRowFastSlots.Length - 1)
            lastUsedColumnInRowFastSlots.[slot] <- entry

let private tryGetLastUsedColumnInRowCache cells mergedCells =
    let slot = cellMergedReferenceHash cells mergedCells &&& (lastUsedColumnInRowFastSlots.Length - 1)
    let fast = lastUsedColumnInRowFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && lastUsedColumnInRowCacheMatches cells mergedCells fast then
        ValueSome fast
    else
        let key = cellMergedReferenceKey cells mergedCells
        let mutable cached = Unchecked.defaultof<LastUsedColumnInRowCache>
        if lastUsedColumnInRowCacheBySource.TryGetValue(key, &cached) then
            lastUsedColumnInRowFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private lastUsedColumnInRowCacheWithoutSource cells mergedCells =
    let key = cellMergedReferenceKey cells mergedCells
    let mutable cached = Unchecked.defaultof<LastUsedColumnInRowCache>
    if lastUsedColumnInRowCacheBySource.TryGetValue(key, &cached) then
        lastUsedColumnInRowCache
        |> List.filter (lastUsedColumnInRowCacheMatches cells mergedCells >> not)
    else
        lastUsedColumnInRowCache

let private removeLastUsedColumnInRowCacheForSource cells mergedCells =
    let entries = lastUsedColumnInRowCacheWithoutSource cells mergedCells
    if not (obj.ReferenceEquals(entries, lastUsedColumnInRowCache)) then
        setLastUsedColumnInRowCache entries

let private lastUsedColumnInRowCacheRetainable (cached: LastUsedColumnInRowCache) =
    cached.CellCount <= lastUsedColumnInRowCacheMaxTotalCells
    && cached.MergedCellCount <= lastUsedColumnInRowCacheMaxTotalMergedCells
    && cached.SourceTextCharCount <= lastUsedColumnInRowCacheMaxTotalSourceTextChars

let private buildLastUsedColumnRowIndex (cells: Map<CellAddress, CellData>) =
    let rows = System.Collections.Generic.Dictionary<int, int>()
    use enumerator = (cells :> seq<_>).GetEnumerator()
    while enumerator.MoveNext() do
        let addr = enumerator.Current.Key
        let mutable lastCol = 0
        if rows.TryGetValue(addr.Row, &lastCol) then
            if addr.Col > lastCol then
                rows.[addr.Row] <- addr.Col
        else
            rows.[addr.Row] <- addr.Col
    rows

let private buildLastUsedColumnMergeRowIndex (mergedCells: MergedCell list) =
    let mutable rowEntries = 0
    let mutable eligible = true
    let mutable remaining = mergedCells
    while eligible && not remaining.IsEmpty do
        let merge = remaining.Head
        if merge.EndRow >= merge.StartRow then
            rowEntries <- rowEntries + merge.EndRow - merge.StartRow + 1
            if rowEntries > lastUsedColumnInRowCacheMaxTotalMergedCells then
                eligible <- false
        remaining <- remaining.Tail

    if eligible then
        let rows = System.Collections.Generic.Dictionary<int, int>()
        for merge in mergedCells do
            if merge.EndRow >= merge.StartRow then
                for row in merge.StartRow .. merge.EndRow do
                    let mutable lastCol = 0
                    if rows.TryGetValue(row, &lastCol) then
                        if merge.EndCol > lastCol then
                            rows.[row] <- merge.EndCol
                    else
                        rows.[row] <- merge.EndCol
        Some rows
    else
        None

let private lastUsedColumnInRowFromIndex row (cellRowLastColumns: System.Collections.Generic.Dictionary<int, int>) (mergeRowLastColumns: System.Collections.Generic.Dictionary<int, int> option) (mergedCells: MergedCell list) =
    let mutable found = false
    let mutable lastCol = 0
    let mutable cellLastCol = 0

    if cellRowLastColumns.TryGetValue(row, &cellLastCol) then
        found <- true
        lastCol <- cellLastCol

    match mergeRowLastColumns with
    | Some mergeRows ->
        let mutable mergeLastCol = 0
        if mergeRows.TryGetValue(row, &mergeLastCol) then
            found <- true
            if mergeLastCol > lastCol then lastCol <- mergeLastCol
    | None ->
        for merge in mergedCells do
            if row >= merge.StartRow && row <= merge.EndRow then
                found <- true
                if merge.EndCol > lastCol then lastCol <- merge.EndCol

    if found then Some lastCol else None

let private scanLastUsedColumnInRow row (model: SpreadsheetModel) =
    let mutable found = false
    let mutable lastCol = 0

    model.Cells
    |> Map.exists (fun addr _ ->
        if addr.Row = row then
            found <- true
            if addr.Col > lastCol then lastCol <- addr.Col
            false
        else
            addr.Row > row)
    |> ignore

    for merge in model.MergedCells do
        if row >= merge.StartRow && row <= merge.EndRow then
            found <- true
            if merge.EndCol > lastCol then lastCol <- merge.EndCol

    if found then Some lastCol else None

let private rememberLastUsedColumnInRowCache cached =
    if not (lastUsedColumnInRowCacheRetainable cached) then
        removeLastUsedColumnInRowCacheForSource cached.Cells cached.MergedCells
    else
        let trim (entries: LastUsedColumnInRowCache list) =
            let rec loop keptCount keptCells keptMergedCells keptSourceTextChars acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateCells = max 1 candidate.CellCount
                    let candidateMergedCells = max 1 candidate.MergedCellCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    if lastUsedColumnInRowCacheRetainable candidate
                       && keptCount < lastUsedColumnInRowCacheMaxSources
                       && keptCells + candidateCells <= lastUsedColumnInRowCacheMaxTotalCells
                       && keptMergedCells + candidateMergedCells <= lastUsedColumnInRowCacheMaxTotalMergedCells
                       && keptSourceTextChars + candidateSourceTextChars <= lastUsedColumnInRowCacheMaxTotalSourceTextChars then
                        loop
                            (keptCount + 1)
                            (keptCells + candidateCells)
                            (keptMergedCells + candidateMergedCells)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptCells keptMergedCells keptSourceTextChars acc tail
            loop 0 0 0 0 [] entries
        (cached :: lastUsedColumnInRowCacheWithoutSource cached.Cells cached.MergedCells)
        |> trim
        |> setLastUsedColumnInRowCache

let private addLastUsedColumnInRowCacheEntry row result (entries: System.Collections.Generic.Dictionary<int, int option>) (order: System.Collections.Generic.Queue<int>) =
    if not (entries.ContainsKey row) then
        order.Enqueue row
    entries.[row] <- result
    while entries.Count > lastUsedColumnInRowCacheMaxEntries && order.Count > 0 do
        let oldest = order.Dequeue()
        if oldest <> row then
            entries.Remove oldest |> ignore
    if entries.Count > lastUsedColumnInRowCacheMaxEntries then
        entries.Clear()
        order.Clear()
        order.Enqueue row
        entries.[row] <- result

let lastUsedColumnInRow (row: int) (model: SpreadsheetModel) : int option =
    match tryGetLastUsedColumnInRowCache model.Cells model.MergedCells with
    | ValueSome cached when lastUsedColumnInRowCacheRetainable cached ->
        let mutable lastColumn = Unchecked.defaultof<int option>
        if cached.Entries.TryGetValue(row, &lastColumn) then
            lastColumn
        else
            let computed = lastUsedColumnInRowFromIndex row cached.CellRowLastColumns cached.MergeRowLastColumns cached.MergedCells
            addLastUsedColumnInRowCacheEntry row computed cached.Entries cached.Order
            computed
    | ValueSome cached ->
        rememberLastUsedColumnInRowCache cached
        let mutable lastColumn = Unchecked.defaultof<int option>
        if cached.Entries.TryGetValue(row, &lastColumn) then
            lastColumn
        else
            lastUsedColumnInRowFromIndex row cached.CellRowLastColumns cached.MergeRowLastColumns cached.MergedCells
    | ValueNone ->
        let mergedCellCount = mergedCellCountCapped lastUsedColumnInRowCacheMaxTotalMergedCells model.MergedCells
        if model.Cells.Count > lastUsedColumnInRowCacheMaxTotalCells
           || mergedCellCount > lastUsedColumnInRowCacheMaxTotalMergedCells then
            scanLastUsedColumnInRow row model
        else
            let cellRowLastColumns = buildLastUsedColumnRowIndex model.Cells
            let mergeRowLastColumns = buildLastUsedColumnMergeRowIndex model.MergedCells
            let result = lastUsedColumnInRowFromIndex row cellRowLastColumns mergeRowLastColumns model.MergedCells
            let entries = System.Collections.Generic.Dictionary<int, int option>()
            let order = System.Collections.Generic.Queue<int>()
            addLastUsedColumnInRowCacheEntry row result entries order
            rememberLastUsedColumnInRowCache
                {
                    Cells = model.Cells
                    CellCount = max 1 model.Cells.Count
                    SourceTextCharCount = structuralCellSourceTextCharCountFor model.Cells
                    MergedCells = model.MergedCells
                    MergedCellCount = mergedCellCount
                    CellRowLastColumns = cellRowLastColumns
                    MergeRowLastColumns = mergeRowLastColumns
                    Entries = entries
                    Order = order
                }
            result

/// Search cells for a query string
let private containsSearchText (query: string) (caseSensitive: bool) (text: string) =
#if FABLE_COMPILER
    if query.Length = 0 then
        true
    elif query.Length > text.Length then
        false
    elif caseSensitive then
        text.IndexOf(query) >= 0
    else
        let equalsIgnoreCase a b =
            a = b || Char.ToUpperInvariant a = Char.ToUpperInvariant b
        let mutable found = false
        let mutable start = 0
        while not found && start <= text.Length - query.Length do
            let mutable offset = 0
            let mutable matches = true
            while matches && offset < query.Length do
                if not (equalsIgnoreCase text.[start + offset] query.[offset]) then
                    matches <- false
                offset <- offset + 1
            if matches then
                found <- true
            start <- start + 1
        found
#else
    let comparison = if caseSensitive then StringComparison.Ordinal else StringComparison.OrdinalIgnoreCase
    text.IndexOf(query, comparison) >= 0
#endif

let private numberDisplayMayContainSearchText (query: string) =
    let mutable possible = true
    let mutable index = 0
    while possible && index < query.Length do
        let ch = Char.ToUpperInvariant query.[index]
        if not ((ch >= '0' && ch <= '9') ||
                ch = '.' || ch = '-' || ch = '+' ||
                ch = 'E' || ch = 'N' || ch = 'A' || ch = 'I' ||
                ch = 'F' || ch = 'T' || ch = 'Y') then
            possible <- false
        index <- index + 1
    possible

let private booleanDisplayMayContainSearchText (query: string) =
    let mutable possible = true
    let mutable index = 0
    while possible && index < query.Length do
        let ch = Char.ToUpperInvariant query.[index]
        if not (ch = 'T' || ch = 'R' || ch = 'U' || ch = 'E' ||
                ch = 'F' || ch = 'A' || ch = 'L' || ch = 'S') then
            possible <- false
        index <- index + 1
    possible

let private errorSearchTextMatcher (query: string) (caseSensitive: bool) =
    if query = "#" then
        fun _ -> true
    elif query.Length > 0 && query.[0] = '#' then
        let tail = query.Substring 1
        if tail.Length = 0 then
            fun _ -> true
        else
            fun code -> containsSearchText tail caseSensitive code
    else
        fun code -> containsSearchText query caseSensitive code

let private replaceLiteralText (findText: string) (replaceText: string) caseSensitive (currentText: string) =
    if caseSensitive then
        currentText.Replace(findText, replaceText)
    else
        let pattern = System.Text.RegularExpressions.Regex.Escape findText
        let replacement = replaceText.Replace("$", "$$")
        System.Text.RegularExpressions.Regex.Replace(
            currentText,
            pattern,
            replacement,
            System.Text.RegularExpressions.RegexOptions.IgnoreCase)

let private caseInsensitiveIdentityReplacementWouldChange (findText: string) (currentText: string) =
    let comparison = StringComparison.OrdinalIgnoreCase
    let mutable start = 0
    let mutable foundChangedMatch = false
    while not foundChangedMatch && start <= currentText.Length - findText.Length do
#if FABLE_COMPILER
        let mutable matchIndex = -1
        let mutable candidate = start
        while matchIndex < 0 && candidate <= currentText.Length - findText.Length do
            let mutable offset = 0
            let mutable matches = true
            while matches && offset < findText.Length do
                let actual = currentText.[candidate + offset]
                let expected = findText.[offset]
                if actual <> expected && Char.ToUpperInvariant actual <> Char.ToUpperInvariant expected then
                    matches <- false
                offset <- offset + 1
            if matches then
                matchIndex <- candidate
            else
                candidate <- candidate + 1
#else
        let matchIndex = currentText.IndexOf(findText, start, comparison)
#endif
        if matchIndex < 0 then
            start <- currentText.Length + 1
        else
            let mutable exact = true
            let mutable offset = 0
            while exact && offset < findText.Length do
                if currentText.[matchIndex + offset] <> findText.[offset] then
                    exact <- false
                offset <- offset + 1
            if exact then
                start <- matchIndex + findText.Length
            else
                foundChangedMatch <- true
    foundChangedMatch

type private SearchCellsCacheEntry = {
    Results: CellAddress array
}

type private SearchCellsCache = {
    Cells: Map<CellAddress, CellData>
    CellCount: int
    SourceTextCharCount: int
    Entries: System.Collections.Generic.Dictionary<struct (string * bool), SearchCellsCacheEntry>
    Order: System.Collections.Generic.Queue<struct (string * bool)>
    mutable TotalResultCount: int
    mutable TotalQueryCharCount: int
}

type private ReplaceIdentityCache = {
    Cells: Map<CellAddress, CellData>
    CellCount: int
    SourceTextCharCount: int
    Entries: System.Collections.Generic.Dictionary<struct (string * string * bool), unit>
    Order: System.Collections.Generic.Queue<struct (string * string * bool)>
    mutable TotalKeyCharCount: int
}

type private ReplaceAllChanges = {
    OldCells: ResizeArray<CellAddress * CellData option>
    NewCells: ResizeArray<CellAddress * CellData option>
}

let mutable private searchCellsCache: SearchCellsCache list = []
let mutable private replaceIdentityCache: ReplaceIdentityCache list = []
let private searchCellsCacheByCells =
    System.Collections.Generic.Dictionary<Map<CellAddress, CellData>, SearchCellsCache>(cellMapReferenceComparer)
let private replaceIdentityCacheByCells =
    System.Collections.Generic.Dictionary<Map<CellAddress, CellData>, ReplaceIdentityCache>(cellMapReferenceComparer)
let private searchCellsFastSlots = Array.zeroCreate<SearchCellsCache> 256
let private replaceIdentityFastSlots = Array.zeroCreate<ReplaceIdentityCache> 256
let private emptySearchResults: CellAddress array = [||]
let private searchCellsCacheMaxEntries = 64
let private searchCellsCacheMaxSources = 64
let private searchCellsCacheMaxResultCount = 200000
let private searchCellsCacheMaxTotalResultCount = 200000
let private searchCellsCacheMaxTotalCells = 500000
let private searchCellsCacheMaxTotalSourceTextChars = 5000000
let private searchCellsCacheMaxTotalQueryChars = 5000000
let private replaceIdentityCacheMaxEntries = 64
let private replaceIdentityCacheMaxSources = 64
let private replaceIdentityCacheMaxTotalCells = 500000
let private replaceIdentityCacheMaxTotalSourceTextChars = 5000000
let private replaceIdentityCacheMaxTotalKeyChars = 5000000

let private searchCellsCacheKeyCharCount (struct (query, _): struct (string * bool)) =
    query.Length

let private replaceIdentityCacheKeyCharCount (struct (findText, replaceText, _): struct (string * string * bool)) =
    findText.Length + replaceText.Length

let private searchCellsCacheMatches cells (cached: SearchCellsCache) =
    obj.ReferenceEquals(cached.Cells, cells)

let private replaceIdentityCacheMatches cells (cached: ReplaceIdentityCache) =
    obj.ReferenceEquals(cached.Cells, cells)

let private setSearchCellsCache entries =
    searchCellsCache <- entries
    searchCellsCacheByCells.Clear()
    clearArray searchCellsFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<SearchCellsCache>
        if not (searchCellsCacheByCells.TryGetValue(entry.Cells, &existing)) then
            searchCellsCacheByCells.[entry.Cells] <- entry
            let slot = referenceSlot entry.Cells searchCellsFastSlots.Length
            searchCellsFastSlots.[slot] <- entry

let private tryGetSearchCellsCache cells =
    let slot = referenceSlot cells searchCellsFastSlots.Length
    let fast = searchCellsFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && searchCellsCacheMatches cells fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<SearchCellsCache>
        if searchCellsCacheByCells.TryGetValue(cells, &cached) then
            searchCellsFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private searchCellsCacheWithoutCells cells =
    let mutable cached = Unchecked.defaultof<SearchCellsCache>
    if searchCellsCacheByCells.TryGetValue(cells, &cached) then
        searchCellsCache
        |> List.filter (searchCellsCacheMatches cells >> not)
    else
        searchCellsCache

let private removeSearchCellsCacheForCells cells =
    let entries = searchCellsCacheWithoutCells cells
    if not (obj.ReferenceEquals(entries, searchCellsCache)) then
        setSearchCellsCache entries

let private setReplaceIdentityCache entries =
    replaceIdentityCache <- entries
    replaceIdentityCacheByCells.Clear()
    clearArray replaceIdentityFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<ReplaceIdentityCache>
        if not (replaceIdentityCacheByCells.TryGetValue(entry.Cells, &existing)) then
            replaceIdentityCacheByCells.[entry.Cells] <- entry
            let slot = referenceSlot entry.Cells replaceIdentityFastSlots.Length
            replaceIdentityFastSlots.[slot] <- entry

let private tryGetReplaceIdentityCache cells =
    let slot = referenceSlot cells replaceIdentityFastSlots.Length
    let fast = replaceIdentityFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && replaceIdentityCacheMatches cells fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<ReplaceIdentityCache>
        if replaceIdentityCacheByCells.TryGetValue(cells, &cached) then
            replaceIdentityFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private replaceIdentityCacheWithoutCells cells =
    let mutable cached = Unchecked.defaultof<ReplaceIdentityCache>
    if replaceIdentityCacheByCells.TryGetValue(cells, &cached) then
        replaceIdentityCache
        |> List.filter (replaceIdentityCacheMatches cells >> not)
    else
        replaceIdentityCache

let private removeReplaceIdentityCacheForCells cells =
    let entries = replaceIdentityCacheWithoutCells cells
    if not (obj.ReferenceEquals(entries, replaceIdentityCache)) then
        setReplaceIdentityCache entries

let private searchCellsCacheRetainable (cached: SearchCellsCache) =
    cached.Entries.Count > 0
    && cached.TotalResultCount <= searchCellsCacheMaxTotalResultCount
    && cached.TotalQueryCharCount <= searchCellsCacheMaxTotalQueryChars
    && cached.CellCount <= searchCellsCacheMaxTotalCells
    && cached.SourceTextCharCount <= searchCellsCacheMaxTotalSourceTextChars

let private replaceIdentityCacheRetainable (cached: ReplaceIdentityCache) =
    cached.Entries.Count > 0
    && cached.TotalKeyCharCount <= replaceIdentityCacheMaxTotalKeyChars
    && cached.CellCount <= replaceIdentityCacheMaxTotalCells
    && cached.SourceTextCharCount <= replaceIdentityCacheMaxTotalSourceTextChars

let private searchCellsSourceTextCharCountFor resultCount queryCharCount cells =
    if resultCount > searchCellsCacheMaxResultCount
       || resultCount > searchCellsCacheMaxTotalResultCount
       || queryCharCount > searchCellsCacheMaxTotalQueryChars then
        searchCellsCacheMaxTotalSourceTextChars + 1
    else
        cellMapSourceTextCharCountFor
            searchCellsCacheMaxTotalCells
            searchCellsCacheMaxTotalSourceTextChars
            cells

let private replaceIdentityCacheSourceTextCharCountFor keyCharCount cells =
    if keyCharCount > replaceIdentityCacheMaxTotalKeyChars then
        replaceIdentityCacheMaxTotalSourceTextChars + 1
    else
        cellMapSourceTextCharCountFor
            replaceIdentityCacheMaxTotalCells
            replaceIdentityCacheMaxTotalSourceTextChars
            cells

let private rememberSearchCellsCache cached =
    if not (searchCellsCacheRetainable cached) then
        removeSearchCellsCacheForCells cached.Cells
    else
        let mutable keptCount = 1
        let mutable keptResults = cached.TotalResultCount
        let mutable keptQueryChars = cached.TotalQueryCharCount
        let mutable keptCells = max 1 cached.CellCount
        let mutable keptSourceTextChars = max 1 cached.SourceTextCharCount
        let retained =
            searchCellsCacheWithoutCells cached.Cells
            |> List.choose (fun candidate ->
                let candidateResults = candidate.TotalResultCount
                let candidateQueryChars = max 1 candidate.TotalQueryCharCount
                let candidateCells = max 1 candidate.CellCount
                let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                if searchCellsCacheRetainable candidate
                   && keptCount < searchCellsCacheMaxSources
                   && keptResults + candidateResults <= searchCellsCacheMaxTotalResultCount
                   && keptQueryChars + candidateQueryChars <= searchCellsCacheMaxTotalQueryChars
                   && keptCells + candidateCells <= searchCellsCacheMaxTotalCells
                   && keptSourceTextChars + candidateSourceTextChars <= searchCellsCacheMaxTotalSourceTextChars then
                    keptCount <- keptCount + 1
                    keptResults <- keptResults + candidateResults
                    keptQueryChars <- keptQueryChars + candidateQueryChars
                    keptCells <- keptCells + candidateCells
                    keptSourceTextChars <- keptSourceTextChars + candidateSourceTextChars
                    Some candidate
                else
                    None)
        cached :: retained
        |> setSearchCellsCache

let private rememberReplaceIdentityCache cached =
    if not (replaceIdentityCacheRetainable cached) then
        removeReplaceIdentityCacheForCells cached.Cells
    else
        let mutable keptCount = 1
        let mutable keptKeyChars = cached.TotalKeyCharCount
        let mutable keptCells = max 1 cached.CellCount
        let mutable keptSourceTextChars = max 1 cached.SourceTextCharCount
        let retained =
            replaceIdentityCacheWithoutCells cached.Cells
            |> List.choose (fun candidate ->
                let candidateKeyChars = max 1 candidate.TotalKeyCharCount
                let candidateCells = max 1 candidate.CellCount
                let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                if replaceIdentityCacheRetainable candidate
                   && keptCount < replaceIdentityCacheMaxSources
                   && keptKeyChars + candidateKeyChars <= replaceIdentityCacheMaxTotalKeyChars
                   && keptCells + candidateCells <= replaceIdentityCacheMaxTotalCells
                   && keptSourceTextChars + candidateSourceTextChars <= replaceIdentityCacheMaxTotalSourceTextChars then
                    keptCount <- keptCount + 1
                    keptKeyChars <- keptKeyChars + candidateKeyChars
                    keptCells <- keptCells + candidateCells
                    keptSourceTextChars <- keptSourceTextChars + candidateSourceTextChars
                    Some candidate
                else
                    None)
        cached :: retained
        |> setReplaceIdentityCache

let private cachedSearchCells query caseSensitive (model: SpreadsheetModel) =
    let key = struct (query, caseSensitive)
    match tryGetSearchCellsCache model.Cells with
    | ValueSome cached when searchCellsCacheRetainable cached ->
        let mutable entry = Unchecked.defaultof<SearchCellsCacheEntry>
        if cached.Entries.TryGetValue(key, &entry) then Some entry.Results else None
    | ValueSome cached ->
        rememberSearchCellsCache cached
        None
    | ValueNone -> None

let private cachedReplaceIdentity findText replaceText caseSensitive (model: SpreadsheetModel) =
    let key = struct (findText, replaceText, caseSensitive)
    match tryGetReplaceIdentityCache model.Cells with
    | ValueSome cached when replaceIdentityCacheRetainable cached ->
        cached.Entries.ContainsKey key
    | ValueSome cached ->
        rememberReplaceIdentityCache cached
        false
    | ValueNone -> false

let private removeSearchCellsCacheEntry key (cache: SearchCellsCache) =
    let mutable existing = Unchecked.defaultof<SearchCellsCacheEntry>
    if cache.Entries.TryGetValue(key, &existing) then
        cache.TotalResultCount <- max 0 (cache.TotalResultCount - existing.Results.Length)
        cache.TotalQueryCharCount <- max 0 (cache.TotalQueryCharCount - searchCellsCacheKeyCharCount key)
        cache.Entries.Remove key |> ignore

let private removeReplaceIdentityCacheEntry key (cache: ReplaceIdentityCache) =
    if cache.Entries.Remove key then
        cache.TotalKeyCharCount <- max 0 (cache.TotalKeyCharCount - replaceIdentityCacheKeyCharCount key)

let private addSearchCellsCacheEntry key (results: CellAddress array) (cache: SearchCellsCache) =
    let mutable existing = Unchecked.defaultof<SearchCellsCacheEntry>
    let keyCharCount = searchCellsCacheKeyCharCount key
    if results.Length > searchCellsCacheMaxResultCount
       || keyCharCount > searchCellsCacheMaxTotalQueryChars then
        removeSearchCellsCacheEntry key cache
    else
        let alreadyCached = cache.Entries.TryGetValue(key, &existing)
        if alreadyCached then
            cache.TotalResultCount <- max 0 (cache.TotalResultCount - existing.Results.Length)
        else
            cache.Order.Enqueue key
            cache.TotalQueryCharCount <- cache.TotalQueryCharCount + keyCharCount
        cache.Entries.[key] <- { Results = results }
        cache.TotalResultCount <- cache.TotalResultCount + results.Length
        while (cache.Entries.Count > searchCellsCacheMaxEntries
               || cache.TotalResultCount > searchCellsCacheMaxResultCount
               || cache.TotalQueryCharCount > searchCellsCacheMaxTotalQueryChars)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if oldest <> key then
                removeSearchCellsCacheEntry oldest cache
        if cache.Entries.Count > searchCellsCacheMaxEntries
           || cache.TotalResultCount > searchCellsCacheMaxResultCount
           || cache.TotalQueryCharCount > searchCellsCacheMaxTotalQueryChars then
            let current = cache.Entries.[key]
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue key
            cache.Entries.[key] <- current
            cache.TotalResultCount <- current.Results.Length
            cache.TotalQueryCharCount <- keyCharCount

let private addReplaceIdentityCacheEntry key (cache: ReplaceIdentityCache) =
    let keyCharCount = replaceIdentityCacheKeyCharCount key
    if keyCharCount > replaceIdentityCacheMaxTotalKeyChars then
        removeReplaceIdentityCacheEntry key cache
    else
        if not (cache.Entries.ContainsKey key) then
            cache.Order.Enqueue key
            cache.TotalKeyCharCount <- cache.TotalKeyCharCount + keyCharCount
        cache.Entries.[key] <- ()
        while (cache.Entries.Count > replaceIdentityCacheMaxEntries
               || cache.TotalKeyCharCount > replaceIdentityCacheMaxTotalKeyChars)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if oldest <> key then
                removeReplaceIdentityCacheEntry oldest cache
        if cache.Entries.Count > replaceIdentityCacheMaxEntries
           || cache.TotalKeyCharCount > replaceIdentityCacheMaxTotalKeyChars then
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue key
            cache.Entries.[key] <- ()
            cache.TotalKeyCharCount <- keyCharCount

let private rememberSearchCellsWithSourceTextCharCount query caseSensitive (model: SpreadsheetModel) results sourceTextCharCountOpt =
    let key = struct (query, caseSensitive)
    let keyCharCount = searchCellsCacheKeyCharCount key
    match tryGetSearchCellsCache model.Cells with
    | ValueSome cached ->
        addSearchCellsCacheEntry key results cached
        rememberSearchCellsCache cached
    | ValueNone ->
        let entries = System.Collections.Generic.Dictionary<struct (string * bool), SearchCellsCacheEntry>()
        let order = System.Collections.Generic.Queue<struct (string * bool)>()
        let sourceTextCharCount =
            match sourceTextCharCountOpt with
            | ValueSome count -> count
            | ValueNone ->
                searchCellsSourceTextCharCountFor
                    results.Length
                    keyCharCount
                    model.Cells
        let cache = {
            Cells = model.Cells
            CellCount = model.Cells.Count
            SourceTextCharCount = sourceTextCharCount
            Entries = entries
            Order = order
            TotalResultCount = 0
            TotalQueryCharCount = 0
        }
        addSearchCellsCacheEntry key results cache
        rememberSearchCellsCache cache

let private rememberSearchCells query caseSensitive (model: SpreadsheetModel) results =
    rememberSearchCellsWithSourceTextCharCount query caseSensitive model results ValueNone

let private rememberReplaceIdentity findText replaceText caseSensitive (model: SpreadsheetModel) =
    let key = struct (findText, replaceText, caseSensitive)
    let keyCharCount = replaceIdentityCacheKeyCharCount key
    if model.Cells.Count <= replaceIdentityCacheMaxTotalCells
       && keyCharCount <= replaceIdentityCacheMaxTotalKeyChars then
        let sourceTextCharCount = replaceIdentityCacheSourceTextCharCountFor keyCharCount model.Cells
        if sourceTextCharCount <= replaceIdentityCacheMaxTotalSourceTextChars then
            let cache =
                match tryGetReplaceIdentityCache model.Cells with
                | ValueSome cached ->
                    cached
                | ValueNone ->
                    {
                        Cells = model.Cells
                        CellCount = max 1 model.Cells.Count
                        SourceTextCharCount = sourceTextCharCount
                        Entries = System.Collections.Generic.Dictionary<struct (string * string * bool), unit>()
                        Order = System.Collections.Generic.Queue<struct (string * string * bool)>()
                        TotalKeyCharCount = 0
                    }
            addReplaceIdentityCacheEntry key cache
            rememberReplaceIdentityCache cache

let private cachedEmptySearchResults query caseSensitive (model: SpreadsheetModel) =
    match cachedSearchCells query caseSensitive model with
    | Some results -> results.Length = 0
    | None -> false

let private rememberEmptySearchResults query caseSensitive (model: SpreadsheetModel) =
    rememberSearchCells query caseSensitive model emptySearchResults

let searchCells (query: string) (caseSensitive: bool) (model: SpreadsheetModel): CellAddress array =
    if String.IsNullOrEmpty query then
        emptySearchResults
    else
        match cachedSearchCells query caseSensitive model with
        | Some results ->
            results
        | _ ->
            let numberQueryPossible = numberDisplayMayContainSearchText query
            let booleanQueryPossible = booleanDisplayMayContainSearchText query
            let errorMatchesSearchText = errorSearchTextMatcher query caseSensitive
            let key = struct (query, caseSensitive)
            let keyCharCount = searchCellsCacheKeyCharCount key
            let trackSourceText =
                model.Cells.Count <= searchCellsCacheMaxTotalCells
                && keyCharCount <= searchCellsCacheMaxTotalQueryChars
            let mutable sourceTextCharCount =
                if trackSourceText then 0 else searchCellsCacheMaxTotalSourceTextChars + 1
            let results = System.Collections.Generic.List<CellAddress>()
            let addResult addr =
                results.Add addr
                if results.Count > searchCellsCacheMaxResultCount then
                    sourceTextCharCount <- searchCellsCacheMaxTotalSourceTextChars + 1
            let trackCellSourceText cellData =
                if sourceTextCharCount <= searchCellsCacheMaxTotalSourceTextChars
                   && results.Count <= searchCellsCacheMaxResultCount then
                    sourceTextCharCount <-
                        sourceTextAddCapped
                            searchCellsCacheMaxTotalSourceTextChars
                            (cellDataTextCharCountFor searchCellsCacheMaxTotalSourceTextChars cellData)
                            sourceTextCharCount
            if numberQueryPossible || booleanQueryPossible then
                model.Cells
                |> Map.iter (fun addr cellData ->
                    trackCellSourceText cellData
                    let cellText =
                        match cellData.Value with
                        | Formula (formula, _) -> formula  // Search in formula text
                        | other -> displayValue other

                    if containsSearchText query caseSensitive cellText then
                        addResult addr)
            else
                model.Cells
                |> Map.iter (fun addr cellData ->
                    trackCellSourceText cellData
                    match cellData.Value with
                    | Formula (formula, _) ->
                        // Search in formula text
                        if containsSearchText query caseSensitive formula then
                            addResult addr
                    | Number _ ->
                        ()
                    | Boolean _ when not booleanQueryPossible ->
                        ()
                    | Error code ->
                        if errorMatchesSearchText code then
                            addResult addr
                    | other ->
                        let cellText = displayValue other
                        if containsSearchText query caseSensitive cellText then
                            addResult addr)
            let results = results.ToArray()
            let sourceTextCharCount =
                if trackSourceText then
                    ValueSome (min sourceTextCharCount (searchCellsCacheMaxTotalSourceTextChars + 1))
                else
                    ValueNone
            rememberSearchCellsWithSourceTextCharCount query caseSensitive model results sourceTextCharCount
            results

/// Initialize search with a query
let startSearch (query: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel =
    let results = searchCells query caseSensitive model
    let sameSearch =
        match model.SearchQuery with
        | Some currentQuery when currentQuery = query ->
            match model.SearchCaseSensitive with
            | Some currentCaseSensitive -> currentCaseSensitive = caseSensitive
            | None -> false
        | _ -> false
    if sameSearch &&
       obj.ReferenceEquals(model.SearchResults, results) &&
       model.SearchIndex = 0 then
        model
    else
        { model with
            SearchQuery = Some query
            SearchCaseSensitive = Some caseSensitive
            SearchResults = results
            SearchIndex = 0
        }

/// Move to next search result
let nextSearchResultValue (model: SpreadsheetModel): SpreadsheetModel voption =
    let resultCount = model.SearchResults.Length
    if resultCount = 0 then
        ValueNone
    else
        let newIndex = (model.SearchIndex + 1) % resultCount
        let nextAddr = model.SearchResults.[newIndex]
        match model.Selection with
        | SingleCell (row, col) when model.SearchIndex = newIndex && row = nextAddr.Row && col = nextAddr.Col ->
            ValueSome model
        | _ ->
            ValueSome {
                model with
                    SearchIndex = newIndex
                    Selection = SingleCell (nextAddr.Row, nextAddr.Col)
            }

/// Move to next search result
let nextSearchResult (model: SpreadsheetModel): SpreadsheetModel option =
    match nextSearchResultValue model with
    | ValueSome next -> Some next
    | ValueNone -> None

/// Replace text at current selection
let replaceAtSelection (findText: string) (replaceText: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel option =
    if String.IsNullOrEmpty findText then
        None
    elif caseSensitive && findText = replaceText then
        None
    else
        match model.Selection with
        | SingleCell (row, col) ->
            let addr = { Row = row; Col = col }
            match getCell addr model with
            | Some cellData ->
                let currentText = displayValue cellData.Value

                if containsSearchText findText caseSensitive currentText then
                    // Perform replacement
                    let newText = replaceLiteralText findText replaceText caseSensitive currentText
                    if newText = currentText then
                        None
                    else
                        let newValue = String newText
                        let newCellData = { cellData with Value = newValue }
                        Some (setCell addr newCellData model)
                else
                    None
            | None -> None
        | _ -> None

let private replaceAllWithChangeCapture (captureChanges: bool) (findText: string) (replaceText: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel * int * ReplaceAllChanges option =
    if String.IsNullOrEmpty findText then
        model, 0, None
    elif caseSensitive && findText = replaceText then
        model, 0, None
    elif cachedEmptySearchResults findText caseSensitive model then
        model, 0, None
    elif cachedReplaceIdentity findText replaceText caseSensitive model then
        model, 0, None
    else
        let numberFindPossible = numberDisplayMayContainSearchText findText
        let booleanFindPossible = booleanDisplayMayContainSearchText findText
        let errorMatchesFindText = errorSearchTextMatcher findText caseSensitive
        let replacement = replaceText.Replace("$", "$$")
        let mutable regex = Unchecked.defaultof<System.Text.RegularExpressions.Regex>
        let mutable foundAnyMatch = false
        let changes =
            if captureChanges then
                Some {
                    OldCells = ResizeArray<CellAddress * CellData option>()
                    NewCells = ResizeArray<CellAddress * CellData option>()
                }
            else
                None

        let recordChange addr oldCell newCell =
            match changes with
            | Some changes ->
                changes.OldCells.Add(addr, Some oldCell)
                changes.NewCells.Add(addr, Some newCell)
            | None -> ()

        let replaceTextIn (currentText: string) =
            if caseSensitive then
                currentText.Replace(findText, replaceText)
            elif findText = replaceText && not (caseInsensitiveIdentityReplacementWouldChange findText currentText) then
                currentText
            else
                if obj.ReferenceEquals(regex, null) then
                    regex <- System.Text.RegularExpressions.Regex(System.Text.RegularExpressions.Regex.Escape findText, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                regex.Replace(currentText, replacement)

        let displayMayContainFindText value =
            match unwrapFormulaValue value with
            | Empty -> false
            | Number _ -> numberFindPossible
            | Boolean _ -> booleanFindPossible
            | Error code -> errorMatchesFindText code
            | String _ -> true
            | Formula _ -> true

        let primitiveSkipLikely value =
            match unwrapFormulaValue value with
            | Empty -> true
            | Number _ -> not numberFindPossible
            | Boolean _ -> not booleanFindPossible
            | Error code -> not (errorMatchesFindText code)
            | _ -> false

        let modelLikelyBenefitsFromPrimitiveSkip firstValue =
            if primitiveSkipLikely firstValue then
                true
            else
                let _, lastCell = Map.maxKeyValue model.Cells
                primitiveSkipLikely lastCell.Value

        let noChangeSearchKey = struct (findText, caseSensitive)
        let noChangeSourceTextLimit = searchCellsCacheMaxTotalSourceTextChars
        let canTrackEmptySearchSourceText =
            findText <> replaceText
            && model.Cells.Count <= searchCellsCacheMaxTotalCells
            && searchCellsCacheKeyCharCount noChangeSearchKey <= searchCellsCacheMaxTotalQueryChars
        let mutable noChangeSourceTextCharCount =
            if canTrackEmptySearchSourceText then
                0
            else
                noChangeSourceTextLimit + 1
        let mutable trackNoChangeSourceText = canTrackEmptySearchSourceText

        let trackNoChangeCellSourceText cellData =
            if trackNoChangeSourceText then
                noChangeSourceTextCharCount <-
                    sourceTextAddCapped
                        noChangeSourceTextLimit
                        (cellDataTextCharCountFor noChangeSourceTextLimit cellData)
                        noChangeSourceTextCharCount
                if noChangeSourceTextCharCount > noChangeSourceTextLimit then
                    trackNoChangeSourceText <- false

        let stopNoChangeSourceTextTracking () =
            trackNoChangeSourceText <- false

        let noChangeSourceTextCharCountOpt () =
            if canTrackEmptySearchSourceText then
                ValueSome (min noChangeSourceTextCharCount (noChangeSourceTextLimit + 1))
            else
                ValueNone

        let rememberNoChange changedCount =
            if changedCount = 0 then
                if not foundAnyMatch then
                    rememberSearchCellsWithSourceTextCharCount
                        findText
                        caseSensitive
                        model
                        emptySearchResults
                        (noChangeSourceTextCharCountOpt ())
                else
                    rememberReplaceIdentity findText replaceText caseSensitive model

        let replaceTargetedFull () =
            let mutable cells = model.Cells
            let mutable changedCount = 0
            if canTrackEmptySearchSourceText then
                model.Cells
                |> Map.iter (fun addr cellData ->
                    if trackNoChangeSourceText then
                        trackNoChangeCellSourceText cellData
                    let currentText = displayValue cellData.Value
                    if containsSearchText findText caseSensitive currentText then
                        foundAnyMatch <- true
                        stopNoChangeSourceTextTracking ()
                        let newText = replaceTextIn currentText
                        if newText <> currentText then
                            let newValue = String newText
                            let newCellData = { cellData with Value = newValue }
                            recordChange addr cellData newCellData
                            cells <- Map.add addr newCellData cells
                            changedCount <- changedCount + 1)
            else
                model.Cells
                |> Map.iter (fun addr cellData ->
                    let currentText = displayValue cellData.Value
                    if containsSearchText findText caseSensitive currentText then
                        foundAnyMatch <- true
                        let newText = replaceTextIn currentText
                        if newText <> currentText then
                            let newValue = String newText
                            let newCellData = { cellData with Value = newValue }
                            recordChange addr cellData newCellData
                            cells <- Map.add addr newCellData cells
                            changedCount <- changedCount + 1)

            rememberNoChange changedCount
            if changedCount > 0 then { model with Cells = cells }, changedCount, changes else model, 0, changes

        let replaceTargetedWithPrimitiveSkip () =
            let mutable cells = model.Cells
            let mutable changedCount = 0
            if canTrackEmptySearchSourceText then
                model.Cells
                |> Map.iter (fun addr cellData ->
                    if trackNoChangeSourceText then
                        trackNoChangeCellSourceText cellData
                    if displayMayContainFindText cellData.Value then
                        let currentText = displayValue cellData.Value
                        if containsSearchText findText caseSensitive currentText then
                            foundAnyMatch <- true
                            stopNoChangeSourceTextTracking ()
                            let newText = replaceTextIn currentText
                            if newText <> currentText then
                                let newValue = String newText
                                let newCellData = { cellData with Value = newValue }
                                recordChange addr cellData newCellData
                                cells <- Map.add addr newCellData cells
                                changedCount <- changedCount + 1)
            else
                model.Cells
                |> Map.iter (fun addr cellData ->
                    if displayMayContainFindText cellData.Value then
                        let currentText = displayValue cellData.Value
                        if containsSearchText findText caseSensitive currentText then
                            foundAnyMatch <- true
                            let newText = replaceTextIn currentText
                            if newText <> currentText then
                                let newValue = String newText
                                let newCellData = { cellData with Value = newValue }
                                recordChange addr cellData newCellData
                                cells <- Map.add addr newCellData cells
                                changedCount <- changedCount + 1)

            rememberNoChange changedCount
            if changedCount > 0 then { model with Cells = cells }, changedCount, changes else model, 0, changes

        if model.Cells.Count >= 4096 then
            let _, firstCell = Map.minKeyValue model.Cells
            let firstMatches =
                displayMayContainFindText firstCell.Value &&
                containsSearchText findText caseSensitive (displayValue firstCell.Value)
            let lastMatches =
                if firstMatches then
                    let _, lastCell = Map.maxKeyValue model.Cells
                    displayMayContainFindText lastCell.Value &&
                    containsSearchText findText caseSensitive (displayValue lastCell.Value)
                else
                    false
            let denseSampleMatches =
                if firstMatches && lastMatches then
                    let mutable checkedCount = 0
                    let mutable allMatch = true
                    use e = (model.Cells :> seq<_>).GetEnumerator()
                    while allMatch && checkedCount < 64 && e.MoveNext() do
                        let pair = e.Current
                        if not (displayMayContainFindText pair.Value.Value &&
                                containsSearchText findText caseSensitive (displayValue pair.Value.Value)) then
                            allMatch <- false
                        checkedCount <- checkedCount + 1
                    allMatch
                else
                    false
            if denseSampleMatches then
                let mutable sampleHasChange = false
                let mutable checkedCount = 0
                use e = (model.Cells :> seq<_>).GetEnumerator()
                while not sampleHasChange && checkedCount < 64 && e.MoveNext() do
                    let pair = e.Current
                    if displayMayContainFindText pair.Value.Value then
                        let currentText = displayValue pair.Value.Value
                        if containsSearchText findText caseSensitive currentText then
                            sampleHasChange <- replaceTextIn currentText <> currentText
                    checkedCount <- checkedCount + 1
                if sampleHasChange then
                    let mutable changedCount = 0
                    let cells =
                        model.Cells
                        |> Map.map (fun addr cellData ->
                            if displayMayContainFindText cellData.Value then
                                let currentText = displayValue cellData.Value
                                if containsSearchText findText caseSensitive currentText then
                                    foundAnyMatch <- true
                                    let newText = replaceTextIn currentText
                                    if newText <> currentText then
                                        stopNoChangeSourceTextTracking ()
                                        changedCount <- changedCount + 1
                                        let newCellData = { cellData with Value = String newText }
                                        recordChange addr cellData newCellData
                                        newCellData
                                    else
                                        cellData
                                else
                                    cellData
                            else
                                cellData)
                    rememberNoChange changedCount
                    if changedCount > 0 then { model with Cells = cells }, changedCount, changes else model, 0, changes
                else
                    replaceTargetedFull ()
            else
                if modelLikelyBenefitsFromPrimitiveSkip firstCell.Value then
                    replaceTargetedWithPrimitiveSkip ()
                else
                    replaceTargetedFull ()
        else
            replaceTargetedFull ()

/// Replace all occurrences and return the number of cells changed.
let replaceAllWithCount (findText: string) (replaceText: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel * int =
    let updated, count, _ = replaceAllWithChangeCapture false findText replaceText caseSensitive model
    updated, count

/// Replace all occurrences and return the changed old/new cells for sparse undo.
let replaceAllWithChanges (findText: string) (replaceText: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel * int * ResizeArray<CellAddress * CellData option> * ResizeArray<CellAddress * CellData option> =
    let updated, count, changes = replaceAllWithChangeCapture true findText replaceText caseSensitive model
    match changes with
    | Some changes ->
        updated, count, changes.OldCells, changes.NewCells
    | None ->
        updated, count, ResizeArray<CellAddress * CellData option>(), ResizeArray<CellAddress * CellData option>()

/// Replace all occurrences
let replaceAll (findText: string) (replaceText: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel =
    let updated, _ = replaceAllWithCount findText replaceText caseSensitive model
    updated

/// Evaluate if a cell value matches a conditional format rule
let evaluateRule (rule: ConditionalFormatRule) (value: CellValue): bool =
    match rule, value with
    | GreaterThan threshold, Number n -> n > threshold
    | LessThan threshold, Number n -> n < threshold
    | Between (min, max), Number n -> n >= min && n <= max
    | Equal target, Number n -> n = target
    | ContainsText text, String s -> s.Contains(text)
    | ContainsText text, Formula (formula, _) -> formula.Contains(text)
    | _ -> false

/// Get conditional format to apply to a cell (returns first matching rule)
let getConditionalFormat (addr: CellAddress) (value: CellValue) (model: SpreadsheetModel): CellFormat option =
    match Map.tryFind addr model.ConditionalFormats with
    | Some formats ->
        let mutable found = None
        let mutable remaining = formats
        while found.IsNone && not remaining.IsEmpty do
            match remaining with
            | cf :: tail ->
                if evaluateRule cf.Rule value then
                    found <- Some cf.Style
                remaining <- tail
            | [] -> ()
        found
    | None -> None

/// Add a conditional format rule to a cell
let addConditionalFormat (addr: CellAddress) (format: ConditionalFormat) (model: SpreadsheetModel): SpreadsheetModel =
    let existingFormats = Map.tryFind addr model.ConditionalFormats |> Option.defaultValue []
    match existingFormats with
    | head :: _ when head = format ->
        model
    | _ ->
        let newFormats = format :: existingFormats
        { model with ConditionalFormats = Map.add addr newFormats model.ConditionalFormats }

/// Remove all conditional formats from a cell
let removeConditionalFormats (addr: CellAddress) (model: SpreadsheetModel): SpreadsheetModel =
    if Map.containsKey addr model.ConditionalFormats then
        { model with ConditionalFormats = Map.remove addr model.ConditionalFormats }
    else
        model

// --- Conditional-format VISUALS: data bars / color scales / icon sets -------
// These compute over a range and bake the result into each cell's CellFormat
// (DataBar / IconSet / BgColor), which the renderers already paint. Applying
// them is a range operation (from the CF dialog or programmatically).

let private unwrapVal (v: CellValue) : CellValue =
    unwrapFormulaValue v

let private iterRangeNumbers sr sc er ec (model: SpreadsheetModel) emit =
    let rowCount = max 0 (er - sr + 1)
    let colCount = max 0 (ec - sc + 1)
    let area = int64 rowCount * int64 colCount
    let scanThreshold = max 4096L (int64 model.Cells.Count / 16L)
    if area > scanThreshold then
        model.Cells
        |> Map.iter (fun addr cellData ->
            if addr.Row >= sr && addr.Row <= er && addr.Col >= sc && addr.Col <= ec then
                match unwrapVal cellData.Value with
                | Number n -> emit addr n cellData
                | Boolean b -> emit addr (if b then 1.0 else 0.0) cellData
                | _ -> ())
    else
        for r in sr .. er do
            for c in sc .. ec do
                let addr = { Row = r; Col = c }
                match getCell addr model with
                | Some cellData ->
                    match unwrapVal cellData.Value with
                    | Number n -> emit addr n cellData
                    | Boolean b -> emit addr (if b then 1.0 else 0.0) cellData
                    | _ -> ()
                | None -> ()

let private rangeNumbersWithBounds sr sc er ec (model: SpreadsheetModel) : (ResizeArray<CellAddress * float * CellData> * float * float) option =
    let cells = ResizeArray<CellAddress * float * CellData>()
    let mutable lo = 0.0
    let mutable hi = 0.0
    let mutable found = false
    iterRangeNumbers sr sc er ec model (fun addr n cellData ->
        cells.Add((addr, n, cellData))
        if found then
            if n < lo then lo <- n
            if n > hi then hi <- n
        else
            lo <- n
            hi <- n
            found <- true)
    if found then Some (cells, lo, hi) else None

let private rangeNumberStats sr sc er ec (model: SpreadsheetModel) : (int * float * float) option =
    let mutable count = 0
    let mutable lo = 0.0
    let mutable hi = 0.0
    iterRangeNumbers sr sc er ec model (fun _ n _ ->
        if count = 0 then
            lo <- n
            hi <- n
        else
            if n < lo then lo <- n
            if n > hi then hi <- n
        count <- count + 1)
    if count = 0 then None else Some (count, lo, hi)

let private rangeNumbersHaveAnyFormatChange sr sc er ec (model: SpreadsheetModel) (formatForValue: float -> CellData -> CellFormat) =
    let rowCount = max 0 (er - sr + 1)
    let colCount = max 0 (ec - sc + 1)
    let area = int64 rowCount * int64 colCount
    let scanThreshold = max 4096L (int64 model.Cells.Count / 16L)
    let mutable changed = false
    let checkCell value cellData =
        if formatForValue value cellData <> cellData.Format then
            changed <- true
    if area > scanThreshold then
        use entries = (model.Cells :> seq<_>).GetEnumerator()
        while not changed && entries.MoveNext() do
            let pair = entries.Current
            let addr = pair.Key
            if addr.Row >= sr && addr.Row <= er && addr.Col >= sc && addr.Col <= ec then
                let cellData = pair.Value
                match unwrapVal cellData.Value with
                | Number n -> checkCell n cellData
                | Boolean b -> checkCell (if b then 1.0 else 0.0) cellData
                | _ -> ()
    else
        let mutable row = sr
        while not changed && row <= er do
            let mutable col = sc
            while not changed && col <= ec do
                let addr = { Row = row; Col = col }
                match getCell addr model with
                | Some cellData ->
                    match unwrapVal cellData.Value with
                    | Number n -> checkCell n cellData
                    | Boolean b -> checkCell (if b then 1.0 else 0.0) cellData
                    | _ -> ()
                | None -> ()
                col <- col + 1
            row <- row + 1
    changed

let private hexDigit (ch: char) =
    if ch >= '0' && ch <= '9' then int ch - int '0'
    elif ch >= 'A' && ch <= 'F' then int ch - int 'A' + 10
    elif ch >= 'a' && ch <= 'f' then int ch - int 'a' + 10
    else -1

let private hexByteAt (text: string) index =
    let hi = hexDigit text.[index]
    let lo = hexDigit text.[index + 1]
    if hi < 0 || lo < 0 then None else Some ((hi <<< 4) ||| lo)

let private hexToRgb (s: string) =
    let h = (if s.StartsWith "#" then s.Substring 1 else s)
    if h.Length = 6 then
        match hexByteAt h 0, hexByteAt h 2, hexByteAt h 4 with
        | Some r, Some g, Some b -> r, g, b
        | _ -> 255, 255, 255
    else (255, 255, 255)

let private rgbToHex r g b =
    let hex = "0123456789ABCDEF"
    System.String [|
        '#'
        hex.[r >>> 4]; hex.[r &&& 15]
        hex.[g >>> 4]; hex.[g &&& 15]
        hex.[b >>> 4]; hex.[b &&& 15] |]

let private lerpRgb ((ar, ag, ab): int * int * int) ((br, bg, bb): int * int * int) (t: float) : string =
    let m x y = int (float x + (float y - float x) * t)
    rgbToHex (m ar br) (m ag bg) (m ab bb)

let private hasAnyFormatChange (cells: ResizeArray<CellAddress * float * CellData>) formatForValue =
    let mutable changed = false
    let mutable index = 0
    while not changed && index < cells.Count do
        let _, value, cell = cells.[index]
        if formatForValue value cell <> cell.Format then
            changed <- true
        index <- index + 1
    changed

type private DataBarApplyCache = {
    Cells: Map<CellAddress, CellData>
    CellCount: int
    SourceTextCharCount: int
    Entries: System.Collections.Generic.Dictionary<DataBarApplyCacheKey, unit>
    Order: System.Collections.Generic.Queue<DataBarApplyCacheKey>
    mutable TotalKeyCharCount: int
}

and private DataBarApplyCacheKey = {
    StartRow: int
    StartCol: int
    EndRow: int
    EndCol: int
    Color: string
}

type private IconSetApplyCache = {
    Cells: Map<CellAddress, CellData>
    CellCount: int
    SourceTextCharCount: int
    Entries: System.Collections.Generic.Dictionary<IconSetApplyCacheKey, unit>
    Order: System.Collections.Generic.Queue<IconSetApplyCacheKey>
}

and private IconSetApplyCacheKey = {
    StartRow: int
    StartCol: int
    EndRow: int
    EndCol: int
    Count: int
}

type private ColorScaleApplyCache = {
    Cells: Map<CellAddress, CellData>
    CellCount: int
    SourceTextCharCount: int
    Entries: System.Collections.Generic.Dictionary<ColorScaleApplyCacheKey, unit>
    Order: System.Collections.Generic.Queue<ColorScaleApplyCacheKey>
    mutable TotalKeyCharCount: int
}

and private ColorScaleApplyCacheKey = {
    StartRow: int
    StartCol: int
    EndRow: int
    EndCol: int
    LoColor: string
    MidColor: string option
    HiColor: string
}

let mutable private dataBarApplyCache: DataBarApplyCache list = []
let mutable private iconSetApplyCache: IconSetApplyCache list = []
let mutable private colorScaleApplyCache: ColorScaleApplyCache list = []
let private dataBarApplyCacheByCells =
    System.Collections.Generic.Dictionary<Map<CellAddress, CellData>, DataBarApplyCache>(cellMapReferenceComparer)
let private dataBarApplyFastSlots = Array.zeroCreate<DataBarApplyCache> 256
let private iconSetApplyCacheByCells =
    System.Collections.Generic.Dictionary<Map<CellAddress, CellData>, IconSetApplyCache>(cellMapReferenceComparer)
let private iconSetApplyFastSlots = Array.zeroCreate<IconSetApplyCache> 256
let private colorScaleApplyCacheByCells =
    System.Collections.Generic.Dictionary<Map<CellAddress, CellData>, ColorScaleApplyCache>(cellMapReferenceComparer)
let private colorScaleApplyFastSlots = Array.zeroCreate<ColorScaleApplyCache> 256
let private conditionalFormatApplyCacheMaxEntries = 64
let private conditionalFormatApplyCacheMaxSources = 64
let private conditionalFormatApplyCacheMaxTotalCells = 200000
let private conditionalFormatApplyCacheMaxTotalSourceTextChars = 5000000
let private conditionalFormatApplyCacheMaxTotalKeyChars = 5000000

let private conditionalFormatApplyCacheCellCount cells =
    max 1 (Map.count cells)

let private conditionalFormatApplyCacheSourceTextCharCount cells =
    cellMapSourceTextCharCountFor
        conditionalFormatApplyCacheMaxTotalCells
        conditionalFormatApplyCacheMaxTotalSourceTextChars
        cells

let private dataBarApplyCacheMatches cells (cached: DataBarApplyCache) =
    obj.ReferenceEquals(cached.Cells, cells)

let private iconSetApplyCacheMatches cells (cached: IconSetApplyCache) =
    obj.ReferenceEquals(cached.Cells, cells)

let private colorScaleApplyCacheMatches cells (cached: ColorScaleApplyCache) =
    obj.ReferenceEquals(cached.Cells, cells)

let private setDataBarApplyCache entries =
    dataBarApplyCache <- entries
    dataBarApplyCacheByCells.Clear()
    clearArray dataBarApplyFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<DataBarApplyCache>
        if not (dataBarApplyCacheByCells.TryGetValue(entry.Cells, &existing)) then
            dataBarApplyCacheByCells.[entry.Cells] <- entry
            let slot = referenceSlot entry.Cells dataBarApplyFastSlots.Length
            dataBarApplyFastSlots.[slot] <- entry

let private setIconSetApplyCache entries =
    iconSetApplyCache <- entries
    iconSetApplyCacheByCells.Clear()
    clearArray iconSetApplyFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<IconSetApplyCache>
        if not (iconSetApplyCacheByCells.TryGetValue(entry.Cells, &existing)) then
            iconSetApplyCacheByCells.[entry.Cells] <- entry
            let slot = referenceSlot entry.Cells iconSetApplyFastSlots.Length
            iconSetApplyFastSlots.[slot] <- entry

let private setColorScaleApplyCache entries =
    colorScaleApplyCache <- entries
    colorScaleApplyCacheByCells.Clear()
    clearArray colorScaleApplyFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<ColorScaleApplyCache>
        if not (colorScaleApplyCacheByCells.TryGetValue(entry.Cells, &existing)) then
            colorScaleApplyCacheByCells.[entry.Cells] <- entry
            let slot = referenceSlot entry.Cells colorScaleApplyFastSlots.Length
            colorScaleApplyFastSlots.[slot] <- entry

let private tryGetDataBarApplyCache cells =
    let slot = referenceSlot cells dataBarApplyFastSlots.Length
    let fast = dataBarApplyFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && dataBarApplyCacheMatches cells fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<DataBarApplyCache>
        if dataBarApplyCacheByCells.TryGetValue(cells, &cached) then
            dataBarApplyFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private tryGetIconSetApplyCache cells =
    let slot = referenceSlot cells iconSetApplyFastSlots.Length
    let fast = iconSetApplyFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && iconSetApplyCacheMatches cells fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<IconSetApplyCache>
        if iconSetApplyCacheByCells.TryGetValue(cells, &cached) then
            iconSetApplyFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private tryGetColorScaleApplyCache cells =
    let slot = referenceSlot cells colorScaleApplyFastSlots.Length
    let fast = colorScaleApplyFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && colorScaleApplyCacheMatches cells fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<ColorScaleApplyCache>
        if colorScaleApplyCacheByCells.TryGetValue(cells, &cached) then
            colorScaleApplyFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private dataBarApplyCacheWithoutCells cells =
    let mutable cached = Unchecked.defaultof<DataBarApplyCache>
    if dataBarApplyCacheByCells.TryGetValue(cells, &cached) then
        dataBarApplyCache
        |> List.filter (dataBarApplyCacheMatches cells >> not)
    else
        dataBarApplyCache

let private removeDataBarApplyCacheForCells cells =
    let entries = dataBarApplyCacheWithoutCells cells
    if not (obj.ReferenceEquals(entries, dataBarApplyCache)) then
        setDataBarApplyCache entries

let private iconSetApplyCacheWithoutCells cells =
    let mutable cached = Unchecked.defaultof<IconSetApplyCache>
    if iconSetApplyCacheByCells.TryGetValue(cells, &cached) then
        iconSetApplyCache
        |> List.filter (iconSetApplyCacheMatches cells >> not)
    else
        iconSetApplyCache

let private removeIconSetApplyCacheForCells cells =
    let entries = iconSetApplyCacheWithoutCells cells
    if not (obj.ReferenceEquals(entries, iconSetApplyCache)) then
        setIconSetApplyCache entries

let private colorScaleApplyCacheWithoutCells cells =
    let mutable cached = Unchecked.defaultof<ColorScaleApplyCache>
    if colorScaleApplyCacheByCells.TryGetValue(cells, &cached) then
        colorScaleApplyCache
        |> List.filter (colorScaleApplyCacheMatches cells >> not)
    else
        colorScaleApplyCache

let private removeColorScaleApplyCacheForCells cells =
    let entries = colorScaleApplyCacheWithoutCells cells
    if not (obj.ReferenceEquals(entries, colorScaleApplyCache)) then
        setColorScaleApplyCache entries

let private dataBarApplyCacheRetainable (cached: DataBarApplyCache) =
    cached.Entries.Count > 0
    && cached.CellCount <= conditionalFormatApplyCacheMaxTotalCells
    && cached.SourceTextCharCount <= conditionalFormatApplyCacheMaxTotalSourceTextChars
    && cached.TotalKeyCharCount <= conditionalFormatApplyCacheMaxTotalKeyChars

let private iconSetApplyCacheRetainable (cached: IconSetApplyCache) =
    cached.Entries.Count > 0
    && cached.CellCount <= conditionalFormatApplyCacheMaxTotalCells
    && cached.SourceTextCharCount <= conditionalFormatApplyCacheMaxTotalSourceTextChars

let private colorScaleApplyCacheRetainable (cached: ColorScaleApplyCache) =
    cached.Entries.Count > 0
    && cached.CellCount <= conditionalFormatApplyCacheMaxTotalCells
    && cached.SourceTextCharCount <= conditionalFormatApplyCacheMaxTotalSourceTextChars
    && cached.TotalKeyCharCount <= conditionalFormatApplyCacheMaxTotalKeyChars

let private rememberDataBarApplyCache cached =
    if not (dataBarApplyCacheRetainable cached) then
        removeDataBarApplyCacheForCells cached.Cells
    else
        let trim (entries: DataBarApplyCache list) =
            let rec loop keptCount keptCells keptSourceTextChars keptKeyChars (acc: DataBarApplyCache list) (remaining: DataBarApplyCache list) : DataBarApplyCache list =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateCells = max 1 candidate.CellCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    let candidateKeyChars = max 1 candidate.TotalKeyCharCount
                    if dataBarApplyCacheRetainable candidate
                       && keptCount < conditionalFormatApplyCacheMaxSources
                       && keptCells + candidateCells <= conditionalFormatApplyCacheMaxTotalCells
                       && keptSourceTextChars + candidateSourceTextChars <= conditionalFormatApplyCacheMaxTotalSourceTextChars
                       && keptKeyChars + candidateKeyChars <= conditionalFormatApplyCacheMaxTotalKeyChars then
                        loop
                            (keptCount + 1)
                            (keptCells + candidateCells)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (keptKeyChars + candidateKeyChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptCells keptSourceTextChars keptKeyChars acc tail
            loop 0 0 0 0 [] entries
        (cached
         :: dataBarApplyCacheWithoutCells cached.Cells)
        |> trim
        |> setDataBarApplyCache

let private rememberIconSetApplyCache cached =
    if not (iconSetApplyCacheRetainable cached) then
        removeIconSetApplyCacheForCells cached.Cells
    else
        let trim (entries: IconSetApplyCache list) =
            let rec loop keptCount keptCells keptSourceTextChars (acc: IconSetApplyCache list) (remaining: IconSetApplyCache list) : IconSetApplyCache list =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateCells = max 1 candidate.CellCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    if iconSetApplyCacheRetainable candidate
                       && keptCount < conditionalFormatApplyCacheMaxSources
                       && keptCells + candidateCells <= conditionalFormatApplyCacheMaxTotalCells
                       && keptSourceTextChars + candidateSourceTextChars <= conditionalFormatApplyCacheMaxTotalSourceTextChars then
                        loop (keptCount + 1) (keptCells + candidateCells) (keptSourceTextChars + candidateSourceTextChars) (candidate :: acc) tail
                    else
                        loop keptCount keptCells keptSourceTextChars acc tail
            loop 0 0 0 [] entries
        (cached
         :: iconSetApplyCacheWithoutCells cached.Cells)
        |> trim
        |> setIconSetApplyCache

let private rememberColorScaleApplyCache cached =
    if not (colorScaleApplyCacheRetainable cached) then
        removeColorScaleApplyCacheForCells cached.Cells
    else
        let trim (entries: ColorScaleApplyCache list) =
            let rec loop keptCount keptCells keptSourceTextChars keptKeyChars (acc: ColorScaleApplyCache list) (remaining: ColorScaleApplyCache list) : ColorScaleApplyCache list =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateCells = max 1 candidate.CellCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    let candidateKeyChars = max 1 candidate.TotalKeyCharCount
                    if colorScaleApplyCacheRetainable candidate
                       && keptCount < conditionalFormatApplyCacheMaxSources
                       && keptCells + candidateCells <= conditionalFormatApplyCacheMaxTotalCells
                       && keptSourceTextChars + candidateSourceTextChars <= conditionalFormatApplyCacheMaxTotalSourceTextChars
                       && keptKeyChars + candidateKeyChars <= conditionalFormatApplyCacheMaxTotalKeyChars then
                        loop
                            (keptCount + 1)
                            (keptCells + candidateCells)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (keptKeyChars + candidateKeyChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptCells keptSourceTextChars keptKeyChars acc tail
            loop 0 0 0 0 [] entries
        (cached
         :: colorScaleApplyCacheWithoutCells cached.Cells)
        |> trim
        |> setColorScaleApplyCache

let private dataBarApplyCacheKeyCharCount (key: DataBarApplyCacheKey) =
    key.Color.Length

let private colorScaleApplyCacheKeyCharCount (key: ColorScaleApplyCacheKey) =
    key.LoColor.Length
    + (match key.MidColor with Some color -> color.Length | None -> 0)
    + key.HiColor.Length

let private addConditionalFormatApplyCacheEntry key (entries: System.Collections.Generic.Dictionary<'Key, unit>) (order: System.Collections.Generic.Queue<'Key>) =
    if not (entries.ContainsKey key) then
        order.Enqueue key
    entries.[key] <- ()
    while entries.Count > conditionalFormatApplyCacheMaxEntries && order.Count > 0 do
        let oldest = order.Dequeue()
        if oldest <> key then
            entries.Remove oldest |> ignore
    if entries.Count > conditionalFormatApplyCacheMaxEntries then
        entries.Clear()
        order.Clear()
        order.Enqueue key
        entries.[key] <- ()

let private removeDataBarApplyCacheEntry key (cache: DataBarApplyCache) =
    if cache.Entries.Remove key then
        cache.TotalKeyCharCount <- max 0 (cache.TotalKeyCharCount - dataBarApplyCacheKeyCharCount key)

let private addDataBarApplyCacheEntry key (cache: DataBarApplyCache) =
    let keyCharCount = dataBarApplyCacheKeyCharCount key
    if keyCharCount > conditionalFormatApplyCacheMaxTotalKeyChars then
        removeDataBarApplyCacheEntry key cache
    else
        if not (cache.Entries.ContainsKey key) then
            cache.Order.Enqueue key
            cache.TotalKeyCharCount <- cache.TotalKeyCharCount + keyCharCount
        cache.Entries.[key] <- ()
        while (cache.Entries.Count > conditionalFormatApplyCacheMaxEntries
               || cache.TotalKeyCharCount > conditionalFormatApplyCacheMaxTotalKeyChars)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if oldest <> key then
                removeDataBarApplyCacheEntry oldest cache
        if cache.Entries.Count > conditionalFormatApplyCacheMaxEntries
           || cache.TotalKeyCharCount > conditionalFormatApplyCacheMaxTotalKeyChars then
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue key
            cache.Entries.[key] <- ()
            cache.TotalKeyCharCount <- keyCharCount

let private removeColorScaleApplyCacheEntry key (cache: ColorScaleApplyCache) =
    if cache.Entries.Remove key then
        cache.TotalKeyCharCount <- max 0 (cache.TotalKeyCharCount - colorScaleApplyCacheKeyCharCount key)

let private addColorScaleApplyCacheEntry key (cache: ColorScaleApplyCache) =
    let keyCharCount = colorScaleApplyCacheKeyCharCount key
    if keyCharCount > conditionalFormatApplyCacheMaxTotalKeyChars then
        removeColorScaleApplyCacheEntry key cache
    else
        if not (cache.Entries.ContainsKey key) then
            cache.Order.Enqueue key
            cache.TotalKeyCharCount <- cache.TotalKeyCharCount + keyCharCount
        cache.Entries.[key] <- ()
        while (cache.Entries.Count > conditionalFormatApplyCacheMaxEntries
               || cache.TotalKeyCharCount > conditionalFormatApplyCacheMaxTotalKeyChars)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if oldest <> key then
                removeColorScaleApplyCacheEntry oldest cache
        if cache.Entries.Count > conditionalFormatApplyCacheMaxEntries
           || cache.TotalKeyCharCount > conditionalFormatApplyCacheMaxTotalKeyChars then
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue key
            cache.Entries.[key] <- ()
            cache.TotalKeyCharCount <- keyCharCount

/// Data bars over a range, baselined at 0 when all values are ≥ 0 (Excel's
/// automatic behaviour), else min→max.
let applyDataBars (sr: int) (sc: int) (er: int) (ec: int) (color: string) (model: SpreadsheetModel) : SpreadsheetModel =
    let key = {
        StartRow = sr
        StartCol = sc
        EndRow = er
        EndCol = ec
        Color = color
    }
    let rememberWithSourceTextCharCount sourceTextCharCount cells =
        let cache =
            match tryGetDataBarApplyCache cells with
            | ValueSome cached -> cached
            | ValueNone ->
                {
                    Cells = cells
                    CellCount = conditionalFormatApplyCacheCellCount cells
                    SourceTextCharCount = sourceTextCharCount
                    Entries = System.Collections.Generic.Dictionary<DataBarApplyCacheKey, unit>()
                    Order = System.Collections.Generic.Queue<DataBarApplyCacheKey>()
                    TotalKeyCharCount = 0
                }
        addDataBarApplyCacheEntry key cache
        rememberDataBarApplyCache cache
    let remember cells =
        rememberWithSourceTextCharCount (conditionalFormatApplyCacheSourceTextCharCount cells) cells
    let rememberChangedResult cells =
        rememberWithSourceTextCharCount (conditionalFormatApplyCacheMaxTotalSourceTextChars + 1) cells
    let compute () =
        match rangeNumberStats sr sc er ec model with
        | None ->
            remember model.Cells
            model
        | Some (numberCount, minValue, maxValue) ->
            let lo = min 0.0 minValue
            let hi = max 0.0 maxValue
            let span = if hi - lo < 1e-12 then 1.0 else hi - lo
            let dataBarFormat value cell =
                { cell.Format with DataBar = Some(max 0.0 (min 1.0 ((value - lo) / span)), color) }
            if int64 numberCount * 2L >= int64 model.Cells.Count then
                if not (rangeNumbersHaveAnyFormatChange sr sc er ec model dataBarFormat) then
                    remember model.Cells
                    model
                else
                    let cellMap =
                        model.Cells
                        |> Map.map (fun addr cell ->
                            if addr.Row >= sr && addr.Row <= er && addr.Col >= sc && addr.Col <= ec then
                                match unwrapVal cell.Value with
                                | Number n ->
                                    { cell with Format = dataBarFormat n cell }
                                | Boolean true ->
                                    { cell with Format = dataBarFormat 1.0 cell }
                                | Boolean false ->
                                    { cell with Format = dataBarFormat 0.0 cell }
                                | _ -> cell
                            else
                                cell)
                    let updated = { model with Cells = cellMap }
                    rememberChangedResult updated.Cells
                    updated
            else
                let cells =
                    match rangeNumbersWithBounds sr sc er ec model with
                    | Some (cells, _, _) -> cells
                    | None -> ResizeArray<CellAddress * float * CellData>()
                if not (hasAnyFormatChange cells dataBarFormat) then
                    remember model.Cells
                    model
                else
                    let mutable cellMap = model.Cells
                    let mutable changed = false
                    for addr, v, cell in cells do
                        let newFormat = dataBarFormat v cell
                        if newFormat <> cell.Format then
                            cellMap <- Map.add addr { cell with Format = newFormat } cellMap
                            changed <- true
                    if changed then
                        let updated = { model with Cells = cellMap }
                        rememberChangedResult updated.Cells
                        updated
                    else
                        remember model.Cells
                        model
    match tryGetDataBarApplyCache model.Cells with
    | ValueSome cached when cached.Entries.ContainsKey key -> model
    | _ ->
        compute ()

/// 2- or 3-color scale (pass `mid = None` for 2-color) → per-cell BgColor.
let applyColorScale (sr: int) (sc: int) (er: int) (ec: int) (loColor: string) (mid: string option) (hiColor: string) (model: SpreadsheetModel) : SpreadsheetModel =
    let key = {
        StartRow = sr
        StartCol = sc
        EndRow = er
        EndCol = ec
        LoColor = loColor
        MidColor = mid
        HiColor = hiColor
    }
    let rememberWithSourceTextCharCount sourceTextCharCount cells =
        let cache =
            match tryGetColorScaleApplyCache cells with
            | ValueSome cached -> cached
            | ValueNone ->
                {
                    Cells = cells
                    CellCount = conditionalFormatApplyCacheCellCount cells
                    SourceTextCharCount = sourceTextCharCount
                    Entries = System.Collections.Generic.Dictionary<ColorScaleApplyCacheKey, unit>()
                    Order = System.Collections.Generic.Queue<ColorScaleApplyCacheKey>()
                    TotalKeyCharCount = 0
                }
        addColorScaleApplyCacheEntry key cache
        rememberColorScaleApplyCache cache
    let remember cells =
        rememberWithSourceTextCharCount (conditionalFormatApplyCacheSourceTextCharCount cells) cells
    let rememberChangedResult cells =
        rememberWithSourceTextCharCount (conditionalFormatApplyCacheMaxTotalSourceTextChars + 1) cells
    let compute () =
        match rangeNumberStats sr sc er ec model with
        | None ->
            remember model.Cells
            model
        | Some (numberCount, lo, hi) ->
            let span = if hi - lo < 1e-12 then 1.0 else hi - lo
            let loRgb = hexToRgb loColor
            let hiRgb = hexToRgb hiColor
            let midRgb = mid |> Option.map hexToRgb
            let scaleColor value =
                let t = (value - lo) / span
                match midRgb with
                | None -> lerpRgb loRgb hiRgb t
                | Some mc -> if t <= 0.5 then lerpRgb loRgb mc (t * 2.0) else lerpRgb mc hiRgb ((t - 0.5) * 2.0)
            let colorScaleFormat value cell =
                { cell.Format with BgColor = Some (scaleColor value) }
            if int64 numberCount * 2L >= int64 model.Cells.Count then
                if not (rangeNumbersHaveAnyFormatChange sr sc er ec model colorScaleFormat) then
                    remember model.Cells
                    model
                else
                    let cellMap =
                        model.Cells
                        |> Map.map (fun addr cell ->
                            if addr.Row >= sr && addr.Row <= er && addr.Col >= sc && addr.Col <= ec then
                                match unwrapVal cell.Value with
                                | Number n ->
                                    { cell with Format = colorScaleFormat n cell }
                                | Boolean true ->
                                    { cell with Format = colorScaleFormat 1.0 cell }
                                | Boolean false ->
                                    { cell with Format = colorScaleFormat 0.0 cell }
                                | _ -> cell
                            else
                                cell)
                    let updated = { model with Cells = cellMap }
                    rememberChangedResult updated.Cells
                    updated
            else
                let cells =
                    match rangeNumbersWithBounds sr sc er ec model with
                    | Some (cells, _, _) -> cells
                    | None -> ResizeArray<CellAddress * float * CellData>()
                if not (hasAnyFormatChange cells colorScaleFormat) then
                    remember model.Cells
                    model
                else
                    let mutable cellMap = model.Cells
                    let mutable changed = false
                    for addr, v, cell in cells do
                        let newFormat = colorScaleFormat v cell
                        if newFormat <> cell.Format then
                            cellMap <- Map.add addr { cell with Format = newFormat } cellMap
                            changed <- true
                    if changed then
                        let updated = { model with Cells = cellMap }
                        rememberChangedResult updated.Cells
                        updated
                    else
                        remember model.Cells
                        model
    match tryGetColorScaleApplyCache model.Cells with
    | ValueSome cached when cached.Entries.ContainsKey key -> model
    | _ ->
        compute ()

/// Icon set over a range: `count` buckets by value (Excel's percent
/// thresholds), each cell tagged with its bucket index for the renderer.
let applyIconSet (sr: int) (sc: int) (er: int) (ec: int) (count: int) (model: SpreadsheetModel) : SpreadsheetModel =
    if count <= 1 then model
    else
        let key = {
            StartRow = sr
            StartCol = sc
            EndRow = er
            EndCol = ec
            Count = count
        }
        let rememberWithSourceTextCharCount sourceTextCharCount cells =
            let cache =
                match tryGetIconSetApplyCache cells with
                | ValueSome cached -> cached
                | ValueNone ->
                    {
                        Cells = cells
                        CellCount = conditionalFormatApplyCacheCellCount cells
                        SourceTextCharCount = sourceTextCharCount
                        Entries = System.Collections.Generic.Dictionary<IconSetApplyCacheKey, unit>()
                        Order = System.Collections.Generic.Queue<IconSetApplyCacheKey>()
                    }
            addConditionalFormatApplyCacheEntry key cache.Entries cache.Order
            rememberIconSetApplyCache cache
        let remember cells =
            rememberWithSourceTextCharCount (conditionalFormatApplyCacheSourceTextCharCount cells) cells
        let rememberChangedResult cells =
            rememberWithSourceTextCharCount (conditionalFormatApplyCacheMaxTotalSourceTextChars + 1) cells
        let compute () =
            match rangeNumberStats sr sc er ec model with
            | None ->
                remember model.Cells
                model
            | Some (numberCount, lo, hi) ->
                let span = if hi - lo < 1e-12 then 1.0 else hi - lo
                let iconBucket value =
                    min (count - 1) (int ((value - lo) / span * float count))
                let iconSetFormat value cell =
                    { cell.Format with IconSet = Some(iconBucket value, count) }
                if int64 numberCount * 2L >= int64 model.Cells.Count then
                    if not (rangeNumbersHaveAnyFormatChange sr sc er ec model iconSetFormat) then
                        remember model.Cells
                        model
                    else
                        let cellMap =
                            model.Cells
                            |> Map.map (fun addr cell ->
                                if addr.Row >= sr && addr.Row <= er && addr.Col >= sc && addr.Col <= ec then
                                    match unwrapVal cell.Value with
                                    | Number n ->
                                        { cell with Format = iconSetFormat n cell }
                                    | Boolean true ->
                                        { cell with Format = iconSetFormat 1.0 cell }
                                    | Boolean false ->
                                        { cell with Format = iconSetFormat 0.0 cell }
                                    | _ -> cell
                                else
                                    cell)
                        let updated = { model with Cells = cellMap }
                        rememberChangedResult updated.Cells
                        updated
                else
                    let cells =
                        match rangeNumbersWithBounds sr sc er ec model with
                        | Some (cells, _, _) -> cells
                        | None -> ResizeArray<CellAddress * float * CellData>()
                    if not (hasAnyFormatChange cells iconSetFormat) then
                        remember model.Cells
                        model
                    else
                        let mutable cellMap = model.Cells
                        let mutable changed = false
                        for addr, v, cell in cells do
                            let newFormat = iconSetFormat v cell
                            if newFormat <> cell.Format then
                                cellMap <- Map.add addr { cell with Format = newFormat } cellMap
                                changed <- true
                        if changed then
                            let updated = { model with Cells = cellMap }
                            rememberChangedResult updated.Cells
                            updated
                        else
                            remember model.Cells
                            model
        match tryGetIconSetApplyCache model.Cells with
        | ValueSome cached when cached.Entries.ContainsKey key -> model
        | _ ->
            compute ()

// --- Data validation --------------------------------------------------------

let private displayText (v: CellValue) : string =
    match unwrapVal v with
    | String s -> s
    | Number n -> if n = floor n && abs n < 1e15 then string (int64 n) else string n
    | Boolean b -> if b then "TRUE" else "FALSE"
    | _ -> ""

let private trimmedEqualsOrdinalIgnoreCase (left: string) (right: string) =
    let trimBounds (value: string) =
        let mutable startIndex = 0
        let mutable endIndex = value.Length - 1
        while startIndex <= endIndex && Char.IsWhiteSpace value.[startIndex] do
            startIndex <- startIndex + 1
        while endIndex >= startIndex && Char.IsWhiteSpace value.[endIndex] do
            endIndex <- endIndex - 1
        startIndex, endIndex

    let leftStart, leftEnd = trimBounds left
    let rightStart, rightEnd = trimBounds right
    let leftLength = leftEnd - leftStart + 1
    let rightLength = rightEnd - rightStart + 1
    if leftLength <> rightLength then
        false
    else
        let mutable matches = true
        let mutable offset = 0
        while matches && offset < leftLength do
            let leftChar = left.[leftStart + offset]
            let rightChar = right.[rightStart + offset]
            if leftChar <> rightChar && Char.ToUpperInvariant leftChar <> Char.ToUpperInvariant rightChar then
                matches <- false
            offset <- offset + 1
        matches

/// Check a value against a validation rule. Blank is always allowed (Excel's
/// default). Returns Ok or a human-readable reason.
let validateValue (rule: ValidationRule) (value: CellValue) : Result<unit, string> =
    let v = unwrapVal value
    let isBlank = match v with Empty -> true | String "" -> true | _ -> false
    let within lo hi x =
        (match lo with Some l -> x >= l - 1e-9 | None -> true) && (match hi with Some h -> x <= h + 1e-9 | None -> true)
    if isBlank then Result.Ok()
    else
        match rule with
        | ListValues vs ->
            let text = displayText v
            let mutable found = false
            for allowed in vs do
                if not found && trimmedEqualsOrdinalIgnoreCase allowed text then
                    found <- true
            if found then Result.Ok()
            else Result.Error(sprintf "Value must be one of: %s" (String.concat ", " vs))
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

let getValidation (addr: CellAddress) (model: SpreadsheetModel) : DataValidation option =
    Map.tryFind addr model.Validations

/// Attach a validation rule to a cell. A list rule also flags the cell so the
/// renderer paints the dropdown arrow.
let setValidation (addr: CellAddress) (dv: DataValidation) (model: SpreadsheetModel) : SpreadsheetModel =
    let mutable existingValidation = Unchecked.defaultof<DataValidation>
    let validationChanged =
        not (model.Validations.TryGetValue(addr, &existingValidation) && existingValidation = dv)
    let modelWithValidation =
        if validationChanged then
            { model with Validations = Map.add addr dv model.Validations }
        else
            model
    match dv.Rule with
    | ListValues _ ->
        let mutable cellData = Unchecked.defaultof<CellData>
        if modelWithValidation.Cells.TryGetValue(addr, &cellData) && cellData.Format.HasListDropdown then
            modelWithValidation
        else
            setCellFormat addr (fun f ->
                if f.HasListDropdown then f
                else { f with HasListDropdown = true }) modelWithValidation
    | _ -> modelWithValidation

/// Validate a prospective edit against the cell's rule (if any). Ok when valid
/// or unconstrained; Error carries the rule's message (or a default).
let validateEdit (addr: CellAddress) (value: CellValue) (model: SpreadsheetModel) : Result<unit, string> =
    match getValidation addr model with
    | None -> Result.Ok()
    | Some dv ->
        match validateValue dv.Rule value with
        | Result.Ok() -> Result.Ok()
        | Result.Error def -> Result.Error(dv.ErrorMessage |> Option.defaultValue def)

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
        rows.Add(resizeArrayToList row)
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

    resizeArrayToList rows

type private ImportCsvCache = {
    CsvText: string
    CsvLength: int
    StartRow: int
    StartCol: int
    SourceCells: Map<CellAddress, CellData>
    ResultCells: Map<CellAddress, CellData>
    SourceCellCount: int
    ResultCellCount: int
    SourceTextCharCount: int
}

[<Struct>]
type private ImportCsvInputKey = {
    CsvText: string
    StartRow: int
    StartCol: int
    Cells: Map<CellAddress, CellData>
}

type private ImportCsvChanges = {
    OldCells: ResizeArray<CellAddress * CellData option>
    NewCells: ResizeArray<CellAddress * CellData option>
}

let mutable private importCsvCache: ImportCsvCache list = []
let private importCsvFastSlots = Array.zeroCreate<ImportCsvCache> 256
let private importCsvCacheMaxSources = 64
let private importCsvCacheMaxTotalChars = 5000000
let private importCsvCacheMaxTotalCells = 250000
let private importCsvCacheMaxTotalSourceTextChars = 5000000
let private importCsvNoChangeCacheMaxChars = 65536
let private importCsvBulkEmptySourceMaxChars = 400000
let private importCsvBulkEmptySourceCellLimit = 8192

let private importCsvCacheSameSource csvText startRow startCol sourceCells (cached: ImportCsvCache) =
    cached.CsvText = csvText
    && cached.StartRow = startRow
    && cached.StartCol = startCol
    && obj.ReferenceEquals(cached.SourceCells, sourceCells)

let private importCsvCacheMatchesInput csvText startRow startCol cells (cached: ImportCsvCache) =
    cached.CsvText = csvText
    && cached.StartRow = startRow
    && cached.StartCol = startCol
    && (obj.ReferenceEquals(cached.SourceCells, cells)
        || obj.ReferenceEquals(cached.ResultCells, cells))

let private importCsvInputHash csvText startRow startCol cells =
    let csvHash = hash csvText
    (((csvHash * 397) ^^^ startRow) * 397 ^^^ startCol) * 397 ^^^ referenceHash cells

let private importCsvInputFastSlot csvText startRow startCol cells =
    (((referenceHash csvText * 397) ^^^ startRow) * 397 ^^^ startCol) * 397 ^^^ referenceHash cells
    &&& (importCsvFastSlots.Length - 1)

let private importCsvInputKey csvText startRow startCol cells : ImportCsvInputKey =
    { CsvText = csvText
      StartRow = startRow
      StartCol = startCol
      Cells = cells }

let private importCsvInputComparer =
    { new System.Collections.Generic.IEqualityComparer<ImportCsvInputKey> with
        member _.Equals(left, right) =
            left.CsvText = right.CsvText
            && left.StartRow = right.StartRow
            && left.StartCol = right.StartCol
            && obj.ReferenceEquals(left.Cells, right.Cells)

        member _.GetHashCode(key) =
            importCsvInputHash key.CsvText key.StartRow key.StartCol key.Cells }

let private importCsvCacheByInput =
    System.Collections.Generic.Dictionary<ImportCsvInputKey, ImportCsvCache>(importCsvInputComparer)

let private setImportCsvCache entries =
    importCsvCache <- entries
    importCsvCacheByInput.Clear()
    clearArray importCsvFastSlots
    let addInput (cells: Map<CellAddress, CellData>) (entry: ImportCsvCache) =
        let key = importCsvInputKey entry.CsvText entry.StartRow entry.StartCol cells
        let mutable existing = Unchecked.defaultof<ImportCsvCache>
        if not (importCsvCacheByInput.TryGetValue(key, &existing)) then
            importCsvCacheByInput.[key] <- entry
            let slot = importCsvInputFastSlot entry.CsvText entry.StartRow entry.StartCol cells
            importCsvFastSlots.[slot] <- entry
    for entry in entries do
        addInput entry.SourceCells entry
        if not (obj.ReferenceEquals(entry.SourceCells, entry.ResultCells)) then
            addInput entry.ResultCells entry

let private tryGetImportCsvCache csvText startRow startCol cells =
    let slot = importCsvInputFastSlot csvText startRow startCol cells
    let fast = importCsvFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && importCsvCacheMatchesInput csvText startRow startCol cells fast then
        ValueSome fast
    else
        let key = importCsvInputKey csvText startRow startCol cells
        let mutable cached = Unchecked.defaultof<ImportCsvCache>
        if importCsvCacheByInput.TryGetValue(key, &cached) then
            importCsvFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private importCsvCacheWithoutSource csvText startRow startCol sourceCells =
    let key = importCsvInputKey csvText startRow startCol sourceCells
    let mutable cached = Unchecked.defaultof<ImportCsvCache>
    if importCsvCacheByInput.TryGetValue(key, &cached) then
        importCsvCache
        |> List.filter (importCsvCacheSameSource csvText startRow startCol sourceCells >> not)
    else
        importCsvCache

let private removeImportCsvCacheForSource csvText startRow startCol sourceCells =
    let entries = importCsvCacheWithoutSource csvText startRow startCol sourceCells
    if not (obj.ReferenceEquals(entries, importCsvCache)) then
        setImportCsvCache entries

let private importCsvCacheRetainable (cached: ImportCsvCache) =
    cached.CsvLength <= importCsvCacheMaxTotalChars
    && cached.SourceCellCount <= importCsvCacheMaxTotalCells
    && cached.ResultCellCount <= importCsvCacheMaxTotalCells
    && cached.SourceTextCharCount <= importCsvCacheMaxTotalSourceTextChars

let private importCsvSourceTextCharCountFor csvLength sourceCellCount resultCellCount sourceCells =
    if csvLength > importCsvCacheMaxTotalChars
       || sourceCellCount > importCsvCacheMaxTotalCells
       || resultCellCount > importCsvCacheMaxTotalCells then
        importCsvCacheMaxTotalSourceTextChars + 1
    else
        cellMapSourceTextCharCountFor
            importCsvCacheMaxTotalCells
            importCsvCacheMaxTotalSourceTextChars
            sourceCells

let private importCsvRetainedCellCount (cached: ImportCsvCache) =
    if obj.ReferenceEquals(cached.SourceCells, cached.ResultCells) then
        cached.ResultCellCount
    else
        cached.SourceCellCount + cached.ResultCellCount

let private rememberImportCsv cached =
    if not (importCsvCacheRetainable cached) then
        removeImportCsvCacheForSource cached.CsvText cached.StartRow cached.StartCol cached.SourceCells
    else
        let trim entries =
            let rec loop keptCount keptChars keptCells keptSourceTextChars acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateChars = max 1 candidate.CsvLength
                    let candidateCells = max 1 (importCsvRetainedCellCount candidate)
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    if importCsvCacheRetainable candidate
                       && keptCount < importCsvCacheMaxSources
                       && keptChars + candidateChars <= importCsvCacheMaxTotalChars
                       && keptCells + candidateCells <= importCsvCacheMaxTotalCells
                       && keptSourceTextChars + candidateSourceTextChars <= importCsvCacheMaxTotalSourceTextChars then
                        loop
                            (keptCount + 1)
                            (keptChars + candidateChars)
                            (keptCells + candidateCells)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptChars keptCells keptSourceTextChars acc tail
            loop 0 0 0 0 [] entries
        (cached
         :: importCsvCacheWithoutSource cached.CsvText cached.StartRow cached.StartCol cached.SourceCells)
        |> trim
        |> setImportCsvCache

let private importCSVWithOptionalChanges (captureChanges: bool) (useCache: bool) (csvText: string) (startRow: int) (startCol: int) (model: SpreadsheetModel): SpreadsheetModel * ImportCsvChanges option =
    let cached =
        if useCache then
            match tryGetImportCsvCache csvText startRow startCol model.Cells with
            | ValueSome cached -> Some cached
            | ValueNone -> None
        else
            None

    match cached with
    | Some cached when importCsvCacheRetainable cached && obj.ReferenceEquals(cached.ResultCells, model.Cells) ->
        model, None
    | Some cached when importCsvCacheRetainable cached ->
        if obj.ReferenceEquals(cached.ResultCells, model.Cells) then model, None
        else { model with Cells = cached.ResultCells }, None
    | Some cached ->
        rememberImportCsv cached
        if obj.ReferenceEquals(cached.ResultCells, model.Cells) then model, None
        else { model with Cells = cached.ResultCells }, None
    | None ->
        let sourceCells = model.Cells
        let sourceCellsEmpty = Map.isEmpty sourceCells
        let mutable cells = model.Cells
        let bulkImportedCells =
            if sourceCellsEmpty && not captureChanges && csvText.Length <= importCsvBulkEmptySourceMaxChars then
                ResizeArray<CellAddress * CellData>()
            else
                null
        let mutable bulkImportOverflowed = false
        let mutable changed = false
        let changes =
            if captureChanges then
                Some {
                    OldCells = ResizeArray<CellAddress * CellData option>()
                    NewCells = ResizeArray<CellAddress * CellData option>()
                }
            else
                None
        let cell = System.Text.StringBuilder()
        let mutable inQuotes = false
        let mutable rowIdx = 0
        let mutable colIdx = 0
        let mutable rowHasCells = false

        let recordChange addr oldCell newCell =
            match changes with
            | Some changes ->
                changes.OldCells.Add(addr, oldCell)
                changes.NewCells.Add(addr, newCell)
            | None -> ()

        let addImportedCell addr cellData =
            if isNull bulkImportedCells || bulkImportOverflowed then
                cells <- Map.add addr cellData cells
            elif bulkImportedCells.Count < importCsvBulkEmptySourceCellLimit then
                bulkImportedCells.Add(addr, cellData)
            else
                for pendingAddr, pendingCellData in bulkImportedCells do
                    cells <- Map.add pendingAddr pendingCellData cells
                bulkImportedCells.Clear()
                bulkImportOverflowed <- true
                cells <- Map.add addr cellData cells

        let applyCell (cellText: string) =
            if sourceCellsEmpty && String.IsNullOrEmpty(cellText) then
                ()
            else
                let addr = { Row = startRow + rowIdx; Col = startCol + colIdx }
                let value =
                    if String.IsNullOrEmpty(cellText) then
                        Empty
                    elif cellText.[0] = '=' then
                        Formula (cellText, Empty)
                    elif System.String.Equals(cellText, "true", System.StringComparison.OrdinalIgnoreCase) then
                        Boolean true
                    elif System.String.Equals(cellText, "false", System.StringComparison.OrdinalIgnoreCase) then
                        Boolean false
                    else
                        match tryParseNumber cellText with
                        | Some n -> Number n
                        | None -> String cellText

                if sourceCellsEmpty then
                    if value <> Empty then
                        let cellData = {
                            Value = value
                            Format = defaultFormat
                            Comment = None
                        }
                        recordChange addr None (Some cellData)
                        addImportedCell addr cellData
                        changed <- true
                elif value = Empty then
                    let mutable existing = Unchecked.defaultof<CellData>
                    if cells.TryGetValue(addr, &existing) then
                        recordChange addr (Some existing) None
                        cells <- Map.remove addr cells
                        changed <- true
                else
                    let mutable existing = Unchecked.defaultof<CellData>
                    let hasExisting = cells.TryGetValue(addr, &existing)
                    if not (hasExisting
                            && existing.Value = value
                            && existing.Format = defaultFormat
                            && Option.isNone existing.Comment) then
                        let cellData = {
                            Value = value
                            Format = defaultFormat
                            Comment = None
                        }
                        recordChange addr (if hasExisting then Some existing else None) (Some cellData)
                        cells <- Map.add addr cellData cells
                        changed <- true

        let finishCell () =
            applyCell (cell.ToString())
            cell.Clear() |> ignore
            rowHasCells <- true
            colIdx <- colIdx + 1

        let finishRow () =
            finishCell()
            rowIdx <- rowIdx + 1
            colIdx <- 0
            rowHasCells <- false

        let mutable i = 0
        while i < csvText.Length do
            let c = csvText.[i]
            if inQuotes then
                match c with
                | '"' ->
                    if i + 1 < csvText.Length && csvText.[i + 1] = '"' then
                        cell.Append('"') |> ignore
                        i <- i + 1
                    else
                        inQuotes <- false
                | _ ->
                    cell.Append(c) |> ignore
            else
                match c with
                | '"' -> inQuotes <- true
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

        if cell.Length > 0 || rowHasCells then
            finishRow()

        if changed && not (isNull bulkImportedCells) && not bulkImportOverflowed then
            cells <- Map.ofSeq bulkImportedCells

        if useCache && (changed || csvText.Length <= importCsvNoChangeCacheMaxChars) then
            rememberImportCsv {
                CsvText = csvText
                CsvLength = csvText.Length
                StartRow = startRow
                StartCol = startCol
                SourceCells = sourceCells
                ResultCells = cells
                SourceCellCount = sourceCells.Count
                ResultCellCount = cells.Count
                SourceTextCharCount =
                    importCsvSourceTextCharCountFor
                        csvText.Length
                        sourceCells.Count
                        cells.Count
                        sourceCells
            }
        let result = if changed then { model with Cells = cells } else model
        result, changes

/// Import CSV data into spreadsheet
let importCSV (csvText: string) (startRow: int) (startCol: int) (model: SpreadsheetModel): SpreadsheetModel =
    let imported, _ = importCSVWithOptionalChanges false true csvText startRow startCol model
    imported

/// Import CSV data and return changed old/new cells for sparse undo.
let importCSVWithChanges (csvText: string) (startRow: int) (startCol: int) (model: SpreadsheetModel): SpreadsheetModel * ResizeArray<CellAddress * CellData option> * ResizeArray<CellAddress * CellData option> =
    let imported, changes = importCSVWithOptionalChanges true false csvText startRow startCol model
    match changes with
    | Some changes ->
        imported, changes.OldCells, changes.NewCells
    | None ->
        imported, ResizeArray<CellAddress * CellData option>(), ResizeArray<CellAddress * CellData option>()

/// Export spreadsheet to CSV string
let exportCSV (model: SpreadsheetModel): string =
    match tryGetExportCsvCache model.Cells model.MergedCells with
    | ValueSome cached when exportCsvCacheRetainable cached ->
        cached.Text
    | ValueSome cached ->
        rememberExportCsv cached
        cached.Text
    | ValueNone ->
        let text =
            match usedBounds model with
            | None ->
                ""
            | Some (maxRow, maxCol) ->
                let sb = System.Text.StringBuilder()
                let appendField (cellData: CellData) =
                    let valueText = displayValueWithFormat cellData.Value cellData.Format
                    let hasQuotes = valueText.Contains("\"")
                    let needsQuotes = hasQuotes || valueText.Contains(",") || valueText.Contains("\n") || valueText.Contains("\r")
                    if needsQuotes then
                        sb.Append('"') |> ignore
                        if hasQuotes then
                            let mutable start = 0
                            for i = 0 to valueText.Length - 1 do
                                if valueText.[i] = '"' then
                                    if i > start then
                                        sb.Append(valueText, start, i - start) |> ignore
                                    sb.Append("\"\"") |> ignore
                                    start <- i + 1
                            if start < valueText.Length then
                                sb.Append(valueText, start, valueText.Length - start) |> ignore
                        else
                            sb.Append(valueText) |> ignore
                        sb.Append('"') |> ignore
                    else
                        sb.Append(valueText) |> ignore

                use cells = (model.Cells :> seq<_>).GetEnumerator()
                let mutable hasCurrent = cells.MoveNext()
                let advance () =
                    hasCurrent <- cells.MoveNext()

                for row in 0 .. maxRow do
                    if row > 0 then sb.Append('\n') |> ignore
                    let mutable nextCol = 0
                    let mutable rowDone = false
                    while hasCurrent && not rowDone do
                        let (KeyValue(addr, cellData)) = cells.Current
                        if addr.Row < row || (addr.Row = row && (addr.Col < 0 || addr.Col > maxCol)) then
                            advance ()
                        elif addr.Row > row then
                            rowDone <- true
                        else
                            if nextCol < addr.Col then
                                let commaCount =
                                    if nextCol = 0 then
                                        addr.Col - 1
                                    else
                                        addr.Col - nextCol
                                if commaCount > 0 then
                                    sb.Append(',', commaCount) |> ignore
                                nextCol <- addr.Col
                            if addr.Col > 0 then sb.Append(',') |> ignore
                            appendField cellData
                            nextCol <- addr.Col + 1
                            advance ()
                    let trailingCommaStart = max 1 nextCol
                    let trailingCommaCount = maxCol - trailingCommaStart + 1
                    if trailingCommaCount > 0 then
                        sb.Append(',', trailingCommaCount) |> ignore
                sb.ToString()
        let mergedCellCount = mergedCellCountCapped exportCsvCacheMaxTotalMergedCells model.MergedCells
        rememberExportCsv {
            Cells = model.Cells
            CellCount = model.Cells.Count
            SourceTextCharCount =
                exportCsvSourceTextCharCountFor
                    text.Length
                    model.Cells.Count
                    mergedCellCount
                    model.Cells
            MergedCells = model.MergedCells
            MergedCellCount = mergedCellCount
            TextLength = text.Length
            Text = text
        }
        text

/// Check if a cell address is within a merged range
let isInMergedRange (addr: CellAddress) (merge: MergedCell): bool =
    addr.Row >= merge.StartRow && addr.Row <= merge.EndRow &&
    addr.Col >= merge.StartCol && addr.Col <= merge.EndCol

type private IndexedMerge = {
    Index: int
    Merge: MergedCell
}

type private MergeLookupIndex = {
    Source: MergedCell list
    SourceCount: int
    Rows: System.Collections.Generic.Dictionary<int, ResizeArray<MergedCell>> option
}

type private MergeRangeIndex = {
    Source: MergedCell list
    SourceCount: int
    ByStartRow: IndexedMerge array
}

type private MergeRangeOrderingCache = {
    Source: MergedCell list
    SourceCount: int
    IsStartRowAscending: bool
}

type private MergeRangeQueryKey = {
    StartRow: int
    EndRowExclusive: int
    StartCol: int
    EndColExclusive: int
}

type private MergeRangeQueryCacheEntry = {
    Results: MergedCell list
    ResultCount: int
}

type private MergeRangeQueryCache = {
    Source: MergedCell list
    SourceCount: int
    Entries: System.Collections.Generic.Dictionary<MergeRangeQueryKey, MergeRangeQueryCacheEntry>
    Order: System.Collections.Generic.Queue<MergeRangeQueryKey>
    mutable TotalResultCount: int
}

let mutable private mergeLookupIndex: MergeLookupIndex list = []
let mutable private mergeRangeIndex: MergeRangeIndex list = []
let mutable private mergeRangeOrderingCache: MergeRangeOrderingCache list = []
let mutable private mergeRangeQueryCache: MergeRangeQueryCache list = []
let private mergedCellListReferenceComparer =
    { new System.Collections.Generic.IEqualityComparer<MergedCell list> with
        member _.Equals(left, right) = obj.ReferenceEquals(left, right)
        member _.GetHashCode(value) = referenceHash value }
let private mergeLookupIndexBySource =
    System.Collections.Generic.Dictionary<MergedCell list, MergeLookupIndex>(mergedCellListReferenceComparer)
let private mergeLookupFastSlots = Array.zeroCreate<MergeLookupIndex> 256
let private mergeRangeIndexBySource =
    System.Collections.Generic.Dictionary<MergedCell list, MergeRangeIndex>(mergedCellListReferenceComparer)
let private mergeRangeFastSlots = Array.zeroCreate<MergeRangeIndex> 256
let private mergeRangeOrderingCacheBySource =
    System.Collections.Generic.Dictionary<MergedCell list, MergeRangeOrderingCache>(mergedCellListReferenceComparer)
let private mergeRangeOrderingFastSlots = Array.zeroCreate<MergeRangeOrderingCache> 256
let private mergeRangeQueryCacheBySource =
    System.Collections.Generic.Dictionary<MergedCell list, MergeRangeQueryCache>(mergedCellListReferenceComparer)
let private mergeRangeQueryFastSlots = Array.zeroCreate<MergeRangeQueryCache> 256
let private mergeLookupRowEntryLimit = 100000
let private mergeIndexCacheMaxSources = 64
let private mergeIndexCacheMaxTotalSourceMerges = 100000
let private mergeRangeQueryCacheMaxEntries = 64
let private mergeRangeQueryCacheMaxTotalResults = 200000
let private mergeRangeQueryCacheMaxTotalSourceMerges = 100000

let private mergeSourceCountCapped (merges: MergedCell list) =
    let mutable count = 0
    let mutable remaining = merges
    while count <= mergeIndexCacheMaxTotalSourceMerges && not remaining.IsEmpty do
        count <- count + 1
        remaining <- remaining.Tail
    count

let private mergeLookupIndexMatches merges (cached: MergeLookupIndex) =
    System.Object.ReferenceEquals(cached.Source, merges)

let private mergeRangeIndexMatches merges (cached: MergeRangeIndex) =
    System.Object.ReferenceEquals(cached.Source, merges)

let private mergeRangeOrderingCacheMatches merges (cached: MergeRangeOrderingCache) =
    System.Object.ReferenceEquals(cached.Source, merges)

let private mergeRangeQueryCacheMatches merges (cached: MergeRangeQueryCache) =
    System.Object.ReferenceEquals(cached.Source, merges)

let private setMergeLookupIndex entries =
    mergeLookupIndex <- entries
    mergeLookupIndexBySource.Clear()
    clearArray mergeLookupFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<MergeLookupIndex>
        if not (mergeLookupIndexBySource.TryGetValue(entry.Source, &existing)) then
            mergeLookupIndexBySource.[entry.Source] <- entry
            let slot = referenceSlot entry.Source mergeLookupFastSlots.Length
            mergeLookupFastSlots.[slot] <- entry

let private setMergeRangeIndex entries =
    mergeRangeIndex <- entries
    mergeRangeIndexBySource.Clear()
    clearArray mergeRangeFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<MergeRangeIndex>
        if not (mergeRangeIndexBySource.TryGetValue(entry.Source, &existing)) then
            mergeRangeIndexBySource.[entry.Source] <- entry
            let slot = referenceSlot entry.Source mergeRangeFastSlots.Length
            mergeRangeFastSlots.[slot] <- entry

let private setMergeRangeOrderingCache entries =
    mergeRangeOrderingCache <- entries
    mergeRangeOrderingCacheBySource.Clear()
    clearArray mergeRangeOrderingFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<MergeRangeOrderingCache>
        if not (mergeRangeOrderingCacheBySource.TryGetValue(entry.Source, &existing)) then
            mergeRangeOrderingCacheBySource.[entry.Source] <- entry
            let slot = referenceSlot entry.Source mergeRangeOrderingFastSlots.Length
            mergeRangeOrderingFastSlots.[slot] <- entry

let private setMergeRangeQueryCache entries =
    mergeRangeQueryCache <- entries
    mergeRangeQueryCacheBySource.Clear()
    clearArray mergeRangeQueryFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<MergeRangeQueryCache>
        if not (mergeRangeQueryCacheBySource.TryGetValue(entry.Source, &existing)) then
            mergeRangeQueryCacheBySource.[entry.Source] <- entry
            let slot = referenceSlot entry.Source mergeRangeQueryFastSlots.Length
            mergeRangeQueryFastSlots.[slot] <- entry

let private tryGetMergeLookupIndex merges =
    let slot = referenceSlot merges mergeLookupFastSlots.Length
    let fast = mergeLookupFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && mergeLookupIndexMatches merges fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<MergeLookupIndex>
        if mergeLookupIndexBySource.TryGetValue(merges, &cached) then
            mergeLookupFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private tryGetMergeRangeIndex merges =
    let slot = referenceSlot merges mergeRangeFastSlots.Length
    let fast = mergeRangeFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && mergeRangeIndexMatches merges fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<MergeRangeIndex>
        if mergeRangeIndexBySource.TryGetValue(merges, &cached) then
            mergeRangeFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private tryGetMergeRangeOrderingCache merges =
    let slot = referenceSlot merges mergeRangeOrderingFastSlots.Length
    let fast = mergeRangeOrderingFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && mergeRangeOrderingCacheMatches merges fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<MergeRangeOrderingCache>
        if mergeRangeOrderingCacheBySource.TryGetValue(merges, &cached) then
            mergeRangeOrderingFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private tryGetMergeRangeQueryCache merges =
    let slot = referenceSlot merges mergeRangeQueryFastSlots.Length
    let fast = mergeRangeQueryFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && mergeRangeQueryCacheMatches merges fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<MergeRangeQueryCache>
        if mergeRangeQueryCacheBySource.TryGetValue(merges, &cached) then
            mergeRangeQueryFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private trimMergeSourceCache sourceCount entries =
    let rec loop keptCount keptSources acc remaining =
        match remaining with
        | [] -> List.rev acc
        | candidate :: tail ->
            let candidateSources = max 1 (sourceCount candidate)
            if sourceCount candidate <= mergeIndexCacheMaxTotalSourceMerges
               && keptCount < mergeIndexCacheMaxSources
               && keptSources + candidateSources <= mergeIndexCacheMaxTotalSourceMerges then
                loop (keptCount + 1) (keptSources + candidateSources) (candidate :: acc) tail
            else
                loop keptCount keptSources acc tail
    loop 0 0 [] entries

let private mergeLookupIndexWithoutSource merges =
    let mutable cached = Unchecked.defaultof<MergeLookupIndex>
    if mergeLookupIndexBySource.TryGetValue(merges, &cached) then
        mergeLookupIndex
        |> List.filter (mergeLookupIndexMatches merges >> not)
    else
        mergeLookupIndex

let private mergeRangeIndexWithoutSource merges =
    let mutable cached = Unchecked.defaultof<MergeRangeIndex>
    if mergeRangeIndexBySource.TryGetValue(merges, &cached) then
        mergeRangeIndex
        |> List.filter (mergeRangeIndexMatches merges >> not)
    else
        mergeRangeIndex

let private mergeRangeOrderingCacheWithoutSource merges =
    let mutable cached = Unchecked.defaultof<MergeRangeOrderingCache>
    if mergeRangeOrderingCacheBySource.TryGetValue(merges, &cached) then
        mergeRangeOrderingCache
        |> List.filter (mergeRangeOrderingCacheMatches merges >> not)
    else
        mergeRangeOrderingCache

let private mergeRangeQueryCacheWithoutSource merges =
    let mutable cached = Unchecked.defaultof<MergeRangeQueryCache>
    if mergeRangeQueryCacheBySource.TryGetValue(merges, &cached) then
        mergeRangeQueryCache
        |> List.filter (mergeRangeQueryCacheMatches merges >> not)
    else
        mergeRangeQueryCache

let private mergeRangeQueryCacheRetainable (cached: MergeRangeQueryCache) =
    cached.Entries.Count > 0
    && cached.TotalResultCount <= mergeRangeQueryCacheMaxTotalResults
    && cached.SourceCount <= mergeRangeQueryCacheMaxTotalSourceMerges

let private removeMergeRangeQueryCacheForSource merges =
    let entries = mergeRangeQueryCacheWithoutSource merges
    if not (obj.ReferenceEquals(entries, mergeRangeQueryCache)) then
        setMergeRangeQueryCache entries

let private rememberMergeLookupIndex cached =
    (cached
     :: mergeLookupIndexWithoutSource cached.Source)
    |> trimMergeSourceCache (fun entry -> entry.SourceCount)
    |> setMergeLookupIndex

let private rememberMergeRangeIndex cached =
    (cached
     :: mergeRangeIndexWithoutSource cached.Source)
    |> trimMergeSourceCache (fun entry -> entry.SourceCount)
    |> setMergeRangeIndex

let private rememberMergeRangeOrdering cached =
    (cached
     :: mergeRangeOrderingCacheWithoutSource cached.Source)
    |> trimMergeSourceCache (fun entry -> entry.SourceCount)
    |> setMergeRangeOrderingCache

let private rememberMergeRangeQueryCache cached =
    if not (mergeRangeQueryCacheRetainable cached) then
        removeMergeRangeQueryCacheForSource cached.Source
    else
        let mutable keptCount = 1
        let mutable keptResults = cached.TotalResultCount
        let mutable keptSources = max 1 cached.SourceCount
        let retained =
            mergeRangeQueryCacheWithoutSource cached.Source
            |> List.choose (fun candidate ->
                let candidateResults = candidate.TotalResultCount
                let candidateSources = max 1 candidate.SourceCount
                if mergeRangeQueryCacheRetainable candidate
                   && keptCount < mergeIndexCacheMaxSources
                   && keptResults + candidateResults <= mergeRangeQueryCacheMaxTotalResults
                   && keptSources + candidateSources <= mergeRangeQueryCacheMaxTotalSourceMerges then
                    keptCount <- keptCount + 1
                    keptResults <- keptResults + candidateResults
                    keptSources <- keptSources + candidateSources
                    Some candidate
                else
                    None)
        cached :: retained
        |> setMergeRangeQueryCache

let private mergeRangeQueryKey startRow endRowExclusive startCol endColExclusive = {
    StartRow = startRow
    EndRowExclusive = endRowExclusive
    StartCol = startCol
    EndColExclusive = endColExclusive
}

let private removeMergeRangeQueryCacheEntry key (cache: MergeRangeQueryCache) =
    let mutable existing = Unchecked.defaultof<MergeRangeQueryCacheEntry>
    if cache.Entries.TryGetValue(key, &existing) then
        cache.TotalResultCount <- max 0 (cache.TotalResultCount - existing.ResultCount)
        cache.Entries.Remove key |> ignore

let private addMergeRangeQueryCacheEntry key entry (cache: MergeRangeQueryCache) =
    let mutable existing = Unchecked.defaultof<MergeRangeQueryCacheEntry>
    if entry.ResultCount > mergeRangeQueryCacheMaxTotalResults then
        removeMergeRangeQueryCacheEntry key cache
    else
        if cache.Entries.TryGetValue(key, &existing) then
            cache.TotalResultCount <- max 0 (cache.TotalResultCount - existing.ResultCount)
        else
            cache.Order.Enqueue key
        cache.Entries.[key] <- entry
        cache.TotalResultCount <- cache.TotalResultCount + entry.ResultCount
        while (cache.Entries.Count > mergeRangeQueryCacheMaxEntries
               || cache.TotalResultCount > mergeRangeQueryCacheMaxTotalResults)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if oldest <> key then
                removeMergeRangeQueryCacheEntry oldest cache
        if cache.Entries.Count > mergeRangeQueryCacheMaxEntries
           || cache.TotalResultCount > mergeRangeQueryCacheMaxTotalResults then
            let current = cache.Entries.[key]
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue key
            cache.Entries.[key] <- current
            cache.TotalResultCount <- current.ResultCount

let private cachedMergeRangeQuery startRow endRowExclusive startCol endColExclusive merges =
    let key = mergeRangeQueryKey startRow endRowExclusive startCol endColExclusive
    match tryGetMergeRangeQueryCache merges with
    | ValueSome cached when mergeRangeQueryCacheRetainable cached ->
        let mutable entry = Unchecked.defaultof<MergeRangeQueryCacheEntry>
        if cached.Entries.TryGetValue(key, &entry) then
            Some entry.Results
        else
            None
    | ValueSome cached ->
        rememberMergeRangeQueryCache cached
        None
    | ValueNone -> None

let private rememberMergeRangeQuery startRow endRowExclusive startCol endColExclusive merges results =
    let key = mergeRangeQueryKey startRow endRowExclusive startCol endColExclusive
    let entry = { Results = results; ResultCount = results.Length }
    let cache =
        match tryGetMergeRangeQueryCache merges with
        | ValueSome cached -> cached
        | ValueNone ->
            {
                Source = merges
                SourceCount = mergeSourceCountCapped merges
                Entries = System.Collections.Generic.Dictionary<MergeRangeQueryKey, MergeRangeQueryCacheEntry>()
                Order = System.Collections.Generic.Queue<MergeRangeQueryKey>()
                TotalResultCount = 0
            }
    addMergeRangeQueryCacheEntry key entry cache
    rememberMergeRangeQueryCache cache

let private buildMergeLookupIndex (merges: MergedCell list) =
    let sourceCount = mergeSourceCountCapped merges
    let mutable rowEntries = 0
    let mutable eligible = sourceCount <= mergeIndexCacheMaxTotalSourceMerges
    let mutable remaining = merges
    while eligible && not remaining.IsEmpty do
        let merge = remaining.Head
        if merge.EndRow >= merge.StartRow then
            rowEntries <- rowEntries + merge.EndRow - merge.StartRow + 1
            if rowEntries > mergeLookupRowEntryLimit then
                eligible <- false
        remaining <- remaining.Tail

    let rows =
        if eligible then
            let index = System.Collections.Generic.Dictionary<int, ResizeArray<MergedCell>>()
            for merge in merges do
                if merge.EndRow >= merge.StartRow then
                    for row in merge.StartRow .. merge.EndRow do
                        match index.TryGetValue row with
                        | true, rowMerges -> rowMerges.Add merge
                        | false, _ ->
                            let rowMerges = ResizeArray<MergedCell>()
                            rowMerges.Add merge
                            index.[row] <- rowMerges
            Some index
        else
            None
    { Source = merges; SourceCount = sourceCount; Rows = rows }

let private mergeLookupRows merges =
    match tryGetMergeLookupIndex merges with
    | ValueSome cached ->
        cached.Rows
    | ValueNone ->
        let cached = buildMergeLookupIndex merges
        rememberMergeLookupIndex cached
        cached.Rows

let private mergeRangeRows merges =
    match tryGetMergeRangeIndex merges with
    | ValueSome cached ->
        cached.ByStartRow
    | ValueNone ->
        let sourceCount = mergeSourceCountCapped merges
        if sourceCount > mergeIndexCacheMaxTotalSourceMerges then
            [||]
        else
            let items = ResizeArray<IndexedMerge>()
            let mutable index = 0
            for merge in merges do
                items.Add { Index = index; Merge = merge }
                index <- index + 1
            let arr = items.ToArray()
            Array.sortInPlaceWith
                (fun left right ->
                    let rowCompare = compare left.Merge.StartRow right.Merge.StartRow
                    if rowCompare <> 0 then rowCompare else compare left.Index right.Index)
                arr
            rememberMergeRangeIndex { Source = merges; SourceCount = sourceCount; ByStartRow = arr }
            arr

let private mergeRangeIsStartRowAscending (merges: MergedCell list) =
    match tryGetMergeRangeOrderingCache merges with
    | ValueSome cached ->
        cached.IsStartRowAscending
    | ValueNone ->
        let sourceCount = mergeSourceCountCapped merges
        let mutable ascending = true
        let mutable lastStart = System.Int32.MinValue
        let mutable remaining = merges
        while ascending && not remaining.IsEmpty do
            let merge = remaining.Head
            if merge.StartRow < lastStart then
                ascending <- false
            else
                lastStart <- merge.StartRow
                remaining <- remaining.Tail
        rememberMergeRangeOrdering { Source = merges; SourceCount = sourceCount; IsStartRowAscending = ascending }
        ascending

let private mergedCellsIntersectingRangeLinear startRow endRowExclusive startCol endColExclusive (merges: MergedCell list) =
    let result = ResizeArray<MergedCell>()
    let mutable remaining = merges
    while not remaining.IsEmpty do
        let merge = remaining.Head
        if merge.EndRow >= startRow &&
           merge.StartRow < endRowExclusive &&
           merge.EndCol >= startCol &&
           merge.StartCol < endColExclusive then
            result.Add merge
        remaining <- remaining.Tail
    resizeArrayToList result

/// Find merged cell containing this address (if any)
let findMergedCell (addr: CellAddress) (model: SpreadsheetModel): MergedCell option =
    match mergeLookupRows model.MergedCells with
    | Some rows ->
        match rows.TryGetValue addr.Row with
        | true, rowMerges ->
            let mutable found = None
            let mutable index = 0
            while found.IsNone && index < rowMerges.Count do
                let merge = rowMerges.[index]
                if addr.Col >= merge.StartCol && addr.Col <= merge.EndCol then
                    found <- Some merge
                index <- index + 1
            found
        | false, _ ->
            None
    | None ->
        let mutable found = None
        let mutable remaining = model.MergedCells
        while found.IsNone && not remaining.IsEmpty do
            match remaining with
            | merge :: tail ->
                if isInMergedRange addr merge then
                    found <- Some merge
                remaining <- tail
            | [] -> ()
        found

/// Find merged cells intersecting a row/column range, preserving merge-list order.
let mergedCellsIntersectingRange startRow endRowExclusive startCol endColExclusive (model: SpreadsheetModel) =
    if startRow >= endRowExclusive || startCol >= endColExclusive then
        []
    else
        match cachedMergeRangeQuery startRow endRowExclusive startCol endColExclusive model.MergedCells with
        | Some cached -> cached
        | None ->
            let intersectsVisibleRange (merge: MergedCell) =
                merge.EndRow >= startRow &&
                merge.StartRow < endRowExclusive &&
                merge.EndCol >= startCol &&
                merge.StartCol < endColExclusive

            let results =
                if mergeRangeIsStartRowAscending model.MergedCells then
                    let result = ResizeArray<MergedCell>()
                    let mutable remaining = model.MergedCells
                    let mutable keepScanning = true
                    while keepScanning && not remaining.IsEmpty do
                        match remaining with
                        | merge :: tail ->
                            if merge.StartRow >= endRowExclusive then
                                keepScanning <- false
                            else
                                if intersectsVisibleRange merge then
                                    result.Add merge
                                remaining <- tail
                        | [] -> ()
                    resizeArrayToList result
                elif mergeSourceCountCapped model.MergedCells > mergeIndexCacheMaxTotalSourceMerges then
                    mergedCellsIntersectingRangeLinear startRow endRowExclusive startCol endColExclusive model.MergedCells
                else
                    let byStartRow = mergeRangeRows model.MergedCells
                    let matches = ResizeArray<IndexedMerge>()
                    let mutable index = 0
                    while index < byStartRow.Length && byStartRow.[index].Merge.StartRow < endRowExclusive do
                        let indexed = byStartRow.[index]
                        if intersectsVisibleRange indexed.Merge then
                            matches.Add indexed
                        index <- index + 1
                    if matches.Count = 0 then
                        []
                    else
                        let arr = matches.ToArray()
                        Array.sortInPlaceWith (fun left right -> compare left.Index right.Index) arr
                        let result = ResizeArray<MergedCell>(arr.Length)
                        for indexed in arr do
                            result.Add indexed.Merge
                        resizeArrayToList result
            rememberMergeRangeQuery startRow endRowExclusive startCol endColExclusive model.MergedCells results
            results

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
        let mutable duplicate = false
        let mutable remaining = model.MergedCells
        while not duplicate && not remaining.IsEmpty do
            match remaining with
            | existing :: tail ->
                duplicate <- existing = merge
                remaining <- tail
            | [] -> ()
        if duplicate then
            model
        else
            { model with MergedCells = merge :: model.MergedCells }

/// Remove merged cell containing this address
let removeMergedCell (addr: CellAddress) (model: SpreadsheetModel): SpreadsheetModel =
    match model.MergedCells with
    | head :: tail when isInMergedRange addr head ->
        { model with MergedCells = tail }
    | _ ->
        let mutable found = false
        let mutable prefixCount = 0
        let mutable scan = model.MergedCells
        while not found && not scan.IsEmpty do
            if isInMergedRange addr scan.Head then
                found <- true
            else
                prefixCount <- prefixCount + 1
                scan <- scan.Tail

        if not found then
            model
        else
            let prefix = ResizeArray<MergedCell>(prefixCount)
            let rec removeFirst remaining =
                match remaining with
                | [] ->
                    model
                | merge :: tail when isInMergedRange addr merge ->
                    let mutable hasLaterMatch = false
                    let mutable scanTail = tail
                    while not hasLaterMatch && not scanTail.IsEmpty do
                        if isInMergedRange addr scanTail.Head then
                            hasLaterMatch <- true
                        scanTail <- scanTail.Tail

                    if not hasLaterMatch then
                        { model with MergedCells = prependResizeArray prefix tail }
                    else
                        let kept = ResizeArray<MergedCell>(prefix.Count)
                        for item in prefix do
                            kept.Add item
                        let mutable rest = tail
                        while not rest.IsEmpty do
                            let item = rest.Head
                            if not (isInMergedRange addr item) then
                                kept.Add item
                            rest <- rest.Tail
                        { model with MergedCells = resizeArrayToList kept }
                | merge :: tail ->
                    prefix.Add merge
                    removeFirst tail
            removeFirst model.MergedCells
