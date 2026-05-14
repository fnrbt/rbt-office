module XlsxReader

open XmlDsl
open OfficeDsl
open ZipInterop
open Fable.Core.JsInterop

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
    Value: CellValue
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
    Height: float option
}

type ColumnInfo = {
    Index: int
    MaxIndex: int
    Width: float option
    Hidden: bool
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
    /// `iconSet` family name (e.g. "3TrafficLights1", "5Arrows").
    /// None for non-iconSet rules.
    IconSet: string option
    /// `iconSet` threshold cfvos — N entries for an N-icon set, each
    /// (type, optional value). The first is always the implicit 0%
    /// floor; icon i shows when the value is ≥ threshold i.
    IconSetCfvo: (string * string option) list
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
    StartRow: int
    StartCol: int
    EndRow: int
    EndCol: int
    HeaderRowCount: int
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

type Worksheet = {
    Name: string
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
    Hyperlinks: Hyperlink list
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
}

// DSL for XLSX-specific patterns
module Xlsx =
    // Namespace-aware helpers
    let inline sElem name node = Ns.childNs name node
    let inline sElems name node = Ns.descendantsNs name node

    // Cell reference parsing DSL
    let parseCellRef (ref: string) : CellReference =
        let letters = ref |> Seq.takeWhile System.Char.IsLetter |> Seq.toArray |> System.String
        let digits = ref |> Seq.skipWhile System.Char.IsLetter |> Seq.toArray |> System.String
        {
            Column = letters
            Row = if digits.Length > 0 then int digits else 0
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
            |> Option.bind (fun v -> try Some (float v) with _ -> None)
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
                let text =
                    Xml.descendants "t" child
                    |> List.choose Xml.text
                    |> String.concat ""
                let props =
                    sElem "rPr" child
                    |> Option.map parseRunProps
                    |> Option.defaultValue emptyRunProps
                Some (text, props)
            | _ -> None)

    // Cell value extraction with pattern matching. Excel formula cells
    // store both the expression and a cached result in <v>; we can't
    // evaluate the expression, so prefer the cached value (matching what
    // any XLSX viewer shows) and only fall back to the formula text when
    // no cache is present (uncalculated workbooks).
    let cellValue (cNode: XmlNode) : CellValue =
        let cellType = Xml.attr "t" cNode
        let formula = sElem "f" cNode |> Option.bind Xml.text
        let value = sElem "v" cNode |> Option.bind Xml.text
        // `t="inlineStr"` cells carry their text directly inside <is><t>…</t></is>
        // instead of pointing at a sharedStrings index. Rare in normal
        // Excel exports but emitted by various code generators (POI, the
        // .NET OpenXmlSdk's row-streaming writer, sql-server-style mass
        // importers) where deduping into sharedStrings isn't worth the
        // cost. Without this case those cells came back as Empty.
        let inlineText =
            sElem "is" cNode
            |> Option.map (fun isNode ->
                Xml.descendants "t" isNode
                |> List.choose Xml.text
                |> String.concat "")

        match cellType, value with
        | Some "inlineStr", _ ->
            inlineText |> Option.defaultValue "" |> String
        | Some "s", Some v -> SharedString (int v)
        | Some "b", Some v -> Boolean (v = "1")
        | Some "str", Some v -> String v
        | _, Some v ->
            match System.Double.TryParse(v) with
            | true, n -> Number n
            | false, _ -> String v
        | _ ->
            match formula with
            | Some f -> Formula f
            | None ->
                // No <v>, no <f> — but an <is>/<t> might still be present
                // even if the cell omits an explicit t="inlineStr".
                match inlineText with
                | Some t when t <> "" -> String t
                | _ -> Empty

    // Cell extraction with applicative style
    let cell (cNode: XmlNode) : Cell option =
        Xml.attr "r" cNode
        |> Option.map (fun ref ->
            let styleIdx = Xml.attr "s" cNode |> Option.map int
            // Inline-string cells with <r> children carry per-run formatting
            // identical in shape to sharedStrings rich text. Lift those runs
            // here so the consumer can route them through the same
            // RichRuns map as SharedString-backed rich text.
            let inlineRuns =
                match Xml.attr "t" cNode with
                | Some "inlineStr" ->
                    sElem "is" cNode
                    |> Option.map richRunsOf
                    |> Option.defaultValue []
                | _ -> []
            {
                Reference = parseCellRef ref
                Value = cellValue cNode
                StyleIndex = styleIdx
                Style = None  // Resolved later from styles array
                InlineRuns = inlineRuns
            })

    // Row extraction
    let row (rowNode: XmlNode) : Row option =
        Xml.attr "r" rowNode
        |> Option.map (fun r ->
            let height = Xml.attr "ht" rowNode |> Option.map float
            {
                Index = int r
                Cells =
                    sElems "c" rowNode
                    |> List.choose cell
                Height = height
            })

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
        |> Option.map (fun minCol ->
            let minIdx = int minCol
            let maxIdx =
                Xml.attr "max" colNode
                |> Option.map int
                |> Option.defaultValue minIdx
            let width = Xml.attr "width" colNode |> Option.map float
            let hidden = Xml.attr "hidden" colNode |> Option.map (fun h -> h = "1") |> Option.defaultValue false
            {
                Index = minIdx
                MaxIndex = maxIdx
                Width = width
                Hidden = hidden
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
    }

    let sheetFormatPr (sheet: XmlNode) : SheetFormat =
        match sElem "sheetFormatPr" sheet with
        | None -> emptySheetFormat
        | Some node ->
            {
                BaseColWidth = Xml.attr "baseColWidth" node |> Option.bind (fun v -> try Some (float v) with _ -> None)
                DefaultColWidth = Xml.attr "defaultColWidth" node |> Option.bind (fun v -> try Some (float v) with _ -> None)
                DefaultRowHeight = Xml.attr "defaultRowHeight" node |> Option.bind (fun v -> try Some (float v) with _ -> None)
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
                let rid = Xml.attr "r:id" node
                let target, isInternal =
                    match location with
                    | Some loc -> loc, true
                    | None ->
                        match rid |> Option.bind resolveRel with
                        | Some url -> url, false
                        | None -> rid |> Option.defaultValue "", false
                Some { Range = range; Target = target; IsInternal = isInternal })

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
                    |> Option.bind (fun s ->
                        match System.Double.TryParse(s, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture) with
                        | true, v -> Some (int v)
                        | _ -> None)
                    |> Option.defaultValue 0
                let x = parseSplit "xSplit"
                let y = parseSplit "ySplit"
                if x <= 0 && y <= 0 then None
                else Some { XSplit = x; YSplit = y })

    let worksheet (name: string) (sheetXml: string) (comments: CellComment list) (condFormats: ConditionalFormat list) (dataVals: DataValidation list) (charts: Chart list) (pivotTables: PivotTable list) (tables: TableDefinition list) (sheetRels: Relationship list) : Worksheet =
        let sheet = XmlDsl.parse sheetXml
        let sheetData = sElem "sheetData" sheet

        let rows =
            sheetData
            |> Option.map (sElems "row")
            |> Option.defaultValue []
            |> List.choose row

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
            Hyperlinks = hyperlinks resolveRel sheet
        }

    /// Parse a column letter run (A, B, ..., AA) into a 1-based column index
    /// matching the rest of the reader's `CellReference.Row/Column` convention.
    let private columnLettersToIndex (letters: string) =
        let mutable acc = 0
        for c in letters do
            if System.Char.IsLetter c then
                acc <- acc * 26 + (int (System.Char.ToUpper c) - int 'A' + 1)
        acc

    /// Parse an `A1:F5` style range into 1-indexed inclusive bounds.
    let private parseRange (ref: string) =
        let parts = ref.Split(':')
        let parseCorner (s: string) =
            let letters = s |> Seq.takeWhile System.Char.IsLetter |> Seq.toArray |> System.String
            let digits = s |> Seq.skipWhile System.Char.IsLetter |> Seq.toArray |> System.String
            let row = if digits.Length > 0 then int digits else 1
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
                        |> Option.bind (fun v -> try Some (int v) with _ -> None)
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
                    Some {
                        StartRow = r1
                        StartCol = c1
                        EndRow = r2
                        EndCol = c2
                        HeaderRowCount = headerCount
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
                Xml.descendants name doc
                |> List.tryHead
                |> Option.bind (Xml.descendants "srgbClr" >> List.tryHead)
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
                Xml.descendants name doc
                |> List.tryHead
                |> Option.bind (fun node ->
                    let srgb = Xml.descendants "srgbClr" node |> List.tryHead |> Option.bind (Xml.attr "val")
                    match srgb with
                    | Some s -> Some s
                    | None ->
                        Xml.descendants "sysClr" node
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
                Xml.descendants slotName doc
                |> List.tryHead
                |> Option.bind (fun slot ->
                    Xml.descendants "latin" slot
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
                match System.Int32.TryParse s with
                | true, n when n >= 0 && n < themeColors.Length -> Some themeColors.[n]
                | _ -> None)
        let indexed =
            Xml.attr "indexed" node
            |> Option.bind (fun s ->
                match System.Int32.TryParse s with
                | true, n -> indexedPaletteColor n
                | _ -> None)
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
                |> Option.bind (fun s ->
                    match System.Double.TryParse(s, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture) with
                    | true, v -> Some v
                    | _ -> None)
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
        let size = sElem "sz" fontNode |> Option.bind (Xml.attr "val") |> Option.map float
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
                        |> Option.bind (fun s ->
                            match System.Double.TryParse(s, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture) with
                            | true, v -> Some v
                            | _ -> None)
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

    let parseNumberFormat (numFmtNode: XmlNode) : string option =
        Xml.attr "formatCode" numFmtNode

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

    let parseCellXf (xfNode: XmlNode) (fonts: (bool * bool * bool * bool * bool * string option * float option * string option * string) array) (fills: string option array) (borders: BorderRecord array) (numFmts: Map<int, string>) (themeMajorFont: string option) (themeMinorFont: string option) : CellStyle =
        let fontId = Xml.attr "fontId" xfNode |> Option.bind (fun id -> Some (int id)) |> Option.defaultValue 0
        let fillId = Xml.attr "fillId" xfNode |> Option.bind (fun id -> Some (int id)) |> Option.defaultValue 0
        let borderId = Xml.attr "borderId" xfNode |> Option.bind (fun id -> Some (int id)) |> Option.defaultValue 0
        let numFmtId = Xml.attr "numFmtId" xfNode |> Option.bind (fun id -> Some (int id))

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

        // Parse fonts
        let fonts =
            sElem "fonts" doc
            |> Option.map (sElems "font")
            |> Option.defaultValue []
            |> List.map (parseFont themeColors)
            |> List.toArray

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
                    let! id = Xml.attr "numFmtId" nf |> Option.map int
                    let! code = parseNumberFormat nf
                    return (id, code)
                })
            |> Map.ofList

        // Parse cell formats (cellXfs)
        sElem "cellXfs" doc
        |> Option.map (sElems "xf")
        |> Option.defaultValue []
        |> List.map (fun xf -> parseCellXf xf fonts fills borders numFmts themeMajorFont themeMinorFont)
        |> List.toArray

    let sharedStrings (sstXml: string) : SharedStringEntry array =
        let sst = XmlDsl.parse sstXml
        sElems "si" sst
        |> List.map (fun si ->
            let runs = richRunsOf si
            if List.isEmpty runs then
                // Plain SI — all text in direct <t> descendants (single or
                // none); preserve original concat-all-<t> behavior.
                let text =
                    Xml.descendants "t" si
                    |> List.choose Xml.text
                    |> String.concat ""
                { Text = text; Runs = [] }
            else
                let text = runs |> List.map fst |> String.concat ""
                { Text = text; Runs = runs })
        |> List.toArray

    // Workbook sheets metadata
    type SheetInfo = {
        Name: string
        SheetId: string
        RelationshipId: string
    }

    let sheets (wbNode: XmlNode) : SheetInfo list =
        let sheetsElem = sElem "sheets" wbNode

        sheetsElem
        |> Option.map (sElems "sheet")
        |> Option.defaultValue []
        |> List.choose (fun sheet ->
            xml {
                let! name = Xml.attr "name" sheet
                let! sheetId = Xml.attr "sheetId" sheet
                let! rId = Xml.attr "r:id" sheet
                return {
                    Name = name
                    SheetId = sheetId
                    RelationshipId = rId
                }
            })

// Cell Comments parsing
module CellComments =
    let parseComment (commentNode: XmlNode) : CellComment option =
        Xml.attr "ref" commentNode
        |> Option.map (fun ref ->
            let cellRef = Xlsx.parseCellRef ref
            let author = Ns.childNs "author" commentNode |> Option.bind Xml.text
            let text =
                Ns.descendantsNs "t" commentNode
                |> List.choose Xml.text
                |> String.concat " "
            {
                CellRef = cellRef
                Author = author
                Text = text
            })

    let parseComments (xml: string) : CellComment list =
        let doc = XmlDsl.parse xml
        Ns.descendantsNs "comment" doc
        |> List.choose parseComment

// Conditional Formatting parsing
module ConditionalFormats =
    let parseRule (ruleNode: XmlNode) : ConditionalFormatRule =
        let ruleType = Xml.attr "type" ruleNode |> Option.defaultValue "cellIs"
        let priority = Xml.attr "priority" ruleNode |> Option.map int
        let stopIfTrue = Xml.attr "stopIfTrue" ruleNode |> Option.map (fun v -> v = "1") |> Option.defaultValue false
        let operator = Xml.attr "operator" ruleNode
        let dxfId = Xml.attr "dxfId" ruleNode |> Option.bind (fun s -> match System.Int32.TryParse s with | true, n -> Some n | _ -> None)
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
                            Xml.attr "rgb" colorNode
                            |> Option.map (fun raw ->
                                let t = raw.TrimStart('#')
                                if t.Length = 8 then t.Substring 2 else t)
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
                    |> Option.bind (Xml.attr "rgb")
                    |> Option.map stripAlpha
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
        {
            Type = ruleType
            Operator = operator
            Formulas = formulas
            DxfId = dxfId
            ColorScale = colorScale
            DataBarColor = dataBarColor
            DataBarCfvo = dataBarCfvo
            IconSet = iconSet
            IconSetCfvo = iconSetCfvo
            Priority = priority
            StopIfTrue = stopIfTrue
        }

    let parseConditionalFormat (cfNode: XmlNode) : ConditionalFormat option =
        Xml.attr "sqref" cfNode
        |> Option.map (fun range ->
            let rules =
                Ns.descendantsNs "cfRule" cfNode
                |> List.map parseRule
            {
                Range = range
                Rules = rules
            })

    let parseConditionalFormats (sheet: XmlNode) : ConditionalFormat list =
        Ns.descendantsNs "conditionalFormatting" sheet
        |> List.choose parseConditionalFormat

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
            let allowBlank = Xml.attr "allowBlank" dvNode |> Option.map (fun v -> v = "1") |> Option.defaultValue true
            let showDropDown = Xml.attr "showDropDown" dvNode |> Option.map (fun v -> v = "0") |> Option.defaultValue true
            let formula1 = Ns.childNs "formula1" dvNode |> Option.bind Xml.text
            let formula2 = Ns.childNs "formula2" dvNode |> Option.bind Xml.text
            let errorMsg = Xml.attr "error" dvNode
            {
                Range = range
                ValidationType = valType
                Operator = operator
                Formula1 = formula1
                Formula2 = formula2
                AllowBlank = allowBlank
                ShowDropDown = showDropDown
                ErrorMessage = errorMsg
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
            let! idx = Xml.attr "fld" fieldNode |> Option.map int
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
    let applyToCell (cell: Cell) : Cell =
        match cell.StyleIndex with
        | Some idx when idx < styles.Length ->
            { cell with Style = Some styles.[idx] }
        | _ -> cell

    let applyToRow (row: Row) : Row =
        { row with Cells = row.Cells |> List.map applyToCell }

    { sheet with Rows = sheet.Rows |> List.map applyToRow }

// Parse XLSX archive
let parseWorkbook (archive: ZipArchive) : Fable.Core.JS.Promise<Workbook option> =
    promise {
        // Load workbook.xml
        let! wbXml = getFileText archive "xl/workbook.xml"

        // Load shared strings
        let! sstXml = getFileText archive "xl/sharedStrings.xml"

        // Load styles
        let! stylesXml = getFileText archive "xl/styles.xml"

        // Load theme (needed before styles so font/fill/border colors can
        // resolve theme="N" + tint references).
        let! themeXml = getFileText archive "xl/theme/theme1.xml"

        // Load relationships
        let! relsXml = getFileText archive "xl/_rels/workbook.xml.rels"

        match wbXml with
        | Some xml ->
            let wb = XmlDsl.parse xml
            let sheetInfos = Xlsx.sheets wb

            let rels =
                relsXml
                |> Option.map XmlDsl.parse
                |> Option.map parseRelationships
                |> Option.defaultValue []

            let sharedStringEntries =
                sstXml
                |> Option.map Xlsx.sharedStrings
                |> Option.defaultValue [||]
            let sharedStrings = sharedStringEntries |> Array.map (fun e -> e.Text)
            let sharedStringRuns = sharedStringEntries |> Array.map (fun e -> e.Runs)

            let themeColors =
                themeXml
                |> Option.map Xlsx.parseThemeColors
                |> Option.defaultValue [|
                    "FFFFFF"; "000000"; "E7E6E6"; "44546A"
                    "4472C4"; "ED7D31"; "A5A5A5"; "FFC000"; "5B9BD5"; "70AD47"
                    "0563C1"; "954F72"
                |]

            let themeMajorFont, themeMinorFont =
                themeXml
                |> Option.map Xlsx.parseThemeFonts
                |> Option.defaultValue (None, None)

            let styles =
                stylesXml
                |> Option.map (Xlsx.parseStyles themeColors themeMajorFont themeMinorFont)
                |> Option.defaultValue [||]

            let dxfs =
                stylesXml
                |> Option.map (XmlDsl.parse >> ConditionalFormats.parseDxfs themeColors)
                |> Option.defaultValue [||]

            let namedRanges = NamedRanges.parseNamedRanges wb

            // Load all worksheets with advanced features
            let! worksheets =
                sheetInfos
                |> List.map (fun info ->
                    promise {
                        let rel = Office.relById info.RelationshipId rels
                        match rel with
                        | Some r ->
                            let sheetPath = "xl/" + ZipInterop.normalizePath r.Target
                            let sheetNum = info.SheetId

                            // Load sheet XML
                            let! sheetXml = getFileText archive sheetPath

                            match sheetXml with
                            | Some xml ->
                                let sheet = XmlDsl.parse xml

                                // Parse conditional formats and data validations from sheet
                                let condFormats = ConditionalFormats.parseConditionalFormats sheet
                                let dataVals = DataValidations.parseDataValidations sheet

                                // Load comments for this sheet
                                let commentsPath = $"xl/comments{sheetNum}.xml"
                                let! commentsXml = getFileText archive commentsPath
                                let comments =
                                    commentsXml
                                    |> Option.map CellComments.parseComments
                                    |> Option.defaultValue []

                                // Load charts for this sheet (from drawings)
                                let drawingsPath = $"xl/drawings/drawing{sheetNum}.xml"
                                let! drawingsXml = getFileText archive drawingsPath
                                let charts =
                                    drawingsXml
                                    |> Option.map (fun _ ->
                                        // Charts are in separate chart*.xml files
                                        // For simplicity, we'll return empty list here
                                        // Full implementation would parse drawing rels and load chart files
                                        [])
                                    |> Option.defaultValue []

                                // Load pivot tables for this sheet
                                let pivotPath = $"xl/pivotTables/pivotTable{sheetNum}.xml"
                                let! pivotXml = getFileText archive pivotPath
                                let pivots =
                                    pivotXml
                                    |> Option.map PivotTables.parsePivotTables
                                    |> Option.defaultValue []

                                // Tables (header/banded styling) aren't wired
                                // through the async loader — only the native
                                // path consumes them today. Pass empty for now.
                                return Some (Xlsx.worksheet info.Name xml comments condFormats dataVals charts pivots [] [] |> applyStyling styles)
                            | None -> return None
                        | None -> return None
                    })
                |> List.toArray
                |> Promise.all

            let sheets = worksheets |> Array.choose id |> Array.toList

            return Some {
                Sheets = sheets
                SharedStrings = sharedStrings
                SharedStringRuns = sharedStringRuns
                Styles = styles
                Dxfs = dxfs
                Relationships = rels
                NamedRanges = namedRanges
                ThemeAccents = [| "4472C4"; "ED7D31"; "A5A5A5"; "FFC000"; "5B9BD5"; "70AD47" |]
            }

        | None -> return None
    }

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
        match System.Double.TryParse(formula) with
        | true, n -> Some n
        | false, _ -> None

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

// Resolve shared string references
let resolveValue (wb: Workbook) (value: CellValue) : CellValue =
    match value with
    | SharedString idx when idx < wb.SharedStrings.Length ->
        String wb.SharedStrings.[idx]
    | other -> other

// Extract sheet as 2D array
let sheetToArray (wb: Workbook) (sheet: Worksheet) : obj array array =
    let maxRow = sheet.Rows |> List.map (fun r -> r.Index) |> List.max
    let maxCol =
        sheet.Rows
        |> List.collect (fun r -> r.Cells)
        |> List.map (fun c ->
            // Simple column letter to number conversion
            c.Reference.Column.ToCharArray()
            |> Array.rev
            |> Array.mapi (fun i ch -> (int ch - int 'A' + 1) * (pown 26 i))
            |> Array.sum)
        |> List.max

    Array.init maxRow (fun r ->
        let rowIdx = r + 1
        let rowCells =
            sheet.Rows
            |> List.tryFind (fun row -> row.Index = rowIdx)
            |> Option.map (fun row -> row.Cells)
            |> Option.defaultValue []
            |> List.map (fun c -> c.Reference.Column, resolveValue wb c.Value)
            |> Map.ofList

        Array.init maxCol (fun c ->
            let colLetter = System.String [| char (c + int 'A') |]
            match Map.tryFind colLetter rowCells with
            | Some (String s) -> box s
            | Some (Number n) -> box n
            | Some (Boolean b) -> box b
            | Some (Date d) -> box d
            | Some (Formula f) -> box f
            | _ -> box ""
        )
    )

#if !FABLE_COMPILER
/// Resolve a relative archive target (e.g. "../tables/table1.xml") against
/// the directory of `basePath`. Walks `..` segments without normalizing back
/// past the archive root.
let private resolveArchivePath (basePath: string) (target: string) =
    let normalized = ZipInterop.normalizePath target
    let baseSegments = basePath.Split('/') |> Array.toList
    let baseDir =
        match baseSegments with
        | [] -> []
        | _ -> baseSegments |> List.take (List.length baseSegments - 1)
    let targetSegments = normalized.Split('/') |> Array.toList
    let rec walk dir parts =
        match parts with
        | [] -> dir
        | ".." :: rest ->
            let trimmed = if List.isEmpty dir then dir else List.take (List.length dir - 1) dir
            walk trimmed rest
        | "." :: rest -> walk dir rest
        | p :: rest -> walk (dir @ [p]) rest
    walk baseDir targetSegments |> String.concat "/"

// Synchronous native loader. Mirrors `parseWorkbook` but uses
// `getFileTextNative` so the native runners don't need a Promise scheduler.
let parseWorkbookNative (archive: ZipArchive) : Workbook option =
    match ZipInterop.getFileTextNative archive "xl/workbook.xml" with
    | None -> None
    | Some wbXml ->
        let wb = XmlDsl.parse wbXml
        let sheetInfos = Xlsx.sheets wb

        let rels =
            ZipInterop.getFileTextNative archive "xl/_rels/workbook.xml.rels"
            |> Option.map (XmlDsl.parse >> parseRelationships)
            |> Option.defaultValue []

        let sharedStringEntries =
            ZipInterop.getFileTextNative archive "xl/sharedStrings.xml"
            |> Option.map Xlsx.sharedStrings
            |> Option.defaultValue [||]
        let sharedStrings = sharedStringEntries |> Array.map (fun e -> e.Text)
        let sharedStringRuns = sharedStringEntries |> Array.map (fun e -> e.Runs)

        let themeXml = ZipInterop.getFileTextNative archive "xl/theme/theme1.xml"

        let themeColors =
            themeXml
            |> Option.map Xlsx.parseThemeColors
            |> Option.defaultValue [|
                "FFFFFF"; "000000"; "E7E6E6"; "44546A"
                "4472C4"; "ED7D31"; "A5A5A5"; "FFC000"; "5B9BD5"; "70AD47"
                "0563C1"; "954F72"
            |]

        let themeMajorFont, themeMinorFont =
            themeXml
            |> Option.map Xlsx.parseThemeFonts
            |> Option.defaultValue (None, None)

        let stylesXmlNative = ZipInterop.getFileTextNative archive "xl/styles.xml"

        let styles =
            stylesXmlNative
            |> Option.map (Xlsx.parseStyles themeColors themeMajorFont themeMinorFont)
            |> Option.defaultValue [||]

        let dxfs =
            stylesXmlNative
            |> Option.map (XmlDsl.parse >> ConditionalFormats.parseDxfs themeColors)
            |> Option.defaultValue [||]

        let themeAccents =
            themeXml
            |> Option.map Xlsx.parseThemeAccents
            |> Option.defaultValue [| "4472C4"; "ED7D31"; "A5A5A5"; "FFC000"; "5B9BD5"; "70AD47" |]

        let namedRanges = NamedRanges.parseNamedRanges wb

        let loadSheetRels (sheetPath: string) : Relationship list =
            let lastSlash = sheetPath.LastIndexOf('/')
            if lastSlash < 0 then []
            else
                let dir = sheetPath.Substring(0, lastSlash)
                let file = sheetPath.Substring(lastSlash + 1)
                let relsPath = sprintf "%s/_rels/%s.rels" dir file
                match ZipInterop.getFileTextNative archive relsPath with
                | None -> []
                | Some relsXml -> relsXml |> XmlDsl.parse |> parseRelationships

        let loadSheetTables (sheetPath: string) (sheetRels: Relationship list) =
            sheetRels
            |> List.filter (fun r -> r.Type.EndsWith("/table"))
            |> List.choose (fun r ->
                let tablePath = resolveArchivePath sheetPath r.Target
                ZipInterop.getFileTextNative archive tablePath
                |> Option.bind Xlsx.parseTable)

        let worksheets =
            sheetInfos
            |> List.choose (fun info ->
                match Office.relById info.RelationshipId rels with
                | Some r ->
                    let sheetPath = "xl/" + ZipInterop.normalizePath r.Target
                    match ZipInterop.getFileTextNative archive sheetPath with
                    | Some xml ->
                        let sheet = XmlDsl.parse xml
                        let condFormats = ConditionalFormats.parseConditionalFormats sheet
                        let dataVals = DataValidations.parseDataValidations sheet
                        let sheetRels = loadSheetRels sheetPath
                        let tables = loadSheetTables sheetPath sheetRels
                        // Cell comments live in xl/comments<sheetNum>.xml. The
                        // sheetN.xml.rels has a Relationship with Type ending
                        // in /comments pointing at the comments XML. Use that
                        // rel to locate the file (sheetNum-based pathing fails
                        // for renamed-on-disk sheets); fall back to the
                        // commonly-conventional comments<N>.xml path when no
                        // rel exists.
                        let comments =
                            sheetRels
                            |> List.tryFind (fun r -> r.Type.EndsWith("/comments"))
                            |> Option.map (fun r ->
                                let p = resolveArchivePath sheetPath r.Target
                                ZipInterop.getFileTextNative archive p
                                |> Option.map CellComments.parseComments
                                |> Option.defaultValue [])
                            |> Option.defaultValue []
                        Some (Xlsx.worksheet info.Name xml comments condFormats dataVals [] [] tables sheetRels)
                    | None -> None
                | None -> None)

        let styledSheets = worksheets |> List.map (applyStyling styles)

        Some {
            Sheets = styledSheets
            SharedStrings = sharedStrings
            SharedStringRuns = sharedStringRuns
            Styles = styles
            Dxfs = dxfs
            Relationships = rels
            NamedRanges = namedRanges
            ThemeAccents = themeAccents
        }

/// Top-level native loader: bytes -> Workbook option.
let loadBytesNative (data: byte[]) : Workbook option =
    let archive = ZipInterop.loadArchiveNative data
    parseWorkbookNative archive
#endif
