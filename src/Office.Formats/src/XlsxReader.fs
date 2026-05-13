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
    FontSize: float option
    NumberFormat: string option
    HorizontalAlign: string option
    VerticalAlign: string option
    WrapText: bool
    BorderTop: CellBorderSide option
    BorderBottom: CellBorderSide option
    BorderLeft: CellBorderSide option
    BorderRight: CellBorderSide option
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

type ConditionalFormatRule = {
    Type: string  // "cellIs", "expression", "colorScale", "dataBar", etc.
    Formula: string option
    Priority: int option
    StopIfTrue: bool
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
}

type Workbook = {
    Sheets: Worksheet list
    SharedStrings: string array
    Styles: CellStyle array
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

    // Cell value extraction with pattern matching. Excel formula cells
    // store both the expression and a cached result in <v>; we can't
    // evaluate the expression, so prefer the cached value (matching what
    // any XLSX viewer shows) and only fall back to the formula text when
    // no cache is present (uncalculated workbooks).
    let cellValue (cNode: XmlNode) : CellValue =
        let cellType = Xml.attr "t" cNode
        let formula = sElem "f" cNode |> Option.bind Xml.text
        let value = sElem "v" cNode |> Option.bind Xml.text

        match cellType, value with
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
            | None -> Empty

    // Cell extraction with applicative style
    let cell (cNode: XmlNode) : Cell option =
        Xml.attr "r" cNode
        |> Option.map (fun ref ->
            let styleIdx = Xml.attr "s" cNode |> Option.map int
            {
                Reference = parseCellRef ref
                Value = cellValue cNode
                StyleIndex = styleIdx
                Style = None  // Resolved later from styles array
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
    let worksheet (name: string) (sheetXml: string) (comments: CellComment list) (condFormats: ConditionalFormat list) (dataVals: DataValidation list) (charts: Chart list) (pivotTables: PivotTable list) (tables: TableDefinition list) : Worksheet =
        let sheet = XmlDsl.parse sheetXml
        let sheetData = sElem "sheetData" sheet

        let rows =
            sheetData
            |> Option.map (sElems "row")
            |> Option.defaultValue []
            |> List.choose row

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
                    Some {
                        StartRow = r1
                        StartCol = c1
                        EndRow = r2
                        EndCol = c2
                        HeaderRowCount = headerCount
                        Style = style
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

    // Style parsing
    let parseFont (fontNode: XmlNode) : (bool * bool * string option * float option * string option) =
        let bold = sElem "b" fontNode |> Option.isSome
        let italic = sElem "i" fontNode |> Option.isSome
        let color = sElem "color" fontNode |> Option.bind (Xml.attr "rgb")
        let size = sElem "sz" fontNode |> Option.bind (Xml.attr "val") |> Option.map float
        let name = sElem "name" fontNode |> Option.bind (Xml.attr "val")
        (bold, italic, color, size, name)

    let parseFill (fillNode: XmlNode) : string option =
        sElem "patternFill" fillNode
        |> Option.bind (sElem "fgColor")
        |> Option.bind (Xml.attr "rgb")

    /// Parse a single <left>/<right>/<top>/<bottom> child of a <border>.
    /// Returns Some when the side has a `style` attribute; None for empty
    /// (unstyled) sides.
    let parseBorderSide (sideNode: XmlNode) : CellBorderSide option =
        Xml.attr "style" sideNode
        |> Option.map (fun style ->
            let color =
                sElem "color" sideNode
                |> Option.bind (Xml.attr "rgb")
            { Style = style; Color = color })

    type BorderRecord = {
        Top: CellBorderSide option
        Bottom: CellBorderSide option
        Left: CellBorderSide option
        Right: CellBorderSide option
    }

    let emptyBorder : BorderRecord = {
        Top = None
        Bottom = None
        Left = None
        Right = None
    }

    let parseBorder (borderNode: XmlNode) : BorderRecord =
        {
            Top = sElem "top" borderNode |> Option.bind parseBorderSide
            Bottom = sElem "bottom" borderNode |> Option.bind parseBorderSide
            Left = sElem "left" borderNode |> Option.bind parseBorderSide
            Right = sElem "right" borderNode |> Option.bind parseBorderSide
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

    let parseCellXf (xfNode: XmlNode) (fonts: (bool * bool * string option * float option * string option) array) (fills: string option array) (borders: BorderRecord array) (numFmts: Map<int, string>) : CellStyle =
        let fontId = Xml.attr "fontId" xfNode |> Option.bind (fun id -> Some (int id)) |> Option.defaultValue 0
        let fillId = Xml.attr "fillId" xfNode |> Option.bind (fun id -> Some (int id)) |> Option.defaultValue 0
        let borderId = Xml.attr "borderId" xfNode |> Option.bind (fun id -> Some (int id)) |> Option.defaultValue 0
        let numFmtId = Xml.attr "numFmtId" xfNode |> Option.bind (fun id -> Some (int id))

        let (bold, italic, fontColor, fontSize, fontName) =
            if fontId < fonts.Length then fonts.[fontId]
            else (false, false, None, None, None)

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
            FontSize = fontSize
            NumberFormat = numFmt
            HorizontalAlign = horizontalAlign
            VerticalAlign = verticalAlign
            WrapText = wrapText
            BorderTop = border.Top
            BorderBottom = border.Bottom
            BorderLeft = border.Left
            BorderRight = border.Right
        }

    let parseStyles (stylesXml: string) : CellStyle array =
        let doc = XmlDsl.parse stylesXml

        // Parse fonts
        let fonts =
            sElem "fonts" doc
            |> Option.map (sElems "font")
            |> Option.defaultValue []
            |> List.map parseFont
            |> List.toArray

        // Parse fills
        let fills =
            sElem "fills" doc
            |> Option.map (sElems "fill")
            |> Option.defaultValue []
            |> List.map parseFill
            |> List.toArray

        // Parse borders
        let borders =
            sElem "borders" doc
            |> Option.map (sElems "border")
            |> Option.defaultValue []
            |> List.map parseBorder
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
        |> List.map (fun xf -> parseCellXf xf fonts fills borders numFmts)
        |> List.toArray

    // Shared strings extraction. An <si> entry can hold either a single
    // <t>text</t> child (plain string) OR a sequence of <r><rPr>…</rPr>
    // <t>part</t></r> runs (rich text, each run carrying its own font/color).
    // Pandas/openpyxl-written workbooks happily mix both forms in the same
    // sharedStrings.xml: phpsp-26template's header cell is rich-text, every
    // other string is plain. We must produce one parsed string per <si> in
    // order, otherwise indices shift and every cell reference past the first
    // rich-text entry loads the wrong text (e.g. A1=5 reads SI 6's content).
    // Collect every descendant <t>'s text and concatenate. Rich-text
    // formatting is dropped — the renderer doesn't support mixed-run text
    // inside a cell yet, but the data round-trips correctly.
    let sharedStrings (sstXml: string) : string array =
        let sst = XmlDsl.parse sstXml

        sElems "si" sst
        |> List.map (fun si ->
            Xml.descendants "t" si
            |> List.choose Xml.text
            |> String.concat "")
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
        let formula = Ns.childNs "formula" ruleNode |> Option.bind Xml.text
        {
            Type = ruleType
            Formula = formula
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

            let sharedStrings =
                sstXml
                |> Option.map Xlsx.sharedStrings
                |> Option.defaultValue [||]

            let styles =
                stylesXml
                |> Option.map Xlsx.parseStyles
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
                                return Some (Xlsx.worksheet info.Name xml comments condFormats dataVals charts pivots [] |> applyStyling styles)
                            | None -> return None
                        | None -> return None
                    })
                |> List.toArray
                |> Promise.all

            let sheets = worksheets |> Array.choose id |> Array.toList

            return Some {
                Sheets = sheets
                SharedStrings = sharedStrings
                Styles = styles
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

        let sharedStrings =
            ZipInterop.getFileTextNative archive "xl/sharedStrings.xml"
            |> Option.map Xlsx.sharedStrings
            |> Option.defaultValue [||]

        let styles =
            ZipInterop.getFileTextNative archive "xl/styles.xml"
            |> Option.map Xlsx.parseStyles
            |> Option.defaultValue [||]

        let themeAccents =
            ZipInterop.getFileTextNative archive "xl/theme/theme1.xml"
            |> Option.map Xlsx.parseThemeAccents
            |> Option.defaultValue [| "4472C4"; "ED7D31"; "A5A5A5"; "FFC000"; "5B9BD5"; "70AD47" |]

        let namedRanges = NamedRanges.parseNamedRanges wb

        let loadSheetTables (sheetPath: string) =
            // xl/worksheets/sheet1.xml -> xl/worksheets/_rels/sheet1.xml.rels
            let lastSlash = sheetPath.LastIndexOf('/')
            if lastSlash < 0 then []
            else
                let dir = sheetPath.Substring(0, lastSlash)
                let file = sheetPath.Substring(lastSlash + 1)
                let relsPath = sprintf "%s/_rels/%s.rels" dir file
                match ZipInterop.getFileTextNative archive relsPath with
                | None -> []
                | Some relsXml ->
                    let sheetRels = relsXml |> XmlDsl.parse |> parseRelationships
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
                        let tables = loadSheetTables sheetPath
                        Some (Xlsx.worksheet info.Name xml [] condFormats dataVals [] [] tables)
                    | None -> None
                | None -> None)

        let styledSheets = worksheets |> List.map (applyStyling styles)

        Some {
            Sheets = styledSheets
            SharedStrings = sharedStrings
            Styles = styles
            Relationships = rels
            NamedRanges = namedRanges
            ThemeAccents = themeAccents
        }

/// Top-level native loader: bytes -> Workbook option.
let loadBytesNative (data: byte[]) : Workbook option =
    let archive = ZipInterop.loadArchiveNative data
    parseWorkbookNative archive
#endif
