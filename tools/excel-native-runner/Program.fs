module ExcelNativeRunner

open System
open System.IO
open Renderer

module Xlsx = XlsxReader

open RedBat.Spreadsheet.Model

let private usage () =
    eprintfn "Usage: dotnet run --project tools/excel-native-runner -- <input.xlsx> <output.png> [width] [height] [sheetIndex|all] [jsonOut]"
    eprintfn "       Pass 'all' as the sheet selector to render every sheet to <output>-sheet<N>.png."
    eprintfn "Defaults: width=1100 height=720 sheetIndex=0"

let private fullPath (path: string) =
    if Path.IsPathRooted path then path
    else Path.GetFullPath(path, Directory.GetCurrentDirectory())

let private existingFile path =
    if File.Exists path then Some path else None

let private findFont () =
    let cwd = Directory.GetCurrentDirectory()
    [
        Environment.GetEnvironmentVariable("RB_WORD_FONT_DIR")
        Path.GetFullPath(Path.Combine(cwd, "artifacts/fonts/windows"))
        Path.GetFullPath(Path.Combine(cwd, "../../artifacts/fonts/windows"))
        // Sibling office repo's font cache, when running from office-spreadsheet/.
        // SpreadsheetGrid.fontDirs already includes this — keep the runner's
        // own font lookup in sync so the default UI font (row/column headers,
        // chrome) loads from the same Calibri Windows TTF Excel rendered with.
        Path.GetFullPath(Path.Combine(cwd, "../office/artifacts/fonts/windows"))
    ]
    |> List.choose (fun dir ->
        if String.IsNullOrWhiteSpace dir then None
        else Path.Combine(dir, "calibri.ttf") |> existingFile)
    |> List.tryHead

let private initFont () =
    match findFont () with
    | Some path -> Composit.Ui.Layout.initFont path
    | None ->
        eprintfn "warning: calibri.ttf not found in RB_WORD_FONT_DIR or known font dirs; row/column headers will fall back to NotoSans"

let private columnIndex (letters: string) : int =
    let mutable acc = 0
    for ch in letters do
        if Char.IsLetter ch then
            acc <- acc * 26 + (int (Char.ToUpper ch) - int 'A' + 1)
    acc - 1

/// Split a `<hyperlink ref="A1:C3"/>` style range string into the
/// inclusive zero-indexed (row, col) corners. Single-cell refs like
/// "B2" collapse to (row, col, row, col). Returns None when the input
/// isn't a recognizable cell-range form.
let private parseCellRange (ref: string) : (int * int * int * int) option =
    let parseCorner (s: string) =
        let letters = s |> Seq.takeWhile Char.IsLetter |> Seq.toArray |> System.String
        let digits = s |> Seq.skipWhile Char.IsLetter |> Seq.toArray |> System.String
        if letters.Length = 0 || digits.Length = 0 then None
        else
            match Int32.TryParse digits with
            | true, r -> Some (r - 1, columnIndex letters)
            | _ -> None
    let parts = ref.Split(':')
    match parts.Length with
    | 1 ->
        parseCorner parts.[0] |> Option.map (fun (r, c) -> r, c, r, c)
    | 2 ->
        match parseCorner parts.[0], parseCorner parts.[1] with
        | Some (r1, c1), Some (r2, c2) -> Some (min r1 r2, min c1 c2, max r1 r2, max c1 c2)
        | _ -> None
    | _ -> None

let private resolveValue (workbook: Xlsx.Workbook) (value: Xlsx.CellValue) : CellValue =
    match value with
    | Xlsx.Number n -> Number n
    | Xlsx.String s -> String s
    | Xlsx.SharedString idx when idx >= 0 && idx < workbook.SharedStrings.Length ->
        String workbook.SharedStrings.[idx]
    | Xlsx.SharedString _ -> Empty
    | Xlsx.Boolean b -> Boolean b
    | Xlsx.Date s -> String s
    | Xlsx.Formula expr -> Formula (expr, Empty)
    | Xlsx.Error e -> Error e
    | Xlsx.Empty -> Empty

/// Translate an Xlsx.RunProps into a RunOverride. Drops empty fields so the
/// renderer can fall back to the cell's base CellFormat.
let private runOverrideOf (props: Xlsx.RunProps) : RunOverride =
    let boolOpt v = if v then Some true else None
    let hexColor (rgb: string) =
        let trimmed = rgb.TrimStart('#')
        if trimmed.Length = 6 then Some ("#" + trimmed) else None
    let vertAlign =
        match props.VertAlign with
        | Xlsx.VertAlignSuperscript -> Some VertAlignSuperscript
        | Xlsx.VertAlignSubscript -> Some VertAlignSubscript
        | Xlsx.VertAlignBaseline -> None
    {
        Bold = boolOpt props.Bold
        Italic = boolOpt props.Italic
        Underline = boolOpt props.Underline
        UnderlineDouble = boolOpt props.UnderlineDouble
        Strikethrough = boolOpt props.Strikethrough
        Color = props.Color |> Option.bind hexColor
        FontName = props.FontName
        FontSize = props.FontSize
        VertAlign = vertAlign
    }

/// Look up rich-text runs for a cell. Two sources feed the renderer's
/// per-run formatting pipeline:
///   * SharedString-backed cells with multi-run entries in the workbook
///     sharedStrings table.
///   * Inline-string cells (`t="inlineStr"`) whose `<is>` element
///     contains `<r>` children with their own `<rPr>` blocks; the runs
///     live on the cell itself rather than in the workbook table.
/// Returns Some runs only when there's actual per-run formatting (>1
/// run, since a single run with no override is fully described by the
/// cell's base format).
let private cellRichRuns (workbook: Xlsx.Workbook) (cell: Xlsx.Cell) : (string * RunOverride) list option =
    let lift runs =
        if List.length runs <= 1 then None
        else Some (runs |> List.map (fun (text, props) -> text, runOverrideOf props))
    if not (List.isEmpty cell.InlineRuns) then
        lift cell.InlineRuns
    else
        match cell.Value with
        | Xlsx.SharedString idx when
            idx >= 0 && idx < workbook.SharedStringRuns.Length
            && not (List.isEmpty workbook.SharedStringRuns.[idx]) ->
            lift workbook.SharedStringRuns.[idx]
        | _ -> None

let private cellFormatFromStyle (style: Xlsx.CellStyle) : CellFormat =
    let hexColor (rgb: string) =
        let trimmed = rgb.TrimStart('#')
        // Excel often prefixes ARGB; drop alpha when 8 chars
        let cleaned = if trimmed.Length = 8 then trimmed.Substring(2) else trimmed
        if cleaned.Length = 6 then Some ("#" + cleaned) else None
    let borderSide (side: Xlsx.CellBorderSide option) : BorderStyle option =
        side
        |> Option.map (fun s ->
            {
                Style = s.Style
                Color = s.Color |> Option.bind hexColor
            })
    {
        Bold = style.Bold
        Italic = style.Italic
        Underline = style.Underline
        UnderlineDouble = style.UnderlineDouble
        Strikethrough = style.Strikethrough
        FontName = style.FontName
        TextColor = style.FontColor |> Option.bind hexColor
        BgColor = style.FillColor |> Option.bind hexColor
        FontSize = style.FontSize |> Option.map int
        HorizontalAlign = style.HorizontalAlign
        VerticalAlign = style.VerticalAlign
        WrapText = style.WrapText
        Indent = style.Indent
        ShrinkToFit = style.ShrinkToFit
        TextRotation = style.TextRotation
        NumberFormat = style.NumberFormat
        BorderTop = borderSide style.BorderTop
        BorderBottom = borderSide style.BorderBottom
        BorderLeft = borderSide style.BorderLeft
        BorderRight = borderSide style.BorderRight
        BorderDiagonal = borderSide style.BorderDiagonal
        DiagonalUp = style.DiagonalUp
        DiagonalDown = style.DiagonalDown
        FilterButton = false
        HasComment = false
        HasListDropdown = false
        DataBar = None
        IconSet = None
    }

    // Layer the workbook's table style (e.g. `TableStyleMedium7`) on top of
    // the cellXf-derived CellFormat. Excel renders the table header row with
    // a solid accent fill + white bold text, and (when ShowRowStripes is set)
    // tints alternating data rows with a very light blend of the same accent.
    // We only override fields the cellXf didn't already set, so explicit
    // colors in the file still win.
let private tintHex (hex: string) (t: float) =
    let parseByte (s: string) =
        try System.Convert.ToByte(s, 16) with _ -> 0uy
    let lerp (c: byte) =
        let f = float c
        byte (int (System.Math.Round(f + t * (255.0 - f))))
    if hex.Length < 6 then hex
    else
        sprintf "#%02X%02X%02X"
            (lerp (parseByte (hex.Substring(0,2))))
            (lerp (parseByte (hex.Substring(2,2))))
            (lerp (parseByte (hex.Substring(4,2))))

let private styleNumber (name: string) =
    let digits = name |> Seq.skipWhile (System.Char.IsLetter) |> Seq.toArray |> System.String
    match System.Int32.TryParse digits with
    | true, n -> Some n
    | _ -> None

/// Map a `TableStyleMediumN` name to an accent color from the workbook theme.
/// Medium2-7 → accent1-6; higher Medium variants cycle through the same six
/// accents (the rendering differences between Medium7/Medium14/Medium21 are
/// header banding details we don't reproduce yet).
let private tableAccentColor (themeAccents: string array) (styleName: string) : string option =
    let n = styleNumber styleName
    match n with
    | Some n when n >= 2 ->
        let idx = (n - 2) % 6
        if idx < themeAccents.Length then Some ("#" + themeAccents.[idx]) else None
    | _ -> None

type TableOverlay = {
    BgColor: string option
    TextColor: string option
    Bold: bool
    /// Set on header cells when the table declares an <autoFilter> covering
    /// this column. Renderer paints a dropdown triangle on the right side.
    FilterButton: bool
}

/// For each cell coordinate inside a table range, return the table-derived
/// overlay (background + bold + font color). Excel's TableStyleMedium
/// variants don't draw heavy inner borders — the accent fill + banded
/// stripes are the only visual separation, with the workbook's default
/// gridlines showing through where they survive the fill. Returns None
/// for cells outside any table. Public so the `spreadsheet-vba-layout`
/// harness can subtract the overlay's `Bold` from its TSV output (Excel
/// COM reports the underlying cell font, not the table-style overlay).
let tableOverrideFor (tables: Xlsx.TableDefinition list) (themeAccents: string array) (row: int) (col: int) : TableOverlay option =
    // row/col are 1-indexed to match Xlsx.TableDefinition.
    tables
    |> List.tryPick (fun t ->
        if row < t.StartRow || row > t.EndRow || col < t.StartCol || col > t.EndCol then None
        else
            match t.Style with
            | None -> None
            | Some style ->
                // Light/Dark variants don't carry the same accent banding;
                // restrict the overlay to the Medium family for now.
                if not (style.Name.StartsWith "TableStyleMedium") then None
                else
                    match tableAccentColor themeAccents style.Name with
                    | None -> None
                    | Some accentHex ->
                        let isHeader = row < t.StartRow + t.HeaderRowCount
                        if isHeader then
                            let hasFilter =
                                match t.AutoFilter with
                                | Some (fc1, fc2) -> col >= fc1 && col <= fc2
                                | None -> false
                            Some {
                                BgColor = Some accentHex
                                TextColor = Some "#FFFFFF"
                                Bold = true
                                FilterButton = hasFilter }
                        elif style.ShowRowStripes then
                            // Excel bands every other data row. Pixel-sampled
                            // against artifacts/excel/pandas-test_decimal-
                            // headings.png: the first data row (immediately
                            // below the header) renders at #E2EFDA (light
                            // accent tint); the second is white; alternates.
                            let dataIdx = row - t.StartRow - t.HeaderRowCount
                            if dataIdx % 2 = 0 then
                                Some {
                                    BgColor = Some (tintHex (accentHex.TrimStart('#')) 0.80)
                                    TextColor = None
                                    Bold = false
                                    FilterButton = false }
                            else None
                        else None)

let toModel (sheet: Xlsx.Worksheet) (workbook: Xlsx.Workbook) : SpreadsheetModel =
    let mutable maxRow = 0
    let mutable maxCol = 0
    let mutable richRunsMap : Map<CellAddress, (string * RunOverride) list> = Map.empty

    // Resolve hyperlink ranges once: a HashSet of (row, col) cells that
    // get the Excel "this is a link" formatting override. Excel's
    // built-in Hyperlink cell style is `<color theme="10"/>` (the theme
    // 10 slot = hlink, default #0563C1) plus an underline. We fold both
    // into the cell's format below.
    let hyperlinkCells = System.Collections.Generic.HashSet<int * int>()
    for hl in sheet.Hyperlinks do
        match parseCellRange hl.Range with
        | None -> ()
        | Some (r1, c1, r2, c2) ->
            for r in r1 .. r2 do
                for c in c1 .. c2 do
                    hyperlinkCells.Add(r, c) |> ignore

    // Cells that have a <comment> attached get a red corner triangle.
    // sheet.Comments lives in a separate xl/comments<N>.xml file, keyed
    // by cell ref in Excel's 1-indexed "B7" form.
    let commentCells = System.Collections.Generic.HashSet<int * int>()
    for c in sheet.Comments do
        let r = c.CellRef.Row - 1
        let col = columnIndex c.CellRef.Column
        if r >= 0 && col >= 0 then commentCells.Add(r, col) |> ignore

    // Cells inside a `<dataValidation type="list">` range get the small
    // downward dropdown arrow. Other validation types (whole, decimal,
    // date, custom formula) don't paint a visible marker — Excel only
    // shows the dropdown for list-style validation.
    let listDropdownCells = System.Collections.Generic.HashSet<int * int>()
    for dv in sheet.DataValidations do
        if dv.ValidationType = "list" then
            match parseCellRange dv.Range with
            | None -> ()
            | Some (r1, c1, r2, c2) ->
                for r in r1 .. r2 do
                    for c in c1 .. c2 do
                        listDropdownCells.Add(r, c) |> ignore

    let cells =
        sheet.Rows
        |> List.collect (fun row ->
            row.Cells
            |> List.choose (fun cell ->
                let r = cell.Reference.Row - 1
                let c = columnIndex cell.Reference.Column
                if r < 0 || c < 0 then None
                else
                    let rawValue = resolveValue workbook cell.Value
                    // `<sheetView showZeros="false"/>` hides numeric-zero
                    // cells. Replace value with Empty so the renderer
                    // skips text emission but keeps fills/borders intact.
                    let value =
                        if not sheet.ShowZeros then
                            match rawValue with
                            | Number 0.0 -> Empty
                            | Formula (expr, Number 0.0) -> Formula (expr, Empty)
                            | other -> other
                        else rawValue
                    let baseFormat =
                        cell.Style
                        |> Option.map cellFormatFromStyle
                        |> Option.defaultValue defaultFormat
                    let withTable =
                        match tableOverrideFor sheet.Tables workbook.ThemeAccents cell.Reference.Row (c + 1) with
                        | None -> baseFormat
                        | Some overlay ->
                            { baseFormat with
                                BgColor = baseFormat.BgColor |> Option.orElse overlay.BgColor
                                TextColor = baseFormat.TextColor |> Option.orElse overlay.TextColor
                                Bold = baseFormat.Bold || overlay.Bold
                                FilterButton = baseFormat.FilterButton || overlay.FilterButton }
                    let withLink =
                        if hyperlinkCells.Contains(r, c) then
                            // Theme-10 hyperlink color #0563C1 + underline.
                            // Override TextColor unconditionally — Excel's
                            // built-in Hyperlink style takes precedence over
                            // the default theme-1 (=text) color that cellXfs
                            // pick up automatically. Cells that explicitly
                            // need a different link color have to use a
                            // Followed Hyperlink cellStyle, which we'd hook
                            // separately.
                            { withTable with
                                TextColor = Some "#0563C1"
                                Underline = true }
                        else withTable
                    let format =
                        { withLink with
                            HasComment = withLink.HasComment || commentCells.Contains(r, c)
                            HasListDropdown = withLink.HasListDropdown || listDropdownCells.Contains(r, c) }
                    let cellData = { Value = value; Format = format; Comment = None }
                    match cellRichRuns workbook cell with
                    | Some runs -> richRunsMap <- Map.add { Row = r; Col = c } runs richRunsMap
                    | None -> ()
                    if r > maxRow then maxRow <- r
                    if c > maxCol then maxCol <- c
                    Some ({ Row = r; Col = c }, cellData)))
        |> Map.ofList

    // ---- Column-default fill for missing-cell positions ----------------
    // Excel applies `<col style="N">` to cells NOT present in <sheetData>
    // (per OOXML §18.3.1.13). Cells that exist with their own `s=`
    // already win; cells that exist without `s=` use cellXf 0. We fill
    // the gaps in the used range with empty-but-styled cells so the
    // column's default fill / borders render. Bounded by maxRow / maxCol
    // to avoid materialising the whole 1M-row grid for sparse sheets.
    let cells =
        let colStyles : Map<int, int> =
            sheet.Columns
            |> List.collect (fun col ->
                match col.Style with
                | None -> []
                | Some sIdx -> [ for c1 in col.Index .. col.MaxIndex -> c1 - 1, sIdx ])
            |> Map.ofList
        if Map.isEmpty colStyles then cells
        else
            let mutable result = cells
            for c0 in 0 .. maxCol do
                match Map.tryFind c0 colStyles with
                | None -> ()
                | Some sIdx when sIdx < 0 || sIdx >= workbook.Styles.Length -> ()
                | Some sIdx ->
                    let style = workbook.Styles.[sIdx]
                    let format = cellFormatFromStyle style
                    // Skip when the style is just the no-op default
                    // (no fill, no borders, no font override).
                    let isInteresting =
                        format.BgColor.IsSome
                        || format.BorderTop.IsSome
                        || format.BorderBottom.IsSome
                        || format.BorderLeft.IsSome
                        || format.BorderRight.IsSome
                    if isInteresting then
                        for r0 in 0 .. maxRow do
                            let addr = { Row = r0; Col = c0 }
                            if not (Map.containsKey addr result) then
                                result <- Map.add addr { Value = Empty; Format = format; Comment = None } result
            result

    // ---- Conditional formatting ----------------------------------------
    // Walk every <conditionalFormatting> block, evaluate its rules against
    // the cells already in the map, and overlay the resolved format. Two
    // rule kinds are handled: cellIs (operator + formula → a <dxf> style
    // overlay) and colorScale (2/3 stops → an interpolated background by
    // the cell's position in the range's min..max). Other kinds (dataBar,
    // iconSet, expression, top10) pass through untouched.
    let preCfCells = cells
    let cells =
        if List.isEmpty sheet.ConditionalFormats then cells
        else
            let numericOf (v: CellValue) : float option =
                match v with
                | Number n -> Some n
                | Boolean b -> Some (if b then 1.0 else 0.0)
                | Formula (_, Number n) -> Some n
                | _ -> None
            // sqref can list several space-separated ranges.
            let expandRanges (sqref: string) =
                sqref.Split([| ' ' |], StringSplitOptions.RemoveEmptyEntries)
                |> Array.choose parseCellRange
                |> Array.collect (fun (r1, c1, r2, c2) ->
                    [| for r in r1 .. r2 do for c in c1 .. c2 -> r, c |])
            let parseNum (s: string) =
                match Double.TryParse(s, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
                | true, n -> Some n
                | _ -> None
            let cellIsMatch (op: string) (formulas: string list) (n: float) =
                let f1 = formulas |> List.tryItem 0 |> Option.bind parseNum
                let f2 = formulas |> List.tryItem 1 |> Option.bind parseNum
                match op, f1, f2 with
                | "greaterThan", Some a, _ -> n > a
                | "lessThan", Some a, _ -> n < a
                | "greaterThanOrEqual", Some a, _ -> n >= a
                | "lessThanOrEqual", Some a, _ -> n <= a
                | "equal", Some a, _ -> n = a
                | "notEqual", Some a, _ -> n <> a
                | "between", Some a, Some b -> n >= min a b && n <= max a b
                | "notBetween", Some a, Some b -> n < min a b || n > max a b
                | _ -> false
            // Blend two RRGGBB hex colors at t ∈ [0,1].
            let blend (a: string) (b: string) (t: float) =
                let hx (s: string) =
                    let s = s.TrimStart('#')
                    if s.Length = 8 then s.Substring 2 else s
                let ah, bh = hx a, hx b
                if ah.Length <> 6 || bh.Length <> 6 then a
                else
                    let chan i =
                        let av = Convert.ToInt32(ah.Substring(i, 2), 16)
                        let bv = Convert.ToInt32(bh.Substring(i, 2), 16)
                        int (round (float av * (1.0 - t) + float bv * t)) |> max 0 |> min 255
                    sprintf "#%02X%02X%02X" (chan 0) (chan 2) (chan 4)
            let mutable acc = cells
            for cf in sheet.ConditionalFormats do
                let coords = expandRanges cf.Range
                // Numeric values present in the range — needed for
                // colorScale min/max anchoring.
                let rangeNums =
                    coords
                    |> Array.choose (fun (r, c) ->
                        Map.tryFind { Row = r; Col = c } acc
                        |> Option.bind (fun cd -> numericOf cd.Value))
                let rangeMin = if rangeNums.Length > 0 then Array.min rangeNums else 0.0
                let rangeMax = if rangeNums.Length > 0 then Array.max rangeNums else 1.0
                // Lower priority number = higher precedence; apply
                // highest-precedence rule last so it wins.
                let rules =
                    cf.Rules
                    |> List.sortByDescending (fun rule -> rule.Priority |> Option.defaultValue 0)
                for rule in rules do
                    match rule.Type with
                    | "cellIs" ->
                        let dxf =
                            rule.DxfId
                            |> Option.bind (fun i ->
                                if i >= 0 && i < workbook.Dxfs.Length then Some workbook.Dxfs.[i] else None)
                        match rule.Operator, dxf with
                        | Some op, Some d ->
                            for (r, c) in coords do
                                let addr = { Row = r; Col = c }
                                match Map.tryFind addr acc with
                                | Some cd ->
                                    match numericOf cd.Value with
                                    | Some n when cellIsMatch op rule.Formulas n ->
                                        let fmt =
                                            { cd.Format with
                                                BgColor = d.FillColor |> Option.orElse cd.Format.BgColor
                                                TextColor = d.FontColor |> Option.orElse cd.Format.TextColor
                                                Bold = (d.Bold |> Option.defaultValue cd.Format.Bold)
                                                Italic = (d.Italic |> Option.defaultValue cd.Format.Italic) }
                                        acc <- Map.add addr { cd with Format = fmt } acc
                                    | _ -> ()
                                | None -> ()
                        | _ -> ()
                    | "colorScale" when rule.ColorScale.Length >= 2 ->
                        let stops = rule.ColorScale
                        let lo = List.head stops
                        let hi = List.last stops
                        let mid = if stops.Length >= 3 then Some (stops.[1]) else None
                        let span = rangeMax - rangeMin
                        for (r, c) in coords do
                            let addr = { Row = r; Col = c }
                            match Map.tryFind addr acc with
                            | Some cd ->
                                match numericOf cd.Value with
                                | Some n ->
                                    let t = if span = 0.0 then 0.5 else (n - rangeMin) / span
                                    let color =
                                        match mid with
                                        | Some m when t <= 0.5 -> blend lo.Color m.Color (t * 2.0)
                                        | Some m -> blend m.Color hi.Color ((t - 0.5) * 2.0)
                                        | None -> blend lo.Color hi.Color t
                                    acc <- Map.add addr { cd with Format = { cd.Format with BgColor = Some color } } acc
                                | None -> ()
                            | None -> ()
                    | "dataBar" ->
                        // Bar length is the cell value's position in the
                        // range min..max, clamped to [0,1]. Excel's default
                        // dataBar leaves a small gap at min; we keep it
                        // simple and let 0 → empty, max → full-width.
                        let barColor = rule.DataBarColor |> Option.defaultValue "638EC6"
                        let span = rangeMax - rangeMin
                        for (r, c) in coords do
                            let addr = { Row = r; Col = c }
                            match Map.tryFind addr acc with
                            | Some cd ->
                                match numericOf cd.Value with
                                | Some n ->
                                    let frac = if span = 0.0 then 1.0 else (n - rangeMin) / span
                                    let frac = max 0.0 (min 1.0 frac)
                                    acc <- Map.add addr { cd with Format = { cd.Format with DataBar = Some (frac, "#" + barColor) } } acc
                                | None -> ()
                            | None -> ()
                    | "iconSet" ->
                        // Choose the icon by the value's percentile within
                        // the range. The iconSet's cfvo list gives N
                        // thresholds (first is the implicit 0% floor); icon
                        // i is shown when value ≥ threshold i. We resolve
                        // each cfvo to an absolute value: percent / percentile
                        // map onto the range, num is literal.
                        let iconCount = max 3 (List.length rule.IconSetCfvo)
                        let span = rangeMax - rangeMin
                        let thresholdAbs (cfvoType: string) (cfvoVal: string option) =
                            let v = cfvoVal |> Option.bind parseNum |> Option.defaultValue 0.0
                            match cfvoType with
                            | "percent" | "percentile" -> rangeMin + span * (v / 100.0)
                            | "num" -> v
                            | "min" -> rangeMin
                            | "max" -> rangeMax
                            | _ -> rangeMin + span * (v / 100.0)
                        let thresholds =
                            rule.IconSetCfvo
                            |> List.map (fun (t, v) -> thresholdAbs t v)
                            |> List.toArray
                        for (r, c) in coords do
                            let addr = { Row = r; Col = c }
                            match Map.tryFind addr acc with
                            | Some cd ->
                                match numericOf cd.Value with
                                | Some n ->
                                    // Highest threshold the value clears.
                                    let mutable idx = 0
                                    for i in 0 .. thresholds.Length - 1 do
                                        if n >= thresholds.[i] then idx <- i
                                    acc <- Map.add addr { cd with Format = { cd.Format with IconSet = Some (idx, iconCount) } } acc
                                | None -> ()
                            | None -> ()
                    | "top10" ->
                        // Match the top (or bottom) N cells of the range.
                        // rank is a count, or a percentage of the numeric
                        // cell count when the percent flag is set.
                        let dxf =
                            rule.DxfId
                            |> Option.bind (fun i ->
                                if i >= 0 && i < workbook.Dxfs.Length then Some workbook.Dxfs.[i] else None)
                        match dxf with
                        | Some d when rangeNums.Length > 0 ->
                            let rank = rule.Rank |> Option.defaultValue 10
                            let n =
                                if rule.Percent then
                                    int (ceil (float rangeNums.Length * float rank / 100.0))
                                else rank
                            let n = max 1 (min rangeNums.Length n)
                            // Cutoff value: the n-th largest (top) or n-th
                            // smallest (bottom). Ties at the cutoff all match
                            // — same as Excel.
                            let sorted =
                                if rule.Bottom then Array.sort rangeNums
                                else Array.sortDescending rangeNums
                            let cutoff = sorted.[n - 1]
                            for (r, c) in coords do
                                let addr = { Row = r; Col = c }
                                match Map.tryFind addr acc with
                                | Some cd ->
                                    match numericOf cd.Value with
                                    | Some v ->
                                        let isMatch =
                                            if rule.Bottom then v <= cutoff else v >= cutoff
                                        if isMatch then
                                            let fmt =
                                                { cd.Format with
                                                    BgColor = d.FillColor |> Option.orElse cd.Format.BgColor
                                                    TextColor = d.FontColor |> Option.orElse cd.Format.TextColor
                                                    Bold = (d.Bold |> Option.defaultValue cd.Format.Bold)
                                                    Italic = (d.Italic |> Option.defaultValue cd.Format.Italic) }
                                            acc <- Map.add addr { cd with Format = fmt } acc
                                    | None -> ()
                                | None -> ()
                        | _ -> ()
                    | "aboveAverage" ->
                        // Match cells above (or below) the range mean,
                        // optionally beyond N standard deviations.
                        let dxf =
                            rule.DxfId
                            |> Option.bind (fun i ->
                                if i >= 0 && i < workbook.Dxfs.Length then Some workbook.Dxfs.[i] else None)
                        match dxf with
                        | Some d when rangeNums.Length > 0 ->
                            let mean = Array.average rangeNums
                            let threshold =
                                match rule.StdDev with
                                | Some sd when sd > 0 ->
                                    let variance =
                                        rangeNums
                                        |> Array.averageBy (fun v -> (v - mean) ** 2.0)
                                    let sigma = sqrt variance
                                    if rule.AboveAverage then mean + float sd * sigma
                                    else mean - float sd * sigma
                                | _ -> mean
                            for (r, c) in coords do
                                let addr = { Row = r; Col = c }
                                match Map.tryFind addr acc with
                                | Some cd ->
                                    match numericOf cd.Value with
                                    | Some v ->
                                        let isMatch =
                                            if rule.AboveAverage then
                                                if rule.EqualAverage then v >= threshold else v > threshold
                                            else
                                                if rule.EqualAverage then v <= threshold else v < threshold
                                        if isMatch then
                                            let fmt =
                                                { cd.Format with
                                                    BgColor = d.FillColor |> Option.orElse cd.Format.BgColor
                                                    TextColor = d.FontColor |> Option.orElse cd.Format.TextColor
                                                    Bold = (d.Bold |> Option.defaultValue cd.Format.Bold)
                                                    Italic = (d.Italic |> Option.defaultValue cd.Format.Italic) }
                                            acc <- Map.add addr { cd with Format = fmt } acc
                                    | None -> ()
                                | None -> ()
                        | _ -> ()
                    | "expression" ->
                        // Per-cell formula rule. The single <formula> is
                        // written relative to the range's top-left; shift
                        // its relative refs onto each target cell, evaluate
                        // against the sheet, and overlay the dxf where the
                        // result is truthy.
                        let dxf =
                            rule.DxfId
                            |> Option.bind (fun i ->
                                if i >= 0 && i < workbook.Dxfs.Length then Some workbook.Dxfs.[i] else None)
                        match rule.Formulas, dxf with
                        | formulaText :: _, Some d ->
                            let anchorR, anchorC =
                                cf.Range.Split([| ' ' |], StringSplitOptions.RemoveEmptyEntries)
                                |> Array.tryHead
                                |> Option.bind parseCellRange
                                |> Option.map (fun (r1, c1, _, _) -> r1, c1)
                                |> Option.defaultValue (0, 0)
                            match RedBat.Spreadsheet.Formula.parseFormula ("=" + formulaText) with
                            | Ok baseExpr ->
                                let evalModel = { createEmpty 1048576 16384 with Cells = acc }
                                let truthy (v: CellValue) =
                                    match v with
                                    | Boolean b -> b
                                    | Number n -> n <> 0.0
                                    | Formula (_, Boolean b) -> b
                                    | Formula (_, Number n) -> n <> 0.0
                                    | _ -> false
                                for (r, c) in coords do
                                    let addr = { Row = r; Col = c }
                                    match Map.tryFind addr acc with
                                    | Some cd ->
                                        let shifted =
                                            RedBat.Spreadsheet.Formula.shiftExpr (r - anchorR) (c - anchorC) baseExpr
                                        if truthy (RedBat.Spreadsheet.Formula.eval shifted (RedBat.Spreadsheet.Formula.contextOf evalModel)) then
                                            let fmt =
                                                { cd.Format with
                                                    BgColor = d.FillColor |> Option.orElse cd.Format.BgColor
                                                    TextColor = d.FontColor |> Option.orElse cd.Format.TextColor
                                                    Bold = (d.Bold |> Option.defaultValue cd.Format.Bold)
                                                    Italic = (d.Italic |> Option.defaultValue cd.Format.Italic) }
                                            acc <- Map.add addr { cd with Format = fmt } acc
                                    | None -> ()
                            | Result.Error _ -> ()
                        | _ -> ()
                    | _ -> ()
            acc

    // Diff CF-modified cells against the pre-CF snapshot to identify
    // which cells got their Bold flag from a CF dxf rather than the
    // underlying cellXf. The layout harness uses this set to report
    // `bold=false` for these cells — matches Excel COM's
    // `cell.Font.Bold` (CF doesn't bake into the font).
    let cfBoldCells =
        if List.isEmpty sheet.ConditionalFormats then Set.empty
        else
            cells
            |> Map.toSeq
            |> Seq.choose (fun (addr, cd) ->
                let preBold =
                    preCfCells
                    |> Map.tryFind addr
                    |> Option.map (fun pre -> pre.Format.Bold)
                    |> Option.defaultValue false
                if cd.Format.Bold && not preBold then Some addr else None)
            |> Set.ofSeq

    let mergedCells =
        sheet.MergedCells
        |> List.map (fun mc ->
            {
                StartRow = mc.StartRef.Row - 1
                StartCol = columnIndex mc.StartRef.Column
                EndRow = mc.EndRef.Row - 1
                EndCol = columnIndex mc.EndRef.Column
            })

    // Excel col.Width is in character units. The chars-to-pixels ratio is
    // the "Maximum Digit Width" (MDW) of the workbook's default font —
    // technically the rendered pixel width of the widest digit 0-9 in that
    // font at the workbook's default size, as Excel computes it. Use the
    // real loaded font instead of a Calibri-anchored linear fit, since
    // fonts vary enough at the same point size to skew column widths
    // (Verdana 10 needs 8 px/char, not the 7 a Calibri-only estimate gives).
    let defaultFontSize =
        if workbook.Styles.Length > 0 then
            workbook.Styles.[0].FontSize |> Option.defaultValue 11.0
        else 11.0
    let defaultFontName =
        if workbook.Styles.Length > 0 then
            workbook.Styles.[0].FontName
        else None
    let mdw =
        try
            let font = Composit.Ui.SpreadsheetGrid.rendererFontFor defaultFontName false false
            let pxSize = defaultFontSize * 96.0 / 72.0
            // Excel's MDW table is the integer pixel advance Excel uses for
            // chars-to-pixels column-width math at 96 DPI. Empirically:
            //   Calibri 11 = 7  (measured 7.43)
            //   Calibri 12 = 8  (measured 8.11)
            //   Calibri 14 = 10 (measured 9.46) — synthetic-vertalign
            //   Calibri 18 = 14 (measured 12.16, Microsoft table value)
            //   Arial   10 = 7  (measured 7.42)
            //   Verdana 10 = 8  (measured 8.48)
            // Floor alone matches Calibri 11/12 (where measured truncates
            // cleanly) but falls short at 14+ because Excel's GDI hinting
            // bumps the glyph advance past the raw font-table value. The
            // observation `Excel_MDW = pt - 4` lines up across all known
            // Calibri sizes ≥ 11; `max(floor(measured), pt-4)` therefore
            // covers both the small-size (floor wins, e.g. Verdana 10 = 8)
            // and large-size (pt-4 wins, e.g. Calibri 14 = 10) regimes.
            let zero = Renderer.Text.measureWidth font pxSize "0"
            let measured = System.Math.Floor zero
            let ptBased = defaultFontSize - 4.0
            max 7.0 (max measured ptBased)
        with _ ->
            // Renderer unavailable (font cache miss, etc.) — fall back to
            // the original Calibri-anchored linear estimate.
            max 7.0 (System.Math.Round(defaultFontSize * 7.0 / 11.0))
    // OOXML §18.3.1.13 width formula:
    //   width_stored = TRUNC((chars * mdw + 5) / mdw * 256) / 256
    // So when we have a STORED <col width="N"> attribute, N already
    // encodes the 5-px padding adjustment: pixels = N * mdw with no
    // additional padding. When the column has NO stored width (uses the
    // sheet default chars value), the padding must be added explicitly:
    // pixels = chars * mdw + 5. Verified against phpsp-26template's col B
    // (stored 86.285 -> 604 px = 453 pt) and default 8.43 chars columns
    // (-> 64 px = 48 pt).
    let charsToPixelsDefault (chars: float) = chars * mdw + 5.0
    // Excel reports `cell.Width` in integer pixels (it rounds away the
    // 1/256 chars granularity from the stored width attr). Verified against
    // pandas-multiindex_no_index_names: stored width 5.41 × MDW 7 = 37.87
    // px which Excel rounds to 38 px (= 28.5 pt). Without rounding our
    // column x-positions drift by ~0.1 px per column, accumulating over a
    // wide sheet (the row 1 cells in col F sat 0.6 pt to the left of where
    // Excel placed them, 5 narrow columns of error).
    let storedWidthToPixels (storedWidth: float) =
        System.Math.Round(storedWidth * mdw)

    // sheetFormatPr defaults. Excel uses these for every column without an
    // explicit <col width="..."> entry. Without seeding the defaults across
    // used columns, our renderer falls back to 100 px per column and drifts
    // every column past A. Excel's implicit default when sheetFormatPr is
    // absent is 8.43 character units (the same value Excel writes when you
    // explicitly request "AutoFit Column Width" on an unsized column).
    let defaultColWidthPx =
        match sheet.Format.DefaultColWidth with
        // defaultColWidth attribute is stored using OOXML §18.3.1.13's
        // encoded form (same as <col width>) — pixels = stored*MDW. Verified
        // against pandas-dimension_large (defaultColWidth=11.55, MDW=7 → 81
        // px = 60.75 pt, matches Excel COM cell.Width).
        | Some w -> storedWidthToPixels w
        | None ->
            match sheet.Format.BaseColWidth with
            // OOXML §18.3.1.13 spec: when only baseColWidth is given,
            // defaultColWidth = baseColWidth + (cellPaddingPx + 1 grid) / MDW
            // with cellPadding = 2 px (Excel) for a total of 3 px. So the
            // effective char count is baseColWidth + 3/MDW. Combined with the
            // chars→pixels formula (chars*MDW + 5) this collapses to
            // baseColWidth*MDW + 8. Verified against pandas-test1
            // (baseColWidth=10, MDW=8 → 88 px = 66 pt, matches Excel COM
            // cell.Width = 66) and synthetic-print (implicit baseColWidth=8,
            // MDW=7 → 8.43 chars → 64 px = 48 pt).
            | Some baseW -> charsToPixelsDefault (baseW + 3.0 / mdw)
            | None -> charsToPixelsDefault 8.43

    let defaultRowHeightPx =
        sheet.Format.DefaultRowHeight |> Option.map (fun h -> h * (96.0 / 72.0))

    // Excel's font-derived auto row height. Derived from the font's
    // OS/2 typographic metrics:
    //
    //   pixels = ceil(typoAscender_px) + ceil(|typoDescender|_px)
    //          + ceil(typoLineGap_px) + 1
    //
    // The trick is that Excel ceils each metric SEPARATELY before summing,
    // not the total winLineHeight. Each metric snaps to its own integer-
    // pixel slot in the line box, then a single 1-px leading sits below.
    //
    // Empirical verification (pixels @ 96 dpi, all match Excel exactly):
    //   - Calibri 11: ceil(11.0)+ceil(3.67)+ceil(3.23)+1  = 11+4+4+1 = 20
    //   - Calibri 12: ceil(12.0)+ceil(4.0) +ceil(3.53)+1  = 12+4+4+1 = 21
    //   - Calibri 14: ceil(14.0)+ceil(4.67)+ceil(4.12)+1  = 14+5+5+1 = 25
    //   - Calibri 18: ceil(18.0)+ceil(6.0) +ceil(5.30)+1  = 18+6+6+1 = 31
    //   - Arial 11:   ceil(10.68)+ceil(3.09)+ceil(2.20)+1 = 11+4+3+1 = 19
    let autoRowHeightPt (familyName: string option) (fontPt: float) : float =
        // Fonts we don't ship a TTF for: hard-code the auto-height ratio.
        // Arial Black has heavier metrics than Arial Regular (typoAscender
        // ~2050 vs Arial's 1854). We don't have ariblk.ttf so we can't
        // derive from the font; the 1.7× ratio matches Excel's reported
        // 18.75 pt at 11 pt and 24 pt at 14 pt.
        let unshippedFontRatio =
            match familyName with
            | Some n when n.Trim().Equals("Arial Black", System.StringComparison.OrdinalIgnoreCase) ->
                Some 1.7
            | _ -> None
        match unshippedFontRatio with
        | Some ratio ->
            // Round to the 0.75-pt grid Excel snaps row heights to.
            let raw = fontPt * ratio
            System.Math.Round(raw * 4.0 / 3.0) * 3.0 / 4.0
        | None ->
        try
            let font = Composit.Ui.SpreadsheetGrid.rendererFontFor familyName false false
            let em = float font.OtFont.Head.UnitsPerEm
            let ptToPx = 96.0 / 72.0
            let unit2px (units: float) = units / em * fontPt * ptToPx
            match font.OtFont.OS2 with
            | Some os2 ->
                let asc  = unit2px (float os2.TypoAscender)
                let desc = unit2px (float -os2.TypoDescender)
                let lg   = unit2px (float os2.TypoLineGap)
                let pixels =
                    System.Math.Ceiling asc
                    + System.Math.Ceiling desc
                    + System.Math.Ceiling lg
                    + 1.0
                pixels * 72.0 / 96.0
            | None ->
                // Fonts without OS/2 fall back to hhea metrics.
                let asc  = unit2px (float font.OtFont.Hhea.Ascender)
                let desc = unit2px (float -font.OtFont.Hhea.Descender)
                let lg   = unit2px (float font.OtFont.Hhea.LineGap)
                let pixels =
                    System.Math.Ceiling asc
                    + System.Math.Ceiling desc
                    + System.Math.Ceiling lg
                    + 1.0
                pixels * 72.0 / 96.0
        with _ ->
            // Last-resort: ~Calibri 11 ratio rounded to 0.75-pt grid.
            System.Math.Round(fontPt * 1.36 * 4.0) / 4.0

    // <col min=M max=N width=W hidden=1/> applies W to every column in M..N
    // (inclusive). Hidden columns collapse to width=0 regardless of the
    // declared width — Excel hides them by zero-width, not by skipping the
    // entry, so a hidden col with width=11 still renders as 0 px. poi-51585
    // has 748 hidden cols in its template; without this they'd push the
    // real data way off-screen.
    let explicitColumns =
        sheet.Columns
        |> List.collect (fun col ->
            let widthPx =
                if col.Hidden then Some 0.0
                else col.Width |> Option.map storedWidthToPixels
            match widthPx with
            | None -> []
            | Some px -> [ for idx in col.Index .. col.MaxIndex -> idx - 1, px ])
        |> Map.ofList

    let explicitRows =
        // Hidden rows collapse to 0 px (matches Excel's screen + print
        // behavior). Otherwise honor an explicit height attribute, snapped
        // DOWN to the nearest pixel — Excel rounds stored `ht` values to
        // the 1-px (0.75-pt at 96 DPI) grid for display, and reports the
        // floored value as `cell.Height` / `cell.RowHeight`. Without the
        // floor pandas-df_empty's ht=19.4 rendered at 19.4 pt while Excel
        // showed 18.75 (= 25 px floor), pulling every subsequent row above
        // Excel's grid.
        sheet.Rows
        |> List.choose (fun row ->
            if row.Hidden then Some (row.Index - 1, 0.0)
            else row.Height |> Option.map (fun h ->
                let px = System.Math.Floor(h * (96.0 / 72.0))
                row.Index - 1, px))
        |> Map.ofList

    // Excel's grid is a fixed 1,048,576 × 16,384 (2^20 × 2^14) — not the
    // used range. The used range (maxRow/maxCol, the <dimension> element)
    // is just metadata; the displayed grid is always the full size, which
    // is why Excel paints gridlines across the whole viewport even past
    // the data. Layout.visibleRange virtualizes to the viewport, so a
    // full-size model costs nothing at paint time.
    let excelRowCount = 1048576
    let excelColCount = 16384

    // Seed row heights / column widths with the sheet default. We only
    // need entries out to roughly the viewport extent — past that no cell
    // is ever visible (visibleRange clamps to the viewport). A fixed,
    // generous margin (covers any 720-1100 px render at typical row/col
    // sizes) keeps the maps tiny while ensuring empty cells past the data
    // still get the sheet default rather than Model's 100 px / 25 px
    // fallback.
    let seedCols = max maxCol 50
    let seedRows = max maxRow 200
    let columnWidths =
        [ 0 .. seedCols ]
        |> List.fold (fun acc c ->
            if Map.containsKey c explicitColumns then acc
            else Map.add c defaultColWidthPx acc) explicitColumns

    // Rows fall into four height-resolution buckets:
    //   1. IMPLICIT  — no <row> XML entry. Use max(defaultRowHeight, fontAuto).
    //      pandas-dimension_large row 2 is the canonical case.
    //   2. EXPLICIT customHeight=true — landed in explicitRows above, the
    //      stored `ht` wins literally (the user manually shrank/grew the row,
    //      so even sub-font-height is respected).
    //   3. EXPLICIT `ht` present, customHeight absent OR false — Excel uses
    //      max(stored ht, fontAuto). The stored `ht` is a hint that gets
    //      raised to fit the font. phpsp-26template row 1 (`ht=27`, TNR 22
    //      → 27) and phpsp-27template row 2 (`ht=17`, Arial Black 11 →
    //      18.75) are the two sides of this rule.
    //   4. EXPLICIT customHeight=true but no ht (rare) — falls through to
    //      max(defaultRowHeight, fontAuto) like bucket 1.
    // Rows with a stored `ht` attribute but `customHeight` not explicitly
    // true. Excel raises these to fit the font (max(ht, fontAuto)) instead
    // of either honoring the stored value literally (the customHeight=true
    // case in `explicitRows`) or ignoring it entirely.
    let rawHeightOnly =
        sheet.Rows
        |> List.choose (fun row ->
            let isCustom = row.CustomHeight = Some true
            if row.Hidden || isCustom then None
            else row.RawHeight |> Option.map (fun h ->
                let px = System.Math.Floor (h * (96.0 / 72.0))
                row.Index - 1, px))
        |> Map.ofList

    let rowHeights =
        // When `<sheetFormatPr customHeight="1"/>` is set, Excel treats
        // `defaultRowHeight` as the explicit row height for every row
        // without its own `ht` attribute — the font-derived auto-height
        // does NOT override it. pandas-ints_spelled_with_decimals does
        // this to lock every row at 15.75 pt even though the cells use
        // Arial 10 (which would otherwise auto-height to 12.75).
        if sheet.Format.CustomDefaultHeight then
            let lockedPx =
                sheet.Format.DefaultRowHeight
                |> Option.map (fun h -> h * (96.0 / 72.0))
                |> Option.defaultValue (autoRowHeightPt defaultFontName defaultFontSize * 96.0 / 72.0)
            [ 0 .. seedRows ]
            |> List.fold (fun acc r ->
                if Map.containsKey r explicitRows then acc
                else Map.add r lockedPx acc) explicitRows
        else
        // Build a per-row max-font map from cells we've collected. Any row
        // present in `explicitRows` (customHeight="true") keeps its stored
        // value; everything else gets Excel's font-derived auto-height,
        // which is the MAX of:
        //   - the workbook-default font's auto-height (acts as a floor —
        //     pandas-test1 has a Calibri 12 workbook default that floors
        //     every row at 15.75 pt even when individual cells are 11 pt)
        //   - the row's own max cell font auto-height (so a Title row with
        //     Calibri 18 pushes the row up to 23.25 pt).
        let maxFontByRow =
            let acc = System.Collections.Generic.Dictionary<int, string option * float>()
            let consider row name pt =
                match acc.TryGetValue row with
                | true, (_, cur) when cur >= pt -> ()
                | _ -> acc.[row] <- (name, pt)
            for KeyValue(addr, cd) in cells do
                let pt = cd.Format.FontSize |> Option.map float |> Option.defaultValue defaultFontSize
                let name = cd.Format.FontName |> Option.orElse defaultFontName
                consider addr.Row name pt
            // Rich-text runs can override the cell's base font size for
            // individual segments. Excel auto-grows the row to fit the
            // tallest run, not just the base font. synthetic-inline-rich's
            // A3 has a <sz val="14"/> run inside an otherwise Calibri 11
            // cell — Excel reports row 3 at 18.75 pt (Calibri 14 auto)
            // while we used to report 15 pt (Calibri 11 auto).
            for KeyValue(addr, runs) in richRunsMap do
                for (_, ovr) in runs do
                    match ovr.FontSize with
                    | Some pt ->
                        let name =
                            ovr.FontName
                            |> Option.orElse (
                                Map.tryFind addr cells
                                |> Option.bind (fun cd -> cd.Format.FontName)
                                |> Option.orElse defaultFontName)
                        consider addr.Row name pt
                    | None -> ()
            acc
        // Track rows that contain a superscript or subscript run. Excel
        // bumps row height to fit the raised/lowered run's ascent above
        // (or descent below) the normal baseline. Empirical bump from
        // synthetic-vertalign: super → +2.25 pt, sub → +1.5 pt for
        // Calibri 14. Scale linearly with the row's max-font pt size.
        let rowVertAlign =
            // 0 = none, 1 = sub only, 2 = super (overrides sub when both).
            // Earlier versions wrote `Some Superscript` / `Some Subscript`
            // which bound the names as fresh variables (FS0049) instead of
            // matching the cases — every run ended up as kind=2, so sub-only
            // rows got the larger super bump.
            let acc = System.Collections.Generic.Dictionary<int, int>()
            for KeyValue(addr, runs) in richRunsMap do
                for (_, ovr) in runs do
                    let kind =
                        match ovr.VertAlign with
                        | Some VertAlignSuperscript -> 2
                        | Some VertAlignSubscript -> 1
                        | _ -> 0
                    if kind > 0 then
                        let cur = match acc.TryGetValue addr.Row with true, v -> v | _ -> 0
                        acc.[addr.Row] <- max cur kind
            acc
        let vertAlignBumpPt (kind: int) (fontPt: float) =
            // Excel bumps the row to fit a super/subscript run's raised /
            // lowered glyph. Empirical bumps from synthetic-vertalign at
            // Calibri 14: super → +2.25 pt (21.0 = 18.75 base + 2.25),
            // sub → +1.5 pt (20.25 = 18.75 base + 1.5). Scale linearly
            // with the row's max-font pt size.
            match kind with
            | 2 -> fontPt * 2.25 / 14.0  // super
            | 1 -> fontPt * 1.5 / 14.0   // sub
            | _ -> 0.0
        // The workbook-default-font auto-height that floors every row
        // (per-row max font lands here when smaller).
        let defAutoPt = autoRowHeightPt defaultFontName defaultFontSize
        let defAutoPx = defAutoPt * 96.0 / 72.0
        // sheetFormatPr defaultRowHeight, floored to integer pixels.
        let defaultRowPx =
            sheet.Format.DefaultRowHeight
            |> Option.map (fun h -> System.Math.Floor (h * (96.0 / 72.0)))
        [ 0 .. seedRows ]
        |> List.fold (fun acc r ->
            if Map.containsKey r explicitRows then acc
            else
                // Take the larger of defaultRowHeight (floored to pixel
                // grid) and the row's font-derived auto-height. Then, if a
                // raw `ht` was stored without customHeight=true (bucket 3
                // above), raise to at least the stored value — Excel
                // honors any explicit `ht` as a floor unless the font is
                // even taller.
                let fontPx =
                    match maxFontByRow.TryGetValue r with
                    | true, (name, pt) ->
                        let basePt = autoRowHeightPt name pt
                        let vaKind = match rowVertAlign.TryGetValue r with true, v -> v | _ -> 0
                        let bump = vertAlignBumpPt vaKind pt
                        max defAutoPx ((basePt + bump) * 96.0 / 72.0)
                    | _ -> defAutoPx
                let baseAuto =
                    match defaultRowPx with
                    | Some d -> max d fontPx
                    | None -> fontPx
                let px =
                    match Map.tryFind r rawHeightOnly with
                    | Some stored -> max baseAuto stored
                    | None -> baseAuto
                Map.add r px acc) explicitRows

    // Auto-grow rows for wrapText cells. Excel auto-sizes the row to
    // accommodate wrapped text whenever a row has no explicit `ht`. We
    // measure the cell's text against its column width (real font metrics
    // via Renderer.Text.measureWidth, with `mdw × len` as the fallback),
    // count wrap lines through Paint.wrapLines (which respects literal
    // `\n` and breaks on word boundaries), and bump rowHeights[r] when
    // the wrapped block exceeds the current height. Skipped when the
    // row carries an explicit height — Excel honors the user-set value
    // even if the text doesn't fit.
    let rowHeights =
        let defaultPxHeight = defaultRowHeightPx |> Option.defaultValue (15.0 * 96.0 / 72.0)
        // Cache one renderer-font per (familyName, bold, italic) tuple.
        // Without caching every wrap-cell would re-resolve the font, which
        // hits the Composit font cache repeatedly.
        let fontCache = System.Collections.Generic.Dictionary<string option * bool * bool, _>()
        let getFont (familyName: string option) (bold: bool) (italic: bool) =
            let key = familyName, bold, italic
            match fontCache.TryGetValue key with
            | true, f -> f
            | _ ->
                let f = Composit.Ui.SpreadsheetGrid.rendererFontFor familyName bold italic
                fontCache.[key] <- f
                f
        let measureCell (familyName: string option) (bold: bool) (italic: bool) (fontPx: float) (text: string) : float =
            try
                Renderer.Text.measureWidth (getFont familyName bold italic) fontPx text
            with _ -> float text.Length * mdw
        let mutable result = rowHeights
        for KeyValue(addr, cd) in cells do
            let text =
                match cd.Value with
                | String s -> s
                | _ -> ""
            // Excel auto-grows the row when EITHER wrapText is set OR the
            // cell text has literal `\n` line breaks (Alt+Enter authored).
            // For the wrap case we count word-wrapped lines; for the
            // hard-break case we just count `\n`+1 lines (no word wrap).
            let hasHardBreak = text.Contains('\n')
            let colWidth =
                Map.tryFind addr.Col columnWidths
                |> Option.defaultValue defaultColWidthPx
            // A hidden column (0 width) never grows the row — Excel doesn't wrap
            // text in a hidden column. Without this, the hidden cell's text
            // wraps to one glyph per line and inflates the row height.
            let needsGrow = (cd.Format.WrapText || hasHardBreak) && text.Length > 0 && colWidth > 0.0
            let alreadyExplicit = Map.containsKey addr.Row explicitRows
            if needsGrow && not alreadyExplicit then
                let fontPt = cd.Format.FontSize |> Option.map float |> Option.defaultValue 11.0
                let fontPx = fontPt * 96.0 / 72.0
                let lineCount =
                    if cd.Format.WrapText then
                        let measure = fun (s: string) -> measureCell cd.Format.FontName cd.Format.Bold cd.Format.Italic fontPx s
                        let lines = RedBat.Spreadsheet.CellText.wrapLines (Some measure) colWidth fontPt text
                        lines.Length
                    else
                        // Hard-break only: split on \n without word-wrap.
                        text.Split('\n').Length
                // Excel's wrapped-line height matches the font's auto row
                // height (OS/2 typo-metric-derived). Calibri 11 → 15 pt =
                // 20 px / line, NOT fontPx*1.2 = 17.6 px which under-counted
                // synthetic-wrap-autoheight A2's 8-line block at 123 px vs
                // Excel's 160 px (= 8 * 20).
                let lineHeightPx = autoRowHeightPt cd.Format.FontName fontPt * 96.0 / 72.0
                let needed = float lineCount * lineHeightPx
                let current =
                    Map.tryFind addr.Row result
                    |> Option.defaultValue defaultPxHeight
                if needed > current then
                    result <- Map.add addr.Row needed result
        result

    // Translate the sheet's Xlsx print metadata into the Model's PageSetup
    // record. Where the source file leaves a field unspecified we fall back
    // to the Model's defaults (US Letter, 0.75 inch margins, …).
    let pageSetup : PageSetup =
        let stripDollars (s: string) = s.Replace("$", "")
        let parseRange (s: string) : CellRange option =
            match (stripDollars s).Split(':') with
            | [| a; b |] ->
                match parseCellRef a, parseCellRef b with
                | Some (r1, c1), Some (r2, c2) ->
                    Some { From = { Row = min r1 r2; Col = min c1 c2 }
                           To = { Row = max r1 r2; Col = max c1 c2 } }
                | _ -> None
            | [| a |] ->
                parseCellRef (stripDollars a)
                |> Option.map (fun (r, c) ->
                    { From = { Row = r; Col = c }; To = { Row = r; Col = c } })
            | _ -> None
        // "$1:$3" → (0, 2). Row-only refs.
        let parseRowRange (s: string) : (int * int) option =
            match (stripDollars s).Split(':') with
            | [| a; b |] ->
                match Int32.TryParse a, Int32.TryParse b with
                | (true, r1), (true, r2) -> Some (min r1 r2 - 1, max r1 r2 - 1)
                | _ -> None
            | _ -> None
        // "$A:$B" → (0, 1). Col-only refs.
        let parseColRange (s: string) : (int * int) option =
            match (stripDollars s).Split(':') with
            | [| a; b |] when a.Length > 0 && b.Length > 0
                              && a |> Seq.forall Char.IsLetter
                              && b |> Seq.forall Char.IsLetter ->
                let c1 = columnIndex a
                let c2 = columnIndex b
                Some (min c1 c2, max c1 c2)
            | _ -> None
        let orientation =
            match sheet.PageSetup with
            | Some ps ->
                match ps.Orientation with
                | Xlsx.OrientationLandscape -> LandscapeOrientation
                | _ -> PortraitOrientation
            | None -> PortraitOrientation
        let pageW, pageH =
            // PaperSize gives a portrait-orientation pair; swap for landscape.
            let baseW, baseH =
                sheet.PageSetup
                |> Option.bind (fun ps -> ps.PaperSize |> Option.map paperSizePoints)
                |> Option.defaultValue (defaultPageSetup.PageWidth, defaultPageSetup.PageHeight)
            if orientation = LandscapeOrientation then baseH, baseW else baseW, baseH
        let margins =
            sheet.PageMargins
            |> Option.map (fun m ->
                { Top = m.Top; Bottom = m.Bottom; Left = m.Left; Right = m.Right
                  Header = m.Header; Footer = m.Footer } : PageMargins)
            |> Option.defaultValue defaultPageMargins
        let printOptions : PrintOptions =
            { GridLines = sheet.PrintOptions.GridLines
              Headings = sheet.PrintOptions.Headings
              HorizontalCentered = sheet.PrintOptions.HorizontalCentered
              VerticalCentered = sheet.PrintOptions.VerticalCentered }
        let hf : HeaderFooter =
            { OddHeader = sheet.HeaderFooter.OddHeader
              OddFooter = sheet.HeaderFooter.OddFooter
              EvenHeader = sheet.HeaderFooter.EvenHeader
              EvenFooter = sheet.HeaderFooter.EvenFooter
              FirstHeader = sheet.HeaderFooter.FirstHeader
              FirstFooter = sheet.HeaderFooter.FirstFooter
              DifferentOddEven = sheet.HeaderFooter.DifferentOddEven
              DifferentFirst = sheet.HeaderFooter.DifferentFirst
              ScaleWithDoc = sheet.HeaderFooter.ScaleWithDoc }
        { defaultPageSetup with
            PageWidth = pageW
            PageHeight = pageH
            Orientation = orientation
            Margins = margins
            FitToPage = sheet.PageSetup |> Option.map (fun ps -> ps.FitToPage) |> Option.defaultValue false
            FitToWidth = sheet.PageSetup |> Option.map (fun ps -> ps.FitToWidth) |> Option.defaultValue 1
            FitToHeight = sheet.PageSetup |> Option.map (fun ps -> ps.FitToHeight) |> Option.defaultValue 1
            Scale = sheet.PageSetup |> Option.map (fun ps -> ps.Scale) |> Option.defaultValue 100
            Options = printOptions
            HeaderFooter = hf
            ManualRowBreaks = sheet.RowBreaks
            ManualColBreaks = sheet.ColBreaks
            PrintArea = sheet.PrintArea |> Option.bind parseRange
            RepeatRows = sheet.PrintTitleRows |> Option.bind parseRowRange
            RepeatCols = sheet.PrintTitleCols |> Option.bind parseColRange
            FirstPageNumber = sheet.PageSetup |> Option.map (fun ps -> ps.FirstPageNumber) |> Option.defaultValue 1
            UseFirstPageNumber = sheet.PageSetup |> Option.map (fun ps -> ps.UseFirstPageNumber) |> Option.defaultValue false
            PageOrder =
                match sheet.PageSetup |> Option.map (fun ps -> ps.PageOrder) with
                | Some Xlsx.OverThenDown -> OverThenDown
                | _ -> DownThenOver
            BlackAndWhite = sheet.PageSetup |> Option.map (fun ps -> ps.BlackAndWhite) |> Option.defaultValue false
            CellComments =
                match sheet.PageSetup |> Option.map (fun ps -> ps.CellComments) with
                | Some Xlsx.CellCommentsAsDisplayed -> CellCommentsAsDisplayed
                | Some Xlsx.CellCommentsAtEnd -> CellCommentsAtEnd
                | _ -> CellCommentsNone }

    // Translate Xlsx.TableDefinition → Model.Table so downstream consumers
    // (renderers, layout harness) can read tables from the model alone
    // without needing the original workbook handle.
    let modelTables : Table list =
        sheet.Tables
        |> List.map (fun t ->
            { StartRow = t.StartRow
              StartCol = t.StartCol
              EndRow = t.EndRow
              EndCol = t.EndCol
              HeaderRowCount = t.HeaderRowCount
              Style =
                t.Style
                |> Option.map (fun s ->
                    ({ Name = s.Name
                       ShowFirstColumn = s.ShowFirstColumn
                       ShowLastColumn = s.ShowLastColumn
                       ShowRowStripes = s.ShowRowStripes
                       ShowColumnStripes = s.ShowColumnStripes } : TableStyle))
              AutoFilter = t.AutoFilter })

    let baseModel = createEmpty excelRowCount excelColCount
    // Excel's printed/screenshot output never shows the active-cell border, so
    // park the selection off-grid for snapshot renders.
    { baseModel with
        Cells = cells
        Selection = SingleCell(-1, -1)
        RowHeights = rowHeights
        ColumnWidths = columnWidths
        MergedCells = mergedCells
        RichRuns = richRunsMap
        FrozenRows = sheet.FrozenPane |> Option.map (fun p -> p.YSplit) |> Option.defaultValue 0
        FrozenCols = sheet.FrozenPane |> Option.map (fun p -> p.XSplit) |> Option.defaultValue 0
        PageSetup = pageSetup
        Tables = modelTables
        CfBoldCells = cfBoldCells
        ShowGridLines = sheet.ShowGridLines }

/// Build a RedBat.Spreadsheet workbook context from the parsed XLSX: every
/// sheet's model plus the workbook's defined names. This is what the
/// formula engine needs to resolve cross-sheet references and named ranges
/// during recalculation. Sheet ids are positional (`sheet1`, `sheet2`, …)
/// so a caller can map an Xlsx sheet index straight to its context entry.
let private buildWorkbookContext (xlsxWb: Xlsx.Workbook) : RedBat.Spreadsheet.Workbook.Workbook =
    let sheets =
        xlsxWb.Sheets
        |> List.mapi (fun i ws ->
            let id = sprintf "sheet%d" (i + 1)
            let s : RedBat.Spreadsheet.Workbook.Sheet =
                { Id = id; Name = ws.Name; Model = toModel ws xlsxWb; Visible = true }
            id, s)
    // A definedName's localSheetId (when present) is a 0-based sheet index
    // — map it to the positional sheet id so the name resolves at the right
    // scope; absent localSheetId means a workbook-global name.
    let names =
        xlsxWb.NamedRanges
        |> List.map (fun nr ->
            let scope =
                nr.Scope
                |> Option.bind (fun s -> match Int32.TryParse s with | true, i -> Some i | _ -> None)
                |> Option.map (fun i -> sprintf "sheet%d" (i + 1))
            scope, nr.Name, nr.RefersTo)
        |> RedBat.Spreadsheet.Workbook.namesOfTriples
    { Sheets = Map.ofList sheets
      SheetOrder = sheets |> List.map fst
      ActiveSheet = (match sheets with (id, _) :: _ -> id | [] -> "sheet1")
      Tables = Map.empty
      Names = names }

let private escapeJson (value: string) =
    let sb = System.Text.StringBuilder(value.Length + 2)
    sb.Append('"') |> ignore
    for ch in value do
        match ch with
        | '"' -> sb.Append("\\\"") |> ignore
        | '\\' -> sb.Append("\\\\") |> ignore
        | '\n' -> sb.Append("\\n") |> ignore
        | '\r' -> sb.Append("\\r") |> ignore
        | '\t' -> sb.Append("\\t") |> ignore
        | c when int c < 0x20 -> sb.AppendFormat("\\u{0:x4}", int c) |> ignore
        | c -> sb.Append(c) |> ignore
    sb.Append('"') |> ignore
    sb.ToString()

let private cellInfoJson (sheetName: string) (model: SpreadsheetModel) =
    let columnLetter col = RedBat.Spreadsheet.Model.columnToLetter col
    let viewport = RedBat.Spreadsheet.Layout.defaultViewport 0.0 0.0
    let records =
        model.Cells
        |> Map.toList
        |> List.map (fun (addr, data) ->
            let rect = RedBat.Spreadsheet.Layout.cellRect model viewport addr.Row addr.Col
            let display =
                match data.Value with
                | Number n -> string n
                | String s -> s
                | Boolean b -> if b then "TRUE" else "FALSE"
                | Formula (_, _) -> ""
                | Empty -> ""
                | Error e -> e
            let fields = [
                "\"Sheet\":" + escapeJson sheetName
                "\"Address\":" + escapeJson (sprintf "%s%d" (columnLetter addr.Col) (addr.Row + 1))
                "\"Row\":" + string (addr.Row + 1)
                "\"Column\":" + string (addr.Col + 1)
                "\"X\":" + sprintf "%.2f" (rect.X - viewport.RowHeaderWidth)
                "\"Y\":" + sprintf "%.2f" (rect.Y - viewport.HeaderHeight)
                "\"Width\":" + sprintf "%.2f" rect.Width
                "\"Height\":" + sprintf "%.2f" rect.Height
                "\"Value\":" + escapeJson display
                "\"Text\":" + escapeJson display
                "\"NumberFormat\":" + escapeJson (data.Format.NumberFormat |> Option.defaultValue "")
            ]
            "{ " + String.concat ", " fields + " }")
    "[\n  " + String.concat ",\n  " records + "\n]\n"

#if EXCEL_NATIVE_RUNNER_MAIN
let private renderSheetToPath
    (inputPath: string) (outputPath: string) (jsonPath: string option)
    (width: int) (height: int)
    (sheetIndex: int) (sheet: Xlsx.Worksheet) (wbContext: RedBat.Spreadsheet.Workbook.Workbook) =
    // Pull this sheet's pre-built model out of the workbook context, then
    // run a formula pass: cells with a cached <v> already arrived as plain
    // values, so this only touches genuinely uncalculated formulas — and
    // now resolves their cross-sheet refs and defined names.
    let sheetId = sprintf "sheet%d" (sheetIndex + 1)
    let baseModel =
        match Map.tryFind sheetId wbContext.Sheets with
        | Some s -> s.Model
        | None -> createEmpty 1048576 16384
    let model = RedBat.Spreadsheet.Formula.recalculateAllFormulas (Some wbContext) (Some sheetId) baseModel

    let widget = Composit.Ui.SpreadsheetGrid.widget (float width) (float height) model
    let backend = Raster.create()
    let framebuffer = Composit.Ui.Render.renderWithBg Color.white backend width height widget
    let png = PngEncode.encode framebuffer

    Directory.CreateDirectory(Path.GetDirectoryName outputPath) |> ignore
    File.WriteAllBytes(outputPath, png)

    match jsonPath with
    | Some path ->
        Directory.CreateDirectory(Path.GetDirectoryName path) |> ignore
        File.WriteAllText(path, cellInfoJson sheet.Name model)
    | None -> ()

    let cellCount = Map.count model.Cells
    printfn "excel native runner ok: %s [sheet=%s cells=%d] -> %s (%dx%d)" inputPath sheet.Name cellCount outputPath width height

/// Insert `suffix` before the file extension so `out/foo.png` → `out/foo-sheet2.png`.
let private withSheetSuffix (path: string) (suffix: string) =
    let dir = Path.GetDirectoryName path
    let stem = Path.GetFileNameWithoutExtension path
    let ext = Path.GetExtension path
    let leaf = stem + "-" + suffix + ext
    if String.IsNullOrEmpty dir then leaf else Path.Combine(dir, leaf)

let private renderXlsx (inputPath: string) (outputPath: string) (width: int) (height: int) (sheetSel: Choice<int, unit>) (jsonPath: string option) =
    initFont ()
    let bytes = File.ReadAllBytes inputPath
    let workbook =
        match Xlsx.loadBytesNative bytes with
        | Some parsed -> parsed
        | None -> failwithf "Could not parse XLSX: %s" inputPath

    if workbook.Sheets.IsEmpty then
        failwithf "No sheets in workbook: %s" inputPath

    // Build the formula-engine context once: every sheet's model + the
    // workbook's defined names. Shared across all sheets so cross-sheet
    // references and named ranges resolve during recalculation.
    let wbContext = buildWorkbookContext workbook

    match sheetSel with
    | Choice1Of2 sheetIndex ->
        let chosenIndex = min (max 0 sheetIndex) (workbook.Sheets.Length - 1)
        let sheet = workbook.Sheets.[chosenIndex]
        renderSheetToPath inputPath outputPath jsonPath width height chosenIndex sheet wbContext
    | Choice2Of2 () ->
        // --all: render every sheet to `<base>-sheet<N>.png` (1-indexed,
        // matching how Excel labels them in the tab strip and how pandas
        // writes sheet names like "Sheet1"). When there's only one sheet,
        // skip the suffix so the output path stays predictable.
        let sheets = workbook.Sheets
        sheets
        |> List.iteri (fun i sheet ->
            let path =
                if sheets.Length = 1 then outputPath
                else withSheetSuffix outputPath (sprintf "sheet%d" (i + 1))
            let json =
                jsonPath |> Option.map (fun p ->
                    if sheets.Length = 1 then p
                    else withSheetSuffix p (sprintf "sheet%d" (i + 1)))
            renderSheetToPath inputPath path json width height i sheet wbContext)

/// Parse the sheet argument: `all` / `-a` selects every sheet; an integer
/// selects that 0-based sheet index. Anything else falls back to sheet 0.
let private parseSheetArg (value: string) : Choice<int, unit> =
    if value = "all" || value = "-a" || value = "--all" then Choice2Of2 ()
    else
        match Int32.TryParse value with
        | true, parsed -> Choice1Of2 parsed
        | false, _ -> Choice1Of2 0

[<EntryPoint>]
let main argv =
    try
        let parseInt fallback (value: string) =
            match Int32.TryParse value with
            | true, parsed -> parsed
            | false, _ -> fallback
        let jsonEnv = Environment.GetEnvironmentVariable("EXCEL_NATIVE_JSON")
        let jsonOpt = if String.IsNullOrWhiteSpace jsonEnv then None else Some (fullPath jsonEnv)
        match argv with
        | [| input; output |] ->
            renderXlsx (fullPath input) (fullPath output) 1100 720 (Choice1Of2 0) jsonOpt
            0
        | [| input; output; width |] ->
            renderXlsx (fullPath input) (fullPath output) (parseInt 1100 width) 720 (Choice1Of2 0) jsonOpt
            0
        | [| input; output; width; height |] ->
            renderXlsx (fullPath input) (fullPath output) (parseInt 1100 width) (parseInt 720 height) (Choice1Of2 0) jsonOpt
            0
        | [| input; output; width; height; sheet |] ->
            renderXlsx (fullPath input) (fullPath output) (parseInt 1100 width) (parseInt 720 height) (parseSheetArg sheet) jsonOpt
            0
        | [| input; output; width; height; sheet; jsonOut |] ->
            renderXlsx (fullPath input) (fullPath output) (parseInt 1100 width) (parseInt 720 height) (parseSheetArg sheet) (Some (fullPath jsonOut))
            0
        | _ ->
            usage()
            2
    with ex ->
        eprintfn "%O" ex
        1
#endif
