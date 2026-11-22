module Rbt.Spreadsheet.ConditionalFormatting

open System
open System.Collections.Generic
open System.Globalization
open Rbt.Spreadsheet.Model

// ---------------------------------------------------------------------------
// Conditional formatting: evaluation of a sheet's rules for a cell, and the
// OOXML shape of the rules (for import and export).
//
// Rules are stored by range (`SpreadsheetModel.ConditionalFormats`). A cell's
// formatting is the merge of every rule that covers it and matches, in
// priority order: for each property the first rule that sets it wins, and a
// matching rule marked stop-if-true ends the walk. Rules that depend on the
// rest of the range (top/bottom N, averages, duplicates, colour scales, data
// bars, icon sets) compute their statistics once per sheet state.
// ---------------------------------------------------------------------------

/// What the conditional formats of a sheet do to one cell. Every property is
/// "leave as it is" when None.
type CellEffect = {
    TextColor: string option
    BgColor: string option
    Bold: bool option
    Italic: bool option
    Underline: bool option
    Strikethrough: bool option
    /// See `CellFormat.DataBar`.
    DataBar: DataBarDraw option
    /// See `CellFormat.IconSet`.
    IconSet: IconDraw option
    /// The lines a rule gives the cell's sides; a line of style "none"
    /// takes the cell's own line away.
    BorderTop: BorderStyle option
    BorderBottom: BorderStyle option
    BorderLeft: BorderStyle option
    BorderRight: BorderStyle option
}

let noEffect : CellEffect =
    { TextColor = None; BgColor = None; Bold = None; Italic = None; Underline = None; Strikethrough = None
      DataBar = None; IconSet = None; BorderTop = None; BorderBottom = None; BorderLeft = None; BorderRight = None }

/// A side's line under a rule's line for it ("none": no line).
let private sideWith (rule: BorderStyle option) (own: BorderStyle option) =
    match rule with
    | Some line when line.Style = "none" -> None
    | Some _ -> rule
    | None -> own

/// Whether a rule's style (a format) has a line for any side: Excel's
/// `<dxf><border>`.
let styleHasBorders (style: CellFormat) =
    style.BorderTop.IsSome || style.BorderBottom.IsSome || style.BorderLeft.IsSome || style.BorderRight.IsSome

/// A cell's format with only the lines of an effect laid over it.
let applyBorderEffect (effect: CellEffect) (format: CellFormat) : CellFormat =
    if effect.BorderTop.IsNone && effect.BorderBottom.IsNone && effect.BorderLeft.IsNone && effect.BorderRight.IsNone then format
    else
        { format with
            BorderTop = sideWith effect.BorderTop format.BorderTop
            BorderBottom = sideWith effect.BorderBottom format.BorderBottom
            BorderLeft = sideWith effect.BorderLeft format.BorderLeft
            BorderRight = sideWith effect.BorderRight format.BorderRight }

/// Lay an effect over a cell's own format.
let applyEffect (effect: CellEffect) (format: CellFormat) : CellFormat =
    { format with
        TextColor = (match effect.TextColor with Some _ -> effect.TextColor | None -> format.TextColor)
        BgColor = (match effect.BgColor with Some _ -> effect.BgColor | None -> format.BgColor)
        Bold = defaultArg effect.Bold format.Bold
        Italic = defaultArg effect.Italic format.Italic
        Underline = defaultArg effect.Underline format.Underline
        Strikethrough = defaultArg effect.Strikethrough format.Strikethrough
        DataBar = (match effect.DataBar with Some _ -> effect.DataBar | None -> format.DataBar)
        IconSet = (match effect.IconSet with Some _ -> effect.IconSet | None -> format.IconSet)
        BorderTop = sideWith effect.BorderTop format.BorderTop
        BorderBottom = sideWith effect.BorderBottom format.BorderBottom
        BorderLeft = sideWith effect.BorderLeft format.BorderLeft
        BorderRight = sideWith effect.BorderRight format.BorderRight }

/// Evaluates a formula written for one cell at another: (formula, rows moved,
/// columns moved, sheet, cell). `Formula.evaluateFormulaShifted workbook sheet`
/// for a sheet of a workbook, `Model.evaluateFormulaAt` for a stand-alone one.
type FormulaEvaluator = string -> int -> int -> SpreadsheetModel -> CellAddress -> CellValue

let private parseNumber (text: string) : float option =
    match Double.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture) with
    | true, number -> Some number
    | _ -> None

let private hexToRgb (color: string) =
    let hex = if color.StartsWith "#" then color.Substring 1 else color
    let hex = if hex.Length = 8 then hex.Substring 2 else hex
    if hex.Length = 6 then
        try Convert.ToInt32(hex.Substring(0, 2), 16), Convert.ToInt32(hex.Substring(2, 2), 16), Convert.ToInt32(hex.Substring(4, 2), 16)
        with _ -> 255, 255, 255
    else 255, 255, 255

let private hexByte (value: int) =
    let digits = "0123456789ABCDEF"
    let v = max 0 (min 255 value)
    string digits.[v / 16] + string digits.[v % 16]

/// The colour `t` of the way from `a` to `b` (0..1).
let blendColors (a: string) (b: string) (t: float) : string =
    let (ar, ag, ab), (br, bg, bb) = hexToRgb a, hexToRgb b
    let mix x y = int (Math.Floor (float x + (float y - float x) * t + 0.5))
    "#" + hexByte (mix ar br) + hexByte (mix ag bg) + hexByte (mix ab bb)

/// What two cells must share to count as the same value (duplicates).
let private valueKey (value: CellValue) : string option =
    match value with
    | Number n -> Some ("n:" + n.ToString("R", CultureInfo.InvariantCulture))
    | String s when s <> "" -> Some ("s:" + s.ToLowerInvariant())
    | Boolean b -> Some (if b then "b:1" else "b:0")
    | _ -> None

/// One rule with what was worked out about its range.
type private PreparedRule = {
    Entry: ConditionalFormatRange
    /// Top-left of the rule's first range: where its formulas are anchored.
    Anchor: CellAddress
    /// The numbers in the rule's ranges, ascending.
    mutable Numbers: float[] option
    /// How often each value occurs in the rule's ranges.
    mutable Counts: Dictionary<string, int> option
    /// Resolved colour-scale / data-bar / icon-set thresholds.
    mutable Thresholds: float option[] option
}

/// The conditional formats of one sheet state, ready to be asked about cells.
type Context = private {
    Model: SpreadsheetModel
    Evaluate: FormulaEvaluator
    Rules: PreparedRule list
}

let private anchorOf (entry: ConditionalFormatRange) : CellAddress =
    match entry.Ranges with
    | first :: _ -> first.From
    | [] -> { Row = 0; Col = 0 }

/// Prepare a sheet's conditional formats for evaluation with `evaluate`.
let createContext (evaluate: FormulaEvaluator) (model: SpreadsheetModel) : Context =
    { Model = model
      Evaluate = evaluate
      Rules =
        model.ConditionalFormats
        |> List.sortBy (fun entry -> entry.Priority)
        |> List.map (fun entry -> { Entry = entry; Anchor = anchorOf entry; Numbers = None; Counts = None; Thresholds = None }) }

let private covers (rule: PreparedRule) (addr: CellAddress) =
    rule.Entry.Ranges |> List.exists (cellRangeContains addr)

/// Walk the stored cells of a rule's ranges. The ranges may be whole columns;
/// only cells that exist are visited.
let private iterRuleCells (ctx: Context) (rule: PreparedRule) (visit: CellValue -> unit) =
    for KeyValue(addr, cellData) in ctx.Model.Cells do
        if covers rule addr then visit (unwrapFormulaValue cellData.Value)

let private numbersOf (ctx: Context) (rule: PreparedRule) : float[] =
    match rule.Numbers with
    | Some numbers -> numbers
    | None ->
        let numbers = ResizeArray<float>()
        iterRuleCells ctx rule (fun value -> match value with Number n -> numbers.Add n | _ -> ())
        let sorted = numbers.ToArray()
        Array.sortInPlace sorted
        rule.Numbers <- Some sorted
        sorted

let private countsOf (ctx: Context) (rule: PreparedRule) : Dictionary<string, int> =
    match rule.Counts with
    | Some counts -> counts
    | None ->
        let counts = Dictionary<string, int>()
        iterRuleCells ctx rule (fun value ->
            match valueKey value with
            | Some key ->
                match Portable.Dict.tryFind key counts with
                | ValueSome n -> counts.[key] <- n + 1
                | ValueNone -> counts.[key] <- 1
            | None -> ())
        rule.Counts <- Some counts
        counts

/// The value at a percentile (0..100) of ascending numbers, interpolated the
/// way PERCENTILE does.
let private percentile (sorted: float[]) (p: float) : float option =
    if sorted.Length = 0 then None
    else
        let position = (max 0.0 (min 100.0 p)) / 100.0 * float (sorted.Length - 1)
        let below = int (Math.Floor position)
        let above = min (sorted.Length - 1) (below + 1)
        Some (sorted.[below] + (sorted.[above] - sorted.[below]) * (position - float below))

let private resolveThreshold (ctx: Context) (rule: PreparedRule) (threshold: CfThreshold) : float option =
    let numbers = numbersOf ctx rule
    let lowest = if numbers.Length > 0 then Some numbers.[0] else None
    let highest = if numbers.Length > 0 then Some numbers.[numbers.Length - 1] else None
    let literal = threshold.Value |> Option.bind parseNumber
    match threshold.Kind with
    | CfvoMin -> lowest
    | CfvoMax -> highest
    | CfvoAutoMin -> lowest |> Option.map (fun value -> min 0.0 value)
    | CfvoAutoMax -> highest |> Option.map (fun value -> max 0.0 value)
    | CfvoNumber -> literal
    | CfvoPercent ->
        (match lowest, highest, literal with
         | Some lo, Some hi, Some p -> Some (lo + (hi - lo) * p / 100.0)
         | _ -> None)
    | CfvoPercentile -> literal |> Option.bind (percentile numbers)
    | CfvoFormula ->
        match literal, threshold.Value with
        | Some number, _ -> Some number
        | None, Some formula ->
            (match ctx.Evaluate formula 0 0 ctx.Model rule.Anchor |> unwrapFormulaValue with
             | Number n -> Some n
             | _ -> None)
        | None, None -> None

let private thresholdsOf (ctx: Context) (rule: PreparedRule) (thresholds: CfThreshold list) : float option[] =
    match rule.Thresholds with
    | Some resolved -> resolved
    | None ->
        let resolved = thresholds |> List.map (resolveThreshold ctx rule) |> List.toArray
        rule.Thresholds <- Some resolved
        resolved

/// Number of icons in an icon-set family ("3Arrows" → 3, "5Quarters" → 5).
let iconCountOf (family: string) : int =
    if family.Length > 0 && Char.IsDigit family.[0] then int family.[0] - int '0' else 3

let private isTruthy (value: CellValue) =
    match unwrapFormulaValue value with
    | Boolean flag -> flag
    | Number n -> n <> 0.0
    | _ -> false

/// Does `x <op> first [second]` hold, numbers compared as Excel's "cell value" rules compare them:
/// as its `=` operator (`Formula.compareNumbers`) - equal when they agree to 15 significant digits
/// (0.1+0.2 = 0.3, between 0.3 and 0.5), otherwise their order, with no tolerance near 0 (1E-10 > 0,
/// 1E-10 <> 0). Checked in Excel 2019 (synthetic-cf-compare.xlsx).
let private numberHolds (op: CompareOperator) (x: float) (first: float) (second: float option) : bool =
    let c = Formula.compareNumbers x first
    match op with
    | OpGreaterThan -> c > 0
    | OpLessThan -> c < 0
    | OpGreaterOrEqual -> c >= 0
    | OpLessOrEqual -> c <= 0
    | OpEqual -> c = 0
    | OpNotEqual -> c <> 0
    | OpBetween | OpNotBetween ->
        let other = defaultArg second first
        let lo, hi = (if first <= other then first, other else other, first)
        let inside = Formula.compareNumbers x lo >= 0 && Formula.compareNumbers x hi <= 0
        if op = OpBetween then inside else not inside

/// Does a cell's value satisfy `value <op> operand(s)`? Numbers compare with
/// numbers (`numberHolds`) and text with text (ignoring case); a blank cell
/// counts as 0, as Excel's "cell value" rules treat it.
let private cellIsHolds (op: CompareOperator) (value: CellValue) (first: CellValue) (second: CellValue option) : bool =
    let asNumber (v: CellValue) =
        match v with
        | Number n -> Some n
        | Boolean flag -> Some (if flag then 1.0 else 0.0)
        | Empty -> Some 0.0
        | _ -> None
    match value, first with
    | String text, String operand ->
        let order (other: string) = String.Compare(text, other, StringComparison.OrdinalIgnoreCase)
        (match op with
         | OpEqual -> order operand = 0
         | OpNotEqual -> order operand <> 0
         | OpGreaterThan -> order operand > 0
         | OpLessThan -> order operand < 0
         | OpGreaterOrEqual -> order operand >= 0
         | OpLessOrEqual -> order operand <= 0
         | OpBetween | OpNotBetween ->
            match second with
            | Some (String other) ->
                let lo, hi = (if String.Compare(operand, other, StringComparison.OrdinalIgnoreCase) <= 0 then operand, other else other, operand)
                let inside = order lo >= 0 && order hi <= 0
                if op = OpBetween then inside else not inside
            | _ -> false)
    | Error _, _ | _, Error _ -> false
    | _ ->
        match asNumber value, asNumber first with
        | Some x, Some a -> numberHolds op x a (second |> Option.bind asNumber)
        // Different kinds of value are never equal, and have no order.
        | _ -> op = OpNotEqual

/// The periods of a "date occurring" rule (`TimePeriod`).
let timePeriods =
    [ "today"; "yesterday"; "tomorrow"; "last7Days"; "thisWeek"; "lastWeek"; "nextWeek"; "thisMonth"; "lastMonth"; "nextMonth" ]

/// The formula that is true when the cell `cell` ("A1") holds a date in a
/// period: what Excel writes into the file beside the period, and what
/// decides the rule here. None for a period that is not one.
let timePeriodFormula (period: string) (cell: string) : string option =
    let day = sprintf "FLOOR(%s,1)" cell
    let month (offset: int) =
        let today = if offset = 0 then "TODAY()" else sprintf "EDATE(TODAY(),%d)" offset
        sprintf "AND(ISNUMBER(%s),MONTH(%s)=MONTH(%s),YEAR(%s)=YEAR(%s))" cell cell today cell today
    match period with
    | "today" -> Some (sprintf "%s=TODAY()" day)
    | "yesterday" -> Some (sprintf "%s=TODAY()-1" day)
    | "tomorrow" -> Some (sprintf "%s=TODAY()+1" day)
    | "last7Days" -> Some (sprintf "AND(TODAY()-%s<=6,%s<=TODAY())" day day)
    | "thisWeek" -> Some (sprintf "AND(TODAY()-%s<=WEEKDAY(TODAY())-1,%s-TODAY()<=7-WEEKDAY(TODAY()))" day day)
    | "lastWeek" -> Some (sprintf "AND(TODAY()-%s>=WEEKDAY(TODAY()),TODAY()-%s<WEEKDAY(TODAY())+7)" day day)
    | "nextWeek" -> Some (sprintf "AND(%s-TODAY()>7-WEEKDAY(TODAY()),%s-TODAY()<15-WEEKDAY(TODAY()))" day day)
    | "thisMonth" -> Some (month 0)
    | "lastMonth" -> Some (month -1)
    | "nextMonth" -> Some (month 1)
    | _ -> None

/// Does this rule match the cell? Visual rules (scale, bar, icons) "match"
/// every number in their range; what they paint is computed separately.
let private ruleMatches (ctx: Context) (rule: PreparedRule) (addr: CellAddress) (raw: CellValue) : bool =
    let value = unwrapFormulaValue raw
    let formulaAt (formula: string) =
        ctx.Evaluate formula (addr.Row - rule.Anchor.Row) (addr.Col - rule.Anchor.Col) ctx.Model addr |> unwrapFormulaValue
    match rule.Entry.Format.Rule with
    | ContainsText _ -> evaluateRule rule.Entry.Format.Rule raw
    | GreaterThan threshold -> (match value with Number n -> numberHolds OpGreaterThan n threshold None | _ -> false)
    | LessThan threshold -> (match value with Number n -> numberHolds OpLessThan n threshold None | _ -> false)
    | Between (lo, hi) -> (match value with Number n -> numberHolds OpBetween n lo (Some hi) | _ -> false)
    | Equal target -> (match value with Number n -> numberHolds OpEqual n target None | _ -> false)
    | CellIs (op, first, second) -> cellIsHolds op value (formulaAt first) (second |> Option.map formulaAt)
    | Expression formula -> isTruthy (formulaAt formula)
    | TimePeriod period ->
        // (Dates are numbers; the period's formula is written for the rule's first cell.)
        (match value, timePeriodFormula period (formatCellRef rule.Anchor.Row rule.Anchor.Col) with
         | Number _, Some formula -> isTruthy (formulaAt formula)
         | _ -> false)
    | TextRule (kind, text) ->
        let shown = displayValue value
        let found = text <> "" && shown.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0
        (match kind with
         | TextContains -> found
         | TextNotContains -> not found
         | TextBeginsWith -> text <> "" && shown.StartsWith(text, StringComparison.OrdinalIgnoreCase)
         | TextEndsWith -> text <> "" && shown.EndsWith(text, StringComparison.OrdinalIgnoreCase))
    | TopBottom (rank, percent, bottom) ->
        (match value with
         | Number n ->
            let numbers = numbersOf ctx rule
            if numbers.Length = 0 || rank <= 0 then false
            else
                let wanted = if percent then max 1 (numbers.Length * rank / 100) else rank
                let wanted = min numbers.Length wanted
                if bottom then n <= numbers.[wanted - 1] else n >= numbers.[numbers.Length - wanted]
         | _ -> false)
    | AboveAverage (above, orEqual, stdDev) ->
        (match value with
         | Number n ->
            let numbers = numbersOf ctx rule
            if numbers.Length = 0 then false
            else
                let mean = Array.average numbers
                let deviation =
                    if stdDev <= 0 then 0.0
                    else float stdDev * sqrt (numbers |> Array.averageBy (fun x -> (x - mean) * (x - mean)))
                if above then (if orEqual then n >= mean + deviation else n > mean + deviation)
                else (if orEqual then n <= mean - deviation else n < mean - deviation)
         | _ -> false)
    | DuplicateValues | UniqueValues ->
        (match valueKey value with
         | Some key ->
            let seen = match (countsOf ctx rule).TryGetValue key with | true, n -> n | false, _ -> 1
            if rule.Entry.Format.Rule = DuplicateValues then seen > 1 else seen = 1
         | None -> false)
    | Blanks wanted ->
        let blank = match value with Empty -> true | String s -> s.Trim() = "" | _ -> false
        blank = wanted
    | Errors wanted ->
        let isError = match value with Error _ -> true | _ -> false
        isError = wanted
    | ColorScale _ | DataBar _ | IconSet _ | DataBarEx _ | IconSetEx _ -> (match value with Number _ -> true | _ -> false)

/// How a data bar of Excel 2010+ lies in the cell of the value `n`, the
/// rule's ends being `lo` and `hi`.
let dataBarDrawOf (options: DataBarOptions) (color: string) (lo: float) (hi: float) (n: float) : DataBarDraw option =
    let css (text: string) = if text.StartsWith "#" then text else "#" + text
    let negative = n < 0.0
    let fill = if negative then (match options.NegativeFill with Some other -> css other | None -> css color) else css color
    let border =
        match options.Border with
        | None -> None
        | Some line -> if negative then Some (css (defaultArg options.NegativeBorder line)) else Some (css line)
    let clamp (value: float) = max 0.0 (min 1.0 value)
    let axisColor = css (defaultArg options.AxisColor "#000000")
    // The axis: where zero is, when the rule's values are on both sides of it.
    let axis =
        match options.Axis with
        | AxisNone -> None
        | AxisMiddle -> Some 0.5
        | AxisAutomatic ->
            if lo >= 0.0 then None
            elif hi <= 0.0 then Some 1.0
            else Some (-lo / (hi - lo))
    let span =
        match axis with
        | None ->
            // From the cell's edge, between the shortest and the longest bar.
            let fraction = if hi - lo < 1e-12 then 1.0 else clamp ((n - lo) / (hi - lo))
            let shortest = float (max 0 (min 100 options.MinLength)) / 100.0
            let longest = float (max 0 (min 100 options.MaxLength)) / 100.0
            let length = shortest + fraction * (max shortest longest - shortest)
            if length <= 0.0 then None else Some (0.0, length, true)
        | Some at ->
            let reach = match options.Axis with AxisMiddle -> max (abs lo) (abs hi) | _ -> (if negative then abs lo else abs hi)
            let room = if negative then at else 1.0 - at
            let length = if reach < 1e-12 then 0.0 else clamp (abs n / reach) * room
            if length <= 0.0 then None
            elif negative then Some (at - length, at, false)
            else Some (at, at + length, true)
    let mirrored = options.Direction = BarRightToLeft
    let flip (value: float) = if mirrored then 1.0 - value else value
    match span, axis with
    | None, None -> None
    | _ ->
        let from, upTo, fadeRight = defaultArg span (0.0, 0.0, true)
        let from, upTo, fadeRight = if mirrored then 1.0 - upTo, 1.0 - from, not fadeRight else from, upTo, fadeRight
        Some { From = from; To = upTo; Color = fill; Gradient = options.Gradient; FadeRight = fadeRight; Border = border
               Axis = axis |> Option.map (fun at -> flip at, axisColor); HideValue = not options.ShowValue }

/// The icon sets of Excel 2010+, which only the x14 form of a rule can name.
let extendedIconSets = [ "3Stars"; "3Triangles"; "5Boxes" ]

/// How an icon is drawn where the icon itself cannot be: a character and
/// its colour. `index` 0 is the icon of the lowest values.
let iconGlyph (family: string) (index: int) : string * string =
    let red, amber, green, grey, black = "#E53935", "#F9A825", "#43A047", "#8E8E8E", "#404040"
    let count = if family.Length > 0 && Char.IsDigit family.[0] then int family.[0] - int '0' else 3
    let index = max 0 (min (count - 1) index)
    // Position of the icon between the lowest (0) and the highest (1).
    let level = if count <= 1 then 1.0 else float index / float (count - 1)
    let trafficColor = if level < 0.34 then red elif level < 0.67 then amber else green
    let name = if family.Length > 1 then family.Substring 1 else family
    match name with
    | "Arrows" | "ArrowsGray" ->
        let glyph =
            if level <= 0.0 then "\u2193" elif level >= 1.0 then "\u2191"
            elif level < 0.5 then "\u2198" elif level > 0.5 then "\u2197" else "\u2192"
        glyph, (if name = "ArrowsGray" then grey else trafficColor)
    | "Triangles" -> (if level <= 0.0 then "\u25BC" elif level >= 1.0 then "\u25B2" else "\u25AC"), trafficColor
    | "Flags" -> "\u2691", trafficColor
    | "Signs" -> (if level <= 0.0 then "\u25C6" elif level >= 1.0 then "\u25CF" else "\u25B2"), trafficColor
    | "Symbols" | "Symbols2" -> (if level <= 0.0 then "\u2718" elif level >= 1.0 then "\u2714" else "!"), trafficColor
    | "Stars" -> (if level <= 0.0 then "\u2606" else "\u2605"), (if level >= 1.0 then amber elif level > 0.0 then "#FBD46B" else grey)
    | "Rating" | "Boxes" | "Quarters" ->
        // More of the mark is filled the higher the value.
        let glyph = if level <= 0.0 then "\u25CB" elif level < 0.34 then "\u25D4" elif level < 0.67 then "\u25D1" elif level < 1.0 then "\u25D5" else "\u25CF"
        glyph, (if name = "Quarters" then black else "#4472C4")
    | "RedToBlack" -> "\u25CF", (if level <= 0.0 then black elif level < 0.5 then grey elif level < 1.0 then "#F28B82" else red)
    | _ -> "\u25CF", trafficColor

/// An icon of an icon set as Excel 2019 draws it, as SVG of 16 x 16 px (its own picture is a bitmap of that size at
/// 100 %): shapes and colours taken from Excel's picture of synthetic-icon-sets.xlsx (tools/excel-parity/
/// build-icon-sets.vbs) - red #D65532, amber #EAC282, green #68A490, grey #808080, dark #505050, blue #4A7DB1, each
/// with a darker edge. `index` 0 is the icon of the lowest values.
let iconSvg (family: string) (index: int) : string =
    let count = if family.Length > 0 && Char.IsDigit family.[0] then int family.[0] - int '0' else 3
    let index = max 0 (min (count - 1) index)
    let name = if family.Length > 1 then family.Substring 1 else family
    let red, amber, green, grey, dark, blue, light = ("#D65532", "#B4482B"), ("#EAC282", "#A4802B"), ("#68A490", "#44806C"), ("#808080", "#646262"), ("#505050", "#3A3A3A"), ("#4A7DB1", "#335F8C"), ("#B3B3B3", "#727272")
    let shape (fill: string, edge: string) (body: string) = body.Replace("FILL", fill).Replace("EDGE", edge)
    let circle color = shape color "<circle cx='8' cy='8' r='7.5' fill='FILL' stroke='EDGE' stroke-width='0.75'/>"
    // An arrow pointing down, turned by `angle` degrees (clockwise from down: 90 left ... ; -90 right, 180 up).
    let arrow color (angle: int) =
        shape color ("<path d='M4.8 0.5 H11.2 V6 H15 L8 15.5 L1 6 H4.8 Z' fill='FILL' stroke='EDGE' stroke-width='0.75' stroke-linejoin='round' transform='rotate(" + string angle + " 8 8)'/>")
    let traffic3 = [| red; amber; green |]
    let body =
        match name with
        | "Arrows" when count = 3 -> arrow traffic3.[index] [| 0; -90; 180 |].[index]
        | "ArrowsGray" when count = 3 -> arrow grey [| 0; -90; 180 |].[index]
        | "Arrows" when count = 4 -> arrow [| red; amber; amber; green |].[index] [| 0; -45; -135; 180 |].[index]
        | "ArrowsGray" when count = 4 -> arrow grey [| 0; -45; -135; 180 |].[index]
        | "Arrows" -> arrow [| red; amber; amber; amber; green |].[index] [| 0; -45; -90; -135; 180 |].[index]
        | "ArrowsGray" -> arrow grey [| 0; -45; -90; -135; 180 |].[index]
        | "Flags" ->
            shape traffic3.[index] "<rect x='1' y='0.5' width='1.5' height='15' fill='#7F7F7F'/><path d='M2.5 1 L14.5 5 L2.5 9 Z' fill='FILL' stroke='EDGE' stroke-width='0.75' stroke-linejoin='round'/>"
        | "TrafficLights1" -> circle traffic3.[index]
        | "TrafficLights2" ->
            "<rect x='0.5' y='0.5' width='15' height='15' rx='3' fill='#505050' stroke='#3A3A3A'/>" + shape traffic3.[index] "<circle cx='8' cy='8' r='4.5' fill='FILL' stroke='EDGE' stroke-width='0.75'/>"
        | "Signs" ->
            match index with
            | 0 -> shape red "<path d='M8 0.5 L15.5 8 L8 15.5 L0.5 8 Z' fill='FILL' stroke='EDGE' stroke-width='0.75' stroke-linejoin='round'/>"
            | 1 -> shape amber "<path d='M8 1 L15.5 15 H0.5 Z' fill='FILL' stroke='EDGE' stroke-width='0.75' stroke-linejoin='round'/>"
            | _ -> circle green
        | "Symbols" ->
            let mark =
                match index with
                | 0 -> "<path d='M5 5 L11 11 M11 5 L5 11' stroke='#FFFFFF' stroke-width='2' stroke-linecap='round'/>"
                | 1 -> "<path d='M8 3.5 V9' stroke='#FFFFFF' stroke-width='2' stroke-linecap='round'/><circle cx='8' cy='12' r='1.2' fill='#FFFFFF'/>"
                | _ -> "<path d='M4.5 8.5 L7 11 L11.5 5.5' fill='none' stroke='#FFFFFF' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'/>"
            circle traffic3.[index] + mark
        | "Symbols2" ->
            match index with
            | 0 -> shape red "<path d='M2.5 2.5 L13.5 13.5 M13.5 2.5 L2.5 13.5' stroke='FILL' stroke-width='4.2'/>"
            | 1 -> shape amber "<path d='M8 0.5 L8 10' stroke='FILL' stroke-width='4'/><rect x='6' y='12' width='4' height='3.5' fill='FILL'/>"
            | _ -> shape green "<path d='M1.5 8 L6 12.5 L14.5 3' fill='none' stroke='FILL' stroke-width='4' stroke-linejoin='miter'/>"
        | "RedToBlack" -> circle [| dark; ("#B1B1B1", "#8C8C8C"); ("#EDB2A3", "#C98A7A"); ("#D86344", "#B14629") |].[index]
        | "TrafficLights" -> circle [| dark; red; amber; green |].[index]
        | "Rating" ->
            // Four bars rising to the right, the first `index` (four: `index + 1`) blue.
            let filled = if count = 4 then index + 1 else index
            [ for k in 0 .. 3 ->
                let x = 1.5 + float k * 3.5
                let h = 4.0 + float k * 3.7
                let fill, edge = if k < filled then blue else light
                "<rect x='" + string x + "' y='" + string (15.5 - h) + "' width='2.5' height='" + string h + "' fill='" + fill + "' stroke='" + edge + "' stroke-width='0.75'/>" ]
            |> String.concat ""
        | "Quarters" ->
            let outline = "<circle cx='8' cy='8' r='7.5' fill='#FFFFFF' stroke='#505050' stroke-width='1'/>"
            match index with
            | 0 -> outline
            | 4 -> circle dark
            | k ->
                let ends = [| ""; "M8 8 V0.5 A7.5 7.5 0 0 1 15.5 8 Z"; "M8 8 V0.5 A7.5 7.5 0 0 1 8 15.5 Z"; "M8 8 V0.5 A7.5 7.5 0 1 1 0.5 8 Z" |]
                outline + "<path d='" + ends.[k] + "' fill='#505050'/>"
        | "Stars" ->
            let star = "M8 0.8 L10.1 5.6 L15.3 6.1 L11.4 9.6 L12.5 14.8 L8 12.1 L3.5 14.8 L4.6 9.6 L0.7 6.1 L5.9 5.6 Z"
            match index with
            | 0 -> "<path d='" + star + "' fill='#FFFFFF' stroke='#727272' stroke-width='0.75' stroke-linejoin='round'/>"
            | 1 -> "<defs><clipPath id='h'><rect x='0' y='0' width='8' height='16'/></clipPath></defs><path d='" + star + "' fill='#FFFFFF' stroke='#A4802B' stroke-width='0.75' stroke-linejoin='round'/><path d='" + star + "' fill='#EAC282' clip-path='url(#h)'/>"
            | _ -> "<path d='" + star + "' fill='#EAC282' stroke='#A4802B' stroke-width='0.75' stroke-linejoin='round'/>"
        | "Triangles" ->
            match index with
            | 0 -> shape red "<path d='M0.5 3.5 H15.5 L8 12.5 Z' fill='FILL' stroke='EDGE' stroke-width='0.75' stroke-linejoin='round'/>"
            | 1 -> shape amber "<rect x='0.5' y='5.5' width='15' height='5' fill='FILL' stroke='EDGE' stroke-width='0.75'/>"
            | _ -> shape green "<path d='M0.5 12.5 H15.5 L8 3.5 Z' fill='FILL' stroke='EDGE' stroke-width='0.75' stroke-linejoin='round'/>"
        | "Boxes" ->
            // Four squares, filled blue one more at each step: bottom left, bottom right, top left, top right.
            let order = [| (0.5, 8.5); (8.5, 8.5); (0.5, 0.5); (8.5, 0.5) |]
            [ for k in 0 .. 3 ->
                let x, y = order.[k]
                let fill, edge = if k < index then blue else light
                "<rect x='" + string x + "' y='" + string y + "' width='7' height='7' rx='1' fill='" + fill + "' stroke='" + edge + "' stroke-width='0.75'/>" ]
            |> String.concat ""
        | _ -> circle traffic3.[min 2 index]
    "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 16 16'>" + body + "</svg>"

/// What a matching rule contributes to the cell.
let private ruleEffect (ctx: Context) (rule: PreparedRule) (raw: CellValue) : CellEffect =
    let value = unwrapFormulaValue raw
    match rule.Entry.Format.Rule, value with
    | ColorScale stops, Number n ->
        let resolved = thresholdsOf ctx rule (stops |> List.map fst)
        let points =
            [| for i in 0 .. resolved.Length - 1 do
                match resolved.[i] with
                | Some at -> yield at, snd stops.[i]
                | None -> () |]
        if points.Length < 2 then noEffect
        else
            let mutable color = snd points.[points.Length - 1]
            if n <= fst points.[0] then color <- snd points.[0]
            else
                let mutable placed = false
                for i in 1 .. points.Length - 1 do
                    if not placed && n <= fst points.[i] then
                        let lo, loColor = points.[i - 1]
                        let hi, hiColor = points.[i]
                        color <- (if hi - lo < 1e-12 then hiColor else blendColors loColor hiColor ((n - lo) / (hi - lo)))
                        placed <- true
            { noEffect with BgColor = Some color }
    | DataBar (low, high, color), Number n ->
        let resolved = thresholdsOf ctx rule [ low; high ]
        (match resolved.[0], resolved.[1] with
         | Some lo, Some hi ->
            // Excel's automatic bar for all-positive data starts at 0, so
            // the smallest value still shows a bar.
            let lo = if low.Kind = CfvoMin && lo >= 0.0 then 0.0 else lo
            let fraction = if hi - lo < 1e-12 then 1.0 else (n - lo) / (hi - lo)
            // (The bar of Excel 2007, a rule without the x14 part: its colour fading to white, as
            // Excel 2019 draws it.)
            { noEffect with DataBar = Some { plainDataBar (max 0.0 (min 1.0 fraction)) (if color.StartsWith "#" then color else "#" + color) with Gradient = true } }
         | _ -> noEffect)
    | DataBarEx (low, high, color, options), Number n ->
        let resolved = thresholdsOf ctx rule [ low; high ]
        (match resolved.[0], resolved.[1] with
         | Some lo, Some hi -> { noEffect with DataBar = dataBarDrawOf options color (min lo hi) (max lo hi) n }
         | _ -> noEffect)
    | IconSet (family, thresholds, reverse), Number n ->
        let count = max (iconCountOf family) thresholds.Length
        let resolved = thresholdsOf ctx rule thresholds
        // The highest step the value reaches; the first is the floor.
        let mutable index = 0
        for i in 1 .. resolved.Length - 1 do
            match resolved.[i] with
            | Some at when n >= at -> index <- i
            | _ -> ()
        { noEffect with IconSet = Some { Family = family; Index = (if reverse then count - 1 - index else index); Count = count; HideValue = false } }
    | IconSetEx (family, thresholds, reverse, options), Number n ->
        let count = max (iconCountOf family) thresholds.Length
        let resolved = thresholdsOf ctx rule thresholds
        let mutable index = 0
        for i in 1 .. resolved.Length - 1 do
            match resolved.[i] with
            | Some at when n >= at -> index <- i
            | _ -> ()
        let hide = not options.ShowValue
        (match List.tryItem index options.Custom with
         | Some (custom, _) when custom = "NoIcons" -> noEffect
         | Some (custom, icon) -> { noEffect with IconSet = Some { Family = custom; Index = icon; Count = iconCountOf custom; HideValue = hide } }
         | None -> { noEffect with IconSet = Some { Family = family; Index = (if reverse then count - 1 - index else index); Count = count; HideValue = hide } })
    | (ColorScale _ | DataBar _ | IconSet _ | DataBarEx _ | IconSetEx _), _ -> noEffect
    | _ ->
        let style = rule.Entry.Format.Style
        { noEffect with
            TextColor = style.TextColor
            BgColor = style.BgColor
            Bold = (if style.Bold then Some true else None)
            Italic = (if style.Italic then Some true else None)
            Underline = (if style.Underline then Some true else None)
            Strikethrough = (if style.Strikethrough then Some true else None)
            BorderTop = style.BorderTop
            BorderBottom = style.BorderBottom
            BorderLeft = style.BorderLeft
            BorderRight = style.BorderRight }

/// The first of two effects wins property by property.
let private mergeEffects (first: CellEffect) (second: CellEffect) : CellEffect =
    let pick a b = match a with Some _ -> a | None -> b
    { TextColor = pick first.TextColor second.TextColor
      BgColor = pick first.BgColor second.BgColor
      Bold = pick first.Bold second.Bold
      Italic = pick first.Italic second.Italic
      Underline = pick first.Underline second.Underline
      Strikethrough = pick first.Strikethrough second.Strikethrough
      DataBar = pick first.DataBar second.DataBar
      IconSet = pick first.IconSet second.IconSet
      BorderTop = pick first.BorderTop second.BorderTop
      BorderBottom = pick first.BorderBottom second.BorderBottom
      BorderLeft = pick first.BorderLeft second.BorderLeft
      BorderRight = pick first.BorderRight second.BorderRight }

/// What the sheet's conditional formats do to a cell; None when no rule
/// covering it matches.
let effectAt (ctx: Context) (addr: CellAddress) : CellEffect option =
    if ctx.Rules.IsEmpty then None
    else
        let raw = getCellValue addr ctx.Model
        let mutable effect = noEffect
        let mutable any = false
        let mutable stopped = false
        for rule in ctx.Rules do
            if not stopped && covers rule addr && ruleMatches ctx rule addr raw then
                effect <- mergeEffects effect (ruleEffect ctx rule raw)
                any <- true
                if rule.Entry.StopIfTrue then stopped <- true
        if any && effect <> noEffect then Some effect else None

// --- One-call access for renderers -----------------------------------------
// A renderer holds only the sheet model. The context for the last sheet state
// asked about is kept, so painting a viewport prepares the rules once.

/// Holds the evaluator in a record: a function kept in a mutable binding
/// loses its calling convention in the JavaScript build.
type private EvaluatorHolder = { Run: FormulaEvaluator }

let mutable private currentEvaluator : EvaluatorHolder = { Run = evaluateFormulaAt }

/// Choose how the formulas in rules are evaluated when a renderer asks. A
/// host with a workbook sets this so rules can read other sheets and defined
/// names; the default evaluates on the sheet alone.
let setFormulaEvaluator (evaluate: FormulaEvaluator) : unit =
    currentEvaluator <- { Run = evaluate }

let mutable private cachedCells : obj = null
let mutable private cachedRules : obj = null
let mutable private cachedEvaluator : obj = null
let mutable private cachedContext : Context option = None

/// Back to the default evaluator, and forget the sheet the last context was
/// made for: a host that goes away must not stay reachable from here.
let resetFormulaEvaluator () : unit =
    currentEvaluator <- { Run = evaluateFormulaAt }
    cachedCells <- null
    cachedRules <- null
    cachedEvaluator <- null
    cachedContext <- None

/// The evaluation context for a sheet (remembered for the latest sheet state).
let contextFor (model: SpreadsheetModel) : Context =
    match cachedContext with
    | Some ctx when obj.ReferenceEquals(cachedCells, model.Cells)
                    && obj.ReferenceEquals(cachedRules, model.ConditionalFormats)
                    && obj.ReferenceEquals(cachedEvaluator, currentEvaluator) -> ctx
    | _ ->
        let ctx = createContext currentEvaluator.Run model
        cachedCells <- box model.Cells
        cachedRules <- box model.ConditionalFormats
        cachedEvaluator <- box currentEvaluator
        cachedContext <- Some ctx
        ctx

let mutable private bordersOf : obj = null
let mutable private rulesHaveBorders = false

/// What ends the style of a line a conditional format gives a cell's side
/// (`withBorderEffects`): such a line is the edge's whatever the cell on the
/// other side has, and its "none" takes the edge's lines away (`Layout`).
let ruleLineMark = "!"

/// A cell with the lines the sheet's conditional formats give it (Excel's
/// `<dxf><border>`), marked as a rule's (`ruleLineMark`): the lines of a cell
/// are drawn from what the cells either side of an edge hold, so they are
/// laid over the cell itself. The cell as it is when no rule with lines
/// covers it.
let withBorderEffects (model: SpreadsheetModel) (addr: CellAddress) (data: CellData option) : CellData option =
    if model.ConditionalFormats.IsEmpty then data
    else
        if not (obj.ReferenceEquals(bordersOf, model.ConditionalFormats)) then
            bordersOf <- box model.ConditionalFormats
            rulesHaveBorders <- model.ConditionalFormats |> List.exists (fun entry -> styleHasBorders entry.Format.Style)
        if not rulesHaveBorders then data
        else
            match effectAt (contextFor model) addr with
            | Some effect when effect.BorderTop.IsSome || effect.BorderBottom.IsSome || effect.BorderLeft.IsSome || effect.BorderRight.IsSome ->
                let marked (rule: BorderStyle option) (own: BorderStyle option) =
                    match rule with
                    | Some line -> Some { line with Style = line.Style + ruleLineMark }
                    | None -> own
                let lay (format: CellFormat) =
                    { format with
                        BorderTop = marked effect.BorderTop format.BorderTop
                        BorderBottom = marked effect.BorderBottom format.BorderBottom
                        BorderLeft = marked effect.BorderLeft format.BorderLeft
                        BorderRight = marked effect.BorderRight format.BorderRight }
                match data with
                | Some cell -> Some { cell with Format = lay cell.Format }
                | None -> Some { Value = CellValue.Empty; Format = lay defaultFormat; Comment = None }
            | _ -> data

/// A cell's format with the sheet's conditional formats applied.
let formatAt (model: SpreadsheetModel) (addr: CellAddress) (format: CellFormat) : CellFormat =
    if model.ConditionalFormats.IsEmpty then format
    else
        match effectAt (contextFor model) addr with
        | Some effect -> applyEffect effect format
        | None -> format

// ---------------------------------------------------------------------------
// OOXML shape. The same record describes a `<cfRule>` read from a file (by
// either reader: the .NET one or the browser's) and one about to be written,
// so both import paths share `ofOoxml` and every writer shares `toOoxml`.
// ---------------------------------------------------------------------------

/// The differential style (`<dxf>`) a rule applies.
type OoxmlStyle = {
    /// The lines of its `<border>` (a side named without a style: "none").
    Borders: BorderStyle option list
    FontColor: string option
    FillColor: string option
    Bold: bool
    Italic: bool
    Underline: bool
    Strikethrough: bool
}

/// The x14 data bar of a rule (`<x14:dataBar>`), attribute by attribute.
type OoxmlDataBarExt = {
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
    /// The two ends with the x14 kinds (autoMin, autoMax, min, max, num, percent, percentile, formula).
    Cfvo: (string * string option) list
    BorderColor: string option
    NegativeFillColor: string option
    NegativeBorderColor: string option
    AxisColor: string option
}

/// What the x14 form of a rule holds beyond the `<cfRule>`.
type OoxmlRuleExt = {
    DataBar: OoxmlDataBarExt option
    /// Custom icons of an icon set: (icon set, icon id) per step.
    IconCustom: (string * int) list
    /// `showValue` of an icon set.
    ShowValue: bool
    /// The rule is in the sheet's x14 list only (an icon set of Excel 2010+,
    /// custom icons, a formula that reads another sheet).
    InExtList: bool
}

let emptyOoxmlRuleExt : OoxmlRuleExt = { DataBar = None; IconCustom = []; ShowValue = true; InExtList = false }

/// One `<cfRule>`: its attributes and children as they are in the file.
type OoxmlRule = {
    /// cellIs, expression, containsText, notContainsText, beginsWith,
    /// endsWith, top10, aboveAverage, duplicateValues, uniqueValues,
    /// containsBlanks, notContainsBlanks, containsErrors, notContainsErrors,
    /// colorScale, dataBar, iconSet.
    Type: string
    Operator: string option
    Formulas: string list
    Text: string option
    Rank: int option
    Percent: bool
    Bottom: bool
    AboveAverage: bool
    EqualAverage: bool
    StdDev: int option
    /// (cfvo type, cfvo value, colour) per stop.
    ColorScale: (string * string option * string) list
    DataBarColor: string option
    DataBarCfvo: (string * string option) list
    IconSet: string option
    IconSetCfvo: (string * string option) list
    IconSetReverse: bool
    /// `timePeriod` of a "date occurring" rule ("today", "last7Days" ...).
    TimePeriod: string option
    Priority: int option
    StopIfTrue: bool
    Style: OoxmlStyle option
    /// The x14 extension of the rule, when it has one.
    Ext: OoxmlRuleExt option
}

/// One `<conditionalFormatting sqref="…">` block.
type OoxmlConditionalFormat = {
    Sqref: string
    Rules: OoxmlRule list
}

let emptyOoxmlRule (kind: string) : OoxmlRule =
    { Type = kind; Operator = None; Formulas = []; Text = None; Rank = None; Percent = false; Bottom = false
      AboveAverage = true; EqualAverage = false; StdDev = None; ColorScale = []; DataBarColor = None; DataBarCfvo = []
      IconSet = None; IconSetCfvo = []; IconSetReverse = false; TimePeriod = None; Priority = None; StopIfTrue = false; Style = None; Ext = None }

[<Literal>]
let private lastRow = 1048575
[<Literal>]
let private lastCol = 16383

/// Parse one area of an sqref (`B2`, `A1:C9`, `$A:$A`, `3:7`) to a range.
/// Whole columns / rows keep their full extent.
let parseArea (text: string) : CellRange option =
    let clean = text.Replace("$", "").Trim()
    let corner (part: string) : (int option * int option) option =
        let mutable i = 0
        let mutable col = 0
        while i < part.Length && Char.IsLetter part.[i] do
            col <- col * 26 + (int (Char.ToUpperInvariant part.[i]) - int 'A' + 1)
            i <- i + 1
        let letters = i
        let mutable row = 0
        while i < part.Length && Char.IsDigit part.[i] do
            row <- row * 10 + (int part.[i] - int '0')
            i <- i + 1
        if i <> part.Length || part.Length = 0 then None
        else
            Some ((if i > letters then Some (row - 1) else None), (if letters > 0 then Some (col - 1) else None))
    let parts = clean.Split(':')
    if parts.Length = 0 || parts.Length > 2 then None
    else
        match corner parts.[0], corner parts.[parts.Length - 1] with
        | Some (r0, c0), Some (r1, c1) ->
            let rowLo, rowHi =
                match r0, r1 with
                | Some a, Some b -> min a b, max a b
                | _ -> 0, lastRow
            let colLo, colHi =
                match c0, c1 with
                | Some a, Some b -> min a b, max a b
                | _ -> 0, lastCol
            if rowLo < 0 || colLo < 0 || rowHi > lastRow || colHi > lastCol then None
            else Some { From = { Row = rowLo; Col = colLo }; To = { Row = rowHi; Col = colHi } }
        | _ -> None

/// Parse an sqref: areas separated by spaces.
let parseSqref (sqref: string) : CellRange list =
    sqref.Split([| ' ' |], StringSplitOptions.RemoveEmptyEntries)
    |> Array.choose parseArea
    |> Array.toList

let private areaText (range: CellRange) : string =
    let wholeColumns = range.From.Row = 0 && range.To.Row >= lastRow
    let wholeRows = range.From.Col = 0 && range.To.Col >= lastCol
    if wholeColumns && not wholeRows then columnToLetter range.From.Col + ":" + columnToLetter range.To.Col
    elif wholeRows && not wholeColumns then string (range.From.Row + 1) + ":" + string (range.To.Row + 1)
    elif range.From = range.To then formatCellRef range.From.Row range.From.Col
    else formatCellRef range.From.Row range.From.Col + ":" + formatCellRef range.To.Row range.To.Col

/// The sqref text of some ranges.
let sqrefText (ranges: CellRange list) : string =
    ranges |> List.map areaText |> String.concat " "

let private operatorOfText (text: string) : CompareOperator option =
    match text with
    | "greaterThan" -> Some OpGreaterThan
    | "lessThan" -> Some OpLessThan
    | "greaterThanOrEqual" -> Some OpGreaterOrEqual
    | "lessThanOrEqual" -> Some OpLessOrEqual
    | "equal" -> Some OpEqual
    | "notEqual" -> Some OpNotEqual
    | "between" -> Some OpBetween
    | "notBetween" -> Some OpNotBetween
    | _ -> None

/// The OOXML name of a comparison operator (shared with data validation).
let operatorText (op: CompareOperator) : string =
    match op with
    | OpGreaterThan -> "greaterThan"
    | OpLessThan -> "lessThan"
    | OpGreaterOrEqual -> "greaterThanOrEqual"
    | OpLessOrEqual -> "lessThanOrEqual"
    | OpEqual -> "equal"
    | OpNotEqual -> "notEqual"
    | OpBetween -> "between"
    | OpNotBetween -> "notBetween"

/// Parse the OOXML name of a comparison operator.
let parseOperator (text: string) : CompareOperator option = operatorOfText text

let private thresholdOf (kind: string, value: string option) : CfThreshold =
    let thresholdKind =
        match kind with
        | "min" -> CfvoMin
        | "max" -> CfvoMax
        | "num" -> CfvoNumber
        | "percent" -> CfvoPercent
        | "percentile" -> CfvoPercentile
        | "formula" -> CfvoFormula
        | "autoMin" -> CfvoAutoMin
        | "autoMax" -> CfvoAutoMax
        | _ -> CfvoNumber
    { Kind = thresholdKind; Value = value }

let private cfvoOf (threshold: CfThreshold) : string * string option =
    let kind =
        match threshold.Kind with
        | CfvoMin -> "min"
        | CfvoMax -> "max"
        | CfvoNumber -> "num"
        | CfvoPercent -> "percent"
        | CfvoPercentile -> "percentile"
        | CfvoFormula -> "formula"
        | CfvoAutoMin -> "autoMin"
        | CfvoAutoMax -> "autoMax"
    kind, threshold.Value

/// The `<cfvo>` of the 2007 form for a threshold: it has no automatic ends.
let private legacyCfvoOf (threshold: CfThreshold) : string * string option =
    match threshold.Kind with
    | CfvoAutoMin -> "min", None
    | CfvoAutoMax -> "max", None
    | _ -> cfvoOf threshold

let private cssColor (color: string) : string =
    let hex = color.TrimStart('#')
    let hex = if hex.Length = 8 then hex.Substring 2 else hex
    "#" + hex.ToUpperInvariant()

let private formulaBody (text: string) = if text.StartsWith "=" then text.Substring 1 else text

/// The model rule for one `<cfRule>`, or None for a rule type the model does
/// not express (x14 extensions).
let private ruleOfOoxml (rule: OoxmlRule) : ConditionalFormatRule option =
    let formula index = rule.Formulas |> List.tryItem index
    match rule.Type with
    | "cellIs" ->
        (match rule.Operator |> Option.bind operatorOfText, formula 0 with
         | Some op, Some first -> Some (CellIs (op, first, formula 1))
         | _ -> None)
    | "expression" -> formula 0 |> Option.map Expression
    | "containsText" -> rule.Text |> Option.map (fun text -> TextRule (TextContains, text))
    | "notContainsText" -> rule.Text |> Option.map (fun text -> TextRule (TextNotContains, text))
    | "beginsWith" -> rule.Text |> Option.map (fun text -> TextRule (TextBeginsWith, text))
    | "endsWith" -> rule.Text |> Option.map (fun text -> TextRule (TextEndsWith, text))
    | "top10" -> Some (TopBottom (defaultArg rule.Rank 10, rule.Percent, rule.Bottom))
    | "aboveAverage" -> Some (AboveAverage (rule.AboveAverage, rule.EqualAverage, defaultArg rule.StdDev 0))
    | "duplicateValues" -> Some DuplicateValues
    | "uniqueValues" -> Some UniqueValues
    | "containsBlanks" -> Some (Blanks true)
    | "notContainsBlanks" -> Some (Blanks false)
    | "containsErrors" -> Some (Errors true)
    | "notContainsErrors" -> Some (Errors false)
    | "timePeriod" -> rule.TimePeriod |> Option.filter (fun period -> List.contains period timePeriods) |> Option.map TimePeriod
    | "colorScale" when rule.ColorScale.Length >= 2 ->
        Some (ColorScale (rule.ColorScale |> List.map (fun (kind, value, color) -> thresholdOf (kind, value), cssColor color)))
    | "dataBar" ->
        let low, high =
            match rule.DataBarCfvo with
            | first :: second :: _ -> thresholdOf first, thresholdOf second
            | _ -> { Kind = CfvoMin; Value = None }, { Kind = CfvoMax; Value = None }
        let color = cssColor (defaultArg rule.DataBarColor "638EC6")
        (match rule.Ext |> Option.bind (fun ext -> ext.DataBar) with
         | Some bar ->
            let low, high =
                match bar.Cfvo with
                | first :: second :: _ -> thresholdOf first, thresholdOf second
                | _ -> low, high
            let options =
                { Gradient = bar.Gradient
                  Border = (if bar.Border then Some (cssColor (defaultArg bar.BorderColor (color.TrimStart '#'))) else None)
                  NegativeFill = (if bar.NegativeBarColorSameAsPositive then None else Some (cssColor (defaultArg bar.NegativeFillColor "FF0000")))
                  NegativeBorder = (if bar.NegativeBarBorderColorSameAsPositive || not bar.Border then None else Some (cssColor (defaultArg bar.NegativeBorderColor "FF0000")))
                  Axis = (match bar.AxisPosition with "middle" -> AxisMiddle | "none" -> AxisNone | _ -> AxisAutomatic)
                  AxisColor = bar.AxisColor |> Option.map cssColor
                  Direction = (match bar.Direction with "leftToRight" -> BarLeftToRight | "rightToLeft" -> BarRightToLeft | _ -> BarContext)
                  ShowValue = bar.ShowValue
                  MinLength = bar.MinLength
                  MaxLength = bar.MaxLength }
            Some (DataBarEx (low, high, color, options))
         | None -> Some (DataBar (low, high, color)))
    | "iconSet" ->
        let family = defaultArg rule.IconSet "3TrafficLights1"
        (match rule.Ext with
         | Some ext when ext.InExtList || not ext.ShowValue || not ext.IconCustom.IsEmpty ->
            Some (IconSetEx (family, rule.IconSetCfvo |> List.map thresholdOf, rule.IconSetReverse, { ShowValue = ext.ShowValue; Custom = ext.IconCustom }))
         | _ -> Some (IconSet (family, rule.IconSetCfvo |> List.map thresholdOf, rule.IconSetReverse)))
    | _ -> None

/// Model rules for conditional formatting read from a file. Priorities keep
/// the file's order of precedence (lower first); rules without one follow in
/// document order.
let ofOoxml (formats: OoxmlConditionalFormat list) : ConditionalFormatRange list =
    let mutable sequence = 0
    [ for format in formats do
        let ranges = parseSqref format.Sqref
        for rule in format.Rules do
            sequence <- sequence + 1
            match ruleOfOoxml rule with
            | Some modelRule when not ranges.IsEmpty ->
                let style =
                    match rule.Style with
                    | Some s ->
                        { defaultFormat with
                            TextColor = s.FontColor |> Option.map cssColor
                            BgColor = s.FillColor |> Option.map cssColor
                            Bold = s.Bold
                            Italic = s.Italic
                            Underline = s.Underline
                            Strikethrough = s.Strikethrough
                            BorderTop = List.tryItem 0 s.Borders |> Option.flatten
                            BorderBottom = List.tryItem 1 s.Borders |> Option.flatten
                            BorderLeft = List.tryItem 2 s.Borders |> Option.flatten
                            BorderRight = List.tryItem 3 s.Borders |> Option.flatten }
                    | None -> defaultFormat
                yield { Id = 0
                        Ranges = ranges
                        Format = { Rule = modelRule; Style = style }
                        Priority = (match rule.Priority with Some p -> p | None -> 100000 + sequence)
                        StopIfTrue = rule.StopIfTrue }
            | _ -> () ]
    // (Ids: 1, 2, 3 ... in order of precedence, the same for the same file.)
    |> List.sortBy (fun entry -> entry.Priority)
    |> List.mapi (fun index entry -> { entry with Id = index + 1 })

let private numberText (value: float) =
    if value = Math.Floor value && abs value < 1e15 then (int64 value).ToString(CultureInfo.InvariantCulture)
    else value.ToString("R", CultureInfo.InvariantCulture)

/// The `<cfRule>` for one model rule. `anchor` is the top-left of the rule's
/// first range, which text rules need for the formula Excel expects.
let private ooxmlOfRule (anchor: CellAddress) (rule: ConditionalFormatRule) : OoxmlRule =
    let cellText = formatCellRef anchor.Row anchor.Col
    let quoted (text: string) = "\"" + text.Replace("\"", "\"\"") + "\""
    match rule with
    | GreaterThan n -> { emptyOoxmlRule "cellIs" with Operator = Some "greaterThan"; Formulas = [ numberText n ] }
    | LessThan n -> { emptyOoxmlRule "cellIs" with Operator = Some "lessThan"; Formulas = [ numberText n ] }
    | Equal n -> { emptyOoxmlRule "cellIs" with Operator = Some "equal"; Formulas = [ numberText n ] }
    | Between (lo, hi) -> { emptyOoxmlRule "cellIs" with Operator = Some "between"; Formulas = [ numberText lo; numberText hi ] }
    | CellIs (op, first, second) ->
        { emptyOoxmlRule "cellIs" with
            Operator = Some (operatorText op)
            Formulas = formulaBody first :: (second |> Option.map formulaBody |> Option.toList) }
    | Expression formula -> { emptyOoxmlRule "expression" with Formulas = [ formulaBody formula ] }
    | ContainsText text | TextRule (TextContains, text) ->
        { emptyOoxmlRule "containsText" with
            Operator = Some "containsText"; Text = Some text
            Formulas = [ sprintf "NOT(ISERROR(SEARCH(%s,%s)))" (quoted text) cellText ] }
    | TextRule (TextNotContains, text) ->
        { emptyOoxmlRule "notContainsText" with
            Operator = Some "notContains"; Text = Some text
            Formulas = [ sprintf "ISERROR(SEARCH(%s,%s))" (quoted text) cellText ] }
    | TextRule (TextBeginsWith, text) ->
        { emptyOoxmlRule "beginsWith" with
            Operator = Some "beginsWith"; Text = Some text
            Formulas = [ sprintf "LEFT(%s,%d)=%s" cellText text.Length (quoted text) ] }
    | TextRule (TextEndsWith, text) ->
        { emptyOoxmlRule "endsWith" with
            Operator = Some "endsWith"; Text = Some text
            Formulas = [ sprintf "RIGHT(%s,%d)=%s" cellText text.Length (quoted text) ] }
    | TopBottom (rank, percent, bottom) -> { emptyOoxmlRule "top10" with Rank = Some rank; Percent = percent; Bottom = bottom }
    | AboveAverage (above, orEqual, stdDev) ->
        { emptyOoxmlRule "aboveAverage" with
            AboveAverage = above; EqualAverage = orEqual; StdDev = (if stdDev > 0 then Some stdDev else None) }
    | DuplicateValues -> emptyOoxmlRule "duplicateValues"
    | UniqueValues -> emptyOoxmlRule "uniqueValues"
    | Blanks true -> { emptyOoxmlRule "containsBlanks" with Formulas = [ sprintf "LEN(TRIM(%s))=0" cellText ] }
    | Blanks false -> { emptyOoxmlRule "notContainsBlanks" with Formulas = [ sprintf "LEN(TRIM(%s))>0" cellText ] }
    | Errors true -> { emptyOoxmlRule "containsErrors" with Formulas = [ sprintf "ISERROR(%s)" cellText ] }
    | Errors false -> { emptyOoxmlRule "notContainsErrors" with Formulas = [ sprintf "NOT(ISERROR(%s))" cellText ] }
    | TimePeriod period ->
        { emptyOoxmlRule "timePeriod" with TimePeriod = Some period; Formulas = timePeriodFormula period cellText |> Option.toList }
    | ColorScale stops ->
        { emptyOoxmlRule "colorScale" with
            ColorScale = stops |> List.map (fun (threshold, color) -> let kind, value = cfvoOf threshold in kind, value, color.TrimStart('#')) }
    | DataBar (low, high, color) ->
        { emptyOoxmlRule "dataBar" with DataBarColor = Some (color.TrimStart('#')); DataBarCfvo = [ cfvoOf low; cfvoOf high ] }
    | IconSet (family, thresholds, reverse) ->
        { emptyOoxmlRule "iconSet" with IconSet = Some family; IconSetCfvo = thresholds |> List.map cfvoOf; IconSetReverse = reverse }
    | DataBarEx (low, high, color, options) ->
        let hex (text: string) = text.TrimStart('#')
        { emptyOoxmlRule "dataBar" with
            DataBarColor = Some (hex color)
            DataBarCfvo = [ legacyCfvoOf low; legacyCfvoOf high ]
            Ext =
                Some { emptyOoxmlRuleExt with
                        DataBar =
                            Some { MinLength = options.MinLength; MaxLength = options.MaxLength; ShowValue = options.ShowValue
                                   Border = options.Border.IsSome; Gradient = options.Gradient
                                   Direction = (match options.Direction with BarLeftToRight -> "leftToRight" | BarRightToLeft -> "rightToLeft" | BarContext -> "context")
                                   NegativeBarColorSameAsPositive = options.NegativeFill.IsNone
                                   NegativeBarBorderColorSameAsPositive = options.NegativeBorder.IsNone
                                   AxisPosition = (match options.Axis with AxisMiddle -> "middle" | AxisNone -> "none" | AxisAutomatic -> "automatic")
                                   Cfvo = [ cfvoOf low; cfvoOf high ]
                                   BorderColor = options.Border |> Option.map hex
                                   NegativeFillColor = options.NegativeFill |> Option.map hex
                                   NegativeBorderColor = options.NegativeBorder |> Option.map hex
                                   AxisColor = options.AxisColor |> Option.map hex } } }
    | IconSetEx (family, thresholds, reverse, options) ->
        { emptyOoxmlRule "iconSet" with
            IconSet = Some family; IconSetCfvo = thresholds |> List.map cfvoOf; IconSetReverse = reverse
            Ext =
                Some { emptyOoxmlRuleExt with
                        IconCustom = options.Custom
                        ShowValue = options.ShowValue
                        InExtList = not options.Custom.IsEmpty || List.contains family extendedIconSets } }

/// The model's conditional formats as `<conditionalFormatting>` blocks, one
/// per distinct range set, with every rule's style attached. Priorities are
/// renumbered 1..n in order of precedence, as Excel requires them unique.
let toOoxml (formats: ConditionalFormatRange list) : OoxmlConditionalFormat list =
    let ordered = formats |> List.sortBy (fun entry -> entry.Priority) |> List.mapi (fun i entry -> i + 1, entry)
    let blocks = ResizeArray<string * ResizeArray<OoxmlRule>>()
    for (priority, entry) in ordered do
        match entry.Ranges with
        | [] -> ()
        | first :: _ ->
            let sqref = sqrefText entry.Ranges
            let style = entry.Format.Style
            let hasStyle =
                style.TextColor.IsSome || style.BgColor.IsSome || style.Bold || style.Italic || style.Underline || style.Strikethrough
                || styleHasBorders style
            let visual =
                match entry.Format.Rule with
                | ColorScale _ | DataBar _ | IconSet _ | DataBarEx _ | IconSetEx _ -> true
                | _ -> false
            let described = ooxmlOfRule first.From entry.Format.Rule
            // A formula that reads another sheet: Excel keeps such a rule in the x14 list.
            let described =
                if described.Ext.IsNone && described.Formulas |> List.exists (fun formula -> formula.Contains "!") then
                    { described with Ext = Some { emptyOoxmlRuleExt with InExtList = true } }
                else described
            let rule =
                { described with
                    Priority = Some priority
                    StopIfTrue = entry.StopIfTrue
                    Style =
                        if visual || not hasStyle then None
                        else Some { FontColor = style.TextColor; FillColor = style.BgColor; Bold = style.Bold; Italic = style.Italic
                                    Underline = style.Underline; Strikethrough = style.Strikethrough
                                    Borders = (if styleHasBorders style then [ style.BorderTop; style.BorderBottom; style.BorderLeft; style.BorderRight ] else []) } }
            match blocks |> Seq.tryFind (fun (key, _) -> key = sqref) with
            | Some (_, rules) -> rules.Add rule
            | None ->
                let rules = ResizeArray<OoxmlRule>()
                rules.Add rule
                blocks.Add((sqref, rules))
    [ for (sqref, rules) in blocks -> { Sqref = sqref; Rules = List.ofSeq rules } ]

// ---------------------------------------------------------------------------
// Data validation in its OOXML shape, shared by both importers and the
// writers the same way the conditional-format records above are.
// ---------------------------------------------------------------------------

/// One `<dataValidation>` (without its sqref).
type OoxmlValidation = {
    /// list, whole, decimal, date, time, textLength, custom.
    Type: string
    Operator: string option
    Formula1: string option
    Formula2: string option
    AllowBlank: bool
    ShowErrorMessage: bool
    ShowInputMessage: bool
    /// A list offers its values in the cell (the file's `showDropDown`
    /// attribute says the opposite: set, it hides the dropdown).
    ShowDropDown: bool
    /// stop (default), warning, information.
    ErrorStyle: string option
    ErrorTitle: string option
    Error: string option
    PromptTitle: string option
    Prompt: string option
}

/// The values of a list validation written inline (`"Yes,No"`); None when
/// the source is a range or a name, which the caller has to resolve.
let literalListValues (source: string) : string list option =
    let source = source.Trim()
    if source.Length >= 2 && source.StartsWith "\"" && source.EndsWith "\"" then
        Some (source.Substring(1, source.Length - 2).Split(',') |> Array.map (fun v -> v.Trim()) |> Array.toList)
    else None

/// The model rule of a validation; None for types the model does not know.
/// A list whose source is not written out keeps its source (`ListSource`).
let validationRuleOfOoxml (validation: OoxmlValidation) : ValidationRule option =
    let op = validation.Operator |> Option.bind parseOperator |> Option.defaultValue OpBetween
    let kind =
        match validation.Type with
        | "whole" -> Some WholeKind
        | "decimal" -> Some DecimalKind
        | "date" -> Some DateKind
        | "time" -> Some TimeKind
        | "textLength" -> Some TextLengthKind
        | _ -> None
    match validation.Type, kind, validation.Formula1 with
    | "list", _, Some source ->
        (match literalListValues source with
         | Some values -> Some (ListValues values)
         | None ->
            let source = source.Trim()
            Some (ListSource (if source.StartsWith "=" then source.Substring 1 else source)))
    | "none", _, _ -> Some AnyValue
    | "custom", _, Some formula -> Some (CustomFormula formula)
    | _, Some kind, Some first ->
        // Inclusive bounds given as plain numbers keep the simple range rules.
        let bounds =
            match op, parseNumber first, validation.Formula2 |> Option.bind parseNumber with
            | OpBetween, Some a, Some b -> Some (Some (min a b), Some (max a b))
            | OpEqual, Some a, _ -> Some (Some a, Some a)
            | OpGreaterOrEqual, Some a, _ -> Some (Some a, None)
            | OpLessOrEqual, Some a, _ -> Some (None, Some a)
            | _ -> None
        match kind, bounds with
        | WholeKind, Some (lo, hi) -> Some (WholeNumber (lo, hi))
        | DecimalKind, Some (lo, hi) -> Some (DecimalRange (lo, hi))
        | TextLengthKind, Some (lo, hi) -> Some (TextLength (lo, hi))
        | DateKind, Some (lo, hi) -> Some (DateRange (lo, hi))
        | _ -> Some (Comparison (kind, op, first, validation.Formula2))
    | _ -> None

/// A model validation for `rule` carrying the alert settings of the file's
/// `<dataValidation>`. `anchor` is the top-left of its range.
let validationWithOoxml (anchor: CellAddress) (validation: OoxmlValidation) (rule: ValidationRule) : DataValidation =
    { Rule = rule
      ErrorMessage = validation.Error
      InputMessage = validation.Prompt
      ErrorTitle = validation.ErrorTitle
      InputTitle = validation.PromptTitle
      Style =
        (match validation.ErrorStyle with
         | Some "warning" -> WarningStyle
         | Some "information" -> InformationStyle
         | _ -> StopStyle)
      // A list keeps its dropdown either way; other rules without an alert
      // are advisory.
      ShowError = validation.ShowErrorMessage
      AllowBlank = validation.AllowBlank
      ShowInput = validation.ShowInputMessage
      ShowDropdown = validation.ShowDropDown
      Anchor = Some anchor }

/// The `<dataValidation>` for a model validation.
let ooxmlOfValidation (validation: DataValidation) : OoxmlValidation =
    let number (value: float) = numberText value
    let bounded (kind: string) (lo: float option) (hi: float option) =
        match lo, hi with
        | Some a, Some b when a = b -> kind, Some "equal", Some (number a), None
        | Some a, Some b -> kind, Some "between", Some (number a), Some (number b)
        | Some a, None -> kind, Some "greaterThanOrEqual", Some (number a), None
        | None, Some b -> kind, Some "lessThanOrEqual", Some (number b), None
        | None, None -> kind, None, None, None
    let kind, op, first, second =
        match validation.Rule with
        | ListValues values -> "list", None, Some ("\"" + String.concat "," values + "\""), None
        | WholeNumber (lo, hi) -> bounded "whole" lo hi
        | DecimalRange (lo, hi) -> bounded "decimal" lo hi
        | TextLength (lo, hi) -> bounded "textLength" lo hi
        | DateRange (lo, hi) -> bounded "date" lo hi
        | Comparison (kind, op, firstOperand, secondOperand) ->
            let name =
                match kind with
                | WholeKind -> "whole"
                | DecimalKind -> "decimal"
                | DateKind -> "date"
                | TimeKind -> "time"
                | TextLengthKind -> "textLength"
            name, Some (operatorText op), Some (formulaBody firstOperand), secondOperand |> Option.map formulaBody
        | CustomFormula formula -> "custom", None, Some (formulaBody formula), None
        | ListSource source -> "list", None, Some (formulaBody source), None
        | AnyValue -> "none", None, None, None
    { Type = kind
      Operator = op
      Formula1 = first
      Formula2 = second
      AllowBlank = validation.AllowBlank
      ShowErrorMessage = validation.ShowError
      ShowInputMessage = validation.ShowInput
      ShowDropDown = validation.ShowDropdown
      ErrorStyle =
        (match validation.Style with
         | StopStyle -> None
         | WarningStyle -> Some "warning"
         | InformationStyle -> Some "information")
      ErrorTitle = validation.ErrorTitle
      Error = validation.ErrorMessage
      PromptTitle = validation.InputTitle
      Prompt = validation.InputMessage }

/// The validations of a sheet for writing: one `<dataValidation>` each, its
/// sqref the validation's ranges (a whole column is `A1:A1048576`). The
/// formulas of a rule are written for the top-left cell of its first range,
/// as the file format has them.
let validationsToOoxml (validations: ValidationRange list) : (string * OoxmlValidation) list =
    [ for entry in validations do
        match entry.Ranges with
        | [] -> ()
        | first :: _ ->
            let anchor = validationAnchor entry
            let deltaRow, deltaCol = first.From.Row - anchor.Row, first.From.Col - anchor.Col
            let validation =
                if deltaRow = 0 && deltaCol = 0 then entry.Validation
                else
                    let shift (formula: string) =
                        let shifted = shiftFormulaReferences deltaRow deltaCol ("=" + formulaBody formula)
                        if formula.StartsWith "=" then shifted else shifted.TrimStart('=')
                    let rule =
                        match entry.Validation.Rule with
                        | Comparison (kind, op, firstOperand, secondOperand) -> Comparison (kind, op, shift firstOperand, secondOperand |> Option.map shift)
                        | CustomFormula formula -> CustomFormula (shift formula)
                        | ListSource source -> ListSource (shift source)
                        | other -> other
                    { entry.Validation with Rule = rule }
            // (Whole columns and rows are written out, as Excel writes them.)
            let area (range: CellRange) =
                if range.From = range.To then formatCellRef range.From.Row range.From.Col
                else formatCellRef range.From.Row range.From.Col + ":" + formatCellRef range.To.Row range.To.Col
            yield (entry.Ranges |> List.map area |> String.concat " "), ooxmlOfValidation validation ]
