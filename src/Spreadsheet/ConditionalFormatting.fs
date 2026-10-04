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
    /// (fill fraction 0..1, colour) — see `CellFormat.DataBar`.
    DataBar: (float * string) option
    /// (icon index, icon count) — see `CellFormat.IconSet`.
    IconSet: (int * int) option
}

let noEffect : CellEffect =
    { TextColor = None; BgColor = None; Bold = None; Italic = None; Underline = None; Strikethrough = None
      DataBar = None; IconSet = None }

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
        IconSet = (match effect.IconSet with Some _ -> effect.IconSet | None -> format.IconSet) }

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
                match counts.TryGetValue key with
                | true, n -> counts.[key] <- n + 1
                | false, _ -> counts.[key] <- 1
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

/// Does a cell's value satisfy `value <op> operand(s)`? Numbers compare with
/// numbers and text with text (ignoring case); a blank cell counts as 0, as
/// Excel's "cell value" rules treat it.
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
        | Some x, Some a -> compareHolds op x a (second |> Option.bind asNumber)
        // Different kinds of value are never equal, and have no order.
        | _ -> op = OpNotEqual

/// Does this rule match the cell? Visual rules (scale, bar, icons) "match"
/// every number in their range; what they paint is computed separately.
let private ruleMatches (ctx: Context) (rule: PreparedRule) (addr: CellAddress) (raw: CellValue) : bool =
    let value = unwrapFormulaValue raw
    let formulaAt (formula: string) =
        ctx.Evaluate formula (addr.Row - rule.Anchor.Row) (addr.Col - rule.Anchor.Col) ctx.Model addr |> unwrapFormulaValue
    match rule.Entry.Format.Rule with
    | ContainsText _ -> evaluateRule rule.Entry.Format.Rule raw
    | GreaterThan _ | LessThan _ | Between _ | Equal _ -> evaluateRule rule.Entry.Format.Rule value
    | CellIs (op, first, second) -> cellIsHolds op value (formulaAt first) (second |> Option.map formulaAt)
    | Expression formula -> isTruthy (formulaAt formula)
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
    | ColorScale _ | DataBar _ | IconSet _ -> (match value with Number _ -> true | _ -> false)

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
            { noEffect with DataBar = Some (max 0.0 (min 1.0 fraction), (if color.StartsWith "#" then color else "#" + color)) }
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
        { noEffect with IconSet = Some ((if reverse then count - 1 - index else index), count) }
    | (ColorScale _ | DataBar _ | IconSet _), _ -> noEffect
    | _ ->
        let style = rule.Entry.Format.Style
        { noEffect with
            TextColor = style.TextColor
            BgColor = style.BgColor
            Bold = (if style.Bold then Some true else None)
            Italic = (if style.Italic then Some true else None)
            Underline = (if style.Underline then Some true else None)
            Strikethrough = (if style.Strikethrough then Some true else None) }

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
      IconSet = pick first.IconSet second.IconSet }

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
    FontColor: string option
    FillColor: string option
    Bold: bool
    Italic: bool
}

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
    Priority: int option
    StopIfTrue: bool
    Style: OoxmlStyle option
}

/// One `<conditionalFormatting sqref="…">` block.
type OoxmlConditionalFormat = {
    Sqref: string
    Rules: OoxmlRule list
}

let emptyOoxmlRule (kind: string) : OoxmlRule =
    { Type = kind; Operator = None; Formulas = []; Text = None; Rank = None; Percent = false; Bottom = false
      AboveAverage = true; EqualAverage = false; StdDev = None; ColorScale = []; DataBarColor = None; DataBarCfvo = []
      IconSet = None; IconSetCfvo = []; IconSetReverse = false; Priority = None; StopIfTrue = false; Style = None }

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
    kind, threshold.Value

let private cssColor (color: string) : string =
    let hex = color.TrimStart('#')
    let hex = if hex.Length = 8 then hex.Substring 2 else hex
    "#" + hex.ToUpperInvariant()

let private formulaBody (text: string) = if text.StartsWith "=" then text.Substring 1 else text

/// The model rule for one `<cfRule>`, or None for a rule type the model does
/// not express (timePeriod, x14 extensions).
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
    | "colorScale" when rule.ColorScale.Length >= 2 ->
        Some (ColorScale (rule.ColorScale |> List.map (fun (kind, value, color) -> thresholdOf (kind, value), cssColor color)))
    | "dataBar" ->
        let low, high =
            match rule.DataBarCfvo with
            | first :: second :: _ -> thresholdOf first, thresholdOf second
            | _ -> { Kind = CfvoMin; Value = None }, { Kind = CfvoMax; Value = None }
        Some (DataBar (low, high, cssColor (defaultArg rule.DataBarColor "638EC6")))
    | "iconSet" ->
        let family = defaultArg rule.IconSet "3TrafficLights1"
        Some (IconSet (family, rule.IconSetCfvo |> List.map thresholdOf, rule.IconSetReverse))
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
                            Italic = s.Italic }
                    | None -> defaultFormat
                yield { Ranges = ranges
                        Format = { Rule = modelRule; Style = style }
                        Priority = (match rule.Priority with Some p -> p | None -> 100000 + sequence)
                        StopIfTrue = rule.StopIfTrue }
            | _ -> () ]

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
    | ColorScale stops ->
        { emptyOoxmlRule "colorScale" with
            ColorScale = stops |> List.map (fun (threshold, color) -> let kind, value = cfvoOf threshold in kind, value, color.TrimStart('#')) }
    | DataBar (low, high, color) ->
        { emptyOoxmlRule "dataBar" with DataBarColor = Some (color.TrimStart('#')); DataBarCfvo = [ cfvoOf low; cfvoOf high ] }
    | IconSet (family, thresholds, reverse) ->
        { emptyOoxmlRule "iconSet" with IconSet = Some family; IconSetCfvo = thresholds |> List.map cfvoOf; IconSetReverse = reverse }

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
                style.TextColor.IsSome || style.BgColor.IsSome || style.Bold || style.Italic
            let visual =
                match entry.Format.Rule with
                | ColorScale _ | DataBar _ | IconSet _ -> true
                | _ -> false
            let rule =
                { ooxmlOfRule first.From entry.Format.Rule with
                    Priority = Some priority
                    StopIfTrue = entry.StopIfTrue
                    Style =
                        if visual || not hasStyle then None
                        else Some { FontColor = style.TextColor; FillColor = style.BgColor; Bold = style.Bold; Italic = style.Italic } }
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

/// The model rule of a validation; None for a list whose source is not a
/// literal and for types the model does not know.
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
    | "list", _, Some source -> literalListValues source |> Option.map ListValues
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
      InputMessage = (if validation.ShowInputMessage then validation.Prompt else None)
      ErrorTitle = validation.ErrorTitle
      InputTitle = (if validation.ShowInputMessage then validation.PromptTitle else None)
      Style =
        (match validation.ErrorStyle with
         | Some "warning" -> WarningStyle
         | Some "information" -> InformationStyle
         | _ -> StopStyle)
      // A list keeps its dropdown either way; other rules without an alert
      // are advisory.
      ShowError = validation.ShowErrorMessage
      AllowBlank = validation.AllowBlank
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
    { Type = kind
      Operator = op
      Formula1 = first
      Formula2 = second
      AllowBlank = validation.AllowBlank
      ShowErrorMessage = validation.ShowError
      ShowInputMessage = validation.InputMessage.IsSome || validation.InputTitle.IsSome
      ErrorStyle =
        (match validation.Style with
         | StopStyle -> None
         | WarningStyle -> Some "warning"
         | InformationStyle -> Some "information")
      ErrorTitle = validation.ErrorTitle
      Error = validation.ErrorMessage
      PromptTitle = validation.InputTitle
      Prompt = validation.InputMessage }

/// Validations grouped back into ranges for writing: cells that share a rule
/// become one `<dataValidation>` whose sqref lists maximal column runs.
let validationsToOoxml (validations: Map<CellAddress, DataValidation>) : (string * OoxmlValidation) list =
    let groups = ResizeArray<DataValidation * ResizeArray<CellAddress>>()
    for KeyValue(addr, validation) in validations do
        match groups |> Seq.tryFind (fun (key, _) -> key = validation) with
        | Some (_, cells) -> cells.Add addr
        | None ->
            let cells = ResizeArray<CellAddress>()
            cells.Add addr
            groups.Add((validation, cells))
    [ for (validation, cells) in groups do
        // Vertical runs per column, then neighbouring columns with the same
        // run merge into one rectangle.
        let runs = ResizeArray<int * int * int>()   // col, first row, last row
        for (col, rows) in cells |> Seq.groupBy (fun cell -> cell.Col) |> Seq.sortBy fst do
            let sorted = rows |> Seq.map (fun cell -> cell.Row) |> Seq.sort |> Seq.toArray
            let mutable start = sorted.[0]
            for i in 1 .. sorted.Length do
                if i = sorted.Length || sorted.[i] <> sorted.[i - 1] + 1 then
                    runs.Add((col, start, sorted.[i - 1]))
                    if i < sorted.Length then start <- sorted.[i]
        let rectangles = ResizeArray<CellRange>()
        for (col, first, last) in runs do
            let joined =
                rectangles |> Seq.tryFindIndex (fun r -> r.To.Col = col - 1 && r.From.Row = first && r.To.Row = last)
            match joined with
            | Some index -> rectangles.[index] <- { rectangles.[index] with To = { Row = last; Col = col } }
            | None -> rectangles.Add { From = { Row = first; Col = col }; To = { Row = last; Col = col } }
        yield sqrefText (List.ofSeq rectangles), ooxmlOfValidation validation ]
