module RedBat.Spreadsheet.Paint

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Layout

type Color = string

type TextAlign =
    | AlignLeft
    | AlignCenter
    | AlignRight

type VerticalAlign =
    | VAlignTop
    | VAlignMiddle
    | VAlignBottom

type PaintText = {
    Text: string
    /// The full cell (or header) rectangle. Per-line offset is derived from
    /// LineIndex/LineCount and font metrics in the renderer — Paint.fs does
    /// not pack baseline math into the rect.
    Rect: Rect
    Color: Color
    /// Font size in points, as written in the source XLSX. Each renderer
    /// converts to its native unit (px for raster, pt for PDF, css-pt for DOM).
    FontSize: float
    FontName: string option
    Bold: bool
    Italic: bool
    /// Render a horizontal stroke under the text (XLSX `<u/>` font flag).
    Underline: bool
    /// Render a horizontal stroke through the text (XLSX `<strike/>` flag).
    Strikethrough: bool
    Align: TextAlign
    /// Vertical alignment within Rect. Excel "General" = VAlignBottom.
    VerticalAlign: VerticalAlign
    /// 0-based line index within a wrapped block; 0 if not wrapped.
    LineIndex: int
    /// Total lines in the wrapped block; 1 if not wrapped.
    LineCount: int
    /// Signed pixel offset the renderer adds to the computed baseline.
    /// Used for superscript (negative — shift up) and subscript (positive
    /// — shift down). Zero for normal text.
    BaselineShift: float
}

type PaintItem =
    | FillRect of Rect * Color
    | StrokeRect of Rect * Color * float
    | DrawText of PaintText
    | GridLine of float * float * float * float * Color * float

type ViewportPaintItem = {
    Zone: ViewportZone
    Item: PaintItem
}

/// Excel's General horizontal alignment puts numbers/booleans on the right,
/// strings on the left, and errors in the center. Explicit alignment on
/// CellFormat overrides this default.
let rec private generalAlign (value: CellValue): TextAlign =
    match value with
    | Number _ -> AlignRight
    | Boolean _ -> AlignCenter
    | Error _ -> AlignCenter
    | Formula (_, cached) -> generalAlign cached
    | CellValue.Empty -> AlignLeft
    | _ -> AlignLeft

let private alignFromFormat (format: CellFormat) (value: CellValue): TextAlign =
    match format.HorizontalAlign with
    | Some "center" -> AlignCenter
    | Some "right" -> AlignRight
    | Some "left" -> AlignLeft
    | _ -> generalAlign value

/// Interpret common Excel date/time format codes (m/d/yyyy, h:mm, etc.) and
/// render the Excel serial number as a date string. Excel serial 1 ==
/// 1900-01-01; the 1900 leap-year bug means dates after 1900-02-28 need an
/// extra day subtracted, but for typical post-1901 fixtures the simple
/// `baseDate.AddDays(serial - 2)` covers it.
/// Excel format codes use `;` to separate positive/negative/zero/text
/// sections (in that order). For date/time interpretation we only need the
/// positive section, but real numeric values need the right section picked
/// by sign so a format like `"TRUE";"TRUE";"FALSE"` renders 0 as FALSE.
/// Splitter respects quoted segments and backslash escapes.
let private splitSections (code: string) : string list =
    let sections = System.Collections.Generic.List<string>()
    let sb = System.Text.StringBuilder()
    let mutable i = 0
    let mutable inQuote = false
    while i < code.Length do
        let c = code.[i]
        if c = '\\' && i + 1 < code.Length then
            sb.Append(c).Append(code.[i + 1]) |> ignore
            i <- i + 2
        elif c = '"' then
            inQuote <- not inQuote
            sb.Append(c) |> ignore
            i <- i + 1
        elif c = ';' && not inQuote then
            sections.Add(sb.ToString())
            sb.Clear() |> ignore
            i <- i + 1
        else
            sb.Append(c) |> ignore
            i <- i + 1
    sections.Add(sb.ToString())
    List.ofSeq sections

let private firstSection (code: string) =
    match splitSections code with
    | first :: _ -> first
    | [] -> code

let private sectionForValue (code: string) (n: float) : string =
    let sections = splitSections code
    let pick =
        if n > 0.0 then sections |> List.tryHead
        elif n < 0.0 then sections |> List.skip (min 1 (sections.Length - 1)) |> List.tryHead
        else sections |> List.skip (min 2 (sections.Length - 1)) |> List.tryHead
    pick |> Option.defaultValue (List.head sections)

let private isDateFormat (code: string) =
    let lower = (firstSection code).ToLowerInvariant()
    lower.Contains "y" || lower.Contains "d" || lower.Contains "m/" || lower.Contains "mmm"

let private isTimeFormat (code: string) =
    let lower = (firstSection code).ToLowerInvariant()
    lower.Contains "h" || (lower.Contains "s" && lower.Contains "m")

let private excelSerialToDate (serial: float) : System.DateTime =
    // Excel's 1900 calendar: serial 1 is 1900-01-01 but Excel also treats
    // 1900 as a leap year (it isn't). For serials >= 60 we subtract one to
    // line up with the real Gregorian calendar.
    let baseDate = System.DateTime(1899, 12, 30)
    let days = if serial >= 60.0 then serial else serial + 1.0
    baseDate.AddDays(days)

/// Tokenize an Excel format code into .NET DateTime.ToString format. Tracks
/// context (after `h`/`s` → minute; after month-only marker → month) for the
/// shared `m` token. Recognizes fractional-second suffixes `.0`/`.00`/`.000`
/// (Excel) and emits the .NET equivalents `.f`/`.ff`/`.fff`. Quoted/escaped
/// literals are emitted unchanged. Unrecognized tokens are wrapped in single
/// quotes so .NET treats them as literals rather than throwing.
let private excelToNetFormat (code: string) : string =
    let lower = (firstSection code).ToLowerInvariant()
    let sb = System.Text.StringBuilder()
    let mutable i = 0
    // mInTimeContext is true after we just emitted h/H or are about to emit
    // s/S — Excel's `m` token resolves to minute in those positions.
    let mutable mInTimeContext = false
    let has12Hour = lower.Contains "am/pm" || lower.Contains "a/p"

    let runOf c (s: string) =
        let mutable n = 0
        while i + n < s.Length && s.[i + n] = c do n <- n + 1
        n

    while i < lower.Length do
        match lower.[i] with
        | 'y' ->
            let n = runOf 'y' lower
            sb.Append(if n >= 4 then "yyyy" elif n >= 2 then "yy" else "yyyy") |> ignore
            i <- i + n
            mInTimeContext <- false
        | 'd' ->
            let n = runOf 'd' lower
            sb.Append(System.String('d', min 4 n)) |> ignore
            i <- i + n
            mInTimeContext <- false
        | 'h' ->
            let n = runOf 'h' lower
            let token = if has12Hour then (if n >= 2 then "hh" else "h") else (if n >= 2 then "HH" else "H")
            sb.Append(token: string) |> ignore
            i <- i + n
            mInTimeContext <- true
        | 's' ->
            let n = runOf 's' lower
            sb.Append(System.String('s', min 2 n)) |> ignore
            i <- i + n
            mInTimeContext <- true
        | 'm' ->
            let n = runOf 'm' lower
            // 3+ m's are always month (mmm/mmmm). 1-2 m's: minute if adjacent
            // to a time token, otherwise month.
            let isMinute =
                n <= 2 && (
                    mInTimeContext ||
                    (let j = i + n
                     j < lower.Length && (lower.[j] = ':' || lower.[j] = 'h' || lower.[j] = 's')))
            if isMinute then
                sb.Append(System.String('m', min 2 n)) |> ignore
                mInTimeContext <- true
            else
                let upper =
                    match n with
                    | 1 -> "M"
                    | 2 -> "MM"
                    | 3 -> "MMM"
                    | 4 -> "MMMM"
                    | _ -> "MMMMM"
                sb.Append(upper) |> ignore
                mInTimeContext <- false
            i <- i + n
        | 'a' when i + 4 < lower.Length && lower.Substring(i, 5) = "am/pm" ->
            sb.Append("tt") |> ignore
            i <- i + 5
        | 'a' when i + 2 < lower.Length && lower.Substring(i, 3) = "a/p" ->
            sb.Append("t") |> ignore
            i <- i + 3
        | '.' when i + 1 < lower.Length && lower.[i + 1] = '0' ->
            let mutable n = 0
            while i + 1 + n < lower.Length && lower.[i + 1 + n] = '0' do n <- n + 1
            sb.Append('.') |> ignore
            sb.Append(System.String('f', min 7 n)) |> ignore
            i <- i + 1 + n
        | '\\' when i + 1 < lower.Length ->
            // Escaped literal char.
            sb.Append('\'').Append(lower.[i + 1]).Append('\'') |> ignore
            i <- i + 2
        | '"' ->
            // Quoted literal.
            let endQuote = lower.IndexOf('"', i + 1)
            if endQuote < 0 then i <- lower.Length
            else
                sb.Append('\'').Append(lower.Substring(i + 1, endQuote - i - 1)).Append('\'') |> ignore
                i <- endQuote + 1
        | '[' ->
            // Elapsed-time bracket [h]/[m]/[s] — .NET can't express elapsed
            // hours past 24 with a stock format; emit a normal token and
            // accept the wraparound for now.
            let endBracket = lower.IndexOf(']', i + 1)
            if endBracket < 0 then i <- i + 1
            else i <- endBracket + 1
        | ':' | '/' | '-' | ' ' | ',' as c ->
            sb.Append(c) |> ignore
            i <- i + 1
        | c ->
            // Unrecognized: wrap as single-quote literal so .NET doesn't
            // throw on it.
            sb.Append('\'').Append(c).Append('\'') |> ignore
            i <- i + 1
    sb.ToString()

/// Excel rounds sub-second time values to the precision implied by the
/// format's `.0+` suffix. .NET's `fff` truncates instead, so we pre-round
/// to that precision before formatting. Returns -1 when no fractional
/// suffix is present (no rounding needed).
let private fractionalSecondPrecision (code: string) : int =
    let lower = (firstSection code).ToLowerInvariant()
    let mutable i = 0
    let mutable maxRun = -1
    while i < lower.Length do
        if lower.[i] = '.' && i + 1 < lower.Length && lower.[i + 1] = '0' then
            let mutable n = 0
            while i + 1 + n < lower.Length && lower.[i + 1 + n] = '0' do n <- n + 1
            if n > maxRun then maxRun <- n
            i <- i + 1 + n
        else
            i <- i + 1
    maxRun

let private formatExcelDate (code: string) (serial: float) : string =
    // Excel allows huge serials (the "testdateoverflow" fixture has values
    // that exceed DateTime.MaxValue). Catch and fall back to the raw float
    // so we never crash a render on a single bad cell.
    let dt =
        try Some (excelSerialToDate serial)
        with _ -> None
    let inv = System.Globalization.CultureInfo.InvariantCulture
    match dt with
    | None -> serial.ToString("G", inv)
    | Some dt ->
        let dtRounded =
            let prec = fractionalSecondPrecision code
            if prec < 0 then dt
            else
                let unitsPerSecond = pown 10L prec
                let ticksPerUnit = System.TimeSpan.TicksPerSecond / unitsPerSecond
                if ticksPerUnit <= 0L then dt
                else
                    let rounded = (dt.Ticks + ticksPerUnit / 2L) / ticksPerUnit * ticksPerUnit
                    try System.DateTime(rounded) with _ -> dt
        let netFormat = excelToNetFormat code
        try dtRounded.ToString(netFormat, inv)
        with _ -> dtRounded.ToString("MM/dd/yyyy", inv)

/// Approximate rendered Maximum Digit Width (MDW) in CSS pixels for a given
/// XLSX font size (in points). Excel's MDW table is flat around the common
/// 10/11/12 pt sizes (all 7 px), then jumps to 8 px at 12-13 pt. Use the
/// `round(fontPt × 7/11)` formula but floor it at 7 — the value Excel uses
/// for every Calibri / Arial size from 8 pt up through 11 pt. Without the
/// floor, size-10 cells (Arial 10pt, the Excel 2003 legacy default in many
/// pandas-test_datetime_mi-style POI fixtures) under-count chars and trip
/// ##### overflow on values that Excel renders inside the cell.
let private mdwPx (fontPt: float) : float =
    max 7.0 (System.Math.Round(fontPt * 7.0 / 11.0))

/// Width-aware Excel "General" number formatting. Picks the longest .NET
/// "G<n>" representation that fits the cell width. If even the shortest
/// representation overflows, returns `#####` like Excel does.
let private generalNumberDisplay (cellWidthPx: float) (fontPt: float) (n: float): string =
    // Leave 4 px padding on each side (matches the cell-text inset in
    // buildCellPaintItems) plus one MDW char of additional right buffer
    // that Excel's General format reserves. Without the extra char, a
    // column sized against the workbook's default-font MDW (e.g. Calibri
    // 12 at 8 px) ends up "fitting" 11 chars of Calibri-11 text (mdw=7),
    // since column-width math and General-format math use different
    // fonts. The extra char keeps the trim aligned with what Excel shows
    // in pandas-test1's Calibri-12-default sheet (B2 = 0.98026851, not
    // 0.9802685138). Try the requested precision *and* the standard set
    // so we surface G9 mid-trim instead of jumping from G10 straight to
    // G8 — Excel uses 9 sig digits in some columns.
    let usableWidth = max 0.0 (cellWidthPx - 8.0 - mdwPx fontPt)
    let maxChars = max 1 (int (usableWidth / mdwPx fontPt))

    let inv = System.Globalization.CultureInfo.InvariantCulture
    let candidates = [ "G11"; "G10"; "G9"; "G8"; "G7"; "G6"; "G5"; "G4"; "G3"; "G2"; "G1" ]
    let fits =
        candidates
        |> List.tryPick (fun fmt ->
            let s = n.ToString(fmt, inv)
            if s.Length <= maxChars then Some s else None)
    match fits with
    | Some s -> s
    | None -> System.String('#', maxChars)

/// Render an Excel fraction format like `# ?/?` (one-digit denominator),
/// `# ??/??` (two), `# ???/???` (three). Splits the number into its
/// integer part and the closest reduced fraction with denominator ≤
/// 10^digits, mirroring Excel's display. Returns None if the code isn't
/// a recognized fraction pattern.
let private tryFormatFraction (code: string) (n: float) : string option =
    // Match `<intPart>?<spaces>?/?` style: any number of ?'s on both
    // sides of the slash. The integer part can be empty (no whole
    // number shown) or "#"/"0" patterns we strip into "show int if
    // non-zero".
    let pattern = System.Text.RegularExpressions.Regex(@"^(?<int>[#0]*)\s*(?<num>\?+)/(?<den>\?+)$")
    let m = pattern.Match code
    if not m.Success then None else
        let denDigits = m.Groups.["den"].Length
        let maxDen = pown 10 denDigits - 1
        let sign = if n < 0.0 then "-" else ""
        let abs = System.Math.Abs n
        let whole = int (System.Math.Floor abs)
        let frac = abs - float whole
        // Stern-Brocot search for the best rational with denominator
        // ≤ maxDen. Exact enough for Excel's display.
        let mutable bestNum, bestDen = 0, 1
        let mutable bestErr = frac
        for d in 1 .. maxDen do
            let nn = int (System.Math.Round (frac * float d))
            let err = System.Math.Abs (frac - float nn / float d)
            if err < bestErr then
                bestErr <- err
                bestNum <- nn
                bestDen <- d
        // Reduce.
        let rec gcd a b = if b = 0 then a else gcd b (a % b)
        let g = gcd bestNum bestDen |> max 1
        let nn, dd = bestNum / g, bestDen / g
        let result =
            if nn = 0 && whole = 0 then "0"
            elif nn = 0 then sprintf "%s%d" sign whole
            elif whole = 0 then sprintf "%s%d/%d" sign nn dd
            else sprintf "%s%d %d/%d" sign whole nn dd
        Some result

/// Strip Excel's column-width placeholders that have no .NET equivalent.
/// Excel uses `_X` to mean "skip a column's worth of width = width(X)"
/// (typically used as `_(` / `_)` to align with paren-bracketed negatives)
/// and `*X` to mean "fill with character X" (typically `* ` to right-
/// justify a "-" placeholder for zero). Neither has a placement in
/// monospaced PNG/PDF output, so collapse `_X` to a single space and
/// drop the `*` filler entirely (rendering the next char once).
///
/// Also normalize Excel's `?` digit-or-space placeholder. .NET has no
/// direct equivalent — `?` looks like a literal in custom-format-string
/// land, which would print as `??` rather than space-padding. The
/// pragmatic mapping is `?` → `#` (digit or nothing): produces the
/// same visual for non-zero values and renders an empty string for
/// zero, which is what Excel's accounting "-" section needs anyway.
/// Skip `?` inside quoted literal segments.
let private stripExcelWidthPlaceholders (code: string) : string =
    let sb = System.Text.StringBuilder(code.Length)
    let mutable i = 0
    let mutable inQuote = false
    while i < code.Length do
        let c = code.[i]
        if c = '"' then
            sb.Append(c) |> ignore
            inQuote <- not inQuote
            i <- i + 1
        elif inQuote then
            sb.Append(c) |> ignore
            i <- i + 1
        elif (c = '_' || c = '*') && i + 1 < code.Length then
            // Render `_X` as " " (single-space alignment) and `*X` as
            // X alone. Excel proper uses these to fill column width;
            // we don't have that concept so the minimal substitute is
            // a single occurrence of the target char.
            if c = '_' then sb.Append(' ') |> ignore
            else sb.Append(code.[i + 1]) |> ignore
            i <- i + 2
        elif c = '?' then
            sb.Append('#') |> ignore
            i <- i + 1
        else
            sb.Append(c) |> ignore
            i <- i + 1
    sb.ToString()

/// Apply an Excel non-date number-format code (`0.000`, `#,##0.00`, `0%`,
/// `0.00E+00`, etc.) using .NET's custom-format-string support. Excel's
/// number-format grammar is largely compatible with .NET's; we strip the
/// section separator first and pass the positive section to `Double.ToString`,
/// falling back to `G` if .NET rejects it.
let private formatExcelNumber (code: string) (n: float) : string =
    let inv = System.Globalization.CultureInfo.InvariantCulture
    let sections = splitSections code
    let pos = sectionForValue code n
    // Excel's section grammar already encodes the sign visually (e.g.
    // `(#,##0.00)` for negatives), so the picked section should format
    // against the *absolute* value — otherwise .NET prepends an extra
    // `-` to a value that's already wrapped in parens. The same goes
    // for the zero section, which usually has a literal "-" placeholder.
    let valueForSection =
        if (n < 0.0 && sections.Length >= 2) || (n = 0.0 && sections.Length >= 3) then
            System.Math.Abs n
        else n
    // Fractions (`# ?/?`, `# ??/??`, `?/??`) have no .NET equivalent and
    // need explicit numerator-denominator search. Try that first.
    match tryFormatFraction pos n with
    | Some s -> s
    | None ->
    // Translate Excel's escape grammar to .NET's. Excel uses `"..."` and
    // `\x` for literal segments; .NET wants `'...'` and `\x` (\x already
    // works). We rewrite double-quoted runs to single-quoted runs so that
    // a format like `"TRUE"` (literal text, no digit placeholders) renders
    // as the literal `TRUE` instead of throwing.
    let withoutPlaceholders = stripExcelWidthPlaceholders pos
    let cleaned =
        let sb = System.Text.StringBuilder(withoutPlaceholders.Length)
        let mutable i = 0
        let mutable inDquote = false
        while i < withoutPlaceholders.Length do
            let c = withoutPlaceholders.[i]
            if c = '\\' && i + 1 < withoutPlaceholders.Length then
                sb.Append('\'').Append(withoutPlaceholders.[i + 1]).Append('\'') |> ignore
                i <- i + 2
            elif c = '"' then
                sb.Append('\'') |> ignore
                inDquote <- not inDquote
                i <- i + 1
            else
                sb.Append(c) |> ignore
                i <- i + 1
        sb.ToString()
    // Strip any `[$-XXX]` locale tag — these tell Excel which locale's
    // date/month names to use; we keep the format string as-is and let
    // the OS culture default decide. Also drop `[Red]` / color tags;
    // section-based color routing isn't wired and the literal `[Red]`
    // would otherwise become part of the output.
    let cleaned =
        System.Text.RegularExpressions.Regex.Replace(
            cleaned, @"\[(?:\$[^\]]*|Red|Black|Green|White|Blue|Magenta|Yellow|Cyan)\]", "")
    try valueForSection.ToString(cleaned, inv)
    with _ ->
        try n.ToString("G", inv)
        with _ -> string n

/// Greedy word-wrap inside a cell. Splits on hard line breaks (\n) first,
/// then on whitespace within each segment, packing as many words onto a line
/// as the cell's usable width allows. Falls back to char-level breaks when
/// a single word exceeds the width.
///
/// When `measure` is Some, line-width budget is computed via the real loaded
/// font (matches Excel's glyph-aware wrapping). When None — Fable web canvas
/// or other contexts without a measurer — fall back to `mdwPx × charCount`,
/// which over-estimates width and wraps earlier than Excel.
let private wrapLines
    (measure: (string -> float) option)
    (cellWidthPx: float) (fontPt: float)
    (text: string) : string list =
    if System.String.IsNullOrEmpty text then [ text ]
    else
        let usable = max 0.0 (cellWidthPx - 8.0)
        // Approximation budget for the measurer-less path.
        let mdw = mdwPx fontPt
        let approxMaxChars = max 1 (int (usable / mdw))
        let fits (s: string) =
            match measure with
            | Some m -> m s <= usable
            | None -> s.Length <= approxMaxChars
        let wrapSegment (segment: string) : string list =
            let words = segment.Split(' ')
            let result = System.Collections.Generic.List<string>()
            let current = System.Text.StringBuilder()
            let breakOversizedWord (word: string) =
                // Char-by-char until we'd overflow; emit each chunk as its
                // own line. Used when a single word is wider than the cell.
                let mutable buf = System.Text.StringBuilder()
                for ch in word do
                    buf.Append(ch) |> ignore
                    if not (fits (buf.ToString())) && buf.Length > 1 then
                        // Last char overflowed; emit buffer minus last char,
                        // then start a new buffer with just that char.
                        let s = buf.ToString()
                        result.Add(s.Substring(0, s.Length - 1))
                        buf <- System.Text.StringBuilder()
                        buf.Append(ch) |> ignore
                buf
            for word in words do
                if current.Length = 0 then
                    if not (fits word) then
                        // word longer than cell — char-break, leaving the
                        // last chunk in `current` to continue with next word.
                        let leftover = breakOversizedWord word
                        if leftover.Length > 0 then current.Append(leftover) |> ignore
                    else current.Append(word) |> ignore
                else
                    let candidate = current.ToString() + " " + word
                    if fits candidate then
                        current.Append(' ').Append(word) |> ignore
                    else
                        result.Add(current.ToString())
                        current.Clear() |> ignore
                        if not (fits word) then
                            let leftover = breakOversizedWord word
                            if leftover.Length > 0 then current.Append(leftover) |> ignore
                        else current.Append(word) |> ignore
            if current.Length > 0 then result.Add(current.ToString())
            List.ofSeq result
        text.Split('\n')
        |> Array.toList
        |> List.collect wrapSegment

/// Excel shows ##### when a number's display string is wider than the cell.
let private fitOrOverflow (cellWidth: float) (fontPt: float) (text: string) : string =
    let usableWidth = max 0.0 (cellWidth - 8.0)
    let maxChars = max 1 (int (usableWidth / mdwPx fontPt))
    if text.Length <= maxChars then text
    else System.String('#', maxChars)

let rec private cellDisplayText (cellWidth: float) (format: CellFormat) (value: CellValue): string =
    let fontPt = format.FontSize |> Option.map float |> Option.defaultValue 11.0
    match value with
    | CellValue.Empty -> ""
    | Number n ->
        match format.NumberFormat with
        | Some code when isDateFormat code || isTimeFormat code ->
            fitOrOverflow cellWidth fontPt (formatExcelDate code n)
        | Some "currency"
        | Some "percentage"
        | Some "date"
        | Some "decimal2"
        | Some "number" ->
            // Legacy named format codes from the Fable web UI's formatting
            // commands. Keep these on the original Model.formatNumber path.
            fitOrOverflow cellWidth fontPt (formatNumber format.NumberFormat n)
        | Some code when System.String.IsNullOrWhiteSpace code ||
                         code.Equals("general", System.StringComparison.OrdinalIgnoreCase) ->
            generalNumberDisplay cellWidth fontPt n
        | Some code ->
            // Real Excel format code (0.000, #,##0.00, 0%, 0.00E+00, etc.).
            fitOrOverflow cellWidth fontPt (formatExcelNumber code n)
        | None -> generalNumberDisplay cellWidth fontPt n
    | String s -> s
    | Boolean b -> if b then "TRUE" else "FALSE"
    | Error e -> sprintf "#%s" e
    | Formula (_, cached) -> cellDisplayText cellWidth format cached

let private isSelected (selection: Selection) row col =
    match selection with
    | SingleCell (r, c) -> r = row && c = col
    | Range (sr, sc, er, ec) ->
        row >= sr && row <= er && col >= sc && col <= ec

let private buildHeaderPaintItems (model: SpreadsheetModel) (viewport: Viewport): PaintItem list =
    [
        for header in headers model viewport do
            match header with
            | CornerHeader rect ->
                yield FillRect (rect, "#e8e8e8")
                yield StrokeRect (rect, "#a6a6a6", 1.0)
            | ColumnHeader (col, rect) ->
                yield FillRect (rect, "#f2f2f2")
                yield StrokeRect (rect, "#a6a6a6", 1.0)
                yield DrawText {
                    Text = columnToLetter col
                    Rect = rect
                    Color = "#444444"
                    FontSize = 11.0
                    FontName = None
                    Bold = false
                    Italic = false
                    Underline = false
                    Strikethrough = false
                    Align = AlignCenter
                    VerticalAlign = VAlignMiddle
                    LineIndex = 0
                    LineCount = 1
                    BaselineShift = 0.0
                }
            | RowHeader (row, rect) ->
                yield FillRect (rect, "#f2f2f2")
                yield StrokeRect (rect, "#a6a6a6", 1.0)
                yield DrawText {
                    Text = string (row + 1)
                    Rect = rect
                    Color = "#444444"
                    FontSize = 11.0
                    FontName = None
                    Bold = false
                    Italic = false
                    Underline = false
                    Strikethrough = false
                    Align = AlignCenter
                    VerticalAlign = VAlignMiddle
                    LineIndex = 0
                    LineCount = 1
                    BaselineShift = 0.0
                }
    ]

/// Translate an OOXML border style name into an approximate stroke width in
/// CSS pixels. Excel has many variants (hair, thin, medium, thick, double,
/// dashed/dotted in three weights, slantDashDot, etc.); we collapse them into
/// thin/medium/thick stroke widths and a default black color.
let private borderStrokeWidth (style: string) : float =
    match style.ToLowerInvariant() with
    | "hair" -> 0.5
    | "thin" | "dashed" | "dotted" | "dashdot" | "dashdotdot" -> 1.0
    | "medium" | "mediumdashed" | "mediumdotted" | "mediumdashdot" | "mediumdashdotdot" -> 1.5
    | "thick" -> 2.5
    | "double" -> 2.0
    | "slantdashdot" -> 1.5
    | _ -> 1.0

let private borderColor (side: BorderStyle) : string =
    side.Color |> Option.defaultValue "#000000"

/// Multi-pass cell rendering. Drawing fill+stroke per cell in a single loop
/// lets the next cell's fill overwrite the previous cell's right/bottom edge
/// — invisible against white but obvious when a cell has a non-default fill
/// (Table styles, fills baked into cellXf). Splitting into three passes
/// (fills → strokes/borders → text) keeps all borders on top of all fills.
///
/// Items are tagged with the ViewportZone of the cell they came from so the
/// interactive renderer can scroll the body without sliding frozen
/// rows/cols. The static (PNG/PDF) renderer ignores the tag — at
/// ScrollX/Y=0 every zone paints at its absolute position anyway, so the
/// only visible effect of frozen panes there is the divider lines emitted
/// below.
let private buildCellPaintItems (model: SpreadsheetModel) (viewport: Viewport): (ViewportZone * PaintItem) list =
    let cells = visibleCells model viewport
    let cellState (cell: _) =
        let addr = cell.Address
        let cellData = getCell addr model
        let value = getCellValue addr model
        let format = cellData |> Option.map (fun c -> c.Format) |> Option.defaultValue defaultFormat
        let selected = isSelected model.Selection addr.Row addr.Col
        let background =
            if selected then "#E8F5E9"
            else format.BgColor |> Option.defaultValue "#ffffff"
        let zone = zoneForCell model addr.Row addr.Col
        cell, format, value, selected, background, zone

    let states = cells |> List.map cellState

    [
        // Pass 1: all backgrounds. Adjacent cells share an edge, so when a
        // cell with a non-default fill is followed by another cell, the
        // second cell's fill would overdraw the first cell's right/bottom
        // stroke if we drew them in the same per-cell loop.
        for (cell, _format, _value, _selected, background, zone) in states do
            yield zone, FillRect (cell.Rect, background)

        // Pass 2: per-cell strokes + cellXf borders, drawn on top of every
        // fill. Match Excel's *interactive* view: a light gray gridline
        // around every cell — not the near-black grid that print mode emits
        // with PrintGridlines on. Our reference PNGs come from print, but the
        // renderer's target is the on-screen view.
        // Emit a styled border side as one (single solid) or two (double)
        // GridLine items. "double" in Excel is two ~0.75 px parallel rules
        // straddling the cell edge with a small gap; emit two 1 px strokes
        // offset by ±1 px from the edge so the gap is visible at 96 dpi.
        // Other styles flow through the existing single-stroke path with
        // the borderStrokeWidth weight.
        let emitBorderSide (isHorizontal: bool) (side: BorderStyle) (x1: float, y1: float) (x2: float, y2: float) =
            let color = borderColor side
            let w = borderStrokeWidth side.Style
            if side.Style.Equals("double", System.StringComparison.OrdinalIgnoreCase) then
                let off = 1.0
                if isHorizontal then
                    [ GridLine (x1, y1 - off, x2, y2 - off, color, 1.0)
                      GridLine (x1, y1 + off, x2, y2 + off, color, 1.0) ]
                else
                    [ GridLine (x1 - off, y1, x2 - off, y2, color, 1.0)
                      GridLine (x1 + off, y1, x2 + off, y2, color, 1.0) ]
            else
                [ GridLine (x1, y1, x2, y2, color, w) ]

        for (cell, format, _value, selected, _background, zone) in states do
            yield zone, StrokeRect (cell.Rect, (if selected then "#217346" else "#a6a6a6"), (if selected then 2.0 else 1.0))
            let rect = cell.Rect
            match format.BorderTop with
            | Some side ->
                for item in emitBorderSide true side (rect.X, rect.Y) (rect.X + rect.Width, rect.Y) do
                    yield zone, item
            | None -> ()
            match format.BorderBottom with
            | Some side ->
                let y = rect.Y + rect.Height
                for item in emitBorderSide true side (rect.X, y) (rect.X + rect.Width, y) do
                    yield zone, item
            | None -> ()
            match format.BorderLeft with
            | Some side ->
                for item in emitBorderSide false side (rect.X, rect.Y) (rect.X, rect.Y + rect.Height) do
                    yield zone, item
            | None -> ()
            match format.BorderRight with
            | Some side ->
                let x = rect.X + rect.Width
                for item in emitBorderSide false side (x, rect.Y) (x, rect.Y + rect.Height) do
                    yield zone, item
            | None -> ()

        // Pass 3: text on top of all strokes/borders.
        for (cell, format, value, _selected, _background, zone) in states do
            let displayText = cellDisplayText cell.Rect.Width format value
            if displayText <> "" then
                let fontPt = format.FontSize |> Option.map float |> Option.defaultValue 11.0
                let textRect = { cell.Rect with X = cell.Rect.X + 4.0; Width = max 0.0 (cell.Rect.Width - 8.0) }
                let color = format.TextColor |> Option.defaultValue "#000000"
                let align = alignFromFormat format value
                let hasHardBreak = displayText.Contains("\n")
                let measure =
                    viewport.Measurer
                    |> Option.map (fun m ->
                        fun (s: string) -> m s fontPt format.FontName format.Bold format.Italic)
                let lines =
                    if format.WrapText || hasHardBreak then
                        wrapLines measure textRect.Width fontPt displayText
                    else [ displayText ]
                let vAlign =
                    match format.VerticalAlign |> Option.map (fun s -> s.ToLowerInvariant()) with
                    | Some "top" -> VAlignTop
                    | Some "center" | Some "middle" -> VAlignMiddle
                    | _ -> VAlignBottom
                let lineCount = lines.Length
                let richRuns = Map.tryFind cell.Address model.RichRuns
                match richRuns, viewport.Measurer with
                | Some runs, Some m when not (List.isEmpty runs) ->
                    // Rich-text rendering. For single-line cells we emit one
                    // DrawText per run. For wrapped/multi-line cells we walk
                    // a position cursor through the concatenated source text
                    // and chunk runs that straddle line boundaries — each
                    // line gets per-run DrawTexts at its own y, with the
                    // breaking whitespace dropped by wrapLines absorbed into
                    // the cursor (so a run that ends with " " before a line
                    // break still aligns with the next line's start).
                    let resolveRun (ovr: RunOverride) =
                        let bold = ovr.Bold |> Option.defaultValue format.Bold
                        let italic = ovr.Italic |> Option.defaultValue format.Italic
                        let underline = ovr.Underline |> Option.defaultValue format.Underline
                        let strike = ovr.Strikethrough |> Option.defaultValue format.Strikethrough
                        let fName = ovr.FontName |> Option.orElse format.FontName
                        let baseSize = ovr.FontSize |> Option.defaultValue fontPt
                        let runColor = ovr.Color |> Option.defaultValue color
                        // vertAlign: shrink to ~60% of the cell font (Excel
                        // uses ~58% empirically, close enough) and shift
                        // the baseline. Negative shift = up (super); we
                        // express shift in pt at the cell's base size so
                        // the renderer's pt→px scaling stays consistent.
                        let fSize, baselineShift =
                            match ovr.VertAlign with
                            | Some VertAlignSuperscript -> baseSize * 0.6, -fontPt * 0.35
                            | Some VertAlignSubscript -> baseSize * 0.6, fontPt * 0.15
                            | _ -> baseSize, 0.0
                        bold, italic, underline, strike, fName, fSize, runColor, baselineShift
                    // Mutable list of remaining runs. We consume from the
                    // head; if a wrap break splits a run, the head is
                    // replaced with the unconsumed suffix.
                    let mutable remaining: (string * RunOverride) list = runs
                    let displayConcat = runs |> List.map fst |> String.concat ""
                    let mutable docPos = 0
                    for i, line in List.indexed lines do
                        let slices = System.Collections.Generic.List<string * RunOverride>()
                        let mutable lineRem = line.Length
                        while lineRem > 0 && not (List.isEmpty remaining) do
                            let (runText, runProps) = List.head remaining
                            let take = min lineRem runText.Length
                            slices.Add(runText.Substring(0, take), runProps)
                            lineRem <- lineRem - take
                            if take = runText.Length then
                                remaining <- List.tail remaining
                            else
                                // Suffix of the current run stays at head.
                                remaining <- (runText.Substring(take), runProps) :: List.tail remaining
                        docPos <- docPos + line.Length
                        // Skip the breaking space wrapLines drops between
                        // non-final lines. Also consume the matching space
                        // from the run queue head so positions stay aligned.
                        if i < lines.Length - 1
                           && docPos < displayConcat.Length
                           && displayConcat.[docPos] = ' ' then
                            docPos <- docPos + 1
                            match remaining with
                            | (rt, rp) :: rest when rt.StartsWith " " ->
                                if rt.Length > 1 then
                                    remaining <- (rt.Substring(1), rp) :: rest
                                else
                                    remaining <- rest
                            | _ -> ()
                        // Measure this line's segments and emit.
                        let metrics =
                            slices
                            |> Seq.map (fun (t, ovr) ->
                                let bold, italic, underline, strike, fName, fSize, runColor, baselineShift = resolveRun ovr
                                let width = m t fSize fName bold italic
                                t, ovr, bold, italic, underline, strike, fName, fSize, runColor, width, baselineShift)
                            |> Seq.toList
                        let totalWidth = metrics |> List.sumBy (fun (_, _, _, _, _, _, _, _, _, w, _) -> w)
                        let baseLineX =
                            match align with
                            | AlignLeft -> textRect.X
                            | AlignCenter -> textRect.X + (textRect.Width - totalWidth) / 2.0
                            | AlignRight -> textRect.X + textRect.Width - totalWidth
                        let mutable cursor = baseLineX
                        for (t, _ovr, bold, italic, underline, strike, fName, fSize, runColor, width, baselineShift) in metrics do
                            yield zone, DrawText {
                                Text = t
                                Rect = { textRect with X = cursor; Width = width }
                                Color = runColor
                                FontSize = fSize
                                FontName = fName
                                Bold = bold
                                Italic = italic
                                Underline = underline
                                Strikethrough = strike
                                Align = AlignLeft
                                VerticalAlign = vAlign
                                LineIndex = i
                                LineCount = lineCount
                                BaselineShift = baselineShift
                            }
                            cursor <- cursor + width
                | _ ->
                    // Plain (or wrapped) cell text — one DrawText per line.
                    for i, line in List.indexed lines do
                        yield zone, DrawText {
                            Text = line
                            Rect = textRect
                            Color = color
                            FontSize = fontPt
                            FontName = format.FontName
                            Bold = format.Bold
                            Italic = format.Italic
                            Underline = format.Underline
                            Strikethrough = format.Strikethrough
                            Align = align
                            VerticalAlign = vAlign
                            LineIndex = i
                            LineCount = lineCount
                            BaselineShift = 0.0
                        }

        // Pass 4: auto-filter dropdown triangles on table header cells. Excel
        // renders these as a small downward triangle on the right edge of each
        // header cell whose column falls within the table's <autoFilter>
        // range. The triangle is drawn over the header's accent fill — same
        // color as the header text (white on a Medium-style accent).
        for (cell, format, _value, _selected, _background, zone) in states do
            if format.FilterButton then
                let color = format.TextColor |> Option.defaultValue "#FFFFFF"
                // 7-px-wide, 4-px-tall triangle centered in a 16-px right-edge
                // button area. Drawn as five horizontal stroked lines (no
                // primitive for filled polygons; this is cheap enough).
                let buttonWidth = 16.0
                let centerX = cell.Rect.X + cell.Rect.Width - buttonWidth / 2.0
                let triHalfWidth = 3.5
                let triTopY = cell.Rect.Y + cell.Rect.Height / 2.0 - 2.0
                for i in 0 .. 3 do
                    let shrink = float i * (triHalfWidth / 4.0)
                    let y = triTopY + float i
                    yield zone, GridLine (centerX - triHalfWidth + shrink, y, centerX + triHalfWidth - shrink, y, color, 1.0)

        // Pass 5: comment-indicator red triangle in the top-right corner.
        // Excel uses a 5-px right-triangle filled with #FF0000 — we
        // approximate with horizontal strokes of decreasing width since
        // the renderer has no fill-polygon primitive. The triangle's
        // hypotenuse goes from (right-5, top) down to (right, top+5).
        for (cell, format, _value, _selected, _background, zone) in states do
            if format.HasComment then
                let color = "#C00000"  // Excel's "note marker" red
                let triSize = 5.0
                let rightX = cell.Rect.X + cell.Rect.Width
                let topY = cell.Rect.Y
                // Row k from top: from x = rightX - (triSize - k) to x = rightX.
                for k in 0 .. int triSize - 1 do
                    let leftX = rightX - (triSize - float k)
                    yield zone, GridLine (leftX, topY + float k + 0.5, rightX, topY + float k + 0.5, color, 1.0)

        // Pass 6: data-validation list dropdown arrow. Excel only shows
        // this when the cell is selected, but for static export we
        // always paint it so the visual cue carries through to PNG /
        // PDF snapshots. Drawn as a small 7-px-wide triangle just
        // inside the cell's right edge, vertically centered, in gray.
        // FilterButton already owns the dropdown styling for table
        // header cells; skip those to avoid doubled triangles.
        for (cell, format, _value, _selected, _background, zone) in states do
            if format.HasListDropdown && not format.FilterButton then
                let color = "#666666"
                let buttonWidth = 16.0
                let centerX = cell.Rect.X + cell.Rect.Width - buttonWidth / 2.0
                let triHalfWidth = 3.5
                let triTopY = cell.Rect.Y + cell.Rect.Height / 2.0 - 2.0
                for i in 0 .. 3 do
                    let shrink = float i * (triHalfWidth / 4.0)
                    let y = triTopY + float i
                    yield zone, GridLine (centerX - triHalfWidth + shrink, y, centerX + triHalfWidth - shrink, y, color, 1.0)
    ]

/// Freeze-divider strokes drawn on top of the cell layer. Excel renders
/// this as a single ~1 px line slightly darker than the regular gridline,
/// at the split's right (cols) / bottom (rows) edge. Matches the line
/// Excel draws in both interactive and print-snapshot views. Coordinates
/// are in the same space as cellRect (post-HeaderHeight / -RowHeaderWidth
/// offset) so the line lands flush with the bottom/right of the last
/// frozen row/column.
let private buildFreezeDividerItems (model: SpreadsheetModel) (viewport: Viewport) : PaintItem list =
    // Slightly darker than the regular gridline. Excel uses a fixed
    // medium-gray (~#595959 in print mode).
    let color = "#595959"
    let w = 1.0
    [
        if model.FrozenRows > 0 then
            let y =
                viewport.HeaderHeight
                + ([ 0 .. model.FrozenRows - 1 ] |> List.sumBy (fun r -> getRowHeight r model))
            let totalWidth =
                viewport.RowHeaderWidth
                + ([ 0 .. model.ColCount - 1 ] |> List.sumBy (fun c -> getColumnWidth c model))
            yield GridLine (viewport.RowHeaderWidth, y, totalWidth, y, color, w)
        if model.FrozenCols > 0 then
            let x =
                viewport.RowHeaderWidth
                + ([ 0 .. model.FrozenCols - 1 ] |> List.sumBy (fun c -> getColumnWidth c model))
            let totalHeight =
                viewport.HeaderHeight
                + ([ 0 .. model.RowCount - 1 ] |> List.sumBy (fun r -> getRowHeight r model))
            yield GridLine (x, viewport.HeaderHeight, x, totalHeight, color, w)
    ]

let buildPaintItems (model: SpreadsheetModel) (viewport: Viewport): PaintItem list =
    [
        yield FillRect ({ X = 0.0; Y = 0.0; Width = viewport.Width; Height = viewport.Height }, "#ffffff")
        yield! buildHeaderPaintItems model viewport
        // Static (PNG/PDF) path: drop the zone tag; everything paints at
        // absolute cell coordinates because ScrollX/Y are zero.
        for (_zone, item) in buildCellPaintItems model viewport do
            yield item
        // Freeze divider lines on top of cell strokes but under any
        // selection / DrawText emitted later. The cell-item zone tag is
        // irrelevant here; we draw the divider in the model coordinate
        // system that's already baked into the cell rects above.
        yield! buildFreezeDividerItems model viewport
    ]

let private translatePaintItem viewport zone item =
    let tx = toViewportRect viewport zone
    match item with
    | FillRect (rect, color) -> FillRect (tx rect, color)
    | StrokeRect (rect, color, width) -> StrokeRect (tx rect, color, width)
    | DrawText text -> DrawText { text with Rect = tx text.Rect }
    | GridLine (x1, y1, x2, y2, color, width) ->
        let p1 = tx { X = x1; Y = y1; Width = 0.0; Height = 0.0 }
        let p2 = tx { X = x2; Y = y2; Width = 0.0; Height = 0.0 }
        GridLine (p1.X, p1.Y, p2.X, p2.Y, color, width)

let buildViewportPaintItems (model: SpreadsheetModel) (viewport: Viewport): ViewportPaintItem list =
    [
        yield { Zone = CornerZone
                Item = FillRect ({ X = 0.0; Y = 0.0; Width = viewport.Width; Height = viewport.Height }, "#ffffff") }

        for header in headers model viewport do
            match header with
            | CornerHeader rect ->
                yield { Zone = CornerZone; Item = FillRect (rect, "#e8e8e8") }
                yield { Zone = CornerZone; Item = StrokeRect (rect, "#a6a6a6", 1.0) }
            | ColumnHeader (col, rect) ->
                let zone = ColumnHeaderZone
                yield { Zone = zone; Item = translatePaintItem viewport zone (FillRect (rect, "#f2f2f2")) }
                yield { Zone = zone; Item = translatePaintItem viewport zone (StrokeRect (rect, "#a6a6a6", 1.0)) }
                yield {
                    Zone = zone
                    Item =
                        translatePaintItem viewport zone (DrawText {
                            Text = columnToLetter col
                            Rect = rect
                            Color = "#444444"
                            FontSize = 11.0
                            FontName = None
                            Bold = false
                            Italic = false
                            Underline = false
                            Strikethrough = false
                            Align = AlignCenter
                            VerticalAlign = VAlignMiddle
                            LineIndex = 0
                            LineCount = 1
                            BaselineShift = 0.0
                        })
                }
            | RowHeader (row, rect) ->
                let zone = RowHeaderZone
                yield { Zone = zone; Item = translatePaintItem viewport zone (FillRect (rect, "#f2f2f2")) }
                yield { Zone = zone; Item = translatePaintItem viewport zone (StrokeRect (rect, "#a6a6a6", 1.0)) }
                yield {
                    Zone = zone
                    Item =
                        translatePaintItem viewport zone (DrawText {
                            Text = string (row + 1)
                            Rect = rect
                            Color = "#444444"
                            FontSize = 11.0
                            FontName = None
                            Bold = false
                            Italic = false
                            Underline = false
                            Strikethrough = false
                            Align = AlignCenter
                            VerticalAlign = VAlignMiddle
                            LineIndex = 0
                            LineCount = 1
                            BaselineShift = 0.0
                        })
                }

        for (zone, item) in buildCellPaintItems model viewport do
            yield { Zone = zone; Item = translatePaintItem viewport zone item }

        // Freeze dividers paint with the same zone as the body they border,
        // so they scroll with the body horizontally (row divider) or
        // vertically (col divider) and disappear under the frozen band when
        // the body scrolls past them. Excel does the same: the divider
        // moves with the scrolling region, not the frozen one.
        for item in buildFreezeDividerItems model viewport do
            let zone =
                match item with
                | GridLine (x1, _, x2, _, _, _) when x1 = x2 ->
                    // Vertical line (cols divider) — scrolls vertically with
                    // body but its X is pinned (no horizontal scroll).
                    FrozenRowsZone
                | _ ->
                    // Horizontal line (rows divider) — scrolls horizontally
                    // with body, pinned vertically.
                    FrozenColsZone
            yield { Zone = zone; Item = translatePaintItem viewport zone item }
    ]
