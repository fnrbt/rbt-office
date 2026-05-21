module RedBat.Spreadsheet.CellText

open RedBat.Spreadsheet.Model

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

/// A leading comparison tag like `[>=100]` or `[<0]` on a format section.
/// May sit before or after a colour tag (`[<0][Red]`, `[Red][>100]`). The
/// `[op N]` shape doesn't collide with `[h]`, `[$…]`, or `[Red]`, so matching
/// anywhere in the section is safe.
let private sectionCondition (section: string) : (string * float) option =
    let m =
        System.Text.RegularExpressions.Regex.Match(
            section, @"\[(<=|>=|<>|=|<|>)(-?[0-9]+(?:\.[0-9]+)?)\]")
    if m.Success then Some (m.Groups.[1].Value, float m.Groups.[2].Value) else None

let private conditionMatches (op: string) (threshold: float) (n: float) : bool =
    match op with
    | ">" -> n > threshold
    | ">=" -> n >= threshold
    | "<" -> n < threshold
    | "<=" -> n <= threshold
    | "=" -> n = threshold
    | "<>" -> n <> threshold
    | _ -> false

let private sectionForValue (code: string) (n: float) : string =
    let sections = splitSections code
    if sections |> List.exists (fun s -> (sectionCondition s).IsSome) then
        // Conditional format: pick the first section whose `[op N]` condition is
        // true; if none match, the last UNconditioned section is the default.
        match sections
              |> List.tryFind (fun s ->
                  match sectionCondition s with
                  | Some (op, th) -> conditionMatches op th n
                  | None -> false) with
        | Some s -> s
        | None ->
            sections |> List.rev |> List.tryFind (fun s -> (sectionCondition s).IsNone)
            |> Option.defaultValue (List.last sections)
    else
        // Excel sign grammar: [pos], [pos;neg], [pos;neg;zero], [pos;neg;zero;text].
        // With < 3 sections, zero falls back to the positive section — NOT the
        // negative one. With < 2 sections, negative also falls back to positive.
        let pick =
            if n > 0.0 then sections |> List.tryHead
            elif n < 0.0 then
                if sections.Length >= 2 then sections |> List.item 1 |> Some
                else sections |> List.tryHead
            else
                if sections.Length >= 3 then sections |> List.item 2 |> Some
                else sections |> List.tryHead
        pick |> Option.defaultValue (List.head sections)

/// Map an Excel format-section colour tag like `[Red]` to a hex string.
/// Returns None when the tag is absent or unrecognised. Excel only honors
/// the eight named colours plus `[Color N]` (palette index 1..56); we map
/// the eight named cases — enough for the common `#,##0.00;[Red]-...`
/// idiom that the parity pipeline cares about.
let private numberFormatSectionColor (section: string) : string option =
    // The colour tag may not be the FIRST bracket — a section can lead with a
    // condition (`[<0][Red]0.0`) or have the colour first (`[Red][>100]0`), so
    // match the colour tag anywhere among the leading tags, not just at ^.
    let m =
        System.Text.RegularExpressions.Regex.Match(
            section,
            @"\[(Red|Black|Green|White|Blue|Magenta|Yellow|Cyan)\]",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase)
    if not m.Success then None
    else
        match m.Groups.[1].Value.ToLowerInvariant() with
        | "red" -> Some "#FF0000"
        | "black" -> Some "#000000"
        | "green" -> Some "#008000"
        | "white" -> Some "#FFFFFF"
        | "blue" -> Some "#0000FF"
        | "magenta" -> Some "#FF00FF"
        | "yellow" -> Some "#FFFF00"
        | "cyan" -> Some "#00FFFF"
        | _ -> None

/// Pick the colour tag for the section of `code` that applies to value `n`.
/// Default (no tag) → None.
let numberFormatColorForValue (code: string) (n: float) : string option =
    numberFormatSectionColor (sectionForValue code n)

/// Strip `[$<currency>-<locale>]` prefixes before sniffing for date /
/// time tokens. Without this, the currency string `[$JPY-411]` triggers
/// `isDateFormat` because of the literal `y`, and the renderer dispatches
/// the value to the date formatter instead of the number formatter.
let private withoutLocaleTags (s: string) =
    System.Text.RegularExpressions.Regex.Replace(s, @"\[\$[^\]]*\]", "")

/// Strip everything that could carry a stray date/time letter without being a
/// real date/time token: quoted literals (`"High"`→h, `"Red"`→d), backslash
/// escapes (`\d`), and ALL bracket tags — locale `[$…]`, colour `[Red]`,
/// condition `[>=100]`, AND elapsed `[h]`/`[m]`/`[s]` (those route via
/// `hasElapsedBracket`, which is checked separately in the dispatch). What
/// remains is the bare format token stream we sniff for date/time chars.
let private dateTimeSniff (code: string) =
    let s = firstSection code
    let s = System.Text.RegularExpressions.Regex.Replace(s, "\"[^\"]*\"", "")
    let s = System.Text.RegularExpressions.Regex.Replace(s, @"\\.", "")
    let s = System.Text.RegularExpressions.Regex.Replace(s, @"\[[^\]]*\]", "")
    s.ToLowerInvariant()

let private isDateFormat (code: string) =
    let lower = dateTimeSniff code
    lower.Contains "y" || lower.Contains "d" || lower.Contains "m/" || lower.Contains "mmm"

let private isTimeFormat (code: string) =
    let lower = dateTimeSniff code
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
    // `lower` drives case-insensitive date/time TOKEN matching (y/m/d/h/s), but
    // quoted/escaped LITERALS must keep their original case — Excel's `"FY"yy`
    // renders "FY18", not "fy18". `orig` and `lower` share indices.
    let orig = firstSection code
    let lower = orig.ToLowerInvariant()
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
            sb.Append('\'').Append(orig.[i + 1]).Append('\'') |> ignore
            i <- i + 2
        | '"' ->
            // Quoted literal.
            let endQuote = lower.IndexOf('"', i + 1)
            if endQuote < 0 then i <- lower.Length
            else
                sb.Append('\'').Append(orig.Substring(i + 1, endQuote - i - 1)).Append('\'') |> ignore
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

/// True when the format's first section contains an elapsed-time bracket
/// like `[h]`, `[mm]`, or `[ss]` (interior is only h/m/s letters). These
/// request a running total of that unit — hours past 24, minutes past 60,
/// seconds without limit — rather than a wall-clock component, which a
/// stock .NET DateTime format string cannot express. Colour / locale tags
/// such as `[Red]` or `[$-409]` are not matched (they carry other letters).
let private hasElapsedBracket (code: string) : bool =
    System.Text.RegularExpressions.Regex.IsMatch(firstSection code, @"\[[hHmMsS]+\]")

/// Format an Excel time serial through an elapsed-time format code
/// (`[h]:mm:ss`, `[m]`, `[s].000`, ...). The leading bracket unit carries
/// the full elapsed total; any trailing `mm` / `ss` tokens are the
/// remainder within the parent unit. Matching Excel (verified via
/// `tools/excel-parity/probe-elapsed.vbs`), the serial is rounded to the
/// format's fractional-second precision, away from zero, *before* it is
/// decomposed — so a 32.5 s value prints `0:00:33` under `[h]:mm:ss` yet
/// `32.500` under `[s].000` and `0:00` under `[h]:mm`.
let private formatElapsedTime (code: string) (serial: float) : string =
    let section = firstSection code
    // Drop leading colour / locale tags so the token walk sees only the
    // time pattern (`[Red][h]:mm` -> `[h]:mm`); keep the elapsed brackets.
    let pattern =
        System.Text.RegularExpressions.Regex.Replace(
            section, @"\[(?![hHmMsS]+\])[^\]]*\]", "")
    let prec = max 0 (fractionalSecondPrecision pattern)
    let scale = System.Math.Pow(10.0, float prec)
    let totalSec =
        System.Math.Round(serial * 86400.0 * scale, System.MidpointRounding.AwayFromZero) / scale
    let intSec = int64 (floor (abs totalSec))
    let frac = abs totalSec - float intSec
    let totalHours = intSec / 3600L
    let totalMinutes = intSec / 60L
    let minWithinHour = (intSec % 3600L) / 60L
    let secWithinMin = intSec % 60L
    let sign = if totalSec < 0.0 then "-" else ""
    let fracDigits () =
        if prec <= 0 then ""
        else
            let scaled = int64 (System.Math.Round(frac * scale, System.MidpointRounding.AwayFromZero))
            "." + scaled.ToString().PadLeft(prec, '0')
    let sb = System.Text.StringBuilder()
    sb.Append(sign) |> ignore
    let lower = pattern
    let mutable i = 0
    while i < lower.Length do
        let c = lower.[i]
        if c = '[' then
            let endB = lower.IndexOf(']', i + 1)
            if endB < 0 then i <- lower.Length
            else
                let inner = lower.Substring(i + 1, endB - i - 1)
                let width = inner.Length
                let v =
                    match System.Char.ToLowerInvariant inner.[0] with
                    | 'h' -> totalHours
                    | 'm' -> totalMinutes
                    | _ -> intSec
                sb.Append(v.ToString().PadLeft(width, '0')) |> ignore
                i <- endB + 1
        elif c = 'h' || c = 'H' then
            let mutable n = 0
            while i + n < lower.Length && (lower.[i + n] = 'h' || lower.[i + n] = 'H') do n <- n + 1
            // Plain (non-bracketed) hours only appear without a [h] lead;
            // fall back to wall-clock hours so we never throw.
            sb.Append((totalHours % 24L).ToString().PadLeft(min 2 n, '0')) |> ignore
            i <- i + n
        elif c = 'm' || c = 'M' then
            let mutable n = 0
            while i + n < lower.Length && (lower.[i + n] = 'm' || lower.[i + n] = 'M') do n <- n + 1
            sb.Append(minWithinHour.ToString().PadLeft(min 2 n, '0')) |> ignore
            i <- i + n
        elif c = 's' || c = 'S' then
            let mutable n = 0
            while i + n < lower.Length && (lower.[i + n] = 's' || lower.[i + n] = 'S') do n <- n + 1
            sb.Append(secWithinMin.ToString().PadLeft(min 2 n, '0')) |> ignore
            i <- i + n
        elif c = '.' && i + 1 < lower.Length && lower.[i + 1] = '0' then
            let mutable n = 0
            while i + 1 + n < lower.Length && lower.[i + 1 + n] = '0' do n <- n + 1
            sb.Append(fracDigits ()) |> ignore
            i <- i + 1 + n
        elif c = '\\' && i + 1 < lower.Length then
            sb.Append(lower.[i + 1]) |> ignore
            i <- i + 2
        elif c = '"' then
            let endQ = lower.IndexOf('"', i + 1)
            if endQ < 0 then i <- lower.Length
            else
                sb.Append(lower.Substring(i + 1, endQ - i - 1)) |> ignore
                i <- endQ + 1
        else
            sb.Append(c) |> ignore
            i <- i + 1
    sb.ToString()

/// Format an Excel serial number as a date string per the given Excel
/// format code. Returns None when the serial is outside DateTime's
/// representable range; caller should render `#####` overflow in that
/// case (Excel's behavior for serials it can't convert).
let private formatExcelDateOpt (code: string) (serial: float) : string option =
    if hasElapsedBracket code then Some (formatElapsedTime code serial) else
    let dt =
        try Some (excelSerialToDate serial)
        with _ -> None
    let inv = System.Globalization.CultureInfo.InvariantCulture
    match dt with
    | None -> None
    | Some dt ->
        let dtRounded =
            // Excel rounds a time to the precision of its finest displayed
            // unit before drawing. With an explicit `.0+` suffix that's the
            // fractional second; otherwise, when seconds are shown at all,
            // round to the whole second (Excel renders 59.9999996 s as the
            // next whole second, e.g. `0:01:00`, where .NET would truncate
            // to `0:00:59`). Date-only / minute-only formats keep `prec < 0`
            // so they are left untouched.
            let prec =
                let p = fractionalSecondPrecision code
                if p >= 0 then p
                elif (withoutLocaleTags (firstSection code)).ToLowerInvariant().Contains "s" then 0
                else -1
            if prec < 0 then dt
            else
                // Keep the exponent in float space. Fable maps int64 `pown`
                // through BigInt, which browser Math.pow rejects.
                let unitsPerSecond = System.Math.Pow(10.0, float prec)
                let ticksPerUnit = int64 (float System.TimeSpan.TicksPerSecond / unitsPerSecond)
                if ticksPerUnit <= 0L then dt
                else
                    let rounded = (dt.Ticks + ticksPerUnit / 2L) / ticksPerUnit * ticksPerUnit
                    try System.DateTime(rounded) with _ -> dt
        let netFormat = excelToNetFormat code
        try Some (dtRounded.ToString(netFormat, inv))
        with _ -> Some (dtRounded.ToString("MM/dd/yyyy", inv))

/// Approximate rendered Maximum Digit Width (MDW) in CSS pixels for a given
/// XLSX font size (in points). Excel's MDW table is flat around the common
/// 10/11/12 pt sizes (all 7 px), then jumps to 8 px at 12-13 pt. Use the
/// `round(fontPt × 7/11)` formula but floor it at 7 — the value Excel uses
/// for every Calibri / Arial size from 8 pt up through 11 pt. Without the
/// floor, size-10 cells (Arial 10pt, the Excel 2003 legacy default in many
/// pandas-test_datetime_mi-style POI fixtures) under-count chars and trip
/// ##### overflow on values that Excel renders inside the cell.
/// Public so the paint layer can size column-width heuristics from the same
/// Maximum-Digit-Width table the formatter uses for ##### overflow.
let mdwPx (fontPt: float) : float =
    max 7.0 (System.Math.Round(fontPt * 7.0 / 11.0))

/// Width-aware Excel "General" number formatting. Picks the longest .NET
/// "G<n>" representation that fits the cell width. If even the shortest
/// representation overflows, returns `#####` like Excel does.
let private generalNumberDisplay (cellWidthPx: float) (fontPt: float) (n: float): string =
    // Excel's General-format digit budget = floor((cellWidthPx - 5) / MDW).
    // The 5 px matches OOXML's column-width padding (2 px each side + 1 px
    // gridline) so pixels and effective-char count round-trip via the same
    // formula. Verified against pandas-test1 col B (88 px / mdw=8 → 10
    // chars → "0.98026851", matches Excel COM cell.Text). Try the full
    // precision sweep so we surface G9 mid-trim instead of jumping from
    // G10 straight to G8.
    let usableWidth = max 0.0 (cellWidthPx - 5.0)
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
            // `_X` reserves the width of one X glyph (Excel uses this
            // for column-aligning open/close parens in accounting
            // formats); `*X` repeats X to fill remaining column width.
            // Without access to the cell width here we substitute a
            // single occurrence. Width-aware accounting fill expands this
            // marker later, after the overflow decision has been made.
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

/// Excel scales a number by 1000 for each comma at the END of the format
/// (after the last digit placeholder) — `#,##0,`→thousands, `0.0,,`→millions.
/// .NET only honours a scaling comma immediately LEFT of the decimal point, so
/// a trailing comma after the decimals (`#,##0.0,`) is ignored. Strip the
/// trailing scaling commas and return the matching divisor so we scale the
/// value ourselves; the grouping commas inside the digit run are left intact.
let private trailingScale (section: string) : string * float =
    let isDigitPh c = c = '0' || c = '#' || c = '?'
    match section |> Seq.tryFindIndexBack isDigitPh with
    | None -> section, 1.0
    | Some idx ->
        let mutable i = idx + 1
        while i < section.Length && section.[i] = ',' do i <- i + 1
        let commas = i - (idx + 1)
        if commas = 0 then section, 1.0
        else section.Substring(0, idx + 1) + section.Substring(i), (1000.0 ** float commas)

/// Excel's General text for a number, used when a `@` text placeholder is
/// applied to a numeric value (Excel shows the number, e.g. `@`→"5"/"-3.5").
let private generalString (inv: System.Globalization.CultureInfo) (n: float) : string =
    if n = floor n && abs n < 1e15 then (int64 n).ToString(inv)
    else n.ToString("G15", inv)

/// Apply an Excel non-date number-format code (`0.000`, `#,##0.00`, `0%`,
/// `0.00E+00`, etc.) using .NET's custom-format-string support. Excel's
/// number-format grammar is largely compatible with .NET's; we strip the
/// section separator first and pass the positive section to `Double.ToString`,
/// falling back to `G` if .NET rejects it.
let private formatExcelNumber (code: string) (n: float) : string =
    let inv = System.Globalization.CultureInfo.InvariantCulture
    let sections = splitSections code
    let pos, scaleDivisor = trailingScale (sectionForValue code n)
    let isConditional = sections |> List.exists (fun s -> (sectionCondition s).IsSome)
    // Excel formats the section against the ABSOLUTE value when the section
    // already accounts for the sign (a dedicated negative/zero section, or any
    // conditioned section), so .NET doesn't prepend a second `-` to something
    // like `(#,##0.00)`. When a NEGATIVE value falls through to a conditional
    // format's default (unconditioned) section, Excel prepends a `-` itself
    // (e.g. `[>=100]0;[>=50]0;0` renders -10 as `-10`, "Low" as "-Low").
    let baseValue, prependMinus =
        if isConditional then
            if (sectionCondition pos).IsSome then System.Math.Abs n, false
            else System.Math.Abs n, (n < 0.0)
        elif (n < 0.0 && sections.Length >= 2) || (n = 0.0 && sections.Length >= 3) then
            System.Math.Abs n, false
        else n, false
    let valueForSection = baseValue / scaleDivisor
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
    // `[$<currency>-<locale>]` carries a currency symbol prefix (the bit
    // before the `-`) and a locale hint (the hex digits after). The locale
    // hint is just metadata for us — we let .NET's invariant culture
    // format the digits — but the currency symbol IS part of the rendered
    // string and must be preserved. Substitute the prefix as a quoted
    // literal so .NET formats it verbatim. Examples:
    //   [$-409]m/d/yyyy           → m/d/yyyy   (no currency, just locale)
    //   [$EUR-407]#,##0.00        → "EUR"#,##0.00
    //   [$€-2]#,##0.00            → "€"#,##0.00
    let cleaned =
        System.Text.RegularExpressions.Regex.Replace(
            cleaned,
            @"\[\$([^\]\-]*)(?:-[0-9A-Fa-f]+)?\]",
            fun m ->
                let prefix = m.Groups.[1].Value
                if prefix.Length = 0 then ""
                else "'" + prefix.Replace("'", "''") + "'")
    // Color tags don't change the format string; they're handled by
    // `numberFormatColorForValue` separately. Drop them — and the `[op N]`
    // condition tags, which only selected the section — so neither ends up
    // in the output.
    let cleaned =
        System.Text.RegularExpressions.Regex.Replace(
            cleaned, @"\[(?:Red|Black|Green|White|Blue|Magenta|Yellow|Cyan)\]", "")
    let cleaned =
        System.Text.RegularExpressions.Regex.Replace(
            cleaned, @"\[(?:<=|>=|<>|=|<|>)-?[0-9]+(?:\.[0-9]+)?\]", "")
    if System.String.IsNullOrWhiteSpace pos then
        // An empty section renders nothing — e.g. `0.00;;` hides zeros and
        // negatives, `0;;` hides zeros.
        ""
    elif pos.Contains "@" then
        // `@` text placeholder on a number → its General text, with quoted
        // literals / escapes unwrapped and colour tags dropped.
        let s = pos.Replace("@", generalString inv n)
        let s = System.Text.RegularExpressions.Regex.Replace(s, "\"([^\"]*)\"", "$1")
        let s = System.Text.RegularExpressions.Regex.Replace(s, @"\\(.)", "$1")
        System.Text.RegularExpressions.Regex.Replace(
            s, @"\[(?:Red|Black|Green|White|Blue|Magenta|Yellow|Cyan)\]", "")
    else
        let formatted =
            try valueForSection.ToString(cleaned, inv)
            with _ ->
                try n.ToString("G", inv)
                with _ -> string n
        if prependMinus then "-" + formatted else formatted

/// Expand Excel's `* ` accounting fill when it follows a literal currency
/// marker. The fill characters are alignment padding, not overflow content:
/// Excel first decides whether the number fits, then repeats the fill so the
/// currency marker stays near the left edge and the number stays rightward.
let private expandAccountingFill
    (measure: (string -> float) option)
    (cellWidth: float)
    (section: string)
    (text: string) : string =
    let isHash =
        text.Length > 0 && text |> Seq.forall (fun c -> c = '#')
    match measure with
    | None -> text
    | Some m when isHash -> text
    | Some m ->
        let star = section.IndexOf("* ")
        if star <= 0 then text
        else
            let prev = section.[star - 1]
            // Covers Excel's common accounting patterns: literal `_($* #,##0)`
            // (prev = '$') and the locale currency tag `_-[$$-409]* #,##0`
            // (prev = ']', the close of a `[$…]` marker). Non-currency star
            // fills stay collapsed — right-aligned drawing already matches.
            if prev <> '$' && prev <> ']' then text
            else
                let marker = "$ "
                let markerAt = text.IndexOf(marker, System.StringComparison.Ordinal)
                if markerAt < 0 then text
                else
                    let prefix = text.Substring(0, markerAt + 1)
                    let suffix = text.Substring(markerAt + marker.Length)
                    let fillWidth = max 1.0 (m " ")
                    let baseWidth = m (prefix + suffix)
                    // Excel's text fill runs slightly past the visible cell
                    // box; cellWidth + 5 px matches the OOXML cell padding
                    // boundary and synthetic-numfmt's COM cell.Text exactly.
                    let target = cellWidth + 5.0
                    let count =
                        int (System.Math.Round((target - baseWidth) / fillWidth))
                        |> max 1
                    prefix + System.String(' ', count) + suffix

/// Greedy word-wrap inside a cell. Splits on hard line breaks (\n) first,
/// then on whitespace within each segment, packing as many words onto a line
/// as the cell's usable width allows. Falls back to char-level breaks when
/// a single word exceeds the width.
///
/// When `measure` is Some, line-width budget is computed via the real loaded
/// font (matches Excel's glyph-aware wrapping). When None — Fable web canvas
/// or other contexts without a measurer — fall back to `mdwPx × charCount`,
/// which over-estimates width and wraps earlier than Excel.
let wrapLines
    (measure: (string -> float) option)
    (cellWidthPx: float) (fontPt: float)
    (text: string) : string list =
    if System.String.IsNullOrEmpty text then [ text ]
    else
        // Excel's effective usable wrap width is cellWidth - 14 px. The
        // older `cellWidth - 8` budget under-broke wide-column wraps:
        // synthetic-wrap-autoheight's 154-px cell wraps to 8 lines in
        // Excel, but our text engine measures `width like the one Excel`
        // at 140.36 px which fits in 146 (= 154 - 8) and bled to a 7-line
        // wrap. With the tighter 14-px inset the budget drops to 140 and
        // that line breaks before "Excel", restoring the 8-line count.
        let usable = max 0.0 (cellWidthPx - 14.0)
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
///
/// `measure` is the cell's real-font width function (Some on the native /
/// PDF paths, None on the Fable web canvas). When present we both decide
/// overflow AND size the `#` run against the *actual* glyph the renderer
/// will draw. The older `mdwPx`-only path under-counted: a date string
/// measured at ~7 px/char but rendered in Calibri (~9 px '#') produced a
/// `#####` run wider than the cell, and since numbers right-align the
/// excess spilled left into the row-header gutter. The None branch keeps
/// the approximation for the web build.
let private fitOrOverflow
    (measure: (string -> float) option)
    (cellWidth: float) (fontName: string option) (fontPt: float) (text: string) : string =
    // Two thresholds: the FIT check (cellWidth - 8) matches the 4-px
    // text inset on each side in buildCellPaintItems. The HASH BUDGET
    // (cellWidth - 5) matches OOXML's column-width padding so the
    // overflow-`#` count round-trips with Excel's chars-per-pixel math.
    // Verified against pandas-test1 Sheet1 col A (95 px / mdw=8 → 11 #s,
    // matches Excel COM cell.Text for the m/d/yyyy h:mm date that
    // overflows).
    //
    // Fit check uses the measured glyph width. Excel's effective fit
    // budget is `cellWidth - 3` (~2-px inset on the right with a tiny
    // left inset) — verified empirically against pandas-testskiprows C4
    // "2015-01-01 00:00:00" which Excel renders in a 128-px cell at our
    // measured 124.2 px. The 3-px budget lets it fit; the older 8-px
    // budget hashed it. pandas-test1 m/d/yyyy h:mm dates still hash
    // because the date's glyph width sits 3-4 px above the threshold
    // for the narrower Calibri-12 95-px cells.
    //
    // Calibri 11 is the one common case where Excel's COM `cell.Text`
    // boundary sits at the full cell width: synthetic-narrow-date fits
    // "1/10/2000 0:00" in a 94-px cell, while the same text at Calibri
    // 12 still hashes in a 95-px cell (pandas-test1). Keep the tighter
    // inset for the other sizes/families.
    let isCalibri11 =
        let name = fontName |> Option.defaultValue "Calibri"
        name.Trim().StartsWith("Calibri", System.StringComparison.OrdinalIgnoreCase)
        && abs (fontPt - 11.0) < 0.01
    let fitInset = if isCalibri11 then 0.0 else 3.0
    let fitWidth = max 0.0 (cellWidth - fitInset)
    let hashBudget = max 0.0 (cellWidth - 5.0)
    // Use the actual font's digit width when we have a measurer — Excel's
    // hash budget rounds to chars-per-MDW where MDW is the digit advance
    // of the cell's font. The nominal `mdwPx fontPt` ignores font family
    // (always returns 7 for size 11), which under-counts hashes for Arial
    // workbooks where MDW=8 (pandas-testdateoverflow's Polish-locale
    // cells resolve to Arial; Excel renders 29 #s = floor(235/8), we used
    // to render 33 = floor(235/7)).
    let mdw =
        match measure with
        | Some m ->
            let zero = m "0"
            if zero <= 0.0 then mdwPx fontPt else max 7.0 (System.Math.Floor zero)
        | None -> mdwPx fontPt
    match measure with
    | Some m ->
        let textWidth = m text
        if textWidth <= fitWidth then text
        else
            let count = max 1 (int (hashBudget / mdw))
            System.String('#', count)
    | None ->
        let fitChars = max 1 (int (fitWidth / mdw))
        let hashChars = max 1 (int (hashBudget / mdw))
        if text.Length <= fitChars then text
        else System.String('#', hashChars)

/// Per-cell display text logic. Public so the VBA-style API surface
/// (Spreadsheet GetVbaLayoutCells) can return the same text Excel's
/// `cell.Text` would produce — including `#####` overflow, format-applied
/// numbers / dates, error markers.
let rec cellDisplayText
    (measure: (string -> float) option)
    (cellWidth: float) (format: CellFormat) (value: CellValue): string =
    let fontPt = format.FontSize |> Option.map float |> Option.defaultValue 11.0
    match value with
    | CellValue.Empty -> ""
    | Number n ->
        match format.NumberFormat with
        | Some code when isDateFormat code || isTimeFormat code || hasElapsedBracket code ->
            // Out-of-range serials (e.g. testdateoverflow's 1E+20) can't be
            // formatted as a date; Excel renders ##### overflow for those
            // rather than falling through to scientific notation. Treat
            // the None branch as "doesn't fit" by passing an oversized
            // text that fitOrOverflow will hash-render.
            match formatExcelDateOpt code n with
            | Some s -> fitOrOverflow measure cellWidth format.FontName fontPt s
            | None -> fitOrOverflow measure cellWidth format.FontName fontPt (System.String('M', 1000))
        | Some "currency"
        | Some "percentage"
        | Some "date"
        | Some "decimal2"
        | Some "number" ->
            // Legacy named format codes from the Fable web UI's formatting
            // commands. Keep these on the original Model.formatNumber path.
            fitOrOverflow measure cellWidth format.FontName fontPt (formatNumber format.NumberFormat n)
        | Some code when System.String.IsNullOrWhiteSpace code ||
                         code.Equals("general", System.StringComparison.OrdinalIgnoreCase) ->
            generalNumberDisplay cellWidth fontPt n
        | Some code ->
            // Real Excel format code (0.000, #,##0.00, 0%, 0.00E+00, etc.).
            let formatted = formatExcelNumber code n
            let fitted = fitOrOverflow measure cellWidth format.FontName fontPt formatted
            expandAccountingFill measure cellWidth (sectionForValue code n) fitted
        | None -> generalNumberDisplay cellWidth fontPt n
    | String s -> s
    | Boolean b -> if b then "TRUE" else "FALSE"
    | Error e -> sprintf "#%s" e
    | Formula (_, cached) -> cellDisplayText measure cellWidth format cached

