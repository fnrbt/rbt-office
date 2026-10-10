module Rbt.Spreadsheet.Input

open System
open System.Globalization
open Rbt.Spreadsheet.Model

// ---------------------------------------------------------------------------
// What the user typed → a cell value (and, for a number written in a
// recognisable style, the number format that keeps showing it that way).
//
// One parser for every way text enters a cell: the in-cell editor, the
// formula bar, the programmatic SetCell and pasted text. It reads input the
// way Excel (en-US) does and never consults the machine's culture, so the
// .NET and browser builds agree:
//
//   =A1+1            formula
//   'text            text, whatever it looks like (the apostrophe is dropped)
//   TRUE / false     boolean
//   1,234.5  (12)    number; thousands separators and accounting negatives
//   12%              0.12, shown as a percentage
//   $1,234.50        1234.5, shown as currency
//   1 1/2            1.5, shown as a fraction
//   3/4/2024  2024-03-04  4-Mar  Mar 2024  March 4, 2024     date serial
//   14:30  2:30 PM  3/4/2024 14:30                           time / date-time
//   #N/A             error value
//   anything else    text
// ---------------------------------------------------------------------------

/// A parsed entry. `NumberFormat` is the format Excel would give a General
/// cell for this entry (None = leave the cell's format alone).
type ParsedInput = {
    Value: CellValue
    NumberFormat: string option
}

let private plain (value: CellValue) : ParsedInput = { Value = value; NumberFormat = None }

let private isDigit (ch: char) = ch >= '0' && ch <= '9'

let private allDigits (text: string) =
    text.Length > 0 && text |> Seq.forall isDigit

/// A number written with optional thousands separators and exponent:
/// (value, had separators, digits after the decimal point). No sign.
let private tryUnsignedNumber (text: string) : (float * bool * int) option =
    let n = text.Length
    let mutable i = 0
    let digits = Text.StringBuilder()
    let mutable grouped = false
    let mutable valid = n > 0
    // Integer part: digits, or groups of three after a first group of 1-3.
    let mutable groupLength = 0
    let mutable groups = 0
    while valid && i < n && (isDigit text.[i] || text.[i] = ',') do
        if text.[i] = ',' then
            // A separator needs digits before it and exactly three after.
            if groupLength = 0 || (groups > 0 && groupLength <> 3) || (groups = 0 && groupLength > 3) then valid <- false
            groups <- groups + 1
            groupLength <- 0
            grouped <- true
        else
            digits.Append(text.[i]) |> ignore
            groupLength <- groupLength + 1
        i <- i + 1
    if grouped && groupLength <> 3 then valid <- false
    let integerDigits = digits.Length
    let mutable decimals = 0
    if valid && i < n && text.[i] = '.' then
        digits.Append('.') |> ignore
        i <- i + 1
        while i < n && isDigit text.[i] do
            digits.Append(text.[i]) |> ignore
            decimals <- decimals + 1
            i <- i + 1
    if integerDigits = 0 && decimals = 0 then valid <- false
    if valid && i < n && (text.[i] = 'e' || text.[i] = 'E') && not grouped then
        digits.Append('e') |> ignore
        i <- i + 1
        if i < n && (text.[i] = '+' || text.[i] = '-') then
            digits.Append(text.[i]) |> ignore
            i <- i + 1
        let exponentStart = i
        while i < n && isDigit text.[i] do
            digits.Append(text.[i]) |> ignore
            i <- i + 1
        if i = exponentStart then valid <- false
    if not valid || i <> n then None
    else
        let literal = digits.ToString()
        let literal = if literal.StartsWith "." then "0" + literal else literal
        match Double.TryParse(literal, NumberStyles.Float, CultureInfo.InvariantCulture) with
        | true, value when not (Double.IsInfinity value || Double.IsNaN value) -> Some (value, grouped, decimals)
        | _ -> None

let private currencySymbols = "$€£¥"

/// A number with the decorations Excel accepts: sign or accounting
/// parentheses, a currency symbol, thousands separators, a trailing percent.
let private tryNumber (input: string) : ParsedInput option =
    let mutable text = input.Trim()
    let mutable negative = false
    let mutable currency = None
    let mutable percent = false
    let mutable valid = text.Length > 0
    if valid && text.StartsWith "(" && text.EndsWith ")" && text.Length > 2 then
        negative <- true
        text <- text.Substring(1, text.Length - 2).Trim()
    // Sign and currency symbol come in either order: -$5, $-5.
    let mutable leading = true
    while valid && leading && text.Length > 0 do
        let ch = text.[0]
        if (ch = '-' || ch = '+') then
            if ch = '-' then
                if negative then valid <- false
                negative <- true
            text <- text.Substring(1).TrimStart()
            // Only one sign.
            if text.Length > 0 && (text.[0] = '-' || text.[0] = '+') then valid <- false
        elif currencySymbols.IndexOf ch >= 0 && currency.IsNone then
            currency <- Some ch
            text <- text.Substring(1).TrimStart()
        else leading <- false
    if valid && text.EndsWith "%" then
        percent <- true
        text <- text.Substring(0, text.Length - 1).TrimEnd()
    if not valid || (percent && currency.IsSome) then None
    else
        match tryUnsignedNumber text with
        | None -> None
        | Some (magnitude, grouped, decimals) ->
            let signed = if negative then -magnitude else magnitude
            let fraction = if decimals > 0 then ".00" else ""
            if percent then
                Some { Value = Number (signed / 100.0); NumberFormat = Some (if decimals > 0 then "0.00%" else "0%") }
            else
                match currency with
                | Some '$' -> Some { Value = Number signed; NumberFormat = Some ("$#,##0" + fraction) }
                | Some symbol -> Some { Value = Number signed; NumberFormat = Some ("\"" + string symbol + "\"#,##0" + fraction) }
                | None ->
                    Some { Value = Number signed
                           NumberFormat = (if grouped then Some ("#,##0" + fraction) else None) }

/// `1 1/2`, `0 3/4`: a whole number, a space and a fraction.
let private tryFraction (input: string) : ParsedInput option =
    let text = input.Trim()
    let negative = text.StartsWith "-"
    let body = if negative then text.Substring 1 else text
    let space = body.IndexOf ' '
    let slash = body.IndexOf '/'
    if space <= 0 || slash <= space + 1 || slash >= body.Length - 1 then None
    else
        let whole = body.Substring(0, space)
        let numerator = body.Substring(space + 1, slash - space - 1).Trim()
        let denominator = body.Substring(slash + 1)
        if allDigits whole && allDigits numerator && allDigits denominator && whole.Length <= 9 && numerator.Length <= 6 && denominator.Length <= 6 then
            let d = float (int denominator)
            if d = 0.0 then None
            else
                let value = float (int whole) + float (int numerator) / d
                Some { Value = Number (if negative then -value else value)
                       NumberFormat = Some (if denominator.Length > 1 then "# ??/??" else "# ?/?") }
        else None

let private monthNames =
    [| "january"; "february"; "march"; "april"; "may"; "june"; "july"; "august"; "september"; "october"; "november"; "december" |]

/// 1..12 for a month name or its abbreviation of three letters or more.
let private monthOf (text: string) : int option =
    let lower = text.ToLowerInvariant().TrimEnd('.')
    if lower.Length < 3 then None
    else
        monthNames
        |> Array.tryFindIndex (fun name -> name = lower || (lower.Length >= 3 && name.StartsWith lower) || (lower = "sept" && name = "september"))
        |> Option.map (fun index -> index + 1)

let private isLeapYear (year: int) = (year % 4 = 0 && year % 100 <> 0) || year % 400 = 0

let private daysInMonth (year: int) (month: int) =
    match month with
    | 2 -> if isLeapYear year then 29 else 28
    | 4 | 6 | 9 | 11 -> 30
    | _ -> 31

/// Days since 1970-01-01 of a proleptic Gregorian date.
let private daysFromCivil (year: int) (month: int) (day: int) : int =
    let y = if month <= 2 then year - 1 else year
    let era = (if y >= 0 then y else y - 399) / 400
    let yearOfEra = y - era * 400
    let dayOfYear = (153 * (month + (if month > 2 then -3 else 9)) + 2) / 5 + day - 1
    let dayOfEra = yearOfEra * 365 + yearOfEra / 4 - yearOfEra / 100 + dayOfYear
    era * 146097 + dayOfEra - 719468

/// The Excel serial of a date (days since 1899-12-30, the engine's epoch).
let dateSerial (year: int) (month: int) (day: int) : float =
    float (daysFromCivil year month day - daysFromCivil 1899 12 30) - dateSystemDays ()

let private trySerial (year: int) (month: int) (day: int) : float option =
    if year < (if date1904 then 1904 else 1900) || year > 9999 || month < 1 || month > 12 || day < 1 || day > daysInMonth year month then None
    else Some (dateSerial year month day)

/// A year as typed: four digits as they are, two digits by Excel's window
/// (00-29 → 2000s, 30-99 → 1900s).
let private yearOf (text: string) : int option =
    if not (allDigits text) then None
    elif text.Length = 4 then Some (int text)
    elif text.Length <= 2 then (let y = int text in Some (if y < 30 then 2000 + y else 1900 + y))
    else None

let private smallInt (text: string) : int option =
    if allDigits text && text.Length <= 2 then Some (int text) else None

/// A date: (serial, number format).
let private tryDate (currentYear: int) (input: string) : (float * string) option =
    let text = input.Trim()
    if text.Length < 3 || text.Length > 40 then None
    else
        let hasMonthName = text |> Seq.exists Char.IsLetter
        if not hasMonthName then
            let separator =
                if text.IndexOf '/' >= 0 then Some '/'
                elif text.IndexOf '-' > 0 then Some '-'
                else None
            match separator with
            | None -> None
            | Some sep ->
                // Mixed separators are not a date.
                if (sep = '/' && text.IndexOf '-' >= 0) then None
                else
                    match text.Split(sep) with
                    | [| first; second; third |] when allDigits first && allDigits second && allDigits third ->
                        if first.Length = 4 then
                            // ISO: year first.
                            match smallInt second, smallInt third with
                            | Some month, Some day -> trySerial (int first) month day |> Option.map (fun serial -> serial, "yyyy-mm-dd")
                            | _ -> None
                        else
                            match smallInt first, smallInt second, yearOf third with
                            | Some month, Some day, Some year -> trySerial year month day |> Option.map (fun serial -> serial, "m/d/yyyy")
                            | _ -> None
                    | [| first; second |] when allDigits first && allDigits second ->
                        (match smallInt first with
                         | None -> None
                         | Some month ->
                            if second.Length = 4 then
                                // 3/2024: the first of the month.
                                trySerial (int second) month 1 |> Option.map (fun serial -> serial, "mmm-yy")
                            else
                                match smallInt second with
                                | Some day when month >= 1 && month <= 12 && day >= 1 && day <= daysInMonth currentYear month ->
                                    trySerial currentYear month day |> Option.map (fun serial -> serial, "d-mmm")
                                // 3/45 is March 1945, as in Excel.
                                | Some year when month >= 1 && month <= 12 && year > 31 ->
                                    trySerial (1900 + year) month 1 |> Option.map (fun serial -> serial, "mmm-yy")
                                | _ -> None)
                    | _ -> None
        else
            let tokens = text.Split([| ' '; '-'; '/'; ',' |], StringSplitOptions.RemoveEmptyEntries)
            let month index = if index < tokens.Length then monthOf tokens.[index] else None
            let day index = if index < tokens.Length then smallInt tokens.[index] else None
            let year index = if index < tokens.Length then yearOf tokens.[index] else None
            match tokens.Length with
            | 2 ->
                (match day 0, month 1, month 0, day 1 with
                 // 4-Mar
                 | Some d, Some m, _, _ -> trySerial currentYear m d |> Option.map (fun serial -> serial, "d-mmm")
                 // Mar 4 — or Mar 2024 / Mar-45 when the number cannot be a day.
                 | _, _, Some m, Some d when d >= 1 && d <= daysInMonth currentYear m ->
                    trySerial currentYear m d |> Option.map (fun serial -> serial, "d-mmm")
                 | _, _, Some m, _ ->
                    (match year 1 with
                     | Some y -> trySerial y m 1 |> Option.map (fun serial -> serial, "mmm-yy")
                     | None -> None)
                 | _ -> None)
            | 3 ->
                (match day 0, month 1, year 2 with
                 // 4-Mar-2024, 4 March 24
                 | Some d, Some m, Some y -> trySerial y m d |> Option.map (fun serial -> serial, "d-mmm-yy")
                 | _ ->
                    // March 4, 2024
                    match month 0, day 1, year 2 with
                    | Some m, Some d, Some y -> trySerial y m d |> Option.map (fun serial -> serial, "d-mmm-yy")
                    | _ -> None)
            | _ -> None

/// A time of day: (fraction of a day, number format).
let private tryTime (input: string) : (float * string) option =
    let mutable text = input.Trim()
    let upper = text.ToUpperInvariant()
    let meridiem =
        if upper.EndsWith "AM" || upper.EndsWith "A" then Some false
        elif upper.EndsWith "PM" || upper.EndsWith "P" then Some true
        else None
    if meridiem.IsSome then
        let cut = if upper.EndsWith "M" then 2 else 1
        text <- text.Substring(0, text.Length - cut).TrimEnd()
    let parts = text.Split(':')
    // A bare number is a time only with AM/PM ("2 PM").
    if parts.Length < 1 || parts.Length > 3 || (parts.Length = 1 && meridiem.IsNone) then None
    elif not (allDigits parts.[0]) || parts.[0].Length > 4 then None
    else
        let minutesText = if parts.Length > 1 then parts.[1] else "0"
        let secondsText = if parts.Length > 2 then parts.[2] else "0"
        let seconds =
            // Seconds may carry a fraction: 1:02:03.5
            match Double.TryParse(secondsText, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture) with
            | true, value when secondsText.Length > 0 && isDigit secondsText.[0] -> Some value
            | _ -> None
        match seconds with
        | None -> None
        | Some seconds when not (allDigits minutesText) || minutesText.Length > 2 || int minutesText > 59 || seconds >= 60.0 -> None
        | Some seconds ->
            let hours = int parts.[0]
            let minutes = int minutesText
            let adjusted =
                match meridiem with
                | Some _ when hours < 1 || hours > 12 -> None
                | Some true -> Some (if hours = 12 then 12 else hours + 12)
                | Some false -> Some (if hours = 12 then 0 else hours)
                | None -> Some hours
            match adjusted with
            | None -> None
            | Some h ->
                let fraction = (float h * 3600.0 + float minutes * 60.0 + seconds) / 86400.0
                let withSeconds = parts.Length > 2
                let format =
                    match meridiem, withSeconds with
                    | Some _, true -> "h:mm:ss AM/PM"
                    | Some _, false -> "h:mm AM/PM"
                    | None, _ when h >= 24 -> "[h]:mm:ss"
                    | None, true -> "h:mm:ss"
                    | None, false -> "h:mm"
                Some (fraction, format)

/// A date, a time, or a date followed by a time.
let private tryDateTime (currentYear: int) (input: string) : ParsedInput option =
    let text = input.Trim()
    let colon = text.IndexOf ':'
    let upper = text.ToUpperInvariant()
    let endsWithMeridiem = upper.EndsWith " AM" || upper.EndsWith " PM"
    if colon < 0 && not endsWithMeridiem then
        tryDate currentYear text |> Option.map (fun (serial, format) -> { Value = Number serial; NumberFormat = Some format })
    else
        // The time starts at the last space before the first colon (or
        // before the AM/PM marker); what precedes it, if anything, is a date.
        let anchor = if colon >= 0 then colon else text.Length - 3
        let split = text.LastIndexOf(' ', max 0 (anchor - 1))
        if split < 0 then
            tryTime text |> Option.map (fun (fraction, format) -> { Value = Number fraction; NumberFormat = Some format })
        else
            match tryTime text with
            | Some (fraction, format) -> Some { Value = Number fraction; NumberFormat = Some format }
            | None ->
                match tryDate currentYear (text.Substring(0, split)), tryTime (text.Substring(split + 1)) with
                | Some (serial, _), Some (fraction, timeFormat) when fraction < 1.0 ->
                    Some { Value = Number (serial + fraction); NumberFormat = Some ("m/d/yyyy " + timeFormat) }
                | _ -> None

let private errorLiterals = [ "#DIV/0!"; "#VALUE!"; "#REF!"; "#NAME?"; "#NUM!"; "#N/A"; "#NULL!"; "#SPILL!"; "#CALC!" ]

/// Parse what was typed into a cell. `currentYear` completes a date typed
/// without a year (`4-Mar`).
let parseInputIn (currentYear: int) (text: string) : ParsedInput =
    if isNull text then plain Empty
    else
        let trimmed = text.Trim()
        if trimmed = "" then plain Empty
        // A leading apostrophe forces text; it is not part of the value.
        elif text.StartsWith "'" then plain (String (text.Substring 1))
        elif trimmed.StartsWith "=" && trimmed.Length > 1 then plain (Formula (trimmed, Empty))
        elif trimmed.Equals("TRUE", StringComparison.OrdinalIgnoreCase) then plain (Boolean true)
        elif trimmed.Equals("FALSE", StringComparison.OrdinalIgnoreCase) then plain (Boolean false)
        else
            match errorLiterals |> List.tryFind (fun literal -> literal.Equals(trimmed, StringComparison.OrdinalIgnoreCase)) with
            | Some literal -> plain (Error (literal.Substring 1))
            | None ->
                match tryNumber trimmed with
                | Some parsed -> parsed
                | None ->
                    match tryFraction trimmed with
                    | Some parsed -> parsed
                    | None ->
                        match tryDateTime currentYear trimmed with
                        | Some parsed -> parsed
                        | None -> plain (String text)

/// Parse what was typed into a cell, as of today.
let parseInput (text: string) : ParsedInput =
    parseInputIn DateTime.UtcNow.Year text

/// The cell that results from typing `text` into a cell that currently holds
/// `existing`: the parsed value, the existing formatting and comment, and —
/// when the cell has no number format of its own — the format the entry
/// implies (typing 12% into a General cell makes it a percentage cell).
let cellFromInput (currentYear: int) (existing: CellData option) (text: string) : CellData =
    let parsed = parseInputIn currentYear text
    let format = existing |> Option.map (fun cell -> cell.Format) |> Option.defaultValue defaultFormat
    let isGeneral =
        match format.NumberFormat with
        | None -> true
        | Some code -> code = "" || code.Equals("General", StringComparison.OrdinalIgnoreCase)
    let format =
        match parsed.NumberFormat with
        | Some code when isGeneral -> { format with NumberFormat = Some code }
        | _ -> format
    { Value = parsed.Value; Format = format; Comment = existing |> Option.bind (fun cell -> cell.Comment) }

let private isDateTimeCode (code: string) = CellText.isDateTimeNumberFormat code

let private numberText (value: float) : string =
    if value = Math.Floor value && abs value < 1e15 then (int64 value).ToString(CultureInfo.InvariantCulture)
    else value.ToString("R", CultureInfo.InvariantCulture)

/// The text to edit for a cell: what, typed back unchanged, gives the same
/// cell. A formula shows its text, a date its date rather than the serial, a
/// percentage its percent figure, and text that would otherwise be read as
/// something else gets the apostrophe.
let editText (cell: CellData option) : string =
    match cell with
    | None -> ""
    | Some data ->
        match data.Value with
        | Empty -> ""
        | Formula (text, _) -> formulaText text
        | Boolean flag -> if flag then "TRUE" else "FALSE"
        | Error code -> if code.StartsWith "#" then code else "#" + code
        | Number n ->
            (match data.Format.NumberFormat with
             | Some code when isDateTimeCode code && n >= 0.0 ->
                let whole = Math.Floor n
                let fraction = n - whole
                if whole = 0.0 && fraction > 0.0 then CellText.formatNumberWithCode "h:mm:ss" n
                elif fraction < 1e-9 then CellText.formatNumberWithCode "m/d/yyyy" n
                else CellText.formatNumberWithCode "m/d/yyyy h:mm:ss" n
             | Some code when code.IndexOf '%' >= 0 && code.IndexOf '"' < 0 ->
                numberText (Math.Round(n * 100.0, 10)) + "%"
             | _ -> numberText n)
        | String text ->
            // Text that reads as a number, date, formula… needs the
            // apostrophe to stay text when the edit is committed.
            match (parseInputIn 2000 text).Value with
            | String same when same = text -> text
            | _ -> "'" + text

// ---------------------------------------------------------------------------
// Autofill: continuing a block of cells along a row or column, the way
// dragging the fill handle does. Each line of the block is continued on its
// own: numbers by their linear trend, a date by one day, weekday and month
// names around their cycle, text ending in a number by counting on, formulas
// by shifting their references, and anything else by repeating the pattern.
// ---------------------------------------------------------------------------

type FillDirection =
    | FillDown
    | FillUp
    | FillRight
    | FillLeft

let private weekdayNames = [| "sunday"; "monday"; "tuesday"; "wednesday"; "thursday"; "friday"; "saturday" |]

/// A name from a cycle (weekdays, months): (cycle, index, abbreviated).
let private cycleEntry (text: string) : (string[] * int * bool) option =
    let lower = text.Trim().ToLowerInvariant()
    let find (names: string[]) =
        match names |> Array.tryFindIndex (fun name -> name = lower) with
        | Some index -> Some (names, index, false)
        | None ->
            if lower.Length <> 3 then None
            else names |> Array.tryFindIndex (fun name -> name.Substring(0, 3) = lower) |> Option.map (fun index -> names, index, true)
    match find weekdayNames with
    | Some entry -> Some entry
    | None -> find monthNames

/// Write a cycle name in the letter case of `sample`.
let private inCaseOf (sample: string) (name: string) (abbreviated: bool) : string =
    let name = if abbreviated then name.Substring(0, 3) else name
    let sample = sample.Trim()
    if sample.Length > 1 && sample = sample.ToUpperInvariant() then name.ToUpperInvariant()
    elif sample.Length > 0 && Char.IsLower sample.[0] then name
    else name.Substring(0, 1).ToUpperInvariant() + name.Substring 1

/// Text ending in digits: (prefix, number, digit count).
let private trailingNumber (text: string) : (string * int * int) option =
    let mutable start = text.Length
    while start > 0 && isDigit text.[start - 1] do start <- start - 1
    let digits = text.Length - start
    if digits = 0 || digits > 9 || start = 0 then None
    else Some (text.Substring(0, start), int (text.Substring start), digits)

let private positiveModulo (value: int) (modulus: int) = ((value % modulus) + modulus) % modulus

/// The continuation of one line of source cells. `shiftFormula delta text`
/// moves a formula `delta` cells along the fill axis. The result maps a
/// position on the line (0 = first source cell; negative = before it) to
/// the cell that belongs there.
let private continueLine (series: bool) (shiftFormula: int -> string -> string) (source: CellData option[]) : int -> CellData option =
    let n = source.Length
    let repeat (k: int) : CellData option =
        let j = positiveModulo k n
        match source.[j] with
        | Some ({ Value = Formula (text, _) } as cell) -> Some { cell with Value = Formula (shiftFormula (k - j) text, Empty) }
        | other -> other
    let values = source |> Array.map (Option.map (fun cell -> cell.Value))
    let formatOf (k: int) = source.[positiveModulo k n] |> Option.map (fun cell -> cell.Format) |> Option.defaultValue defaultFormat
    let numbers = values |> Array.choose (fun v -> match v with Some (Number x) -> Some x | _ -> None)
    let texts = values |> Array.choose (fun v -> match v with Some (String t) -> Some t | _ -> None)
    if not series || n = 0 then repeat
    elif numbers.Length = n then
        if n = 1 then
            // A lone number repeats; a lone date counts on by days.
            match source.[0] with
            | Some cell when (match cell.Format.NumberFormat with Some code -> CellText.isDateTimeNumberFormat code | None -> false)
                             && numbers.[0] >= 1.0 ->
                fun k -> Some { cell with Value = Number (numbers.[0] + float k) }
            | _ -> repeat
        else
            // Least-squares line through the values (an exact step when they
            // are evenly spaced), as Excel's linear trend.
            let count = float n
            let meanX = (count - 1.0) / 2.0
            let meanY = Array.average numbers
            let mutable sxy = 0.0
            let mutable sxx = 0.0
            for i in 0 .. n - 1 do
                sxy <- sxy + (float i - meanX) * (numbers.[i] - meanY)
                sxx <- sxx + (float i - meanX) * (float i - meanX)
            let slope = sxy / sxx
            let intercept = meanY - slope * meanX
            fun k ->
                if k >= 0 && k < n then source.[k]
                else Some { Value = Number (Math.Round(intercept + slope * float k, 10)); Format = formatOf k; Comment = None }
    elif texts.Length = n then
        let entries = texts |> Array.map cycleEntry
        let numbered = texts |> Array.map trailingNumber
        if entries |> Array.forall Option.isSome then
            let names, first, abbreviated = entries.[0].Value
            let sameCycle = entries |> Array.forall (fun entry -> let cycle, _, _ = entry.Value in obj.ReferenceEquals(cycle, names))
            let step = if n >= 2 then (let _, second, _ = entries.[1].Value in positiveModulo (second - first) names.Length) else 1
            let step = if step = 0 then names.Length else step
            let steady =
                entries |> Array.mapi (fun i entry -> let _, index, _ = entry.Value in index = positiveModulo (first + step * i) names.Length)
                |> Array.forall id
            if sameCycle && steady then
                fun k -> Some { Value = String (inCaseOf texts.[0] names.[positiveModulo (first + step * k) names.Length] abbreviated); Format = formatOf k; Comment = None }
            else repeat
        elif numbered |> Array.forall Option.isSome then
            let prefix, first, digits = numbered.[0].Value
            let samePrefix = numbered |> Array.forall (fun entry -> let p, _, _ = entry.Value in p = prefix)
            let step = if n >= 2 then (let _, second, _ = numbered.[1].Value in second - first) else 1
            let steady = numbered |> Array.mapi (fun i entry -> let _, value, _ = entry.Value in value = first + step * i) |> Array.forall id
            if samePrefix && steady && step <> 0 then
                fun k ->
                    let value = abs (first + step * k)
                    // Keep leading zeros: "Item 01" → "Item 02".
                    let text = (string value).PadLeft((if texts.[0].Length > prefix.Length && texts.[0].[prefix.Length] = '0' then digits else 1), '0')
                    Some { Value = String (prefix + text); Format = formatOf k; Comment = None }
            else repeat
        else repeat
    else repeat

/// Continue the block `(top, left, bottom, right)` by `count` rows or columns
/// in `direction`. With `series` the lines are continued as series (the fill
/// handle); without, the block is repeated and formulas shifted (Ctrl+D /
/// Ctrl+R). Returns the model; the caller recalculates.
let fillCells (top: int, left: int, bottom: int, right: int) (direction: FillDirection) (count: int) (series: bool) (model: SpreadsheetModel) : SpreadsheetModel =
    if count <= 0 || top > bottom || left > right then model
    else
        let vertical = direction = FillDown || direction = FillUp
        let forward = direction = FillDown || direction = FillRight
        let lineLength = if vertical then bottom - top + 1 else right - left + 1
        let mutable result = model
        let lines = if vertical then [ left .. right ] else [ top .. bottom ]
        for line in lines do
            let addressAt (k: int) : CellAddress =
                if vertical then { Row = top + k; Col = line } else { Row = line; Col = left + k }
            let source = [| for k in 0 .. lineLength - 1 -> getCell (addressAt k) model |]
            let shiftFormula (delta: int) (text: string) =
                if vertical then shiftFormulaReferences delta 0 text else shiftFormulaReferences 0 delta text
            let continuation = continueLine series shiftFormula source
            for step in 1 .. count do
                let k = if forward then lineLength - 1 + step else -step
                let addr = addressAt k
                // (A row the AutoFilter hides is not filled, as in Excel.)
                if addr.Row >= 0 && addr.Col >= 0 && addr.Row < model.RowCount && addr.Col < model.ColCount && not (isFilterHiddenRow addr.Row model) then
                    match continuation k with
                    | Some cell -> result <- setCell addr { cell with Comment = None } result
                    | None -> result <- clearCell addr result
        result

/// The cells `fillCells` writes for the same arguments.
let fillTargets (top: int, left: int, bottom: int, right: int) (direction: FillDirection) (count: int) : CellAddress list =
    [ match direction with
      | FillDown -> for r in bottom + 1 .. bottom + count do for c in left .. right -> { Row = r; Col = c }
      | FillUp -> for r in top - count .. top - 1 do for c in left .. right -> { Row = r; Col = c }
      | FillRight -> for r in top .. bottom do for c in right + 1 .. right + count -> { Row = r; Col = c }
      | FillLeft -> for r in top .. bottom do for c in left - count .. left - 1 -> { Row = r; Col = c } ]
    |> List.filter (fun addr -> addr.Row >= 0 && addr.Col >= 0)

/// Where dragging the fill handle of a block to the cell `(row, col)` fills:
/// the direction and how many rows/columns, or None inside the block.
let fillTowards (top: int, left: int, bottom: int, right: int) (row: int) (col: int) : (FillDirection * int) option =
    let down, up = row - bottom, top - row
    let rightward, leftward = col - right, left - col
    let vertical = max down up
    let horizontal = max rightward leftward
    if vertical <= 0 && horizontal <= 0 then None
    elif vertical >= horizontal then (if down > 0 then Some (FillDown, down) else Some (FillUp, up))
    else (if rightward > 0 then Some (FillRight, rightward) else Some (FillLeft, leftward))
