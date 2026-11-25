/// Number-to-text formatting with .NET invariant-culture semantics that also
/// works under Fable. Fable does not implement `Double.ToString` with custom
/// patterns (`#,##0.00`, `0%`, `0.00E+00`), thousands grouping or `G<n>`, so
/// the browser build rendered every such format as the raw double. The
/// formatter here works on decimal digit strings and needs only an exponential
/// rendering of the number from the host, which it rounds itself: JS and .NET
/// disagree on ties (2.5 to zero decimals is "3" in JS and "2" in .NET).
module Rbt.Spreadsheet.NumberText

open System
open System.Text

let private inv = Globalization.CultureInfo.InvariantCulture

#if FABLE_COMPILER
[<Fable.Core.Emit("$0.toExponential($1)")>]
let private toExponential (x: float) (fractionDigits: int) : string = Fable.Core.Util.jsNative

[<Fable.Core.Emit("$0.toExponential()")>]
let private toShortest (x: float) : string = Fable.Core.Util.jsNative
#else
let private toExponential (x: float) (fractionDigits: int) : string =
    x.ToString("E" + string fractionDigits, inv)

let private toShortest (x: float) : string = x.ToString("R", inv)
#endif

/// A non-negative decimal: value = 0.DIGITS x 10^Scale. Digits carry no
/// leading or trailing zeros; zero is the empty string with Scale 0.
[<Struct>]
type private Digits = { Text: string; Scale: int }

/// Read the digits out of plain or exponential notation ("123.45", "1.2e-7").
let private parseDigits (text: string) : Digits =
    let body = text.TrimStart('-', '+')
    let ePos = body.IndexOfAny [| 'e'; 'E' |]
    let mantissa, exponent =
        if ePos < 0 then body, 0
        else body.Substring(0, ePos), int (body.Substring(ePos + 1).TrimStart('+'))
    let dot = mantissa.IndexOf '.'
    let intPart, fracPart =
        if dot < 0 then mantissa, ""
        else mantissa.Substring(0, dot), mantissa.Substring(dot + 1)
    let all = intPart + fracPart
    let mutable leading = 0
    while leading < all.Length && all.[leading] = '0' do
        leading <- leading + 1
    let digits = all.Substring(leading).TrimEnd('0')
    if digits.Length = 0 then { Text = ""; Scale = 0 }
    else { Text = digits; Scale = intPart.Length - leading + exponent }

/// Keep the first `count` digits. A dropped remainder of exactly one half
/// rounds up when `tiesToEven` is false, to the even neighbour otherwise.
let private roundDigits (tiesToEven: bool) (d: Digits) (count: int) : Digits =
    if count >= d.Text.Length then d
    elif count < 0 then { Text = ""; Scale = 0 }
    else
        let kept = d.Text.Substring(0, count).ToCharArray()
        let isTie = d.Text.[count] = '5' && d.Text.Length = count + 1
        let mutable carry =
            if tiesToEven && isTie then count > 0 && (int kept.[count - 1] - int '0') % 2 = 1
            else d.Text.[count] >= '5'
        let mutable i = count - 1
        while carry && i >= 0 do
            if kept.[i] = '9' then
                kept.[i] <- '0'
            else
                kept.[i] <- char (int kept.[i] + 1)
                carry <- false
            i <- i - 1
        let text, scale =
            if carry then "1" + String(kept), d.Scale + 1
            else String(kept), d.Scale
        let trimmed = text.TrimEnd('0')
        if trimmed.Length = 0 then { Text = ""; Scale = 0 }
        else { Text = trimmed; Scale = scale }

/// The decimal expansion of |x|, exact for any double whose expansion fits in
/// 101 significant digits (which covers every tie the formats can round).
let private exactDigits (x: float) : Digits = parseDigits (toExponential (abs x) 100)

/// Correctly rounded significant digits, as the runtime produces them.
let private significant (x: float) (precision: int) : Digits =
    roundDigits true (exactDigits x) precision

let private appendExponent (sb: StringBuilder) (exponent: int) (marker: char) (minDigits: int) (positiveSign: bool) =
    sb.Append(marker) |> ignore
    if exponent < 0 then sb.Append('-') |> ignore
    elif positiveSign then sb.Append('+') |> ignore
    sb.Append((string (abs exponent)).PadLeft(minDigits, '0')) |> ignore

let private nonFinite (x: float) : string voption =
    if Double.IsNaN x then ValueSome "NaN"
    elif Double.IsPositiveInfinity x then ValueSome "Infinity"
    elif Double.IsNegativeInfinity x then ValueSome "-Infinity"
    else ValueNone

/// "G" / "G<precision>": fixed notation while the exponent is within the
/// precision, scientific beyond it. Precision 0 means shortest round-trip.
let private formatGeneral (x: float) (precision: int) (marker: char) : string =
    let digits, limit =
        if precision <= 0 then
            // Shortest round-trip digits stay in fixed notation up to 17 places.
            let d = parseDigits (toShortest (abs x))
            d, max d.Text.Length 17
        else significant x precision, precision
    let sb = StringBuilder()
    if x < 0.0 && digits.Text.Length > 0 then sb.Append('-') |> ignore
    if digits.Text.Length = 0 then
        sb.Append('0') |> ignore
    elif digits.Scale > limit || digits.Scale < -3 then
        sb.Append(digits.Text.[0]) |> ignore
        if digits.Text.Length > 1 then
            sb.Append('.').Append(digits.Text.Substring 1) |> ignore
        appendExponent sb (digits.Scale - 1) marker 2 true
    elif digits.Scale <= 0 then
        sb.Append("0.").Append(String('0', -digits.Scale)).Append(digits.Text) |> ignore
    elif digits.Scale >= digits.Text.Length then
        sb.Append(digits.Text).Append(String('0', digits.Scale - digits.Text.Length)) |> ignore
    else
        sb.Append(digits.Text.Substring(0, digits.Scale)).Append('.').Append(digits.Text.Substring digits.Scale) |> ignore
    sb.ToString()

/// "F<decimals>" / "N<decimals>": fixed decimals, optionally grouped.
let private formatFixed (x: float) (decimals: int) (grouped: bool) : string =
    let exact = exactDigits x
    let digits = roundDigits true exact (exact.Scale + decimals)
    // Lay the digits out around the decimal point, zero-filled on both sides.
    let intPart =
        if digits.Scale <= 0 then "0"
        elif digits.Scale >= digits.Text.Length then digits.Text + String('0', digits.Scale - digits.Text.Length)
        else digits.Text.Substring(0, digits.Scale)
    let fracPart =
        let frac =
            if digits.Scale <= 0 then String('0', -digits.Scale) + digits.Text
            elif digits.Scale >= digits.Text.Length then ""
            else digits.Text.Substring digits.Scale
        frac.PadRight(decimals, '0')
    let sb = StringBuilder()
    if x < 0.0 then sb.Append('-') |> ignore
    if grouped then
        for i in 0 .. intPart.Length - 1 do
            if i > 0 && (intPart.Length - i) % 3 = 0 then sb.Append(',') |> ignore
            sb.Append(intPart.[i]) |> ignore
    else
        sb.Append(intPart) |> ignore
    if decimals > 0 then sb.Append('.').Append(fracPart) |> ignore
    sb.ToString()

/// Custom numeric pattern (`0`, `#`, `.`, `,`, `%`, `E+0`, quoted and escaped
/// literals), following .NET's two-pass algorithm: scan the pattern for digit
/// placeholders and scaling, round the number to fit, then emit.
let private formatCustom (x: float) (format: string) : string =
    let mutable digitCount = 0
    let mutable decimalPos = -1
    let mutable firstDigit = Int32.MaxValue
    let mutable lastDigit = 0
    let mutable scientific = false
    let mutable thousandPos = -1
    let mutable thousandCount = 0
    let mutable thousandSeps = false
    let mutable scaleAdjust = 0
    let exponentFollows (src: int) =
        (src < format.Length && format.[src] = '0')
        || (src + 1 < format.Length && (format.[src] = '+' || format.[src] = '-') && format.[src + 1] = '0')

    let mutable src = 0
    while src < format.Length && format.[src] <> ';' do
        let ch = format.[src]
        src <- src + 1
        match ch with
        | '#' -> digitCount <- digitCount + 1
        | '0' ->
            if firstDigit = Int32.MaxValue then firstDigit <- digitCount
            digitCount <- digitCount + 1
            lastDigit <- digitCount
        | '.' -> if decimalPos < 0 then decimalPos <- digitCount
        | ',' ->
            if digitCount > 0 && decimalPos < 0 then
                if thousandPos >= 0 && thousandPos = digitCount then
                    thousandCount <- thousandCount + 1
                else
                    if thousandPos >= 0 then thousandSeps <- true
                    thousandPos <- digitCount
                    thousandCount <- 1
        | '%' -> scaleAdjust <- scaleAdjust + 2
        | '‰' -> scaleAdjust <- scaleAdjust + 3
        | '\'' | '"' ->
            while src < format.Length && format.[src] <> ch do
                src <- src + 1
            if src < format.Length then src <- src + 1
        | '\\' -> if src < format.Length then src <- src + 1
        | 'E' | 'e' ->
            if exponentFollows src then
                src <- src + 1
                while src < format.Length && format.[src] = '0' do
                    src <- src + 1
                scientific <- true
        | _ -> ()

    if decimalPos < 0 then decimalPos <- digitCount
    if thousandPos >= 0 then
        if thousandPos = decimalPos then scaleAdjust <- scaleAdjust - thousandCount * 3
        else thousandSeps <- true

    // .NET formats doubles through custom patterns at 15 significant digits.
    let number =
        let d = significant x 15
        if d.Text.Length = 0 then d
        else
            let scaled = { d with Scale = d.Scale + scaleAdjust }
            roundDigits false scaled (if scientific then digitCount else scaled.Scale + digitCount - decimalPos)
    let isZero = number.Text.Length = 0

    let firstDigit = if firstDigit < decimalPos then decimalPos - firstDigit else 0
    let lastDigit = if lastDigit > decimalPos then decimalPos - lastDigit else 0
    let mutable digPos = 0
    let mutable adjust = 0
    if scientific then
        digPos <- decimalPos
    else
        digPos <- if number.Scale > decimalPos then number.Scale else decimalPos
        adjust <- number.Scale - decimalPos

    let sb = StringBuilder()
    let mutable cur = 0
    let nextDigit () =
        if cur < number.Text.Length then
            let c = number.Text.[cur]
            cur <- cur + 1
            ValueSome c
        else ValueNone
    let appendDigit (c: char) =
        sb.Append(c) |> ignore
        if thousandSeps && digPos > 1 && (digPos - 1) % 3 = 0 then sb.Append(',') |> ignore
    let mutable decimalWritten = false

    src <- 0
    while src < format.Length && format.[src] <> ';' do
        let ch = format.[src]
        src <- src + 1
        // Integer digits beyond the pattern's placeholders go out before the
        // first placeholder or the decimal point.
        if adjust > 0 && (ch = '#' || ch = '0' || ch = '.') then
            while adjust > 0 do
                appendDigit (match nextDigit () with ValueSome c -> c | ValueNone -> '0')
                digPos <- digPos - 1
                adjust <- adjust - 1
        match ch with
        | '#' | '0' ->
            let digit =
                if adjust < 0 then
                    adjust <- adjust + 1
                    if digPos <= firstDigit then ValueSome '0' else ValueNone
                else
                    match nextDigit () with
                    | ValueSome c -> ValueSome c
                    | ValueNone -> if digPos > lastDigit then ValueSome '0' else ValueNone
            match digit with
            | ValueSome c -> appendDigit c
            | ValueNone -> ()
            digPos <- digPos - 1
        | '.' ->
            if digPos = 0 && not decimalWritten then
                if lastDigit < 0 || (decimalPos < digitCount && cur < number.Text.Length) then
                    sb.Append('.') |> ignore
                    decimalWritten <- true
        | '‰' -> sb.Append('‰') |> ignore
        | '%' -> sb.Append('%') |> ignore
        | ',' -> ()
        | '\'' | '"' ->
            while src < format.Length && format.[src] <> ch do
                sb.Append(format.[src]) |> ignore
                src <- src + 1
            if src < format.Length then src <- src + 1
        | '\\' ->
            if src < format.Length then
                sb.Append(format.[src]) |> ignore
                src <- src + 1
        | 'E' | 'e' ->
            if scientific && exponentFollows src then
                let positiveSign = format.[src] = '+'
                if format.[src] <> '0' then src <- src + 1
                let mutable zeros = 0
                while src < format.Length && format.[src] = '0' do
                    zeros <- zeros + 1
                    src <- src + 1
                appendExponent sb (if isZero then 0 else number.Scale - decimalPos) ch (min zeros 10) positiveSign
                scientific <- false
            else
                sb.Append(ch) |> ignore
                if src < format.Length && (format.[src] = '+' || format.[src] = '-') then
                    sb.Append(format.[src]) |> ignore
                    src <- src + 1
                while src < format.Length && format.[src] = '0' do
                    sb.Append('0') |> ignore
                    src <- src + 1
        | c -> sb.Append(c) |> ignore
    // A negative that rounds to zero keeps its sign ("-0.00") unless the
    // pattern then renders nothing at all, as in .NET.
    if x < 0.0 && (not isZero || sb.Length > 0) then "-" + sb.ToString()
    else sb.ToString()

/// "E<decimals>": one integer digit, fixed decimals, three exponent digits.
let private formatExponential (x: float) (decimals: int) (marker: char) : string =
    let digits = significant x (decimals + 1)
    let sb = StringBuilder()
    if x < 0.0 then sb.Append('-') |> ignore
    let padded = digits.Text.PadRight(decimals + 1, '0')
    sb.Append(padded.[0]) |> ignore
    if decimals > 0 then sb.Append('.').Append(padded.Substring 1) |> ignore
    appendExponent sb (if digits.Text.Length = 0 then 0 else digits.Scale - 1) marker 3 true
    sb.ToString()

/// `Double.ToString(format, InvariantCulture)` for the general (G), fixed (F),
/// number (N) and exponential (E) standard formats and for custom patterns.
let formatInvariant (format: string) (x: float) : string =
    match nonFinite x with
    | ValueSome text -> text
    | ValueNone ->
        let isStandard =
            format.Length >= 1 && format.Length <= 3
            && Char.IsLetter format.[0]
            && (format.Length = 1 || Seq.forall Char.IsDigit (format.Substring 1))
        if String.IsNullOrEmpty format then formatGeneral x 0 'E'
        elif isStandard then
            let precision = if format.Length = 1 then -1 else int (format.Substring 1)
            match format.[0] with
            | 'G' -> formatGeneral x (max precision 0) 'E'
            | 'g' -> formatGeneral x (max precision 0) 'e'
            | 'F' | 'f' -> formatFixed x (if precision < 0 then 2 else precision) false
            | 'N' | 'n' -> formatFixed x (if precision < 0 then 2 else precision) true
            | 'E' | 'e' -> formatExponential x (if precision < 0 then 6 else precision) format.[0]
            | _ -> formatCustom x format
        else formatCustom x format

/// Format with invariant-culture .NET semantics on every target: the runtime
/// formatter on .NET, the portable one under Fable.
let numberToString (format: string) (x: float) : string =
#if FABLE_COMPILER
    formatInvariant format x
#else
    x.ToString(format, inv)
#endif

/// Whether x and y are the same number to `precision` significant digits, each rounded from its
/// exact binary value, a half away from zero - Excel's comparison of numbers (`=`, `<`, ...,
/// COUNTIF's criteria) at 15 digits: 1+22 ulp is 1.00000000000000 and 1+23 ulp 1.00000000000001;
/// 1E15+5 is 1.00000000000001E15, 1E14+2.5 is 100000000000003 (checked in Excel 2019).
let sameToSignificantDigits (precision: int) (x: float) (y: float) : bool =
    if x = y then true
    elif (x < 0.0) <> (y < 0.0) || x = 0.0 || y = 0.0 then false
    else
        let a = roundDigits false (exactDigits x) precision
        let b = roundDigits false (exactDigits y) precision
        a.Text = b.Text && a.Scale = b.Scale

/// x rounded to `precision` significant digits, a half away from zero (CEILING, FLOOR and MROUND round
/// their quotient so before they take its integer). x itself when it is 0, not finite, or an integer
/// of at most 15 digits.
let roundToSignificantDigits (precision: int) (x: float) : float =
    if x = 0.0 || Double.IsNaN x || Double.IsInfinity x || (x = Math.Floor x && abs x < 1e15) then x
    else
        let d = roundDigits false (exactDigits x) precision
        if d.Text.Length = 0 then 0.0
        else
            let v = float ("0." + d.Text + "e" + string d.Scale)
            if x < 0.0 then -v else v

/// A number's text as Excel reads it - typed into a cell, written in a formula, or text turned into
/// a number (VALUE, `"1.5"*1`, a criteria's number): its first 15 significant digits, the rest cut,
/// not rounded (`1.0000000000000099` is 1, `123456789012345.6` is 123456789012345; checked in
/// Excel 2019). The text as it is when it has no more digits than that.
let cutToSignificantDigits (text: string) : string =
    let mutable digits = 0
    let mutable significant = 0
    let mutable i = 0
    while i < text.Length && text.[i] <> 'e' && text.[i] <> 'E' do
        let c = text.[i]
        if c >= '0' && c <= '9' then
            digits <- digits + 1
            if c <> '0' || significant > 0 then significant <- significant + 1
        i <- i + 1
    if digits <= 15 || significant <= 15 then text
    else
        let chars = text.ToCharArray()
        let mutable seen = 0
        let mutable j = 0
        while j < chars.Length && chars.[j] <> 'e' && chars.[j] <> 'E' do
            let c = chars.[j]
            if c >= '0' && c <= '9' then
                if c <> '0' || seen > 0 then
                    seen <- seen + 1
                    if seen > 15 then chars.[j] <- '0'
            j <- j + 1
        String(chars)
