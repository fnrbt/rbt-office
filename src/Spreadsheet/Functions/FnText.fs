/// Text functions of the function table: values and arrays as text, DOLLAR,
/// ENCODEURL, and the byte-counting variants of the text functions.
module Rbt.Spreadsheet.Functions.FnText

open System
open Rbt.Spreadsheet
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Functions.Kit

let moduleName = "Text"

/// A value as VALUETOTEXT and ARRAYTOTEXT write it: `strict` puts text in
/// quotes (doubling the quotes in it), as a formula would write it.
let private valueText (strict: bool) (value: CellValue) =
    match value with
    | CellValue.Empty -> ""
    | CellValue.Number n -> numberToText n
    | CellValue.Boolean b -> if b then "TRUE" else "FALSE"
    | CellValue.Error e -> "#" + e
    | CellValue.String s -> if strict then "\"" + s.Replace("\"", "\"\"") + "\"" else s
    | CellValue.Formula _ -> ""

/// The format argument of VALUETOTEXT / ARRAYTOTEXT: 0 concise, 1 strict.
let private formatArg (a: Args) (i: int) =
    res {
        let! format = a.NumberOr i 0.0
        let format = truncate format
        do! check (format = 0.0 || format = 1.0) Err.value
        return format = 1.0 }

/// The same text function under its byte-counting name. Every character is
/// one byte here, as in Excel with a language of single-byte characters.
let private byteVariant (name: string) (minArgs: int) (maxArgs: int) (target: string) =
    { Name = name; MinArgs = minArgs; MaxArgs = maxArgs; Signature = "|S"; Volatile = false
      Impl = ScalarFunction (fun ev args ctx -> ev.Call target args ctx) }

/// Text percent-encoded as UTF-8: all but letters, digits and - _ . ~.
let private encodeUrl (text: string) =
    let sb = Text.StringBuilder()
    for b in Text.Encoding.UTF8.GetBytes text do
        let c = char b
        if (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c = '-' || c = '_' || c = '.' || c = '~' then
            sb.Append c |> ignore
        else sb.Append('%').Append("0123456789ABCDEF".[int b / 16]).Append("0123456789ABCDEF".[int b % 16]) |> ignore
    sb.ToString()

let private roundHalfAway (x: float) = if x < 0.0 then -Math.Round (-x, MidpointRounding.AwayFromZero) else Math.Round (x, MidpointRounding.AwayFromZero)

// ----- BAHTTEXT ----------------------------------------------------------------

let private thaiDigits = [| "ศูนย์"; "หนึ่ง"; "สอง"; "สาม"; "สี่"; "ห้า"; "หก"; "เจ็ด"; "แปด"; "เก้า" |]
let private thaiPlaces = [| ""; "สิบ"; "ร้อย"; "พัน"; "หมื่น"; "แสน" |]

/// The Thai words of a whole number written as decimal digits (no leading
/// zeros, not "0"): six-digit groups joined by ล้าน (a million), so 10^12 is
/// หนึ่งล้านล้าน; a ten of one is สิบ, a ten of two ยี่สิบ; a unit of one is
/// เอ็ด after any non-zero digit of the number (101 หนึ่งร้อยเอ็ด, 1001
/// หนึ่งพันเอ็ด, 10^12 + 1 หนึ่งล้านล้านเอ็ด), else หนึ่ง.
let private thaiWords (digits: string) =
    let sb = Text.StringBuilder()
    let mutable seen = false
    let n = digits.Length
    for i in 0 .. n - 1 do
        let d = int digits.[i] - int '0'
        let p = n - 1 - i
        let place = p % 6
        if d <> 0 then
            if place = 0 && d = 1 && seen then sb.Append "เอ็ด" |> ignore
            elif place = 1 && d = 1 then sb.Append "สิบ" |> ignore
            elif place = 1 && d = 2 then sb.Append "ยี่สิบ" |> ignore
            else sb.Append(thaiDigits.[d]).Append(thaiPlaces.[place]) |> ignore
            seen <- true
        if place = 0 && p > 0 && seen then sb.Append "ล้าน" |> ignore
    sb.ToString()

/// BAHTTEXT: the number in Thai words with บาท (baht) and สตางค์ (satang).
/// Measured in Excel 2019 (make-bahttext-probe.py): the number is taken to 15
/// significant digits (1E+300 is หนึ่ง and fifty ล้าน), then rounded to the
/// satang half up in decimal (1.005 is one satang though its double is below
/// it); no satang is ถ้วน ("even"), no baht leaves the baht out (0.5 is
/// ห้าสิบสตางค์), zero is ศูนย์บาทถ้วน; a negative number has ลบ in front,
/// even when it rounds to zero (-0.004 is ลบศูนย์บาทถ้วน).
let private bahtText (x: float) =
    let text = NumberText.numberToString "E14" (abs x)
    let e = text.IndexOf 'E'
    let mantissa = text.Substring(0, e).Replace(".", "")
    let exponent = int (text.Substring(e + 1))
    // Scaled to satang: the digits with the decimal point after position exponent + 3.
    let pointAt = exponent + 3
    let whole, carry =
        if pointAt <= 0 then "", (pointAt = 0 && mantissa.[0] >= '5')
        elif pointAt >= mantissa.Length then mantissa + System.String('0', pointAt - mantissa.Length), false
        else mantissa.Substring(0, pointAt), mantissa.[pointAt] >= '5'
    let satangs =
        let digits = whole.TrimStart('0')
        let digits = if digits = "" then "0" else digits
        if not carry then digits
        else
            let chars = digits.ToCharArray()
            let mutable i = chars.Length - 1
            let mutable c = true
            while c && i >= 0 do
                if chars.[i] = '9' then chars.[i] <- '0'; i <- i - 1
                else chars.[i] <- char (int chars.[i] + 1); c <- false
            (if c then "1" else "") + System.String(chars)
    let padded = satangs.PadLeft(3, '0')
    let baht = padded.Substring(0, padded.Length - 2).TrimStart('0')
    let satang = padded.Substring(padded.Length - 2).TrimStart('0')
    let sb = Text.StringBuilder()
    if x < 0.0 then sb.Append "ลบ" |> ignore
    if baht <> "" then sb.Append(thaiWords baht).Append("บาท") |> ignore
    if satang <> "" then sb.Append(thaiWords satang).Append("สตางค์") |> ignore
    elif baht = "" then sb.Append("ศูนย์บาทถ้วน") |> ignore
    else sb.Append "ถ้วน" |> ignore
    sb.ToString()

let functions : FunctionDef list =
    [ scalar "VALUETOTEXT" 1 2 "|S" (fun a -> res {
        let! strict = formatArg a 1
        return CellValue.String (valueText strict (a.Value 0)) })
      scalar "ARRAYTOTEXT" 1 2 "RS" (fun a -> res {
        let! strict = formatArg a 1
        let rows = a.Array 0
        if strict then
            let body = rows |> Array.map (fun row -> String.Join(",", row |> Array.map (valueText true)))
            return CellValue.String ("{" + String.Join(";", body) + "}")
        else return CellValue.String (String.Join(", ", rows |> Array.concat |> Array.map (valueText false))) })
      scalar "DOLLAR" 1 2 "|S" (fun a -> res {
        let! x = a.Number 0
        let! decimals = a.NumberOr 1 2.0
        let decimals = truncate decimals
        do! check (decimals <= 127.0) Err.value
        let factor = Math.Pow (10.0, decimals)
        let rounded = roundHalfAway (x * factor) / factor
        let! digits = coerceText (a.Ev.Call "FIXED" [ Number (abs rounded); Number (max 0.0 decimals) ] a.Ctx)
        return CellValue.String (if rounded < 0.0 then "($" + digits + ")" else "$" + digits) })
      scalar "BAHTTEXT" 1 1 "|S" (fun a -> res {
        let! x = a.Number 0
        return CellValue.String (bahtText x) })
      scalar "ENCODEURL" 1 1 "|S" (fun a -> res {
        let! text = a.Text 0
        return CellValue.String (encodeUrl text) })
      byteVariant "LENB" 1 1 "LEN"
      byteVariant "LEFTB" 1 2 "LEFT"
      byteVariant "RIGHTB" 1 2 "RIGHT"
      byteVariant "MIDB" 3 3 "MID"
      byteVariant "FINDB" 2 3 "FIND"
      byteVariant "SEARCHB" 2 3 "SEARCH"
      byteVariant "REPLACEB" 4 4 "REPLACE" ]

let install () = registerFunctions moduleName functions
