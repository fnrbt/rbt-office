/// Text functions of the function table: values and arrays as text, DOLLAR,
/// ENCODEURL, and the byte-counting variants of the text functions.
module Rbt.Spreadsheet.Functions.FnText

open System
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
