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

// ---------------------------------------------------------------------------
// ASC and DBCS: full-width (double-byte) and half-width (single-byte) forms, in an Excel whose editing
// language is a double-byte one (`isDoubleByteLanguage`); otherwise the text as it is. The pairs are
// Unicode's <wide> and <narrow> compatibility forms - what Windows' LCMapString maps - with the
// katakana's voiced marks: ASC splits ガ into ｶﾞ, DBCS joins ｶﾞ into ガ (a mark it cannot join
// becomes the full-width ゛). (Not checked against an Excel with a Japanese editing language: the
// parity VM's has none.)
// ---------------------------------------------------------------------------

/// (half-width, full-width) pairs besides ASCII's (U+0021..U+007E with U+FF01..U+FF5E, the space with U+3000).
let private widthPairs =
    [| 0x2985, 0xFF5F; 0x2986, 0xFF60; 0xFF61, 0x3002; 0xFF62, 0x300C; 0xFF63, 0x300D; 0xFF64, 0x3001; 0xFF65, 0x30FB; 0xFF66, 0x30F2
       0xFF67, 0x30A1; 0xFF68, 0x30A3; 0xFF69, 0x30A5; 0xFF6A, 0x30A7; 0xFF6B, 0x30A9; 0xFF6C, 0x30E3; 0xFF6D, 0x30E5; 0xFF6E, 0x30E7
       0xFF6F, 0x30C3; 0xFF70, 0x30FC; 0xFF71, 0x30A2; 0xFF72, 0x30A4; 0xFF73, 0x30A6; 0xFF74, 0x30A8; 0xFF75, 0x30AA; 0xFF76, 0x30AB
       0xFF77, 0x30AD; 0xFF78, 0x30AF; 0xFF79, 0x30B1; 0xFF7A, 0x30B3; 0xFF7B, 0x30B5; 0xFF7C, 0x30B7; 0xFF7D, 0x30B9; 0xFF7E, 0x30BB
       0xFF7F, 0x30BD; 0xFF80, 0x30BF; 0xFF81, 0x30C1; 0xFF82, 0x30C4; 0xFF83, 0x30C6; 0xFF84, 0x30C8; 0xFF85, 0x30CA; 0xFF86, 0x30CB
       0xFF87, 0x30CC; 0xFF88, 0x30CD; 0xFF89, 0x30CE; 0xFF8A, 0x30CF; 0xFF8B, 0x30D2; 0xFF8C, 0x30D5; 0xFF8D, 0x30D8; 0xFF8E, 0x30DB
       0xFF8F, 0x30DE; 0xFF90, 0x30DF; 0xFF91, 0x30E0; 0xFF92, 0x30E1; 0xFF93, 0x30E2; 0xFF94, 0x30E4; 0xFF95, 0x30E6; 0xFF96, 0x30E8
       0xFF97, 0x30E9; 0xFF98, 0x30EA; 0xFF99, 0x30EB; 0xFF9A, 0x30EC; 0xFF9B, 0x30ED; 0xFF9C, 0x30EF; 0xFF9D, 0x30F3; 0xFF9E, 0x309B
       0xFF9F, 0x309C; 0xFFA0, 0x3164; 0xFFA1, 0x3131; 0xFFA2, 0x3132; 0xFFA3, 0x3133; 0xFFA4, 0x3134; 0xFFA5, 0x3135; 0xFFA6, 0x3136
       0xFFA7, 0x3137; 0xFFA8, 0x3138; 0xFFA9, 0x3139; 0xFFAA, 0x313A; 0xFFAB, 0x313B; 0xFFAC, 0x313C; 0xFFAD, 0x313D; 0xFFAE, 0x313E
       0xFFAF, 0x313F; 0xFFB0, 0x3140; 0xFFB1, 0x3141; 0xFFB2, 0x3142; 0xFFB3, 0x3143; 0xFFB4, 0x3144; 0xFFB5, 0x3145; 0xFFB6, 0x3146
       0xFFB7, 0x3147; 0xFFB8, 0x3148; 0xFFB9, 0x3149; 0xFFBA, 0x314A; 0xFFBB, 0x314B; 0xFFBC, 0x314C; 0xFFBD, 0x314D; 0xFFBE, 0x314E
       0xFFC2, 0x314F; 0xFFC3, 0x3150; 0xFFC4, 0x3151; 0xFFC5, 0x3152; 0xFFC6, 0x3153; 0xFFC7, 0x3154; 0xFFCA, 0x3155; 0xFFCB, 0x3156
       0xFFCC, 0x3157; 0xFFCD, 0x3158; 0xFFCE, 0x3159; 0xFFCF, 0x315A; 0xFFD2, 0x315B; 0xFFD3, 0x315C; 0xFFD4, 0x315D; 0xFFD5, 0x315E
       0xFFD6, 0x315F; 0xFFD7, 0x3160; 0xFFDA, 0x3161; 0xFFDB, 0x3162; 0xFFDC, 0x3163; 0xA2, 0xFFE0; 0xA3, 0xFFE1; 0xAC, 0xFFE2
       0xAF, 0xFFE3; 0xA6, 0xFFE4; 0xA5, 0xFFE5; 0x20A9, 0xFFE6; 0xFFE8, 0x2502; 0xFFE9, 0x2190; 0xFFEA, 0x2191; 0xFFEB, 0x2192
       0xFFEC, 0x2193; 0xFFED, 0x25A0; 0xFFEE, 0x25CB |]

/// Katakana with a voiced (゛) or semi-voiced (゜) mark: (it, its letter, the mark U+3099 / U+309A).
let private voiced =
    [| 0x30AC, 0x30AB, 0x3099; 0x30AE, 0x30AD, 0x3099; 0x30B0, 0x30AF, 0x3099; 0x30B2, 0x30B1, 0x3099; 0x30B4, 0x30B3, 0x3099
       0x30B6, 0x30B5, 0x3099; 0x30B8, 0x30B7, 0x3099; 0x30BA, 0x30B9, 0x3099; 0x30BC, 0x30BB, 0x3099; 0x30BE, 0x30BD, 0x3099
       0x30C0, 0x30BF, 0x3099; 0x30C2, 0x30C1, 0x3099; 0x30C5, 0x30C4, 0x3099; 0x30C7, 0x30C6, 0x3099; 0x30C9, 0x30C8, 0x3099
       0x30D0, 0x30CF, 0x3099; 0x30D1, 0x30CF, 0x309A; 0x30D3, 0x30D2, 0x3099; 0x30D4, 0x30D2, 0x309A; 0x30D6, 0x30D5, 0x3099
       0x30D7, 0x30D5, 0x309A; 0x30D9, 0x30D8, 0x3099; 0x30DA, 0x30D8, 0x309A; 0x30DC, 0x30DB, 0x3099; 0x30DD, 0x30DB, 0x309A
       0x30F4, 0x30A6, 0x3099; 0x30F7, 0x30EF, 0x3099; 0x30F8, 0x30F0, 0x3099; 0x30F9, 0x30F1, 0x3099; 0x30FA, 0x30F2, 0x3099
       0x30FE, 0x30FD, 0x3099 |]

let private toFull = lazy (dict [ for (half, full) in widthPairs -> half, full ])
let private toHalf = lazy (dict [ for (half, full) in widthPairs -> full, half ])
let private splitVoiced = lazy (dict [ for (k, letter, mark) in voiced -> k, (letter, mark) ])
let private joinVoiced = lazy (dict [ for (k, letter, mark) in voiced -> letter * 4 + (if mark = 0x3099 then 0 else 1), k ])

let private halfOf (cp: int) =
    if cp >= 0xFF01 && cp <= 0xFF5E then Some (cp - 0xFEE0)
    elif cp = 0x3000 then Some 0x20
    else
        match toHalf.Force().TryGetValue cp with
        | true, h -> Some h
        | _ -> None

let private fullOf (cp: int) =
    if cp >= 0x21 && cp <= 0x7E then Some (cp + 0xFEE0)
    elif cp = 0x20 then Some 0x3000
    else
        match toFull.Force().TryGetValue cp with
        | true, f -> Some f
        | _ -> None

/// A text as code points (all of these are of the Basic Multilingual Plane: a surrogate stays as it is).
let private unitsOf (s: string) = [| for ch in s -> int ch |]
let private textOfUnits (cps: int seq) =
    let sb = Text.StringBuilder()
    for cp in cps do sb.Append(char cp) |> ignore
    sb.ToString()

/// ASC's conversion: full-width forms to half-width; `onlyKana` keeps all but the katakana (and their
/// marks) as they are - the half-width katakana of a furigana of type halfwidthKatakana.
let toHalfWidth (onlyKana: bool) (text: string) =
    let out = ResizeArray<int>()
    for cp in unitsOf text do
        let kana = (cp >= 0x3099 && cp <= 0x30FF) || cp = 0x3001 || cp = 0x3002 || cp = 0x300C || cp = 0x300D
        if onlyKana && not kana then out.Add cp
        else
            match splitVoiced.Force().TryGetValue cp with
            | true, (letter, mark) when (halfOf letter).IsSome ->
                out.Add (halfOf letter).Value
                out.Add (if mark = 0x3099 then 0xFF9E else 0xFF9F)
            | _ ->
                if cp = 0x3099 then out.Add 0xFF9E
                elif cp = 0x309A then out.Add 0xFF9F
                else out.Add (defaultArg (halfOf cp) cp)
    textOfUnits out

/// DBCS's conversion: half-width forms to full-width, a half-width katakana and the voiced mark after
/// it to the katakana with the mark.
let toFullWidth (text: string) =
    let cps = unitsOf text
    let out = ResizeArray<int>()
    let mutable i = 0
    while i < cps.Length do
        let cp = cps.[i]
        let full = defaultArg (fullOf cp) cp
        let next = if i + 1 < cps.Length then cps.[i + 1] else -1
        let joined =
            if (next = 0xFF9E || next = 0xFF9F) && cp >= 0xFF66 && cp <= 0xFF9D then
                match joinVoiced.Force().TryGetValue (full * 4 + (if next = 0xFF9E then 0 else 1)) with
                | true, k -> Some k
                | _ -> None
            else None
        match joined with
        | Some k ->
            out.Add k
            i <- i + 2
        | None ->
            out.Add full
            i <- i + 1
    textOfUnits out

/// Hiragana to katakana (`toKatakana`) or the other way.
let private kanaCase (toKatakana: bool) (text: string) =
    textOfUnits (
        unitsOf text
        |> Array.map (fun cp ->
            if toKatakana && cp >= 0x3041 && cp <= 0x3096 then cp + 0x60
            elif not toKatakana && cp >= 0x30A1 && cp <= 0x30F6 then cp - 0x60
            else cp))

/// PHONETIC: the furigana of the cells of a reference, one after another (a cell without one gives its
/// text; numbers, logicals, errors and blank cells nothing), in the kana of each reading's type; a
/// reference of several areas is #N/A. In an Excel without a double-byte editing language every
/// reference is #N/A (the parity VM's Excel 2019, furigana in the file or not).
let private phonetic (a: Args) : Result<CellValue, string> =
    match a.Reference 0 with
    | Some (ErrorLit e) -> Result.Error e
    | Some reference when isDoubleByteLanguage () ->
        let area =
            match reference with
            | CellRef r -> Some (r.Sheet, r.Row, r.Col, r.Row, r.Col)
            | RangeRef (x, y) -> Some (x.Sheet, min x.Row y.Row, min x.Col y.Col, max x.Row y.Row, max x.Col y.Col)
            | _ -> None
        match area with
        | None -> Result.Error Err.na
        | Some (sheet, top, left, bottom, right) ->
            match referenceModel sheet a.Ctx with
            | None -> Result.Error Err.ref
            | Some m ->
                let sb = Text.StringBuilder()
                for KeyValue(addr, cell) in m.Cells do
                    if addr.Row >= top && addr.Row <= bottom && addr.Col >= left && addr.Col <= right then
                        match cell.Value with
                        | CellValue.String s ->
                            match Map.tryFind addr m.Phonetics with
                            | Some p when p.Text = s ->
                                let reading =
                                    match p.Kind with
                                    | "Hiragana" -> kanaCase false p.Reading
                                    | "halfwidthKatakana" -> toHalfWidth true (kanaCase true p.Reading)
                                    | "noConversion" -> p.Reading
                                    | _ -> kanaCase true p.Reading
                                sb.Append(reading) |> ignore
                            | _ -> sb.Append(s) |> ignore
                        | CellValue.Formula (_, cached) ->
                            match unwrapValue cached with
                            | CellValue.String s -> sb.Append(s) |> ignore
                            | _ -> ()
                        | _ -> ()
                Ok (CellValue.String (sb.ToString()))
    | Some _ -> Result.Error Err.na
    | None ->
        // not a reference (Excel does not take such a formula): the value's error, else #N/A
        match a.Value 0 with
        | CellValue.Error e -> Result.Error e
        | _ -> Result.Error Err.na

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
      scalar "ASC" 1 1 "|S" (fun a -> res {
        let! text = a.Text 0
        return CellValue.String (if isDoubleByteLanguage () then toHalfWidth false text else text) })
      scalar "DBCS" 1 1 "|S" (fun a -> res {
        let! text = a.Text 0
        return CellValue.String (if isDoubleByteLanguage () then toFullWidth text else text) })
      scalar "PHONETIC" 1 1 "R" phonetic
      byteVariant "LENB" 1 1 "LEN"
      byteVariant "LEFTB" 1 2 "LEFT"
      byteVariant "RIGHTB" 1 2 "RIGHT"
      byteVariant "MIDB" 3 3 "MID"
      byteVariant "FINDB" 2 3 "FIND"
      byteVariant "SEARCHB" 2 3 "SEARCH"
      byteVariant "REPLACEB" 4 4 "REPLACE" ]

let install () = registerFunctions moduleName functions
