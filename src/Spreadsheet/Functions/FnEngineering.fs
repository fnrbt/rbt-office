/// Engineering functions of the function table: conversions between the
/// binary, octal and hexadecimal notations of ten digits, bit shifts, DELTA
/// and GESTEP. (BIN2DEC, DEC2BIN and the bitwise AND/OR/XOR are Formula.fs's.)
module Rbt.Spreadsheet.Functions.FnEngineering

open System
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Functions.Kit

let moduleName = "Engineering"

/// The digits a value stands for: a number as it is written ("1100100"), text as it is.
let private digitsText (value: CellValue) : Result<string, string> =
    match value with
    | CellValue.Boolean _ -> Result.Error Err.value
    | CellValue.Error e -> Result.Error e
    | CellValue.Number n -> Ok (numberToText n)
    | CellValue.Empty -> Ok ""
    | other -> coerceText other

/// The number written in `radix` with at most ten digits, a ten-digit one with
/// its top digit's high bit set being negative (two's complement).
let private parseTenDigits (radix: int) (text: string) : Result<float, string> =
    let text = text.Trim()
    if text.Length > 10 then Result.Error Err.num
    else
        let digit (c: char) =
            let d =
                if c >= '0' && c <= '9' then int c - int '0'
                elif c >= 'A' && c <= 'F' then int c - int 'A' + 10
                elif c >= 'a' && c <= 'f' then int c - int 'a' + 10
                else 99
            if d < radix then Some d else None
        let digits = text.ToCharArray() |> Array.map digit
        if digits |> Array.exists Option.isNone then Result.Error Err.num
        else
            let value = digits |> Array.fold (fun total d -> total * float radix + float d.Value) 0.0
            let full = Math.Pow (float radix, 10.0)
            Ok (if text.Length = 10 && value >= full / 2.0 then value - full else value)

/// A number written in `radix` within ten digits: a negative one as its two's
/// complement of ten digits (`places` then does not count), a positive one
/// padded to `places` digits when given.
let private writeTenDigits (radix: int) (value: float) (places: float option) : Result<CellValue, string> =
    let full = Math.Pow (float radix, 10.0)
    if value < -full / 2.0 || value >= full / 2.0 then Result.Error Err.num
    else
        let toText (v: float) =
            let mutable text = ""
            let mutable rest = v
            while rest >= 1.0 do
                let d = int (rest - Math.Floor (rest / float radix) * float radix)
                text <- string "0123456789ABCDEF".[d] + text
                rest <- Math.Floor (rest / float radix)
            if text = "" then "0" else text
        if value < 0.0 then Ok (CellValue.String (toText (full + value)))
        else
            let text = toText value
            match places with
            | None -> Ok (CellValue.String text)
            | Some p ->
                let p = truncate p
                if p < 0.0 || p > 10.0 || float text.Length > p then Result.Error Err.num
                else Ok (CellValue.String (text.PadLeft(int p, '0')))

let private convert (name: string) (fromRadix: int) (toRadix: int) =
    scalar name 1 2 "|S" (fun a -> res {
        let! text = digitsText (a.Value 0)
        let! value = parseTenDigits fromRadix text
        let! places = if a.Supplied 1 then a.Number 1 |> Result.map Some else Ok None
        return! writeTenDigits toRadix value places })

/// The largest number the bit functions take: 2^48 - 1.
let private bitLimit = 281474976710655.0

let private shift (name: string) (left: bool) =
    scalar name 2 2 "|S" (fun a -> res {
        let! number = a.Number 0
        let! amount = a.Number 1
        do! check (number >= 0.0 && number <= bitLimit && number = Math.Floor number) Err.num
        let amount = truncate amount
        do! check (abs amount <= 53.0) Err.num
        let amount = if left then amount else -amount
        let result = if amount >= 0.0 then number * Math.Pow (2.0, amount) else Math.Floor (number / Math.Pow (2.0, -amount))
        do! check (result <= bitLimit) Err.num
        return CellValue.Number result })

let functions : FunctionDef list =
    [ convert "BIN2HEX" 2 16
      convert "BIN2OCT" 2 8
      convert "HEX2BIN" 16 2
      convert "HEX2OCT" 16 8
      convert "OCT2BIN" 8 2
      convert "OCT2HEX" 8 16
      shift "BITLSHIFT" true
      shift "BITRSHIFT" false
      scalar "DELTA" 1 2 "|S" (fun a -> res {
        let! x = a.Number 0
        let! y = a.NumberOr 1 0.0
        return CellValue.Number (if x = y then 1.0 else 0.0) })
      scalar "GESTEP" 1 2 "|S" (fun a -> res {
        let! x = a.Number 0
        let! step = a.NumberOr 1 0.0
        return CellValue.Number (if x >= step then 1.0 else 0.0) }) ]

let install () = registerFunctions moduleName functions
