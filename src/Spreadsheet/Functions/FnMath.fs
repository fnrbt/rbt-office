/// Math and trigonometry functions of the function table (see Formula.fs, "The function table").
module Rbt.Spreadsheet.Functions.FnMath

open System
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Functions.Kit

let moduleName = "Math"

/// Excel's trigonometric functions refuse arguments of 2^27 and more.
let private trigLimit = 134217728.0

/// x / step as a whole number when it is one but for the error of binary fractions.
let private quotient (x: float) (step: float) =
    let q = x / step
    let r = Math.Round q
    if abs (q - r) <= 1e-12 * max 1.0 (abs q) then r else q

let private ceilingTo (x: float) (step: float) = Math.Ceiling (quotient x step) * step
let private floorTo (x: float) (step: float) = Math.Floor (quotient x step) * step

/// Multiples of |significance|: `up` toward +infinity (else toward -infinity);
/// `awayFromZero` turns it round for a negative number (the MATH variants' mode).
let private roundToMultiple (up: bool) (a: Args) (modeArg: int option) =
    res {
        let! x = a.Number 0
        let! significance = a.NumberOr 1 1.0
        let! mode =
            match modeArg with
            | Some i -> a.NumberOr i 0.0
            | None -> Ok 0.0
        let step = abs significance
        if step = 0.0 || x = 0.0 then return CellValue.Number 0.0
        else
            let reverse = x < 0.0 && mode <> 0.0
            let goUp = if reverse then not up else up
            return! numOk (if goUp then ceilingTo x step else floorTo x step)
    }

let private romanDigits = [| 'M'; 'D'; 'C'; 'L'; 'X'; 'V'; 'I' |]
let private romanValues = [| 1000; 500; 100; 50; 10; 5; 1 |]

/// ROMAN's text of a number of 0..3999 in one of its five forms (0 classic ...
/// 4 most concise): the subtraction of a digit before a larger one is let
/// reach further the larger the form.
let private romanText (number: int) (form: int) =
    let sb = Text.StringBuilder()
    let mutable value = number
    let maxIndex = romanValues.Length - 1
    for i in 0 .. maxIndex / 2 do
        let mutable index = 2 * i
        let digit = value / romanValues.[index]
        if digit % 5 = 4 then
            let index2 = if digit = 4 then index - 1 else index - 2
            let mutable steps = 0
            while steps < form && index < maxIndex do
                steps <- steps + 1
                if romanValues.[index2] - romanValues.[index + 1] <= value then index <- index + 1
                else steps <- form
            sb.Append(romanDigits.[index]).Append(romanDigits.[index2]) |> ignore
            value <- value + romanValues.[index] - romanValues.[index2]
        else
            if digit > 4 then sb.Append(romanDigits.[index - 1]) |> ignore
            sb.Append(String.replicate (digit % 5) (string romanDigits.[index])) |> ignore
            value <- value % romanValues.[index]
    sb.ToString()

let private romanValue (c: char) =
    match Char.ToUpperInvariant c with
    | 'I' -> 1 | 'V' -> 5 | 'X' -> 10 | 'L' -> 50 | 'C' -> 100 | 'D' -> 500 | 'M' -> 1000
    | _ -> 0

let private digitValue (c: char) =
    if c >= '0' && c <= '9' then int c - int '0'
    elif c >= 'A' && c <= 'Z' then int c - int 'A' + 10
    elif c >= 'a' && c <= 'z' then int c - int 'a' + 10
    else 99

/// A whole number of 0 or more in base `radix` (upper-case digits).
let private digitsIn (radix: int) (number: float) =
    let mutable text = ""
    let mutable rest = number
    while rest >= 1.0 do
        let d = int (rest - Math.Floor (rest / float radix) * float radix)
        text <- string "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ".[d] + text
        rest <- Math.Floor (rest / float radix)
    if text = "" then "0" else text

/// (n1+n2+...)! / (n1! n2! ...), as a running product of binomials.
let private multinomial (numbers: float[]) =
    let mutable total = 0.0
    let mutable result = 1.0
    for x in numbers do
        let mutable i = 1.0
        while i <= x do
            total <- total + 1.0
            result <- result * total / i
            i <- i + 1.0
    result

/// A sum over the pairs of two arrays (SUMX2MY2, SUMX2PY2, SUMXMY2).
let private pairSum (name: string) (term: float -> float -> float) =
    scalar name 2 2 "RR" (fun a -> res {
        let! pairs = pairedNumbers a 0 1
        return! numOk (pairs |> Array.sumBy (fun (x, y) -> term x y)) })

let functions : FunctionDef list =
    [ numeric1 "ACOT" (fun x -> numOk (Math.PI / 2.0 - Math.Atan x))
      numeric1 "ACOTH" (fun x -> if abs x <= 1.0 then fail Err.num else numOk (0.5 * Math.Log ((x + 1.0) / (x - 1.0))))
      numeric1 "COTH" (fun x -> if x = 0.0 then fail Err.div0 elif abs x >= trigLimit then fail Err.num else numOk (1.0 / Math.Tanh x))
      numeric1 "CSC" (fun x -> if abs x >= trigLimit then fail Err.num elif x = 0.0 then fail Err.div0 else (let v = Rbt.Spreadsheet.ExcelTrig.sin x in if v = 0.0 then fail Err.div0 else numOk (1.0 / v)))
      numeric1 "CSCH" (fun x -> if abs x >= trigLimit then fail Err.num elif x = 0.0 then fail Err.div0 else numOk (1.0 / Math.Sinh x))
      numeric1 "SEC" (fun x -> if abs x >= trigLimit then fail Err.num else (let v = Rbt.Spreadsheet.ExcelTrig.cos x in if v = 0.0 then fail Err.div0 else numOk (1.0 / v)))
      numeric1 "SECH" (fun x -> if abs x >= trigLimit then fail Err.num else numOk (1.0 / Math.Cosh x))
      numeric1 "SQRTPI" (fun x -> if x < 0.0 then fail Err.num else numOk (Math.Sqrt (x * Math.PI)))
      numeric1 "FACTDOUBLE" (fun x ->
        let n = truncate x
        if n < -1.0 then fail Err.num
        else
            let mutable product = 1.0
            let mutable k = n
            while k > 1.0 do
                product <- product * k
                k <- k - 2.0
            numOk product)
      scalar "ARABIC" 1 1 "|S" (fun a -> res {
        let! text = a.Text 0
        let text = text.Trim()
        do! check (text.Length <= 255) Err.value
        let negative = text.StartsWith "-"
        let digits = if negative then text.Substring 1 else text
        let values = digits.ToCharArray() |> Array.map romanValue
        do! check (values |> Array.forall (fun v -> v > 0)) Err.value
        let total =
            values |> Array.mapi (fun i v -> if i + 1 < values.Length && v < values.[i + 1] then -v else v) |> Array.sum
        return CellValue.Number (float (if negative then -total else total)) })
      scalar "BASE" 2 3 "|S" (fun a -> res {
        let! number = a.Number 0
        let! radix = a.Number 1
        let! minLength = a.NumberOr 2 0.0
        let number = truncate number
        let radix = int (truncate radix)
        let minLength = int (truncate minLength)
        do! check (number >= 0.0 && number < 9007199254740992.0) Err.num
        do! check (radix >= 2 && radix <= 36) Err.num
        do! check (minLength >= 0 && minLength <= 255) Err.num
        let text = digitsIn radix number
        return CellValue.String (if text.Length < minLength then text.PadLeft(minLength, '0') else text) })
      scalar "DECIMAL" 2 2 "|S" (fun a -> res {
        let! text = a.Text 0
        let! radix = a.Number 1
        let radix = int (truncate radix)
        do! check (radix >= 2 && radix <= 36) Err.num
        do! check (text.Length <= 255) Err.num
        let text = text.Trim()
        let digits = text.ToCharArray() |> Array.map digitValue
        do! check (digits |> Array.forall (fun d -> d < radix)) Err.num
        return! numOk (digits |> Array.fold (fun total d -> total * float radix + float d) 0.0) })
      scalar "CEILING.MATH" 1 3 "|S" (fun a -> roundToMultiple true a (Some 2))
      scalar "FLOOR.MATH" 1 3 "|S" (fun a -> roundToMultiple false a (Some 2))
      scalar "CEILING.PRECISE" 1 2 "|S" (fun a -> roundToMultiple true a None)
      scalar "ISO.CEILING" 1 2 "|S" (fun a -> roundToMultiple true a None)
      // ECMA.CEILING is CEILING (Excel 2019: the same value or error as CEILING on 154 probe pairs; it takes
      // both arguments, a formula with one Excel does not accept).
      scalar "ECMA.CEILING" 2 2 "|S" (fun a -> res {
        let! x = a.Number 0
        let! significance = a.Number 1
        return a.Ev.Call "CEILING" [ Expr.Number x; Expr.Number significance ] a.Ctx })
      scalar "FLOOR.PRECISE" 1 2 "|S" (fun a -> roundToMultiple false a None)
      numeric2 "COMBINA" (fun n k ->
        let n = truncate n
        let k = truncate k
        if n < 0.0 || k < 0.0 then fail Err.num
        elif k = 0.0 then Ok (CellValue.Number 1.0)
        elif n = 0.0 then Ok (CellValue.Number 0.0)
        else
            // C(n+k-1, k), multiplied out with the smaller of k and n-1 factors.
            let top = n + k - 1.0
            let r = min k (top - k)
            let mutable c = 1.0
            let mutable i = 1.0
            while i <= r do
                c <- c * (top - r + i) / i
                i <- i + 1.0
            numOk (Math.Round c))
      scalar "MULTINOMIAL" 1 -1 "|R" (fun a -> res {
        let! numbers = a.Numbers 0
        let numbers = numbers |> Array.map truncate
        do! check (numbers |> Array.forall (fun x -> x >= 0.0)) Err.num
        return! numOk (Math.Round (multinomial numbers)) })
      arrayed "MUNIT" 1 1 "" (fun a -> res {
        let! n = a.Number 0
        let n = int (truncate n)
        do! check (n >= 1) Err.value
        do! check (n <= 16384) Err.value
        return [| for r in 0 .. n - 1 -> [| for c in 0 .. n - 1 -> CellValue.Number (if r = c then 1.0 else 0.0) |] |] })
      scalar "PERCENTOF" 2 2 "RR" (fun a -> res {
        let! part = coerceNumber (a.Ev.Call "SUM" [ a.Expr 0 ] a.Ctx)
        let! whole = coerceNumber (a.Ev.Call "SUM" [ a.Expr 1 ] a.Ctx)
        if whole = 0.0 then return! fail Err.div0 else return! numOk (part / whole) })
      scalar "ROMAN" 1 2 "|S" (fun a -> res {
        let! number = a.Number 0
        let! form =
            if not (a.Supplied 1) then Ok 0.0
            else
                match a.Value 1 with
                | CellValue.Boolean b -> Ok (if b then 0.0 else 4.0)
                | v -> coerceNumber v
        let number = truncate number
        let form = int (truncate form)
        do! check (number >= 0.0 && number <= 3999.0) Err.value
        do! check (form >= 0 && form <= 4) Err.value
        return CellValue.String (romanText (int number) form) })
      scalar "SERIESSUM" 4 4 "SSSR" (fun a -> res {
        let! x = a.Number 0
        let! n = a.Number 1
        let! m = a.Number 2
        let coefficients = a.Flat 3
        match coefficients |> Array.tryPick (fun v -> match v with CellValue.Number _ -> None | CellValue.Error e -> Some e | _ -> Some Err.value) with
        | Some e -> return! fail e
        | None ->
            let terms = coefficients |> Array.mapi (fun i v -> match v with CellValue.Number c -> c * Math.Pow (x, n + float i * m) | _ -> 0.0)
            return! numOk (Array.sum terms) })
      pairSum "SUMX2MY2" (fun x y -> x * x - y * y)
      pairSum "SUMX2PY2" (fun x y -> x * x + y * y)
      pairSum "SUMXMY2" (fun x y -> (x - y) * (x - y)) ]

let install () = registerFunctions moduleName functions
