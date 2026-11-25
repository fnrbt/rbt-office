/// Complex numbers of the function table: COMPLEX and the IM* functions.
/// A complex number is text ("3+4i", "-2.5j", "i"); what is Excel-specific is
/// how that text is read and written, as found against Excel 2019
/// (tools/excel-parity/make-complex-probe.py):
/// - reading: the whole text trimmed of spaces; a last character i or j
///   (lower case) makes the part before it the imaginary coefficient, split
///   from a real part at the first + or - after the first character that
///   does not follow an e or E; each part is a number as VALUE reads it
///   ("1,000", "$3", "3%", "(4)", a date, a time), an empty or lone sign
///   coefficient 1 or -1. Anything else is #NUM!, as is a number beyond
///   9.99999999999999E+307 or one too small to have an exponent; a number
///   below the smallest normal double is 0.
/// - writing: each part to 15 significant digits, written out in full
///   unless that takes more than 21 characters (then 1.5E-19, 1E+21); a zero
///   part is left out (both zero: "0"), a coefficient of exactly 1 or -1 is
///   just the suffix, which is the arguments' ("j" when they all have it).
module Rbt.Spreadsheet.Functions.FnComplex

open System
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Functions.Kit

let moduleName = "Complex"

/// A complex number and its suffix: 'i', 'j', or ' ' when it had none (a
/// real number), which goes with either.
[<Struct>]
type Complex = { Re: float; Im: float; Suffix: char }

let private smallestNormal = 2.2250738585072014E-308
let private largest = 9.99999999999999E+307

/// A number as Excel holds one: none beyond the largest, a denormal is 0.
let private held (x: float) : Result<float, string> =
    if Double.IsNaN x || Double.IsInfinity x || abs x > largest then Result.Error Err.num
    elif abs x < smallestNormal then Ok 0.0
    else Ok x

/// The decimal exponent written in a number's text ("1.5e-400" -400 + 0),
/// to tell a number too small to hold (#NUM!) from one that is 0.
let private tooSmall (text: string) (value: float) =
    value = 0.0 && text |> Seq.exists (fun c -> c >= '1' && c <= '9')
    && (let i = text.IndexOfAny [| 'e'; 'E' |] in i >= 0)

/// One part of a complex text as VALUE reads it.
let private partValue (text: string) : Result<float, string> =
    if text.IndexOfAny [| '\t'; '\n'; '\r'; ' ' |] >= 0 then Result.Error Err.num
    else
        match coerceNumber (CellValue.String text) with
        | Ok v ->
            let v = if tooSmall (text.Trim()) v then nan else v
            if Double.IsNaN v then Result.Error Err.num
            elif abs v > largest || Double.IsInfinity v then Result.Error Err.num
            elif abs v < smallestNormal then
                // 5e-324 is no number Excel can read; 1e-308 is 0.
                let t = text.Trim()
                let e = t.IndexOfAny [| 'e'; 'E' |]
                if e > 0 && v <> 0.0 && (match Int32.TryParse(t.Substring(e + 1)) with | true, x -> x < -308 | _ -> false) then Result.Error Err.num
                else Ok 0.0
            else Ok v
        | Result.Error _ -> Result.Error Err.num

/// The coefficient before the suffix: nothing or a lone + is 1, a lone - is -1.
let private coefficient (text: string) : Result<float, string> =
    match text with
    | "" | "+" -> Ok 1.0
    | "-" -> Ok -1.0
    | _ -> partValue text

/// A complex number written as text.
let parseComplex (raw: string) : Result<Complex, string> =
    // Arabic-Indic digits are digits to Excel ("٣+4i" is 3+4i).
    let raw =
        if raw |> Seq.exists (fun c -> (c >= '\u0660' && c <= '\u0669') || (c >= '\u06F0' && c <= '\u06F9')) then
            raw |> String.map (fun c ->
                if c >= '\u0660' && c <= '\u0669' then char (int '0' + int c - 0x0660)
                elif c >= '\u06F0' && c <= '\u06F9' then char (int '0' + int c - 0x06F0)
                else c)
        else raw
    let text = raw.Trim(' ')
    if text = "" then Result.Error Err.num
    else
        let last = text.[text.Length - 1]
        if last <> 'i' && last <> 'j' then
            partValue text |> Result.map (fun re -> { Re = re; Im = 0.0; Suffix = ' ' })
        else
            let body = text.Substring(0, text.Length - 1)
            // The operator between the parts: the first + or - after the first
            // character that is not an exponent's sign.
            let mutable split = -1
            let mutable k = 1
            while split < 0 && k < body.Length do
                let c = body.[k]
                if (c = '+' || c = '-') && not (body.[k - 1] = 'e' || body.[k - 1] = 'E') then split <- k
                k <- k + 1
            if split < 0 then
                coefficient body |> Result.map (fun im -> { Re = 0.0; Im = im; Suffix = last })
            else
                match partValue (body.Substring(0, split)) with
                | Result.Error e -> Result.Error e
                | Ok re ->
                    let rest = body.Substring(split + 1)
                    let sign = if body.[split] = '-' then -1.0 else 1.0
                    // After the operator the coefficient may carry one sign of its own.
                    let valid = not (rest.StartsWith "+" || rest.StartsWith "-") || not (rest.Length > 1 && (rest.[1] = '+' || rest.[1] = '-'))
                    if not valid then Result.Error Err.num
                    else
                        match coefficient rest with
                        | Ok im -> Ok { Re = re; Im = sign * im; Suffix = last }
                        | Result.Error e -> Result.Error e

/// A complex argument: text read as above, a number its real part, a blank 0;
/// a logical is #VALUE!, an argument left empty #N/A.
let private complexOf (value: CellValue) : Result<Complex, string> =
    match value with
    | CellValue.Error e -> Result.Error e
    | CellValue.Boolean _ -> Result.Error Err.value
    | CellValue.Empty -> Ok { Re = 0.0; Im = 0.0; Suffix = ' ' }
    | CellValue.Number n -> held n |> Result.map (fun re -> { Re = re; Im = 0.0; Suffix = ' ' })
    | CellValue.String s -> parseComplex s
    | other ->
        match coerceText other with
        | Ok s -> parseComplex s
        | Result.Error e -> Result.Error e

let private complexAt (a: Args) (i: int) : Result<Complex, string> =
    if not (a.Supplied i) then Result.Error Err.na else complexOf (a.Value i)

/// A number to 15 significant digits as COMPLEX writes it: in full when that
/// takes at most 21 characters (the sign not counted), else 1.5E-19.
let formatPart (x: float) : string =
    if x = 0.0 then "0"
    else
        let s = Rbt.Spreadsheet.NumberText.numberToString "E14" (abs x)
        let e = s.IndexOf 'E'
        let digits = (string s.[0] + s.Substring(2, e - 2)).TrimEnd('0')
        let exponent = int (s.Substring(e + 1))
        let digits = if digits = "" then "0" else digits
        let full =
            if exponent >= 0 then
                if digits.Length <= exponent + 1 then digits + System.String('0', exponent + 1 - digits.Length)
                else digits.Substring(0, exponent + 1) + "." + digits.Substring(exponent + 1)
            else "0." + System.String('0', -exponent - 1) + digits
        let text =
            if full.Length <= 21 then full
            else
                let mantissa = if digits.Length > 1 then string digits.[0] + "." + digits.Substring 1 else digits
                let ex = abs exponent
                mantissa + "E" + (if exponent < 0 then "-" else "+") + (if ex < 10 then "0" + string ex else string ex)
        if x < 0.0 then "-" + text else text

/// A complex number as text.
let formatComplex (re: float) (im: float) (suffix: char) : string =
    let suffix = if suffix = 'j' then "j" else "i"
    let imText () =
        if im = 1.0 then suffix
        elif im = -1.0 then "-" + suffix
        else formatPart im + suffix
    if im = 0.0 then formatPart re
    elif re = 0.0 then imText ()
    else formatPart re + (if im > 0.0 then "+" else "") + imText ()

/// A complex result: a part that overflowed (or reached the largest double)
/// or is no number is #NUM!; one below the smallest normal number is 0.
let private result (re: float) (im: float) (suffix: char) : Result<CellValue, string> =
    let part (x: float) =
        if Double.IsNaN x || abs x >= Double.MaxValue then Result.Error Err.num
        elif abs x < smallestNormal then Ok 0.0
        else Ok x
    res {
        let! re = part re
        let! im = part im
        return CellValue.String (formatComplex re im suffix) }

/// The suffix of two arguments: one with none takes the other's; i with j is #VALUE!.
let private joinSuffix (s: char) (t: char) : Result<char, string> =
    if s = ' ' then Ok t
    elif t = ' ' || s = t then Ok s
    else Result.Error Err.value

let private unary (name: string) (f: Complex -> Result<CellValue, string>) =
    scalar name 1 1 "|S" (fun a -> res {
        let! z = complexAt a 0
        return! f z })

/// A function of one complex number giving one.
let private complexFn (name: string) (f: float -> float -> Result<float * float, string>) =
    unary name (fun z -> res {
        let! re, im = f z.Re z.Im
        return! result re im z.Suffix })

let private hypot (x: float) (y: float) =
    let x = abs x
    let y = abs y
    let big = max x y
    if big = 0.0 then 0.0
    else
        let small = min x y
        let r = small / big
        big * Math.Sqrt (1.0 + r * r)

/// |z| as IMSQRT, IMPOWER and IMLN take it: sqrt(x^2 + y^2), which
/// overflows past 1E154 and underflows (a square below the smallest normal
/// number is 0) below 1E-154.
let private naiveAbs (x: float) (y: float) =
    let flush (v: float) = if abs v < smallestNormal then 0.0 else v
    Math.Sqrt (flush (flush (x * x) + flush (y * y)))

let private argument (x: float) (y: float) = Math.Atan2 (y, x)

// Excel's sines and cosines are the x87's: its argument reduction takes pi
// as the 64-bit pi of FLDPI, 5.0166E-20 above the true pi, which shows in a
// result near zero - IMSQRT("-1") is 6.1257422745431E-17+i, not
// 6.12323399573677E-17+i - and it takes no argument of 2^63 or more (#NUM!).
// To first order sin(x * pi / pi64).
let private piSkew = 5.016557612668332e-20 / Math.PI
let private x87Limit = 9223372036854775808.0
let private sinX (x: float) = if abs x >= x87Limit then nan else Math.Sin x - Math.Cos x * x * piSkew
let private cosX (x: float) = if abs x >= x87Limit then nan else Math.Cos x + Math.Sin x * x * piSkew

/// sinh and cosh from e^x and e^-x, as IMSIN and IMCOS compute them (the
/// difference loses digits near zero: IMSIN("1e-5+1e-5i")).
let private sinhOfExp (x: float) = (Math.Exp x - Math.Exp (-x)) / 2.0
let private coshOfExp (x: float) = (Math.Exp x + Math.Exp (-x)) / 2.0

/// e^x held to the largest double, as IMEXP has it: IMEXP("800+1i") is
/// 9.71297746009434E+307+1.51270661257499E+308i, IMEXP("710") #NUM!.
let private expHeld (x: float) = let e = Math.Exp x in if Double.IsInfinity e then Double.MaxValue else e

let private cexp (x: float) (y: float) = let m = expHeld x in m * cosX y, m * sinX y
let private csin (x: float) (y: float) = sinX x * coshOfExp y, cosX x * sinhOfExp y
let private ccos (x: float) (y: float) = cosX x * coshOfExp y, -(sinX x * sinhOfExp y)
/// sinh and cosh as IMSINH, IMCOSH, IMTAN and the reciprocals take them:
/// accurate near zero, but no number once e^|x| overflows (IMSINH("710") is
/// #NUM! though sinh 710 is a double).
let private sinhAcc (x: float) = if Double.IsInfinity (Math.Exp (abs x)) then nan else Math.Sinh x
let private coshAcc (x: float) = if Double.IsInfinity (Math.Exp (abs x)) then nan else Math.Cosh x
let private csinh (x: float) (y: float) = sinhAcc x * cosX y, coshAcc x * sinX y
let private ccosh (x: float) (y: float) = coshAcc x * cosX y, sinhAcc x * sinX y

/// tan z = (sin 2x + i sinh 2y) / (cos 2x + cosh 2y), i sign(y) once e^2|y|
/// overflows (IMTAN("1+800i") is i); IMCOT is its reciprocal.
let private ctan (x: float) (y: float) =
    let ch = Math.Cosh (2.0 * y)
    if Double.IsInfinity (Math.Exp (abs (2.0 * y))) then 0.0, float (sign y)
    else
        let d = cosX (2.0 * x) + ch
        sinX (2.0 * x) / d, Math.Sinh (2.0 * y) / d
/// (a + bi) / (c + di) by Smith's method (no square of the divisor to
/// overflow or underflow: IMDIV("1","1e-200") is 1E+200): #NUM! when c + di is 0.
let private divide (a: float) (b: float) (c: float) (d: float) : Result<float * float, string> =
    if c = 0.0 && d = 0.0 then Result.Error Err.num
    elif abs c >= abs d then
        let r = d / c
        let den = c + d * r
        Ok ((a + b * r) / den, (b - a * r) / den)
    else
        let r = c / d
        let den = c * r + d
        Ok ((a * r + b) / den, (b * r - a) / den)

/// 1 / (c + di) as the reciprocal functions compute it: 1/c when d is 0,
/// else (c - di) / (c^2 + d^2), #NUM! when that square underflows
/// (IMCSC("1e-300") is 1E+300, IMCSC("-1e-300i") #NUM!).
let private invert (c: float) (d: float) : Result<float * float, string> =
    if d = 0.0 then
        if c = 0.0 then Result.Error Err.num else Ok (1.0 / c, 0.0)
    else
        let flush (v: float) = if abs v < smallestNormal then 0.0 else v
        let den = flush (flush (c * c) + flush (d * d))
        if den = 0.0 then Result.Error Err.num else Ok (c / den, -d / den)

/// 1 / f(z), 0 once the hyperbolic part of f(z) - cosh of `big` - overflows
/// (IMCSC("1+800i") and IMCSC("1e300+1e300i") are 0).
let private reciprocal (f: float -> float -> float * float) (big: float) (x: float) (y: float) =
    if Double.IsInfinity (Math.Exp (abs big)) then Ok (0.0, 0.0)
    else
        let re, im = f x y
        invert re im
let private csinAcc (x: float) (y: float) = sinX x * coshAcc y, cosX x * sinhAcc y
let private ccosAcc (x: float) (y: float) = cosX x * coshAcc y, -(sinX x * sinhAcc y)

/// ln z: of a real or imaginary z ln of its one part, else of the naive |z|
/// (IMLN("1e200+1e200i") is #NUM!, IMLN("1e300") 690.775527898214).
let private ln (x: float) (y: float) : Result<float * float, string> =
    let r = if y = 0.0 then abs x elif x = 0.0 then abs y else naiveAbs x y
    if r = 0.0 || Double.IsInfinity r then Result.Error Err.num
    else Ok (Math.Log r, argument x y)

/// z to the power n in polar form.
let private power (x: float) (y: float) (n: float) : Result<float * float, string> =
    if x = 0.0 && y = 0.0 then
        if n > 0.0 then Ok (0.0, 0.0) else Result.Error Err.num
    else
        let r = Math.Pow (naiveAbs x y, n)
        let t = argument x y * n
        Ok (r * cosX t, r * sinX t)

/// The complex numbers of an aggregate's arguments (IMSUM, IMPRODUCT): of a
/// range or array every value but blank cells; their suffix.
let private aggregateArgs (a: Args) : Result<Complex list * char, string> =
    let values = ResizeArray<Complex>()
    let mutable failure = None
    for i in 0 .. a.Count - 1 do
        if failure.IsNone then
            if not (a.Supplied i) then
                if i = 0 then failure <- Some Err.na
            elif a.IsRangeLike i then
                for v in a.Flat i do
                    if failure.IsNone then
                        match v with
                        | v ->
                            match complexOf v with
                            | Ok z -> values.Add z
                            | Result.Error e -> failure <- Some e
            else
                match complexAt a i with
                | Ok z -> values.Add z
                | Result.Error e -> failure <- Some e
    match failure with
    | Some e -> Result.Error e
    | None ->
        let mutable suffix = Ok ' '
        for z in values do suffix <- suffix |> Result.bind (fun s -> joinSuffix s z.Suffix)
        suffix |> Result.map (fun s -> List.ofSeq values, s)

/// IMDIV and IMSUB: two complex numbers of different suffixes are #NUM!.
let private binary (name: string) (f: Complex -> Complex -> Result<float * float, string>) =
    scalar name 2 2 "|S" (fun a -> res {
        let! z = complexAt a 0
        let! w = complexAt a 1
        let! suffix = joinSuffix z.Suffix w.Suffix |> Result.mapError (fun _ -> Err.num)
        let! re, im = f z w
        return! result re im suffix })

let functions : FunctionDef list =
    [ scalar "COMPLEX" 2 3 "|S" (fun a -> res {
        let! re = a.AtpNumber 0
        let! im = a.AtpNumber 1
        let! suffix =
            if not (a.Supplied 2) then Ok 'i'
            else
                match a.Value 2 with
                | CellValue.String "i" | CellValue.String "" | CellValue.Empty -> Ok 'i'
                | CellValue.String "j" -> Ok 'j'
                | CellValue.Error e -> Result.Error e
                | _ -> Result.Error Err.value
        return! result re im suffix })
      unary "IMREAL" (fun z -> Ok (CellValue.Number z.Re))
      unary "IMAGINARY" (fun z -> Ok (CellValue.Number z.Im))
      unary "IMABS" (fun z -> numOk (hypot z.Re z.Im))
      unary "IMARGUMENT" (fun z ->
        if z.Re = 0.0 && z.Im = 0.0 then Result.Error Err.div0 else numOk (argument z.Re z.Im))
      complexFn "IMCONJUGATE" (fun x y -> Ok (x, -y))
      complexFn "IMEXP" (fun x y -> Ok (cexp x y))
      complexFn "IMLN" ln
      complexFn "IMLOG10" (fun x y -> ln x y |> Result.map (fun (re, im) -> re / Math.Log 10.0, im / Math.Log 10.0))
      complexFn "IMLOG2" (fun x y -> ln x y |> Result.map (fun (re, im) -> re / Math.Log 2.0, im / Math.Log 2.0))
      complexFn "IMSQRT" (fun x y ->
        let r = Math.Sqrt (naiveAbs x y)
        if r = 0.0 then Ok (0.0, 0.0)
        else
            let t = argument x y / 2.0
            Ok (r * cosX t, r * sinX t))
      complexFn "IMSIN" (fun x y -> Ok (csin x y))
      complexFn "IMCOS" (fun x y -> Ok (ccos x y))
      complexFn "IMSINH" (fun x y -> Ok (csinh x y))
      complexFn "IMCOSH" (fun x y -> Ok (ccosh x y))
      complexFn "IMTAN" (fun x y -> Ok (ctan x y))
      complexFn "IMCOT" (fun x y -> let re, im = ctan x y in invert re im)
      complexFn "IMSEC" (fun x y -> reciprocal ccosAcc y x y)
      complexFn "IMCSC" (fun x y -> reciprocal csinAcc y x y)
      complexFn "IMSECH" (fun x y -> reciprocal ccosh x x y)
      complexFn "IMCSCH" (fun x y -> reciprocal csinh x x y)
      scalar "IMPOWER" 2 2 "|S" (fun a -> res {
        let! z = complexAt a 0
        let! n = a.AtpNumber 1
        let! re, im = power z.Re z.Im n
        return! result re im z.Suffix })
      binary "IMDIV" (fun z w -> divide z.Re z.Im w.Re w.Im)
      binary "IMSUB" (fun z w -> Ok (z.Re - w.Re, z.Im - w.Im))
      scalar "IMSUM" 1 255 "|R" (fun a -> res {
        let! values, suffix = aggregateArgs a
        let re = values |> List.sumBy (fun z -> z.Re)
        let im = values |> List.sumBy (fun z -> z.Im)
        return! result re im suffix })
      scalar "IMPRODUCT" 1 255 "|R" (fun a -> res {
        let! values, suffix = aggregateArgs a
        let re, im = values |> List.fold (fun (re, im) z -> re * z.Re - im * z.Im, re * z.Im + im * z.Re) (1.0, 0.0)
        if Double.IsInfinity re || Double.IsInfinity im || Double.IsNaN re || Double.IsNaN im then return! fail Err.value
        else return! result re im suffix }) ]

let install () = registerFunctions moduleName functions
