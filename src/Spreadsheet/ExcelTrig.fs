/// SIN, COS, TAN, COT, SEC and CSC as Excel computes them (Excel 2019, 462 probe cells bit for bit,
/// tools/excel-parity/trig-*.tsv):
///
///  - The argument is reduced modulo pi/2 where pi is the x87's 64-bit FLDPI (0xC90FDAA22168C235 x 2^-62,
///    5.0e-20 above the true pi) - the reason SIN(1E+6) and the like differ from the C library from the
///    10th digit on - and the sine, cosine or tangent of the reduced argument is then correctly rounded.
///  - An argument of 2^27 or more in size is #NUM!.
///  - COT, SEC and CSC are 1 / TAN, 1 / COS and 1 / SIN, each of the rounded double.
///
/// The reduction is exact (pi/2 split in three parts of 26, 26 and 12 bits, so every product with the
/// quotient, below 2^27, is exact) and the series are summed in double-double arithmetic, so the result
/// is the correctly rounded value on both runtimes.
module Rbt.Spreadsheet.ExcelTrig

open System

[<Struct>]
type private DD = { Hi: float; Lo: float }

let private twoSum (a: float) (b: float) =
    let s = a + b
    let bb = s - a
    { Hi = s; Lo = (a - (s - bb)) + (b - bb) }
let private quick (a: float) (b: float) =
    let s = a + b
    { Hi = s; Lo = b - (s - a) }
let private splitter = 134217729.0
let private twoProd (a: float) (b: float) =
    let p = a * b
    let t = splitter * a
    let ah = t - (t - a)
    let al = a - ah
    let u = splitter * b
    let bh = u - (u - b)
    let bl = b - bh
    { Hi = p; Lo = ((ah * bh - p) + ah * bl + al * bh) + al * bl }
let private add (x: DD) (y: DD) =
    let s = twoSum x.Hi y.Hi
    let t = twoSum x.Lo y.Lo
    let a = quick s.Hi (s.Lo + t.Hi)
    quick a.Hi (a.Lo + t.Lo)
let private neg (x: DD) = { Hi = -x.Hi; Lo = -x.Lo }
let private mul (x: DD) (y: DD) =
    let p = twoProd x.Hi y.Hi
    quick p.Hi (p.Lo + (x.Hi * y.Lo + x.Lo * y.Hi))
let private div (x: DD) (y: DD) =
    let q1 = x.Hi / y.Hi
    let r = add x (neg (mul y { Hi = q1; Lo = 0.0 }))
    let q2 = r.Hi / y.Hi
    let r = add r (neg (mul y { Hi = q2; Lo = 0.0 }))
    let q3 = r.Hi / y.Hi
    add (quick q1 q2) { Hi = q3; Lo = 0.0 }

// 1/n! for n = 0 .. 27 in double-double: the head the rounded double, the tail the rest, computed once
// from the exact product (each 1/n! is formed from 1/(n-1)! in double-double and is good to ~1e-32).
let private inverseFactorials : DD[] =
    let table = Array.zeroCreate 28
    table.[0] <- { Hi = 1.0; Lo = 0.0 }
    for n in 1 .. 27 do
        table.[n] <- div table.[n - 1] { Hi = float n; Lo = 0.0 }
    table

// FLDPI / 2 = 0xC90FDAA22168C235 x 2^-63, in three parts.
let private p1 = 0x3243F6A |> float |> fun m -> m * Math.Pow(2.0, -25.0)                // top 26 bits
let private p2 = 0x222168C |> float |> fun m -> m * Math.Pow(2.0, -51.0)                // next 26
let private p3 = 0x235 |> float |> fun m -> m * Math.Pow(2.0, -63.0)                    // last 12
let private halfPi64 = p1 + p2 + p3

/// The largest argument (in size) Excel takes: 2^27.
let limit = 134217728.0

/// x - q (FLDPI / 2), exactly, as double-double, and q mod 4.
let private reduce (x: float) =
    let q = Math.Round(x / halfPi64, MidpointRounding.AwayFromZero)
    let t1 = x - q * p1
    let a = twoSum t1 (-(q * p2))
    let b = twoSum a.Hi (-(q * p3))
    let r = twoSum b.Hi (a.Lo + b.Lo)
    let k = int (q - 4.0 * Math.Floor(q / 4.0))
    r, k

let private sinSeries (r: DD) =
    let r2 = mul r r
    let mutable term = r
    let mutable s = r
    for k in 1 .. 13 do
        term <- mul term r2
        let t = mul term inverseFactorials.[2 * k + 1]
        s <- add s (if k % 2 = 0 then t else neg t)
    s
let private cosSeries (r: DD) =
    let r2 = mul r r
    let mutable term = { Hi = 1.0; Lo = 0.0 }
    let mutable s = term
    for k in 1 .. 13 do
        term <- mul term r2
        let t = mul term inverseFactorials.[2 * k]
        s <- add s (if k % 2 = 0 then t else neg t)
    s

let private sinCos (x: float) =
    let r, k = reduce x
    let s = sinSeries r
    let c = cosSeries r
    match k with
    | 0 -> s, c
    | 1 -> c, neg s
    | 2 -> neg s, neg c
    | _ -> neg c, s

let private value (x: DD) = x.Hi + x.Lo

/// sin x as Excel (x below the limit in size).
let sin (x: float) = if x = 0.0 then x else value (fst (sinCos x))
/// cos x as Excel.
let cos (x: float) = value (snd (sinCos x))
/// tan x as Excel; NaN where the cosine is 0.
let tan (x: float) =
    if x = 0.0 then x
    else
        let s, c = sinCos x
        if c.Hi = 0.0 then nan else value (div s c)
