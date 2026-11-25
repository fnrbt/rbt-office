/// Engineering functions of the function table: conversions between the
/// binary, octal and hexadecimal notations of ten digits, bit shifts, DELTA
/// and GESTEP, the error function, the Bessel functions and CONVERT's units.
/// (BIN2DEC, DEC2BIN and the bitwise AND/OR/XOR are Formula.fs's.) Checked
/// against Excel 2019 with tools/excel-parity/make-engineering-probe.py.
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
        // `places` first, 1 to 10 (as DEC2BIN and the others in Formula.fs): OCT2HEX(TRUE,11) is #NUM!.
        let! places =
            if a.Supplied 1 then
                (match a.Value 1 with CellValue.Boolean _ -> Result.Error Err.value | v -> numberOf v)
                |> Result.bind (fun p -> let p = truncate p in if p < 1.0 || p > 10.0 then Result.Error Err.num else Ok (Some p))
            else Ok None
        let! text = digitsText (a.Value 0)
        let! value = parseTenDigits fromRadix text
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

/// A result as Excel holds it: no denormal numbers (ERFC(27) is 0).
let private held (x: float) = if abs x < 2.2250738585072014E-308 then 0.0 else x

// --- The error function ------------------------------------------------------------------
// fdlibm's erf and erfc (s_erf.c), accurate to an ulp, as Excel's are.

let private erfPoly (c: float[]) (x: float) = Array.foldBack (fun k acc -> k + x * acc) c 0.0

let private pp = [| 1.28379167095512558561e-01; -3.25042107247001499370e-01; -2.84817495755985104766e-02; -5.77027029648944159157e-03; -2.37630166566501626084e-05 |]
let private qq = [| 1.0; 3.97917223959155352819e-01; 6.50222499887672944485e-02; 5.08130628187576562776e-03; 1.32494738004321644526e-04; -3.96022827877536812320e-06 |]
let private pa = [| -2.36211856075265944077e-03; 4.14856118683748331666e-01; -3.72207876035701323847e-01; 3.18346619901161753674e-01; -1.10894694282396677476e-01; 3.54783043256182359371e-02; -2.16637559486879084300e-03 |]
let private qa = [| 1.0; 1.06420880400844228286e-01; 5.40397917702171048937e-01; 7.18286544141962662868e-02; 1.26171219808761642112e-01; 1.36370839120290507362e-02; 1.19844998467991074170e-02 |]
let private ra = [| -9.86494403484714822705e-03; -6.93858572707181764372e-01; -1.05586262253232909814e+01; -6.23753324503260060396e+01; -1.62396669462573470355e+02; -1.84605092906711035994e+02; -8.12874355063065934246e+01; -9.81432934416914548592e+00 |]
let private sa = [| 1.0; 1.96512716674392571292e+01; 1.37657754143519042600e+02; 4.34565877475229228821e+02; 6.45387271733267880336e+02; 4.29008140027567833386e+02; 1.08635005541779435134e+02; 6.57024977031928170135e+00; -6.04244152148580987438e-02 |]
let private rb = [| -9.86494292470009928597e-03; -7.99283237680523006574e-01; -1.77579549177547519889e+01; -1.60636384855821916062e+02; -6.37566443368389627722e+02; -1.02509513161107724954e+03; -4.83519191608651397019e+02 |]
let private sb = [| 1.0; 3.03380607434824582924e+01; 3.25792512996573918826e+02; 1.53672958608443695994e+03; 3.19985821950859553908e+03; 2.55305040643316442583e+03; 4.74528541206955367215e+02; -2.24409524465858183362e+01 |]
let private erx = 8.45062911510467529297e-01
let private efx = 1.28379167095512586316e-01

/// e^(-x^2) R(x) for the tails: x^2 split so its high part squares exactly.
let private tail (x: float) =
    let s = 1.0 / (x * x)
    let r, q =
        if x < 1.0 / 0.35 then erfPoly ra s, erfPoly sa s
        else erfPoly rb s, erfPoly sb s
    let c = 134217729.0 * x
    let z = c - (c - x)
    Math.Exp (-z * z - 0.5625) * Math.Exp ((z - x) * (z + x) + r / q) / x

/// erf and erfc of |x| < 0.84375 or 1.25 (the small and middle pieces).
let private erfSmall (x: float) = x + x * (erfPoly pp (x * x) / erfPoly qq (x * x))
let private erfMiddle (ax: float) = erfPoly pa (ax - 1.0) / erfPoly qa (ax - 1.0)

let erf (x: float) : float =
    let ax = abs x
    let r =
        if ax < 0.84375 then
            if ax < 3.7252902984e-09 then ax + efx * ax else erfSmall ax
        elif ax < 1.25 then erx + erfMiddle ax
        elif ax >= 6.0 then 1.0
        else 1.0 - tail ax
    if x < 0.0 then -r else r

let erfc (x: float) : float =
    let ax = abs x
    if ax < 0.84375 then 1.0 - (if ax < 1.3877787807814457e-17 then x else erfSmall ax * float (sign x))
    elif ax < 1.25 then
        let p = erfMiddle ax
        if x > 0.0 then 1.0 - erx - p else 1.0 + (erx + p)
    elif x < -6.0 then 2.0
    elif x > 28.0 then 0.0
    elif x > 0.0 then tail ax
    else 2.0 - tail ax

/// ERF(a, b) = erf b - erf a (taken as it is: ERF(5,6) is the difference of
/// two numbers near 1, as Excel's is).
let private erfBetween (a: float) (b: float) = erf b - erf a

// --- Bessel functions ----------------------------------------------------------------------
// Excel's are Numerical Recipes' (bessj0, bessj1, bessj, bessy*, bessi*,
// bessk*): their polynomial and rational approximations (good to about 1e-8,
// so BESSELJ(0,0) is 1.00000000283141), Miller's downward recurrence for J
// and I of order 2 and up, the upward one for Y and K.

let private besselJ0 (x: float) =
    let ax = abs x
    if ax < 8.0 then
        let y = x * x
        let a1 = 57568490574.0 + y * (-13362590354.0 + y * (651619640.7 + y * (-11214424.18 + y * (77392.33017 + y * (-184.9052456)))))
        let a2 = 57568490411.0 + y * (1029532985.0 + y * (9494680.718 + y * (59272.64853 + y * (267.8532712 + y * 1.0))))
        a1 / a2
    else
        let z = 8.0 / ax
        let y = z * z
        let xx = ax - 0.785398164
        let a1 = 1.0 + y * (-0.1098628627e-2 + y * (0.2734510407e-4 + y * (-0.2073370639e-5 + y * 0.2093887211e-6)))
        let a2 = -0.1562499995e-1 + y * (0.1430488765e-3 + y * (-0.6911147651e-5 + y * (0.7621095161e-6 - y * 0.934935152e-7)))
        Math.Sqrt (0.636619772 / ax) * (Math.Cos xx * a1 - z * Math.Sin xx * a2)

let private besselJ1 (x: float) =
    let ax = abs x
    if ax < 8.0 then
        let y = x * x
        let a1 = x * (72362614232.0 + y * (-7895059235.0 + y * (242396853.1 + y * (-2972611.439 + y * (15704.48260 + y * (-30.16036606))))))
        let a2 = 144725228442.0 + y * (2300535178.0 + y * (18583304.74 + y * (99447.43394 + y * (376.9991397 + y * 1.0))))
        a1 / a2
    else
        let z = 8.0 / ax
        let y = z * z
        let xx = ax - 2.356194491
        let a1 = 1.0 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * (-0.240337019e-6))))
        let a2 = 0.04687499995 + y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)))
        let r = Math.Sqrt (0.636619772 / ax) * (Math.Cos xx * a1 - z * Math.Sin xx * a2)
        if x < 0.0 then -r else r

let private besselJ (n: int) (x: float) =
    if n = 0 then besselJ0 x
    elif n = 1 then besselJ1 x
    else
        let ax = abs x
        if ax = 0.0 then 0.0
        else
            let r =
                if ax > float n then
                    let tox = 2.0 / ax
                    let mutable bjm = besselJ0 ax
                    let mutable bj = besselJ1 ax
                    for j in 1 .. n - 1 do
                        let bjp = float j * tox * bj - bjm
                        bjm <- bj
                        bj <- bjp
                    bj
                else
                    let tox = 2.0 / ax
                    let m = 2 * ((n + int (Math.Sqrt (40.0 * float n))) / 2)
                    let mutable jsum = false
                    let mutable bjp = 0.0
                    let mutable ans = 0.0
                    let mutable sum = 0.0
                    let mutable bj = 1.0
                    for j in m .. -1 .. 1 do
                        let bjm = float j * tox * bj - bjp
                        bjp <- bj
                        bj <- bjm
                        if abs bj > 1.0e10 then
                            bj <- bj * 1.0e-10
                            bjp <- bjp * 1.0e-10
                            ans <- ans * 1.0e-10
                            sum <- sum * 1.0e-10
                        if jsum then sum <- sum + bj
                        jsum <- not jsum
                        if j = n then ans <- bjp
                    sum <- 2.0 * sum - bj
                    ans / sum
            if x < 0.0 && n % 2 = 1 then -r else r

let private besselY0 (x: float) =
    if x < 8.0 then
        let y = x * x
        let a1 = -2957821389.0 + y * (7062834065.0 + y * (-512359803.6 + y * (10879881.29 + y * (-86327.92757 + y * 228.4622733))))
        let a2 = 40076544269.0 + y * (745249964.8 + y * (7189466.438 + y * (47447.26470 + y * (226.1030244 + y * 1.0))))
        a1 / a2 + 0.636619772 * besselJ0 x * Math.Log x
    else
        let z = 8.0 / x
        let y = z * z
        let xx = x - 0.785398164
        let a1 = 1.0 + y * (-0.1098628627e-2 + y * (0.2734510407e-4 + y * (-0.2073370639e-5 + y * 0.2093887211e-6)))
        let a2 = -0.1562499995e-1 + y * (0.1430488765e-3 + y * (-0.6911147651e-5 + y * (0.7621095161e-6 + y * (-0.934945152e-7))))
        Math.Sqrt (0.636619772 / x) * (Math.Sin xx * a1 + z * Math.Cos xx * a2)

let private besselY1 (x: float) =
    if x < 8.0 then
        let y = x * x
        let a1 = x * (-0.4900604943e13 + y * (0.1275274390e13 + y * (-0.5153438139e11 + y * (0.7349264551e9 + y * (-0.4237922726e7 + y * 0.8511937935e4)))))
        let a2 = 0.2499580570e14 + y * (0.4244419664e12 + y * (0.3733650367e10 + y * (0.2245904002e8 + y * (0.1020426050e6 + y * (0.3549632885e3 + y)))))
        a1 / a2 + 0.636619772 * (besselJ1 x * Math.Log x - 1.0 / x)
    else
        let z = 8.0 / x
        let y = z * z
        let xx = x - 2.356194491
        let a1 = 1.0 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * (-0.240337019e-6))))
        let a2 = 0.04687499995 + y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)))
        Math.Sqrt (0.636619772 / x) * (Math.Sin xx * a1 + z * Math.Cos xx * a2)

let private besselY (n: int) (x: float) =
    if n = 0 then besselY0 x
    elif n = 1 then besselY1 x
    else
        let tox = 2.0 / x
        let mutable by = besselY1 x
        let mutable bym = besselY0 x
        for j in 1 .. n - 1 do
            let byp = float j * tox * by - bym
            bym <- by
            by <- byp
        by

let private besselI0 (x: float) =
    let ax = abs x
    if ax < 3.75 then
        let y = (x / 3.75) * (x / 3.75)
        1.0 + y * (3.5156229 + y * (3.0899424 + y * (1.2067492 + y * (0.2659732 + y * (0.360768e-1 + y * 0.45813e-2)))))
    else
        let y = 3.75 / ax
        (Math.Exp ax / Math.Sqrt ax) * (0.39894228 + y * (0.1328592e-1 + y * (0.225319e-2 + y * (-0.157565e-2 + y * (0.916281e-2 + y * (-0.2057706e-1 + y * (0.2635537e-1 + y * (-0.1647633e-1 + y * 0.392377e-2))))))))

let private besselI1 (x: float) =
    let ax = abs x
    let r =
        if ax < 3.75 then
            let y = (x / 3.75) * (x / 3.75)
            ax * (0.5 + y * (0.87890594 + y * (0.51498869 + y * (0.15084934 + y * (0.2658733e-1 + y * (0.301532e-2 + y * 0.32411e-3))))))
        else
            let y = 3.75 / ax
            let a = 0.2282967e-1 + y * (-0.2895312e-1 + y * (0.1787654e-1 - y * 0.420059e-2))
            let a = 0.39894228 + y * (-0.3988024e-1 + y * (-0.362018e-2 + y * (0.163801e-2 + y * (-0.1031555e-1 + y * a))))
            a * (Math.Exp ax / Math.Sqrt ax)
    if x < 0.0 then -r else r

let private besselI (n: int) (x: float) =
    if n = 0 then besselI0 x
    elif n = 1 then besselI1 x
    elif x = 0.0 then 0.0
    else
        let tox = 2.0 / abs x
        let mutable bip = 0.0
        let mutable ans = 0.0
        let mutable bi = 1.0
        for j in 2 * (n + int (Math.Sqrt (40.0 * float n))) .. -1 .. 1 do
            let bim = bip + float j * tox * bi
            bip <- bi
            bi <- bim
            if abs bi > 1.0e10 then
                ans <- ans * 1.0e-10
                bi <- bi * 1.0e-10
                bip <- bip * 1.0e-10
            if j = n then ans <- bip
        let r = ans * besselI0 x / bi
        if x < 0.0 && n % 2 = 1 then -r else r

let private besselK0 (x: float) =
    if x <= 2.0 then
        let y = x * x / 4.0
        (-Math.Log (x / 2.0) * besselI0 x) + (-0.57721566 + y * (0.42278420 + y * (0.23069756 + y * (0.3488590e-1 + y * (0.262698e-2 + y * (0.10750e-3 + y * 0.74e-5))))))
    else
        let y = 2.0 / x
        (Math.Exp (-x) / Math.Sqrt x) * (1.25331414 + y * (-0.7832358e-1 + y * (0.2189568e-1 + y * (-0.1062446e-1 + y * (0.587872e-2 + y * (-0.251540e-2 + y * 0.53208e-3))))))

let private besselK1 (x: float) =
    if x <= 2.0 then
        let y = x * x / 4.0
        (Math.Log (x / 2.0) * besselI1 x) + (1.0 / x) * (1.0 + y * (0.15443144 + y * (-0.67278579 + y * (-0.18156897 + y * (-0.1919402e-1 + y * (-0.110404e-2 + y * (-0.4686e-4)))))))
    else
        let y = 2.0 / x
        (Math.Exp (-x) / Math.Sqrt x) * (1.25331414 + y * (0.23498619 + y * (-0.3655620e-1 + y * (0.1504268e-1 + y * (-0.780353e-2 + y * (0.325614e-2 + y * (-0.68245e-3)))))))

let private besselK (n: int) (x: float) =
    if n = 0 then besselK0 x
    elif n = 1 then besselK1 x
    else
        let tox = 2.0 / x
        let mutable bkm = besselK0 x
        let mutable bk = besselK1 x
        for j in 1 .. n - 1 do
            let bkp = bkm + float j * tox * bk
            bkm <- bk
            bk <- bkp
        bk

/// A BESSEL function: x and the order n (truncated; below 0 #NUM!); for J, Y
/// and I an x of 2^27 or more #NUM! (Excel's limit; K is 0 there).
let private bessel (name: string) (positiveX: bool) (f: int -> float -> float) =
    scalar name 2 2 "|S" (fun a -> res {
        let! x = a.AtpNumber 0
        let! n = a.AtpNumber 1
        do! check (n >= 0.0) Err.num
        let n = truncate n
        do! check (name = "BESSELK" || abs x < 134217728.0) Err.num
        do! check (not positiveX || x > 0.0) Err.num
        return! numOk (held (f (int n) x)) })

// --- CONVERT ---------------------------------------------------------------------------------
// Excel's unit table: each unit's category, its size in the category's base
// unit (grams, metres, seconds, pascals, newtons, joules, watts, teslas,
// cubic metres, square metres, bits, metres a second), and whether it takes
// a decimal prefix (k, da, u, ...) - bit and byte the binary ones too (ki,
// Mi, ...), a squared or cubed unit the prefix squared or cubed.

type private Unit = { Category: string; Size: float; Prefixed: bool; Power: int }

let private units : Collections.Generic.Dictionary<string, Unit> =
    let d = Collections.Generic.Dictionary<string, Unit>()
    let add (cat: string) (prefixed: bool) (power: int) (size: float) (names: string list) =
        for n in names do d.[n] <- { Category = cat; Size = size; Prefixed = prefixed; Power = power }
    let inch = 0.0254
    let foot = 0.3048
    let lbm = 453.59237
    let lbf = 4.4482216152605
    let gal = 231.0 * inch * inch * inch
    let ly = 9460730472580800.0
    let pica = inch / 72.0
    let mile = 1609.344
    let nmi = 1852.0
    let yd = 0.9144
    let ang = 1e-10
    let surveyFoot = 1200.0 / 3937.0
    // Mass (grams).
    add "mass" true 1 1.0 [ "g" ]
    add "mass" false 1 (lbf / foot * 1000.0) [ "sg" ]
    add "mass" false 1 lbm [ "lbm" ]
    add "mass" true 1 1.660538782e-24 [ "u" ]
    add "mass" false 1 (lbm / 16.0) [ "ozm" ]
    add "mass" false 1 0.06479891 [ "grain" ]
    add "mass" false 1 (100.0 * lbm) [ "cwt"; "shweight" ]
    add "mass" false 1 (112.0 * lbm) [ "uk_cwt"; "lcwt"; "hweight" ]
    add "mass" false 1 (14.0 * lbm) [ "stone" ]
    add "mass" false 1 (2000.0 * lbm) [ "ton" ]
    add "mass" false 1 (2240.0 * lbm) [ "uk_ton"; "LTON"; "brton" ]
    // Distance (metres).
    add "distance" true 1 1.0 [ "m" ]
    add "distance" false 1 mile [ "mi" ]
    add "distance" false 1 nmi [ "Nmi" ]
    add "distance" false 1 inch [ "in" ]
    add "distance" false 1 foot [ "ft" ]
    add "distance" false 1 yd [ "yd" ]
    add "distance" true 1 ang [ "ang" ]
    add "distance" false 1 1.143 [ "ell" ]
    add "distance" true 1 ly [ "ly" ]
    add "distance" true 1 3.0856775812815532e16 [ "parsec"; "pc" ]
    add "distance" false 1 pica [ "Picapt"; "Pica" ]
    add "distance" false 1 (inch / 6.0) [ "pica" ]
    add "distance" false 1 (5280.0 * surveyFoot) [ "survey_mi" ]
    // Time (seconds).
    add "time" true 1 1.0 [ "sec"; "s" ]
    add "time" false 1 31557600.0 [ "yr" ]
    add "time" false 1 86400.0 [ "day"; "d" ]
    add "time" false 1 3600.0 [ "hr" ]
    add "time" false 1 60.0 [ "mn"; "min" ]
    // Pressure (pascals).
    add "pressure" true 1 1.0 [ "Pa"; "p" ]
    add "pressure" true 1 101325.0 [ "atm"; "at" ]
    add "pressure" true 1 133.322 [ "mmHg" ]
    add "pressure" false 1 (lbf / (inch * inch)) [ "psi" ]
    add "pressure" false 1 (101325.0 / 760.0) [ "Torr" ]
    // Force (newtons).
    add "force" true 1 1.0 [ "N" ]
    add "force" true 1 1e-5 [ "dyn"; "dy" ]
    add "force" false 1 lbf [ "lbf" ]
    add "force" true 1 0.00980665 [ "pond" ]
    // Energy (joules).
    add "energy" true 1 1.0 [ "J" ]
    add "energy" true 1 1e-7 [ "e" ]
    add "energy" true 1 4.184 [ "c" ]
    add "energy" true 1 4.1868 [ "cal" ]
    add "energy" true 1 1.602176487e-19 [ "eV"; "ev" ]
    add "energy" false 1 (550.0 * foot * lbf * 3600.0) [ "HPh"; "hh" ]
    add "energy" true 1 3600.0 [ "Wh"; "wh" ]
    add "energy" false 1 (foot * lbf) [ "flb" ]
    add "energy" false 1 1055.05585262 [ "BTU"; "btu" ]
    // Power (watts).
    add "power" true 1 1.0 [ "W"; "w" ]
    add "power" false 1 (550.0 * foot * lbf) [ "HP"; "h" ]
    add "power" false 1 735.49875 [ "PS" ]
    // Magnetism (teslas).
    add "magnetism" true 1 1.0 [ "T" ]
    add "magnetism" true 1 1e-4 [ "ga" ]
    // Volume (cubic metres).
    add "volume" true 3 1.0 [ "m3"; "m^3" ]
    add "volume" false 1 (gal / 768.0) [ "tsp" ]
    add "volume" false 1 5e-6 [ "tspm" ]
    add "volume" false 1 (gal / 256.0) [ "tbs" ]
    add "volume" false 1 (gal / 128.0) [ "oz" ]
    add "volume" false 1 (gal / 16.0) [ "cup" ]
    add "volume" false 1 (gal / 8.0) [ "pt"; "us_pt" ]
    add "volume" true 1 (0.00454609 / 8.0) [ "uk_pt" ]
    add "volume" false 1 (gal / 4.0) [ "qt" ]
    add "volume" false 1 (0.00454609 / 4.0) [ "uk_qt" ]
    add "volume" false 1 gal [ "gal" ]
    add "volume" false 1 0.00454609 [ "uk_gal" ]
    add "volume" true 1 0.001 [ "l"; "L"; "lt" ]
    add "volume" true 3 (ang * ang * ang) [ "ang3"; "ang^3" ]
    add "volume" false 1 (42.0 * gal) [ "barrel" ]
    add "volume" false 1 (2150.42 * inch * inch * inch) [ "bushel" ]
    add "volume" false 1 (foot * foot * foot) [ "ft3"; "ft^3" ]
    add "volume" false 1 (inch * inch * inch) [ "in3"; "in^3" ]
    add "volume" false 1 (ly * ly * ly) [ "ly3"; "ly^3" ]
    add "volume" false 1 (mile * mile * mile) [ "mi3"; "mi^3" ]
    add "volume" false 1 (yd * yd * yd) [ "yd3"; "yd^3" ]
    add "volume" false 1 (nmi * nmi * nmi) [ "Nmi3"; "Nmi^3" ]
    add "volume" false 1 (pica * pica * pica) [ "Picapt3"; "Picapt^3"; "Pica3"; "Pica^3" ]
    add "volume" false 1 (100.0 * foot * foot * foot) [ "GRT"; "regton" ]
    add "volume" false 1 (40.0 * foot * foot * foot) [ "MTON" ]
    // Area (square metres).
    add "area" true 2 1.0 [ "m2"; "m^2" ]
    add "area" false 1 4046.8564224 [ "uk_acre" ]
    add "area" false 1 (43560.0 * surveyFoot * surveyFoot) [ "us_acre" ]
    add "area" true 2 (ang * ang) [ "ang2"; "ang^2" ]
    add "area" true 1 100.0 [ "ar" ]
    add "area" false 1 (foot * foot) [ "ft2"; "ft^2" ]
    add "area" false 1 10000.0 [ "ha" ]
    add "area" false 1 (inch * inch) [ "in2"; "in^2" ]
    add "area" false 1 (ly * ly) [ "ly2"; "ly^2" ]
    add "area" false 1 2500.0 [ "Morgen" ]
    add "area" false 1 (mile * mile) [ "mi2"; "mi^2" ]
    add "area" false 1 (nmi * nmi) [ "Nmi2"; "Nmi^2" ]
    add "area" false 1 (pica * pica) [ "Picapt2"; "Pica2"; "Pica^2"; "Picapt^2" ]
    add "area" false 1 (yd * yd) [ "yd2"; "yd^2" ]
    // Information (bits).
    add "information" true 1 1.0 [ "bit" ]
    add "information" true 1 8.0 [ "byte" ]
    // Speed (metres a second).
    add "speed" true 1 1.0 [ "m/s"; "m/sec" ]
    add "speed" false 1 (6080.0 * foot / 3600.0) [ "admkn" ]
    add "speed" false 1 (nmi / 3600.0) [ "kn" ]
    add "speed" true 1 (1.0 / 3600.0) [ "m/h"; "m/hr" ]
    add "speed" true 1 (mile / 3600.0) [ "mph" ]
    d

let private decimalPrefixes =
    dict [ "Y", 1e24; "Z", 1e21; "E", 1e18; "P", 1e15; "T", 1e12; "G", 1e9; "M", 1e6; "k", 1e3; "h", 1e2
           "da", 1e1; "e", 1e1; "d", 1e-1; "c", 1e-2; "m", 1e-3; "u", 1e-6; "n", 1e-9; "p", 1e-12; "f", 1e-15
           "a", 1e-18; "z", 1e-21; "y", 1e-24 ]
let private binaryPrefixes =
    dict [ "Yi", 1208925819614629174706176.0; "Zi", 1180591620717411303424.0; "Ei", 1152921504606846976.0
           "Pi", 1125899906842624.0; "Ti", 1099511627776.0; "Gi", 1073741824.0; "Mi", 1048576.0; "ki", 1024.0 ]

/// A unit name: its category and its size in the base unit (a temperature
/// its scale's name and the prefix's factor). A name is first looked up
/// whole ("mmHg", "Pa", "pc"), then as a prefix of one or two letters and a
/// unit that takes it.
let private lookupUnit (name: string) : (Unit * float) option =
    match units.TryGetValue name with
    | true, u -> Some (u, 1.0)
    | _ ->
        let tryPrefix (len: int) =
            if name.Length <= len then None
            else
                let p = name.Substring(0, len)
                match units.TryGetValue (name.Substring len) with
                | true, u when u.Prefixed ->
                    match decimalPrefixes.TryGetValue p with
                    | true, f -> Some (u, Math.Pow (f, float u.Power))
                    | _ ->
                        match binaryPrefixes.TryGetValue p with
                        | true, f when u.Category = "information" -> Some (u, f)
                        | _ -> None
                | _ -> None
        match tryPrefix 1 with
        | Some r -> Some r
        | None -> tryPrefix 2

/// The temperature scales by name, with the prefix factor of kelvin (the one
/// scale that takes a prefix: mK, kkel).
let private temperatureOf (name: string) : (string * float) option =
    let scale (n: string) =
        match n with
        | "K" | "kel" -> Some "K"
        | "C" | "cel" -> Some "C"
        | "F" | "fah" -> Some "F"
        | "Rank" -> Some "Rank"
        | "Reau" -> Some "Reau"
        | _ -> None
    match scale name with
    | Some s -> Some (s, 1.0)
    | None ->
        let tryPrefix (len: int) =
            if name.Length <= len then None
            else
                match decimalPrefixes.TryGetValue (name.Substring(0, len)), scale (name.Substring len) with
                | (true, f), Some "K" -> Some ("K", f)
                | _ -> None
        match tryPrefix 1 with
        | Some t -> Some t
        | None -> tryPrefix 2

/// A temperature from one scale to another as Excel computes it: through
/// Celsius ((F - 32) / 1.8, Reaumur / 0.8, Rankine * 5 / 9 - 273.15, back
/// * 1.8 + 32, * 0.8, (C + 273.15) * 1.8), but kelvin and Rankine to each
/// other by 1.8 and Fahrenheit and Rankine by 459.67, a scale to itself
/// untouched. A prefix of kelvin is dropped going to Rankine and either way
/// with Reaumur (CONVERT(1000,"kK","Rank") is 1800).
let private convertTemperature (x: float) (fromScale: string, fromFactor: float) (toScale: string, toFactor: float) =
    let fromFactor, toFactor =
        if fromScale = "Reau" || toScale = "Reau" then 1.0, 1.0
        elif toScale = "Rank" then 1.0, toFactor
        else fromFactor, toFactor
    let x = x * fromFactor
    let r =
        match fromScale, toScale with
        | a, b when a = b -> x
        | "K", "Rank" -> x * 1.8
        | "Rank", "K" -> x / 1.8
        | "F", "Rank" -> x + 459.67
        | "Rank", "F" -> x - 459.67
        | _ ->
            let c =
                match fromScale with
                | "K" -> x - 273.15
                | "F" -> (x - 32.0) / 1.8
                | "Rank" -> x * 5.0 / 9.0 - 273.15
                | "Reau" -> x / 0.8
                | _ -> x
            match toScale with
            | "K" -> c + 273.15
            | "F" -> c * 1.8 + 32.0
            | "Rank" -> (c + 273.15) * 1.8
            | "Reau" -> c * 0.8
            | _ -> c
    r / toFactor

let private convertUnits (x: float) (fromName: string) (toName: string) : Result<float, string> =
    match temperatureOf fromName, temperatureOf toName with
    | Some t, Some u -> Ok (convertTemperature x t u)
    | Some _, None | None, Some _ -> Result.Error Err.na
    | None, None ->
        match lookupUnit fromName, lookupUnit toName with
        | Some (u, p), Some (v, q) when u.Category = v.Category -> Ok (x * (u.Size * p) / (v.Size * q))
        | _ -> Result.Error Err.na

/// A unit argument of CONVERT: its text (a number or a logical is no unit: #N/A).
let private unitAt (a: Args) (i: int) : Result<string, string> =
    if not (a.Supplied i) then Result.Error Err.na
    else
        match a.Value i with
        | CellValue.String t -> Ok t
        | CellValue.Error e -> Result.Error e
        | _ -> Result.Error Err.na

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
      scalar "ERF" 1 2 "|S" (fun a -> res {
        let! lower = a.AtpNumber 0
        if a.Supplied 1 then
            let! upper = a.AtpNumber 1
            return! numOk (held (erfBetween lower upper))
        elif a.Count > 1 then return! fail Err.na
        else return! numOk (held (erf lower)) })
      scalar "ERF.PRECISE" 1 1 "|S" (fun a -> res {
        let! x = a.AtpNumber 0
        return! numOk (held (erf x)) })
      scalar "ERFC" 1 1 "|S" (fun a -> res {
        let! x = a.AtpNumber 0
        return! numOk (held (erfc x)) })
      scalar "ERFC.PRECISE" 1 1 "|S" (fun a -> res {
        let! x = a.AtpNumber 0
        return! numOk (held (erfc x)) })
      bessel "BESSELI" false besselI
      bessel "BESSELJ" false besselJ
      bessel "BESSELK" true besselK
      bessel "BESSELY" true besselY
      scalar "CONVERT" 3 3 "|S" (fun a -> res {
        let! x = a.AtpNumber 0
        let! fromUnit = unitAt a 1
        let! toUnit = unitAt a 2
        let! r = convertUnits x fromUnit toUnit
        return! numOk r })
      scalar "GESTEP" 1 2 "|S" (fun a -> res {
        let! x = a.Number 0
        let! step = a.NumberOr 1 0.0
        return CellValue.Number (if x >= step then 1.0 else 0.0) }) ]

let install () = registerFunctions moduleName functions
