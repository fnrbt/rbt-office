/// The standard normal distribution as Excel computes it, for NORM.S.DIST and
/// every function built on it (NORM.DIST, the legacy names, GAUSS, Z.TEST,
/// CONFIDENCE.NORM, LOGNORM.*). Measured against Excel 2019 with
/// tools/excel-parity/make-normal-probe.py:
///
///  - The distribution function is 0.5 erfc(-z/sqrt 2), the argument formed as
///    x = -z * 0.7071067811865476 (rounded), erfc(x) for |x| > 0.46875 as
///    exp(-fl(x x)) times the rational approximations of W. J. Cody's CALERF.
///    The rounded square in the exponent is what makes Excel's far tail differ
///    from the exact value by 0, 1 or 2 units in the last place of z^2/2 (a
///    relative error up to 1e-13 at z = -37); reproducing it reproduces Excel's
///    digits there. Cody's split exp(-ysq^2) exp(-del), the accurate form, does
///    not.
///  - The density is exp(-0.5 z z) * (1 / sqrt(2 pi)).
///  - The quantile is Wichura's AS 241 (PPND16), unrefined.
///  - Excel has no subnormal results: below 2.2250738585072014e-308 a
///    probability or density is 0 (NORM.S.DIST(-37.5) is 4.6e-308,
///    NORM.S.DIST(-37.525) 0), and a probability that small is no argument of
///    the quantile (#NUM!).
module Rbt.Spreadsheet.NormalDistribution

open System

/// The least normal double; Excel flushes what is below it to 0.
let minNormal = 2.2250738585072014e-308

let private flush (x: float) = if x < minNormal && x > -minNormal then 0.0 else x

let private a = [| 3.16112374387056560e00; 1.13864154151050156e02; 3.77485237685302021e02; 3.20937758913846947e03; 1.85777706184603153e-1 |]
let private b = [| 2.36012909523441209e01; 2.44024637934444173e02; 1.28261652607737228e03; 2.84423683343917062e03 |]
let private c =
    [| 5.64188496988670089e-1; 8.88314979438837594e00; 6.61191906371416295e01; 2.98635138197400131e02; 8.81952221241769090e02
       1.71204761263407058e03; 2.05107837782607147e03; 1.23033935479799725e03; 2.15311535474403846e-8 |]
let private d =
    [| 1.57449261107098347e01; 1.17693950891312499e02; 5.37181101862009858e02; 1.62138957456669019e03; 3.29079923573345963e03
       4.36261909014324716e03; 3.43936767414372164e03; 1.23033935480374942e03 |]
let private p =
    [| 3.05326634961232344e-1; 3.60344899949804439e-1; 1.25781726111229246e-1; 1.60837851487422766e-2; 6.58749161529837803e-4
       1.63153871373020978e-2 |]
let private q = [| 2.56852019228982242e00; 1.87295284992346725e00; 5.27905102951428412e-1; 6.05183413124413191e-2; 2.33520497626869185e-3 |]
let private oneOverSqrtPi = 5.6418958354775628695e-1

/// erfc x the way Excel's normal distribution takes it (see the module notes).
let erfc (x: float) : float =
    let y = abs x
    if Double.IsNaN x then nan
    elif y <= 0.46875 then
        let ysq = if y > 1.11e-16 then y * y else 0.0
        let mutable xnum = a.[4] * ysq
        let mutable xden = ysq
        for i in 0 .. 2 do
            xnum <- (xnum + a.[i]) * ysq
            xden <- (xden + b.[i]) * ysq
        1.0 - x * (xnum + a.[3]) / (xden + b.[3])
    else
        let r =
            if y <= 4.0 then
                let mutable xnum = c.[8] * y
                let mutable xden = y
                for i in 0 .. 6 do
                    xnum <- (xnum + c.[i]) * y
                    xden <- (xden + d.[i]) * y
                (xnum + c.[7]) / (xden + d.[7])
            else
                let ysq = 1.0 / (y * y)
                let mutable xnum = p.[5] * ysq
                let mutable xden = ysq
                for i in 0 .. 3 do
                    xnum <- (xnum + p.[i]) * ysq
                    xden <- (xden + q.[i]) * ysq
                (oneOverSqrtPi - ysq * (xnum + p.[4]) / (xden + q.[4])) / y
        let r = Math.Exp (-(y * y)) * r
        if x < 0.0 then 2.0 - r else r

/// P(Z <= z), Excel's NORM.S.DIST(z, TRUE).
let cdf (z: float) : float = flush (0.5 * erfc (-z * 0.7071067811865476))

/// The density, Excel's NORM.S.DIST(z, FALSE); NaN when z^2 overflows (#NUM! in Excel).
let pdf (z: float) : float =
    let s = z * z
    if Double.IsInfinity s then nan else flush (Math.Exp (-0.5 * s) * 0.3989422804014327)

/// z with P(Z <= z) = p, 0 < p < 1 (NaN outside, and for a p below the normal doubles).
let quantile (prob: float) : float =
    if not (prob >= minNormal && prob < 1.0) then nan
    else
        let qq = prob - 0.5
        if abs qq <= 0.425 then
            let r = 0.180625 - qq * qq
            qq * (((((((r * 2509.0809287301226727 + 33430.575583588128105) * r + 67265.770927008700853) * r + 45921.953931549871457) * r + 13731.693765509461125) * r + 1971.5909503065514427) * r + 133.14166789178437745) * r + 3.387132872796366608)
            / (((((((r * 5226.495278852545925 + 28729.085735721942674) * r + 39307.89580009271061) * r + 21213.794301586595867) * r + 5394.1960214247511077) * r + 687.1870074920579083) * r + 42.313330701600911252) * r + 1.0)
        else
            let r = Math.Sqrt (-(Math.Log (min prob (1.0 - prob))))
            let v =
                if r <= 5.0 then
                    let r = r - 1.6
                    (((((((r * 7.7454501427834140764e-4 + 0.0227238449892691845833) * r + 0.24178072517745061177) * r + 1.27045825245236838258) * r + 3.64784832476320460504) * r + 5.7694972214606914055) * r + 4.6303378461565452959) * r + 1.42343711074968357734)
                    / (((((((r * 1.05075007164441684324e-9 + 5.475938084995344946e-4) * r + 0.0151986665636164571966) * r + 0.14810397642748007459) * r + 0.68976733498510000455) * r + 1.6763848301838038494) * r + 2.05319162663775882187) * r + 1.0)
                else
                    let r = r - 5.0
                    (((((((r * 2.01033439929228813265e-7 + 2.71155556874348757815e-5) * r + 0.0012426609473880784386) * r + 0.026532189526576123093) * r + 0.29656057182850489123) * r + 1.7848265399172913358) * r + 5.4637849111641143699) * r + 6.6579046435011037772)
                    / (((((((r * 2.04426310338993978564e-15 + 1.4215117583164458887e-7) * r + 1.8463183175100546818e-5) * r + 7.868691311456132591e-4) * r + 0.0148753612908506148525) * r + 0.13692988092273580531) * r + 0.59983220655588793769) * r + 1.0)
            if qq < 0.0 then -v else v
