/// The probability distributions of the function table (t, F, chi-squared,
/// beta, gamma, binomial, Poisson and their kin, with Excel's compatibility
/// names) and the tests built on them (T.TEST, F.TEST, CHISQ.TEST).
///
/// One numerical core serves them all: log-gamma, the regularized incomplete
/// gamma and beta functions, their inverses. The densities and the prefactors
/// of the incomplete functions are Loader's saddle-point forms (stirlerr and
/// bd0, as R computes dbinom and dpois), which keep their relative accuracy
/// where exp(a log x - x - lgamma a) would lose digits to cancellation.
/// Excel's argument rules (truncation of degrees of freedom, the domains, the
/// legacy names' tails) are those measured with
/// tools/excel-parity/make-statdist-probe.py.
module Rbt.Spreadsheet.Functions.FnStatDist

open System
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Functions.Kit

let moduleName = "Distributions"

// ===== numerical core ======================================================

let private ls2pi = 0.918938533204672741780329736406 // log(sqrt(2 pi))
let private eulerGamma = 0.577215664901532860606512090082

let private log1p (x: float) =
    let u = 1.0 + x
    if u = 1.0 then x else Math.Log u * x / (u - 1.0)

let private expm1 (x: float) =
    if abs x < 1e-5 then x + x * x / 2.0 + x * x * x / 6.0
    else Math.Exp x - 1.0

/// zeta(k) - 1 for k = 0 .. 63 (entries 0 and 1 unused): the sum to 19 and an
/// Euler-Maclaurin tail from 20.
let private zetaMinus1 : float[] =
    let n = 20.0
    let table = Array.zeroCreate 64
    for k in 2 .. 63 do
        let kf = float k
        let mutable s = 0.0
        for i in 19 .. -1 .. 2 do s <- s + Math.Pow (float i, -kf)
        let tail =
            Math.Pow (n, 1.0 - kf) / (kf - 1.0) + Math.Pow (n, -kf) / 2.0
            + kf * Math.Pow (n, -kf - 1.0) / 12.0
            - kf * (kf + 1.0) * (kf + 2.0) * Math.Pow (n, -kf - 3.0) / 720.0
            + kf * (kf + 1.0) * (kf + 2.0) * (kf + 3.0) * (kf + 4.0) * Math.Pow (n, -kf - 5.0) / 30240.0
            - kf * (kf + 1.0) * (kf + 2.0) * (kf + 3.0) * (kf + 4.0) * (kf + 5.0) * (kf + 6.0) * Math.Pow (n, -kf - 7.0) / 1209600.0
        table.[k] <- s + tail
    table

/// log Gamma(1 + t) for |t| <= 0.5: -gamma t + sum_{k>=2} (-1)^k zeta(k) t^k / k.
let private lgammaSeries1 (t: float) =
    let mutable sum = 0.0
    let mutable power = t
    let mutable k = 2
    let mutable fin = false
    while not fin && k < 64 do
        power <- power * t
        let term = (1.0 + zetaMinus1.[k]) * power / float k
        sum <- sum + (if k % 2 = 0 then term else -term)
        if abs term < 1e-18 * abs sum then fin <- true
        k <- k + 1
    -eulerGamma * t + sum

/// log Gamma(2 + t) for |t| <= 0.5: (1 - gamma) t + sum_{k>=2} (-1)^k (zeta(k) - 1) t^k / k.
let private lgammaSeries2 (t: float) =
    let mutable sum = 0.0
    let mutable power = t
    let mutable k = 2
    let mutable fin = false
    while not fin && k < 64 do
        power <- power * t
        let term = zetaMinus1.[k] * power / float k
        sum <- sum + (if k % 2 = 0 then term else -term)
        if abs term < 1e-18 * abs sum then fin <- true
        k <- k + 1
    (1.0 - eulerGamma) * t + sum

/// Stirling's error log Gamma(n+1) - (n + 1/2) log n + n - log sqrt(2 pi),
/// n > 0: the asymptotic series from 15 on, below it the exact differences
/// stirlerr(t) - stirlerr(t+1) = sum u^2k / (2k+1), u = 1/(2t+1).
let rec stirlerr (n: float) : float =
    if n <= 0.0 then 0.0
    elif n >= 15.0 then
        let nn = n * n
        (1.0 / 12.0 - (1.0 / 360.0 - (1.0 / 1260.0 - (1.0 / 1680.0 - (1.0 / 1188.0 - (691.0 / 360360.0 - (1.0 / 156.0 - 3617.0 / 122400.0 / nn) / nn) / nn) / nn) / nn) / nn) / nn) / n
    else
        let steps = int (Math.Ceiling (15.0 - n))
        let mutable sum = 0.0
        for j in 0 .. steps - 1 do
            let t = n + float j
            if t >= 0.5 then
                let u = 1.0 / (2.0 * t + 1.0)
                let u2 = u * u
                let mutable power = u2
                let mutable k = 1
                let mutable s = 0.0
                while power > 1e-20 * (s + 1e-300) && k < 400 do
                    s <- s + power / float (2 * k + 1)
                    power <- power * u2
                    k <- k + 1
                sum <- sum + s
            else
                sum <- sum + ((t + 0.5) * (log1p t - Math.Log t) - 1.0)
        sum + stirlerr (n + float steps)

/// log Gamma(x), x > 0.
let lgamma (x: float) : float =
    if x <= 0.0 || Double.IsNaN x then nan
    elif x < 0.5 then lgammaSeries1 x - Math.Log x
    elif x < 1.5 then lgammaSeries1 (x - 1.0)
    elif x < 2.5 then lgammaSeries2 (x - 2.0)
    elif x < 3.5 then Math.Log (x - 1.0) + lgammaSeries2 (x - 3.0)
    else stirlerr x + (x - 0.5) * Math.Log x - x + ls2pi

// ----- double-double products, for GAMMA's long products ---------------------

let private splitter = 134217729.0
let private twoProd (a: float) (b: float) =
    let p = a * b
    let ca = splitter * a
    let ah = ca - (ca - a)
    let al = a - ah
    let cb = splitter * b
    let bh = cb - (cb - b)
    let bl = b - bh
    p, ((ah * bh - p) + ah * bl + al * bh) + al * bl

let private ddMul (hi: float, lo: float) (d: float) =
    let p, e = twoProd hi d
    let e = e + lo * d
    let s = p + e
    s, e - (s - p)

/// The product start (start + 1) ... (start + count - 1) as a double-double
/// (hi, lo) times 2^scale: rescaled by powers of two (exactly) so that it
/// neither overflows nor leaves the range where the splitting of twoProd works.
let private factorProduct (start: float) (count: int) =
    let down = Math.Pow (2.0, -800.0)
    let up = Math.Pow (2.0, 800.0)
    let mutable acc = (1.0, 0.0)
    let mutable scale = 0
    for j in 0 .. count - 1 do
        acc <- ddMul acc (start + float j)
        let h, l = acc
        if abs h > 1e250 then
            acc <- (h * down, l * down)
            scale <- scale + 800
        elif h <> 0.0 && abs h < 1e-250 then
            acc <- (h * up, l * up)
            scale <- scale - 800
    let h, l = acc
    h, l, scale

/// x times 2^k, in two steps so that neither overflows on the way.
let private scaleBy (x: float) (k: int) =
    let half = k / 2
    x * Math.Pow (2.0, float half) * Math.Pow (2.0, float (k - half))

/// Gamma(x) for x not 0 or a negative integer: Gamma of [1, 2) times (or
/// divided by) the exact double-double product of the factors between.
let gammaFn (x: float) : float =
    if x > 171.7 then infinity
    elif x >= 1.0 then
        let r = x - Math.Floor x + 1.0
        let h, l, scale = factorProduct r (int (Math.Round (x - r)))
        let s, e = ddMul (h, l) (Math.Exp (lgamma r))
        scaleBy (s + e) scale
    else
        // Gamma(x) = Gamma(x + n) / (x (x+1) ... (x+n-1)), x + n in [1, 2).
        let n = int (Math.Ceiling (1.0 - x))
        let n = if x + float n >= 2.0 then n - 1 else n
        let dh, dl, scale = factorProduct x n
        let g = Math.Exp (lgamma (x + float n))
        let q = g / dh
        let ph, pl = twoProd q dh
        let q = q + ((g - ph) - pl - q * dl) / dh
        // Below the normal doubles Excel's GAMMA is 0.
        let g = scaleBy q -scale
        if abs g < 2.2250738585072014e-308 then 0.0 else g

/// x log(x / np) + np - x, accurately for x near np (Loader).
let bd0 (x: float) (np: float) =
    if abs (x - np) < 0.1 * (x + np) then
        let mutable v = (x - np) / (x + np)
        let mutable s = (x - np) * v
        if abs s < 1e-300 then s
        else
            let mutable ej = 2.0 * x * v
            v <- v * v
            let mutable j = 1
            let mutable fin = false
            while not fin && j < 1000 do
                ej <- ej * v
                let s1 = s + ej / float (2 * j + 1)
                if s1 = s then fin <- true
                s <- s1
                j <- j + 1
            s
    else x * Math.Log (x / np) + np - x

/// A product of factors that are each normal doubles, when it is one too
/// (else NaN, which sends the caller to its logarithmic form).
let private productOf (factors: float list) =
    let mutable v = 1.0
    let mutable ok = true
    for f in factors do
        if not (f > 1e-290 && f < 1e290) then ok <- false
        v <- v * f
    if ok && v > 1e-290 && v < 1e290 then v else nan

/// lambda^x e^-lambda / Gamma(x + 1), x >= 0 real. Where every factor is a
/// double it is their product (a few roundings, at any distance from the
/// mode); else Loader's saddle-point form.
let dpoisRaw (x: float) (lambda: float) =
    let direct =
        if x > 0.0 && x <= 170.0 && lambda > 0.0 && lambda < 700.0 then
            productOf [ Math.Pow (lambda, x); Math.Exp -lambda; 1.0 / gammaFn (x + 1.0) ]
        else nan
    if not (Double.IsNaN direct) then direct
    elif lambda = 0.0 then (if x = 0.0 then 1.0 else 0.0)
    elif x = 0.0 then Math.Exp -lambda
    elif x < 0.0 then 0.0
    elif x < 1.0 then
        let lg = if x <= 0.5 then lgammaSeries1 x else lgamma (x + 1.0)
        Math.Exp (x * Math.Log lambda - lambda - lg)
    else Math.Exp (-stirlerr x - bd0 x lambda) / Math.Sqrt (2.0 * Math.PI * x)

/// Gamma(n+1) / (Gamma(x+1) Gamma(n-x+1)) p^x q^(n-x), q = 1 - p: the product
/// where its factors are doubles, else R's dbinom_raw (Loader).
let dbinomRaw (x: float) (n: float) (p: float) (q: float) =
    let direct =
        if x > 0.0 && x < n && n <= 170.0 && p > 0.0 && q > 0.0 then
            productOf [ gammaFn (n + 1.0) / (gammaFn (x + 1.0) * gammaFn (n - x + 1.0)); Math.Pow (p, x); Math.Pow (q, n - x) ]
        else nan
    if not (Double.IsNaN direct) then direct
    elif p = 0.0 then (if x = 0.0 then 1.0 else 0.0)
    elif q = 0.0 then (if x = n then 1.0 else 0.0)
    elif x = 0.0 then
        if n = 0.0 then 1.0
        elif p < 0.1 then Math.Exp (-bd0 n (n * q) - n * p)
        else Math.Pow (q, n)
    elif x = n then
        if q < 0.1 then Math.Exp (-bd0 n (n * p) - n * q) else Math.Pow (p, n)
    elif x < 0.0 || x > n then 0.0
    else
        let lc = stirlerr n - stirlerr x - stirlerr (n - x) - bd0 x (n * p) - bd0 (n - x) (n * q)
        // log(2 pi x (n - x) / n): n - x is exact, 1 - x/n would not be.
        let lf = Math.Log (2.0 * Math.PI) + Math.Log x + Math.Log (n - x) - Math.Log n
        Math.Exp (lc - 0.5 * lf)

/// The regularized incomplete gamma functions (P, Q) of a > 0 at x.
let gammaPQ (a: float) (x: float) : float * float =
    if x <= 0.0 then 0.0, 1.0
    elif Double.IsPositiveInfinity x then 1.0, 0.0
    else
        let prefix = dpoisRaw a x // x^a e^-x / Gamma(a + 1)
        if x < a + 1.0 && not (a <= 1.0 && x > 1.0) then
            let mutable sum = 1.0
            let mutable term = 1.0
            let mutable n = 1.0
            while term > sum * 1e-17 && n < 1e8 do
                term <- term * x / (a + n)
                sum <- sum + term
                n <- n + 1.0
            let p = prefix * sum
            p, 1.0 - p
        else
            // Legendre's continued fraction for Q, 1/(x+1-a- 1(1-a)/(x+3-a- ...)):
            // Lentz's forward evaluation finds how many terms it needs, the
            // value is then evaluated backward, which does not accumulate the
            // rounding of the forward product.
            let tiny = 1e-300
            let mutable b = x + 1.0 - a
            let mutable c = 1.0 / tiny
            let mutable d = 1.0 / b
            let mutable i = 1.0
            let mutable fin = false
            while not fin && i < 1e8 do
                let an = -i * (i - a)
                b <- b + 2.0
                d <- an * d + b
                if abs d < tiny then d <- tiny
                c <- b + an / c
                if abs c < tiny then c <- tiny
                d <- 1.0 / d
                if abs (d * c - 1.0) < 1e-16 then fin <- true
                i <- i + 1.0
            let mutable f = 0.0
            let mutable k = 2.0 * i + 20.0
            while k >= 1.0 do
                f <- -k * (k - a) / (x + 1.0 - a + 2.0 * k + f)
                k <- k - 1.0
            let h = 1.0 / (x + 1.0 - a + f)
            let q = a * prefix * h
            1.0 - q, q

/// x^a y^b / B(a, b), y = 1 - x.
let betaPrefix (a: float) (b: float) (x: float) (y: float) =
    dbinomRaw a (a + b) x y * (a * b / (a + b))

/// I_x(a, b) by the continued fraction of Didonato and Morris (TOMS 708's
/// BFRAC), lambda = (a + b) y - b >= 0: written with lambda rather than with
/// 1 - (a+b) x / (a+1), so that no step cancels when x is near 1.
let private betaFraction (a: float) (b: float) (x: float) (y: float) (lambda: float) =
    let c = lambda + 1.0
    let c0 = b / a
    let c1 = 1.0 / a + 1.0
    let yp1 = y + 1.0
    let mutable n = 0.0
    let mutable p = 1.0
    let mutable s = a + 1.0
    let mutable an = 0.0
    let mutable bn = 1.0
    let mutable anp1 = 1.0
    let mutable bnp1 = c / c1
    let mutable r = c1 / c
    let mutable fin = false
    while not fin && n < 1e6 do
        n <- n + 1.0
        let t = n / a
        let w = n * (b - n) * x
        let e = a / s
        let alpha = p * (p + c0) * e * e * (w * x)
        let e = (t + 1.0) / (c1 + t + t)
        let beta = n + w / s + e * (c + n * yp1)
        p <- t + 1.0
        s <- s + 2.0
        let t = alpha * an + beta * anp1
        an <- anp1
        anp1 <- t
        let t = alpha * bn + beta * bnp1
        bn <- bnp1
        bnp1 <- t
        let r0 = r
        r <- anp1 / bnp1
        if abs (r - r0) <= 1e-16 * r then fin <- true
        else
            an <- an / bnp1
            bn <- bn / bnp1
            anp1 <- r
            bnp1 <- 1.0
    betaPrefix a b x y * r

/// The regularized incomplete beta function and its complement (I_x(a,b),
/// 1 - I_x(a,b)), y = 1 - x given separately so that neither loses digits.
/// The fraction is evaluated on the side of the mean where it converges
/// (lambda >= 0) and the other value is its complement.
let betaIJ (a: float) (b: float) (x: float) (y: float) : float * float =
    if x <= 0.0 then 0.0, 1.0
    elif y <= 0.0 then 1.0, 0.0
    else
        let lambda = if a > b then (a + b) * y - b else a - (a + b) * x
        if lambda < 0.0 then
            let j = betaFraction b a y x -lambda
            1.0 - j, j
        else
            let i = betaFraction a b x y lambda
            i, 1.0 - i

// ----- densities -------------------------------------------------------------

/// The density of the gamma distribution of shape a, scale 1, at x > 0.
let private gammaDensity (a: float) (x: float) =
    if a < 1.0 then dpoisRaw a x * a / x else dpoisRaw (a - 1.0) x

/// The density of the beta distribution at 0 < x < 1.
let private betaDensity (a: float) (b: float) (x: float) =
    betaPrefix a b x (1.0 - x) / (x * (1.0 - x))

/// Student's t density (R's dt).
let private tDensity (n: float) (x: float) =
    let t = -bd0 (n / 2.0) ((n + 1.0) / 2.0) + stirlerr ((n + 1.0) / 2.0) - stirlerr (n / 2.0)
    let x2n = x * x / n
    let large = x2n > 1.0 / 2.220446049250313e-16
    let u =
        if large then n * (Math.Log (abs x) - Math.Log n / 2.0)
        elif x2n > 0.2 then n * Math.Log (1.0 + x2n) / 2.0
        else -bd0 (n / 2.0) ((n + x * x) / 2.0) + x * x / 2.0
    Math.Exp (t - u) * (if large then Math.Sqrt n / abs x else 1.0 / Math.Sqrt (1.0 + x2n)) / Math.Sqrt (2.0 * Math.PI)

/// The F density of (m, n) degrees of freedom at x > 0 (R's df).
let private fDensity (m: float) (n: float) (x: float) =
    let f = 1.0 / (n + x * m)
    let q = n * f
    let p = x * m * f
    if m >= 2.0 then m * q / 2.0 * dbinomRaw ((m - 2.0) / 2.0) ((m + n - 2.0) / 2.0) p q
    else m * m * q / (2.0 * p * (m + n)) * dbinomRaw (m / 2.0) ((m + n) / 2.0) p q

// ----- distribution functions --------------------------------------------------

/// P(T > t) for Student's t of n degrees of freedom (the upper tail).
let tUpper (n: float) (t: float) =
    if t = 0.0 then 0.5
    else
        let t2 = t * t
        let x, y =
            if t2 > 1e300 then (n / abs t) / abs t, 1.0
            else n / (n + t2), t2 / (n + t2)
        // Excel works with x = n / (n + t^2): where that is 1 in doubles its
        // tails are exactly 1/2 (T.DIST(1E-8, 1, TRUE) = 0.5).
        if x = 1.0 then 0.5
        // One degree of freedom: Excel's closed form I_x(1/2, 1/2) = (2/pi) asin(sqrt x).
        elif n = 1.0 then
            let i = 2.0 * Math.Asin (Math.Sqrt x) / Math.PI
            if t > 0.0 then i / 2.0 else 1.0 - i / 2.0
        else
            let i, _ = betaIJ (n / 2.0) 0.5 x y
            if t > 0.0 then i / 2.0 else 1.0 - i / 2.0

/// P(0 < T <= t) for t >= 0: 1/2 - P(T > t), without the cancellation.
let tCentral (n: float) (t: float) =
    let t2 = t * t
    if t2 > 1e300 then 0.5
    else snd (betaIJ (n / 2.0) 0.5 (n / (n + t2)) (t2 / (n + t2))) / 2.0

/// P(F <= x) and P(F > x) for F of (m, n) degrees of freedom.
let fTails (m: float) (n: float) (x: float) =
    if x <= 0.0 then 0.0, 1.0
    else
        let mx = m * x
        betaIJ (m / 2.0) (n / 2.0) (mx / (mx + n)) (n / (mx + n))

// ----- inverses ----------------------------------------------------------------------

/// Acklam's rational approximation of the standard normal quantile (a start).
let private normalQuantileStart (p: float) =
    let a = [| -3.969683028665376e+01; 2.209460984245205e+02; -2.759285104469687e+02; 1.383577518672690e+02; -3.066479806614716e+01; 2.506628277459239e+00 |]
    let b = [| -5.447609879822406e+01; 1.615858368580409e+02; -1.556989798598866e+02; 6.680131188771972e+01; -1.328068155288572e+01 |]
    let c = [| -7.784894002430293e-03; -3.223964580411365e-01; -2.400758277161838e+00; -2.549732539343734e+00; 4.374664141464968e+00; 2.938163982698783e+00 |]
    let d = [| 7.784695709041462e-03; 3.224671290700398e-01; 2.445134137142996e+00; 3.754408661907416e+00 |]
    if p <= 0.0 then -40.0
    elif p >= 1.0 then 40.0
    elif p < 0.02425 then
        let q = Math.Sqrt (-2.0 * Math.Log p)
        (((((c.[0] * q + c.[1]) * q + c.[2]) * q + c.[3]) * q + c.[4]) * q + c.[5]) / ((((d.[0] * q + d.[1]) * q + d.[2]) * q + d.[3]) * q + 1.0)
    elif p <= 1.0 - 0.02425 then
        let q = p - 0.5
        let r = q * q
        (((((a.[0] * r + a.[1]) * r + a.[2]) * r + a.[3]) * r + a.[4]) * r + a.[5]) * q / (((((b.[0] * r + b.[1]) * r + b.[2]) * r + b.[3]) * r + b.[4]) * r + 1.0)
    else
        let q = Math.Sqrt (-2.0 * Math.Log (1.0 - p))
        -(((((c.[0] * q + c.[1]) * q + c.[2]) * q + c.[3]) * q + c.[4]) * q + c.[5]) / ((((d.[0] * q + d.[1]) * q + d.[2]) * q + d.[3]) * q + 1.0)

/// Solve tail(x) = target on (0, upper) for a monotone tail (increasing when
/// `increasing`), slope = d tail / dx: Newton steps on log tail, first in
/// log x (where power and exponential tails are nearly straight), then in x,
/// inside a bracket that bisection falls back to.
let private solvePositive (tail: float -> float) (slope: float -> float) (increasing: bool) (target: float) (upper: float) (start: float) : float =
    let logTarget = Math.Log target
    let mutable lo = -infinity
    let mutable hi = if Double.IsPositiveInfinity upper then infinity else Math.Log upper
    let mutable u = Math.Log (max start 1e-300)
    if u >= hi then u <- hi - 0.5
    let mutable fin = false
    let mutable iterations = 0
    let mutable result = nan
    while not fin && iterations < 400 do
        iterations <- iterations + 1
        let x = Math.Exp u
        let t = tail x
        if t = target then
            fin <- true
            result <- x
        else
            if (t < target) = increasing then lo <- u else hi <- u
            let s = slope x
            let step =
                if t > 0.0 && s <> 0.0 && not (Double.IsInfinity s) then (logTarget - Math.Log t) * t / (s * x) else nan
            let mutable next = u + step
            if not (next > lo && next < hi) then
                next <-
                    if Double.IsNegativeInfinity lo then hi - max 1.0 (2.0 * abs hi)
                    elif Double.IsPositiveInfinity hi then lo + max 1.0 (2.0 * abs lo)
                    else (lo + hi) / 2.0
            if abs (next - u) < 1e-15 || next = lo || next = hi then
                fin <- true
                result <- Math.Exp next
            u <- next
    if Double.IsNaN result then result <- Math.Exp u
    // Polish in x: the last digits that the logarithm's rounding hides.
    let xlo = Math.Exp lo
    let xhi = if Double.IsPositiveInfinity hi then infinity else Math.Exp hi
    let mutable x = result
    let mutable k = 0
    let mutable stop = false
    while not stop && k < 8 && x > 0.0 && not (Double.IsInfinity x) do
        k <- k + 1
        let t = tail x
        let s = slope x
        if t <= 0.0 || s = 0.0 || Double.IsInfinity s || t = target then stop <- true
        else
            let next = x + (logTarget - Math.Log t) * t / s
            if next = x || not (next >= xlo && next <= xhi) || Double.IsNaN next then stop <- true
            else
                if abs (next - x) <= 2e-16 * abs x then stop <- true
                x <- next
    x

/// x with P(a, x) = p (lower) or Q(a, x) = q (upper): whichever of p and
/// q = 1 - p is the smaller is solved for.
let gammaInverse (a: float) (p: float) (q: float) : float =
    if p <= 0.0 then 0.0
    elif q <= 0.0 then infinity
    else
        let lower = p <= q
        let z = normalQuantileStart p
        let wh = a * Math.Pow (1.0 - 1.0 / (9.0 * a) + z / (3.0 * Math.Sqrt a), 3.0)
        let start =
            if wh > 0.0 && not (Double.IsNaN wh) then wh
            else Math.Exp ((Math.Log p + lgamma (a + 1.0)) / a)
        let start = if start > 0.0 && not (Double.IsInfinity start) then start else a
        let x =
            if lower then solvePositive (fun x -> fst (gammaPQ a x)) (fun x -> gammaDensity a x) true p infinity start
            else solvePositive (fun x -> snd (gammaPQ a x)) (fun x -> -(gammaDensity a x)) false q infinity start
        // A quantile below the normal doubles is 0 (as Excel's GAMMA.INV gives it).
        if x < 2.2250738585072014e-308 then 0.0 else x

/// x with I_x(a, b) = p (q = 1 - p), solving for whichever tail is the
/// smaller.
let private betaSolve (a: float) (b: float) (p: float) (q: float) : float =
    let mean = a / (a + b)
    let small = Math.Exp ((Math.Log (min p 0.5) + Math.Log a + lgamma a + lgamma b - lgamma (a + b)) / a)
    let start = if small > 0.0 && small < mean then small else mean
    if p <= q then solvePositive (fun x -> fst (betaIJ a b x (1.0 - x))) (fun x -> betaDensity a b x) true p 1.0 start
    else solvePositive (fun x -> snd (betaIJ a b x (1.0 - x))) (fun x -> -(betaDensity a b x)) false q 1.0 start

/// (x, 1 - x) with I_x(a, b) = p, q = 1 - p: a solution above 1/2 is solved
/// again as 1 - x of the mirrored distribution, so that both keep their digits.
let betaInverse (a: float) (b: float) (p: float) (q: float) : float * float =
    if p <= 0.0 then 0.0, 1.0
    elif q <= 0.0 then 1.0, 0.0
    else
        let x = betaSolve a b p q
        if x > 0.5 then
            let y = betaSolve b a q p
            1.0 - y, y
        else x, 1.0 - x

/// t >= 0 with P(T > t) = p (p <= 1/2), n degrees of freedom. For n = 1 it is
/// Excel's closed form 1 / tan(pi p) (its T.INV(0, 1) is #DIV/0!); otherwise
/// the tail is solved for from the better of the normal quantile and the
/// power-law asymptote P(T > t) ~ c t^-n.
let tInverseUpper (n: float) (p: float) : float =
    if p >= 0.5 then 0.0
    elif n = 1.0 then 1.0 / Math.Tan (Math.PI * p)
    else
        let z = -(normalQuantileStart p)
        let logC = lgamma ((n + 1.0) / 2.0) - lgamma (n / 2.0) - 0.5 * Math.Log (n * Math.PI) + ((n - 1.0) / 2.0) * Math.Log n
        let t0 = Math.Exp ((logC - Math.Log p) / n)
        let miss t = abs (Math.Log (max (tUpper n t) 1e-300) - Math.Log p)
        let start =
            if z <= 0.0 then 1e-3
            elif Double.IsInfinity t0 || Double.IsNaN t0 || miss z <= miss t0 then z
            else t0
        if p > 0.25 then solvePositive (fun t -> tCentral n t) (fun t -> tDensity n t) true (0.5 - p) infinity start
        else solvePositive (fun t -> tUpper n t) (fun t -> -(tDensity n t)) false p infinity start

/// x with P(F <= x) = p, P(F > x) = q.
let fInverse (m: float) (n: float) (p: float) (q: float) : float =
    let w, v = betaInverse (m / 2.0) (n / 2.0) p q
    if v <= 0.0 then infinity else (n / m) * (w / v)

// ----- discrete distributions ----------------------------------------------------

/// C(n, k) for integers 0 <= k <= n by the product of the ratios (each step
/// exact while the coefficient stays below 2^53).
let private choose (n: float) (k: float) =
    let k = min k (n - k)
    let mutable c = 1.0
    let mutable i = 1.0
    while i <= k do
        c <- c * (n - k + i) / i
        i <- i + 1.0
    c

/// c p^s (1 - p)^f, with 1 - p carried as a double-double (its rounding would
/// count f times); NaN when a factor leaves the normal doubles.
let private bernoulliProduct (c: float) (p: float) (s: float) (f: float) =
    let qh = 1.0 - p
    let ql = (1.0 - qh) - p
    let qf = Math.Pow (qh, f) * (1.0 + f * ql / qh)
    productOf [ c; (if s = 0.0 then 1.0 else Math.Pow (p, s)); (if f = 0.0 then 1.0 else qf) ]

let private smallCount = 1029.0

let binomialPmf (k: float) (n: float) (p: float) =
    let direct =
        if n <= smallCount && p > 0.0 && p < 1.0 && k >= 0.0 && k <= n then bernoulliProduct (choose n k) p k (n - k)
        else nan
    if Double.IsNaN direct then dbinomRaw k n p (1.0 - p) else direct

/// P(X <= k) for the binomial of n trials, success p: the sum of the
/// probabilities for a small n (as Excel's exact dyadic cases need), the
/// incomplete beta function beyond.
let binomialCdf (k: float) (n: float) (p: float) =
    if k < 0.0 then 0.0
    elif k >= n then 1.0
    elif n <= smallCount then
        let mutable sum = 0.0
        let mutable i = 0.0
        while i <= k do
            sum <- sum + binomialPmf i n p
            i <- i + 1.0
        min 1.0 sum
    else fst (betaIJ (n - k) (k + 1.0) (1.0 - p) p)

let poissonCdf (k: float) (lambda: float) =
    if k < 0.0 then 0.0
    elif lambda = 0.0 then 1.0
    else snd (gammaPQ (k + 1.0) lambda)

/// Hypergeometric probability of x successes in a sample of n from a
/// population of r successes and b failures (R's dhyper).
let hypergeometricPmf (x: float) (r: float) (b: float) (n: float) =
    if x < 0.0 || x < n - b || x > r || x > n then 0.0
    elif n = 0.0 then (if x = 0.0 then 1.0 else 0.0)
    elif r + b <= smallCount && choose (r + b) n < 1e300 then
        choose r x * choose b (n - x) / choose (r + b) n
    else
        let p = n / (r + b)
        let q = (r + b - n) / (r + b)
        let p1 = dbinomRaw x r p q
        let p2 = dbinomRaw (n - x) b p q
        let p3 = dbinomRaw n (r + b) p q
        p1 * p2 / p3

let negativeBinomialPmf (f: float) (s: float) (p: float) =
    // C(f + s - 1, f) p^s (1 - p)^f = s / (s + f) dbinom(s; s + f, p)
    let direct = if f + s <= smallCount then bernoulliProduct (choose (f + s - 1.0) f) p s f else nan
    if Double.IsNaN direct then s / (s + f) * dbinomRaw s (s + f) p (1.0 - p) else direct

/// The sum of f k for k = first, first + 1, ..., last.
let private sumOver (first: float) (last: float) (f: float -> float) =
    let mutable sum = 0.0
    let mutable k = first
    while k <= last do
        sum <- sum + f k
        k <- k + 1.0
    sum

/// The smallest k with P(X <= k) >= alpha, searched from the normal guess.
let private binomialQuantile (n: float) (p: float) (alpha: float) =
    let sd = Math.Sqrt (n * p * (1.0 - p))
    let guess = Math.Floor (n * p + sd * normalQuantileStart alpha)
    let mutable k = max 0.0 (min n (if Double.IsNaN guess then 0.0 else guess))
    while k > 0.0 && binomialCdf (k - 1.0) n p >= alpha do k <- k - 1.0
    while k < n && binomialCdf k n p < alpha do k <- k + 1.0
    k

/// Pearson's statistic sum (o - e)^2 / e of two tables of as many cells: an
/// expected 0 is #DIV/0!, a negative one #NUM!.
let private chiSquareStatistic (actual: CellValue[]) (expected: CellValue[]) : Result<float, string> =
    let mutable stat = 0.0
    let mutable failure = None
    for i in 0 .. actual.Length - 1 do
        match actual.[i], expected.[i] with
        | CellValue.Error e, _ | _, CellValue.Error e -> if failure.IsNone then failure <- Some e
        | CellValue.Number o, CellValue.Number e ->
            if e < 0.0 then (if failure.IsNone then failure <- Some Err.num)
            elif e = 0.0 then (if failure.IsNone then failure <- Some Err.div0)
            else stat <- stat + (o - e) * (o - e) / e
        | _ -> ()
    match failure with
    | Some e -> Result.Error e
    | None -> Ok stat

// ===== the functions ===========================================================

/// Degrees of freedom: truncated, from 1 to 10^10.
let private degrees (x: float) : Result<float, string> =
    let d = truncate x
    if d < 1.0 || d > 1e10 then fail Err.num else Ok d

let private probability (p: float) = check (p > 0.0 && p < 1.0) Err.num

let private tDist =
    scalar "T.DIST" 3 3 "|S" (fun a -> res {
        let! x = a.Number 0
        let! df = a.Number 1
        let! cumulative = a.Bool 2
        let! n = degrees df
        if cumulative then
            return! numOk (if x > 0.0 then 1.0 - tUpper n x else tUpper n -x)
        else return! numOk (tDensity n x) })

let private tDist2T name =
    scalar name 2 2 "|S" (fun a -> res {
        let! x = a.Number 0
        let! df = a.Number 1
        let! n = degrees df
        do! check (x >= 0.0) Err.num
        return! numOk (min 1.0 (2.0 * tUpper n x)) })

let private tDistRT =
    scalar "T.DIST.RT" 2 2 "|S" (fun a -> res {
        let! x = a.Number 0
        let! df = a.Number 1
        let! n = degrees df
        return! numOk (tUpper n x) })

let private tInv =
    scalar "T.INV" 2 2 "|S" (fun a -> res {
        let! p = a.Number 0
        let! df = a.Number 1
        do! check (p >= 0.0 && p < 1.0) Err.num
        let! n = degrees df
        // Excel's closed form for one degree of freedom divides by tan(0).
        do! check (p > 0.0) (if n = 1.0 then Err.div0 else Err.num)
        if p = 0.5 then return CellValue.Number 0.0
        elif p < 0.5 then return! numOk (-(tInverseUpper n p))
        else return! numOk (tInverseUpper n (1.0 - p)) })

/// T.INV.2T; TINV (`legacy`) also takes 1 < p < 2, the negative t whose
/// two-tailed probability is 2 - p.
let private tInv2T name (legacy: bool) =
    scalar name 2 2 "|S" (fun a -> res {
        let! p = a.Number 0
        let! df = a.Number 1
        do! check (p >= 0.0 && (p <= 1.0 || legacy && p < 2.0)) Err.num
        let! n = degrees df
        do! check (p > 0.0) (if n = 1.0 then Err.div0 else Err.num)
        if p = 1.0 then return CellValue.Number 0.0
        elif p > 1.0 then return! numOk (-(tInverseUpper n ((2.0 - p) / 2.0)))
        else return! numOk (tInverseUpper n (p / 2.0)) })

let private tDistLegacy =
    scalar "TDIST" 3 3 "|S" (fun a -> res {
        let! x = a.Number 0
        let! df = a.Number 1
        let! tails = a.Number 2
        let! n = degrees df
        let tails = truncate tails
        do! check (tails = 1.0 || tails = 2.0) Err.num
        do! check (x >= 0.0) Err.num
        return! numOk (min 1.0 (tails * tUpper n x)) })

let private fArgs (a: Args) = res {
    let! x = a.Number 0
    let! d1 = a.Number 1
    let! d2 = a.Number 2
    let! m = degrees d1
    let! n = degrees d2
    return x, m, n }

let private fDist =
    scalar "F.DIST" 4 4 "|S" (fun a -> res {
        let! x, m, n = fArgs a
        let! cumulative = a.Bool 3
        do! check (x >= 0.0) Err.num
        if cumulative then return! numOk (fst (fTails m n x))
        elif x = 0.0 then
            if m = 1.0 then return! fail Err.num
            elif m = 2.0 then return CellValue.Number 1.0
            else return CellValue.Number 0.0
        else return! numOk (fDensity m n x) })

let private fDistRT name =
    scalar name 3 3 "|S" (fun a -> res {
        let! x, m, n = fArgs a
        do! check (x >= 0.0) Err.num
        return! numOk (snd (fTails m n x)) })

let private fInv name (rightTail: bool) =
    scalar name 3 3 "|S" (fun a -> res {
        let! p, m, n = fArgs a
        if rightTail then
            do! check (p > 0.0 && p <= 1.0) Err.num
            if p = 1.0 then return CellValue.Number 0.0
            else return! numOk (fInverse m n (1.0 - p) p)
        else
            do! check (p >= 0.0 && p < 1.0) Err.num
            if p = 0.0 then return CellValue.Number 0.0
            else return! numOk (fInverse m n p (1.0 - p)) })

let private chiArgs (a: Args) = res {
    let! x = a.Number 0
    let! df = a.Number 1
    let! n = degrees df
    return x, n }

let private chisqDist =
    scalar "CHISQ.DIST" 3 3 "|S" (fun a -> res {
        let! x, n = chiArgs a
        let! cumulative = a.Bool 2
        do! check (x >= 0.0) Err.num
        if cumulative then return! numOk (fst (gammaPQ (n / 2.0) (x / 2.0)))
        elif x = 0.0 then
            if n = 1.0 then return! fail Err.num
            elif n = 2.0 then return CellValue.Number 0.5
            else return CellValue.Number 0.0
        else return! numOk (gammaDensity (n / 2.0) (x / 2.0) / 2.0) })

let private chisqDistRT name =
    scalar name 2 2 "|S" (fun a -> res {
        let! x, n = chiArgs a
        do! check (x >= 0.0) Err.num
        return! numOk (snd (gammaPQ (n / 2.0) (x / 2.0))) })

let private chisqInv name (rightTail: bool) =
    scalar name 2 2 "|S" (fun a -> res {
        let! p, n = chiArgs a
        if rightTail then
            do! check (p > 0.0 && p <= 1.0) Err.num
            if p = 1.0 then return CellValue.Number 0.0
            else return! numOk (2.0 * gammaInverse (n / 2.0) (1.0 - p) p)
        else
            do! check (p >= 0.0 && p < 1.0) Err.num
            if p = 0.0 then return CellValue.Number 0.0
            else return! numOk (2.0 * gammaInverse (n / 2.0) p (1.0 - p)) })

/// BETA.DIST/BETADIST's bounds A and B: 0 and 1 unless given (an empty B is 1).
let private bounds (a: Args) (i: int) = res {
    let! lower = a.NumberOr i 0.0
    let! upper = a.NumberOr (i + 1) 1.0
    do! check (lower < upper) Err.num
    return lower, upper }

let private shapes (a: Args) (i: int) = res {
    let! alpha = a.Number i
    let! beta = a.Number (i + 1)
    do! check (alpha > 0.0 && beta > 0.0) Err.num
    return alpha, beta }

let private betaDistAt (x: float) (alpha: float) (beta: float) (lower: float) (upper: float) (cumulative: bool) =
    if x < lower || x > upper then fail Err.num
    else
        let width = upper - lower
        let t = (x - lower) / width
        let s = (upper - x) / width
        if cumulative then numOk (fst (betaIJ alpha beta t s))
        elif t = 0.0 then (if alpha <= 1.0 then fail Err.num else Ok (CellValue.Number 0.0))
        elif s = 0.0 then (if beta <= 1.0 then fail Err.num else Ok (CellValue.Number 0.0))
        else numOk (betaPrefix alpha beta t s / (t * s) / width)

let private betaDist =
    scalar "BETA.DIST" 4 6 "|S" (fun a -> res {
        let! x = a.Number 0
        let! alpha, beta = shapes a 1
        let! cumulative = a.Bool 3
        let! lower, upper = bounds a 4
        return! betaDistAt x alpha beta lower upper cumulative })

let private betaDistLegacy =
    scalar "BETADIST" 3 5 "|S" (fun a -> res {
        let! x = a.Number 0
        let! alpha, beta = shapes a 1
        let! lower, upper = bounds a 3
        return! betaDistAt x alpha beta lower upper true })

let private betaInv name =
    scalar name 3 5 "|S" (fun a -> res {
        let! p = a.Number 0
        let! alpha, beta = shapes a 1
        let! lower, upper = bounds a 3
        do! probability p
        let x, _ = betaInverse alpha beta p (1.0 - p)
        // A quantile below the normal doubles: Excel's search gives up (#NUM!).
        do! check (x >= 2.2250738585072014e-308) Err.num
        return! numOk (lower + x * (upper - lower)) })

let private gammaArgs (a: Args) = res {
    let! x = a.Number 0
    let! alpha = a.Number 1
    let! beta = a.Number 2
    do! check (alpha > 0.0 && beta > 0.0) Err.num
    return x, alpha, beta }

let private gammaDist name =
    scalar name 4 4 "|S" (fun a -> res {
        let! x, alpha, beta = gammaArgs a
        let! cumulative = a.Bool 3
        do! check (x >= 0.0) Err.num
        if cumulative then return! numOk (fst (gammaPQ alpha (x / beta)))
        elif x = 0.0 then
            if alpha <= 1.0 then return! fail Err.num else return CellValue.Number 0.0
        else return! numOk (gammaDensity alpha (x / beta) / beta) })

let private gammaInv name =
    scalar name 3 3 "|S" (fun a -> res {
        let! p, alpha, beta = gammaArgs a
        do! check (p >= 0.0 && p < 1.0) Err.num
        if p = 0.0 then return CellValue.Number 0.0
        else return! numOk (beta * gammaInverse alpha p (1.0 - p)) })

let private gammaLn name =
    numeric1 name (fun x -> if x <= 0.0 then fail Err.num else numOk (lgamma x))

let private gammaFunction =
    numeric1 "GAMMA" (fun x ->
        if x <= 0.0 && x = Math.Floor x then fail Err.num
        else numOk (gammaFn x))

let private binomDist name =
    scalar name 4 4 "|S" (fun a -> res {
        let! s = a.Number 0
        let! trials = a.Number 1
        let! p = a.Number 2
        let! cumulative = a.Bool 3
        let s = truncate s
        let n = truncate trials
        do! check (s >= 0.0 && s <= n && p >= 0.0 && p <= 1.0) Err.num
        // Excel: the probability of no success when p = 1 is #NUM! (of 3 of 10 it is 0).
        do! check (cumulative || p < 1.0 || s > 0.0 || n = 0.0) Err.num
        return! numOk (if cumulative then binomialCdf s n p else binomialPmf s n p) })

let private binomDistRange =
    scalar "BINOM.DIST.RANGE" 3 4 "|S" (fun a -> res {
        let! trials = a.Number 0
        let! p = a.Number 1
        let! s1 = a.Number 2
        let! s2 = a.NumberOr 3 s1
        let n = truncate trials
        let s1 = truncate s1
        let s2 = truncate s2
        do! check (n >= 0.0 && p >= 0.0 && p <= 1.0 && s1 >= 0.0 && s1 <= n && s2 >= s1 && s2 <= n) Err.num
        return! numOk (sumOver s1 s2 (fun k -> binomialPmf k n p)) })

let private binomInv name =
    scalar name 3 3 "|S" (fun a -> res {
        let! trials = a.Number 0
        let! p = a.Number 1
        let! alpha = a.Number 2
        let n = truncate trials
        do! check (n >= 0.0 && p > 0.0 && p < 1.0 && alpha > 0.0 && alpha < 1.0) Err.num
        return CellValue.Number (binomialQuantile n p alpha) })

let private poissonDist name =
    scalar name 3 3 "|S" (fun a -> res {
        let! x = a.Number 0
        let! mean = a.Number 1
        let! cumulative = a.Bool 2
        let k = truncate x
        do! check (k >= 0.0 && mean >= 0.0) Err.num
        return! numOk (if cumulative then poissonCdf k mean else dpoisRaw k mean) })

let private weibullDist name =
    scalar name 4 4 "|S" (fun a -> res {
        let! x = a.Number 0
        let! alpha = a.Number 1
        let! beta = a.Number 2
        let! cumulative = a.Bool 3
        do! check (x >= 0.0 && alpha > 0.0 && beta > 0.0) Err.num
        let power = Math.Pow (x / beta, alpha)
        if cumulative then return! numOk (-(expm1 -power))
        // Excel's density at 0 is 0 for every shape.
        elif x = 0.0 then return CellValue.Number 0.0
        else return! numOk (alpha / beta * Math.Pow (x / beta, alpha - 1.0) * Math.Exp -power) })

let private exponDist name =
    scalar name 3 3 "|S" (fun a -> res {
        let! x = a.Number 0
        let! lambda = a.Number 1
        let! cumulative = a.Bool 2
        do! check (x >= 0.0 && lambda > 0.0) Err.num
        if cumulative then return! numOk (-(expm1 (-lambda * x)))
        else return! numOk (lambda * Math.Exp (-lambda * x)) })

let private lognormArgs (a: Args) = res {
    let! x = a.Number 0
    let! mean = a.Number 1
    let! sd = a.Number 2
    return x, mean, sd }

let private lognormDist =
    scalar "LOGNORM.DIST" 4 4 "|S" (fun a -> res {
        let! x, mean, sd = lognormArgs a
        let! cumulative = a.Bool 3
        do! check (x > 0.0 && sd > 0.0) Err.num
        let z = (Math.Log x - mean) / sd
        if cumulative then return! numOk (NormalDistribution.cdf z)
        else return! numOk (Math.Exp (-0.5 * z * z) / (Math.Sqrt (2.0 * Math.PI) * sd * x)) })

let private lognormDistLegacy =
    scalar "LOGNORMDIST" 3 3 "|S" (fun a -> res {
        let! x, mean, sd = lognormArgs a
        do! check (x > 0.0 && sd > 0.0) Err.num
        return! numOk (NormalDistribution.cdf ((Math.Log x - mean) / sd)) })

let private lognormInv name =
    scalar name 3 3 "|S" (fun a -> res {
        let! p, mean, sd = lognormArgs a
        do! probability p
        do! check (sd > 0.0) Err.num
        return! numOk (Math.Exp (mean + sd * NormalDistribution.quantile p)) })

let private hypgeomArgs (a: Args) = res {
    let! sample = a.Number 0
    let! sampleSize = a.Number 1
    let! successes = a.Number 2
    let! population = a.Number 3
    let x = truncate sample
    let n = truncate sampleSize
    let r = truncate successes
    let total = truncate population
    // A sample outside the support is probability 0 (Excel), not an error.
    do! check (x >= 0.0 && n >= 0.0 && r >= 0.0 && total > 0.0 && n <= total && r <= total) Err.num
    return x, n, r, total }

let private hypgeomDist =
    scalar "HYPGEOM.DIST" 5 5 "|S" (fun a -> res {
        let! x, n, r, total = hypgeomArgs a
        let! cumulative = a.Bool 4
        let b = total - r
        if cumulative then
            return! numOk (min 1.0 (sumOver (max 0.0 (n - b)) (min x (min n r)) (fun k -> hypergeometricPmf k r b n)))
        else return! numOk (hypergeometricPmf x r b n) })

let private hypgeomDistLegacy =
    scalar "HYPGEOMDIST" 4 4 "|S" (fun a -> res {
        let! x, n, r, total = hypgeomArgs a
        return! numOk (hypergeometricPmf x r (total - r) n) })

let private negbinomArgs (a: Args) = res {
    let! failures = a.Number 0
    let! successes = a.Number 1
    let! p = a.Number 2
    let f = truncate failures
    let s = truncate successes
    do! check (f >= 0.0 && s >= 1.0 && p > 0.0 && p < 1.0) Err.num
    return f, s, p }

let private negbinomDist =
    scalar "NEGBINOM.DIST" 4 4 "|S" (fun a -> res {
        let! f, s, p = negbinomArgs a
        let! cumulative = a.Bool 3
        if cumulative then
            if f + s <= smallCount then return! numOk (min 1.0 (sumOver 0.0 f (fun k -> negativeBinomialPmf k s p)))
            else return! numOk (fst (betaIJ s (f + 1.0) p (1.0 - p)))
        else return! numOk (negativeBinomialPmf f s p) })

let private negbinomDistLegacy =
    scalar "NEGBINOMDIST" 3 3 "|S" (fun a -> res {
        let! f, s, p = negbinomArgs a
        return! numOk (negativeBinomialPmf f s p) })

let private confidenceT =
    scalar "CONFIDENCE.T" 3 3 "|S" (fun a -> res {
        let! alpha = a.Number 0
        let! sd = a.Number 1
        let! size = a.Number 2
        let n = truncate size
        do! check (alpha > 0.0 && alpha < 1.0 && sd > 0.0 && n >= 1.0) Err.num
        do! check (n >= 2.0) Err.div0
        return! numOk (tInverseUpper (n - 1.0) (alpha / 2.0) * sd / Math.Sqrt n) })

// ----- tests -----------------------------------------------------------------------

let private sampleStats (xs: float[]) =
    let n = float xs.Length
    let m = Array.average xs
    let ss = xs |> Array.sumBy (fun x -> (x - m) * (x - m))
    n, m, ss

let private tTest name =
    scalar name 4 4 "RRSS" (fun a -> res {
        let! tailsArg = a.Number 2
        let! typeArg = a.Number 3
        let tails = truncate tailsArg
        let kind = truncate typeArg
        do! check (tails = 1.0 || tails = 2.0) Err.num
        do! check (kind >= 1.0 && kind <= 3.0) Err.num
        let! t, df =
            if kind = 1.0 then res {
                let! pairs = pairedNumbers a 0 1
                let d = pairs |> Array.map (fun (x, y) -> x - y)
                do! check (d.Length >= 2) Err.div0
                let n, m, ss = sampleStats d
                do! check (ss > 0.0) Err.div0
                return m / Math.Sqrt (ss / (n - 1.0) / n), n - 1.0 }
            else res {
                let! xs = a.NumbersBetween 0 0
                let! ys = a.NumbersBetween 1 1
                do! check (xs.Length >= 2 && ys.Length >= 2) Err.div0
                let n1, m1, ss1 = sampleStats xs
                let n2, m2, ss2 = sampleStats ys
                if kind = 2.0 then
                    let pooled = (ss1 + ss2) / (n1 + n2 - 2.0)
                    do! check (pooled > 0.0) Err.div0
                    return (m1 - m2) / Math.Sqrt (pooled * (1.0 / n1 + 1.0 / n2)), n1 + n2 - 2.0
                else
                    let v1 = ss1 / (n1 - 1.0) / n1
                    let v2 = ss2 / (n2 - 1.0) / n2
                    do! check (v1 + v2 > 0.0) Err.div0
                    let df = (v1 + v2) * (v1 + v2) / (v1 * v1 / (n1 - 1.0) + v2 * v2 / (n2 - 1.0))
                    return (m1 - m2) / Math.Sqrt (v1 + v2), df }
        return! numOk (tails * tUpper df (abs t)) })

let private fTest name =
    scalar name 2 2 "RR" (fun a -> res {
        let! xs = a.NumbersBetween 0 0
        let! ys = a.NumbersBetween 1 1
        do! check (xs.Length > 0 && ys.Length > 0) Err.value
        do! check (xs.Length >= 2 && ys.Length >= 2) Err.div0
        let n1, _, ss1 = sampleStats xs
        let n2, _, ss2 = sampleStats ys
        let v1 = ss1 / (n1 - 1.0)
        let v2 = ss2 / (n2 - 1.0)
        do! check (v1 > 0.0 && v2 > 0.0) Err.div0
        let lower, upper = fTails (n1 - 1.0) (n2 - 1.0) (v1 / v2)
        return! numOk (2.0 * min lower upper) })

let private chisqTest name =
    scalar name 2 2 "RR" (fun a -> res {
        let actual = a.Array 0
        let expected = a.Array 1
        let rows = actual.Length
        let cols = if rows > 0 then actual.[0].Length else 0
        // Tables of the same number of cells are paired row by row (Excel
        // pairs a row with a column of the same length).
        let actualCells = Array.concat actual
        let expectedCells = Array.concat expected
        do! check (rows > 0 && actualCells.Length = expectedCells.Length) Err.na
        let! stat = chiSquareStatistic actualCells expectedCells
        let df = if rows > 1 && cols > 1 then float ((rows - 1) * (cols - 1)) else float (rows * cols - 1)
        do! check (df >= 1.0) Err.na
        return! numOk (snd (gammaPQ (df / 2.0) (stat / 2.0))) })

let functions : FunctionDef list =
    [ tDist
      tDist2T "T.DIST.2T"
      tDistRT
      tInv
      tInv2T "T.INV.2T" false
      tInv2T "TINV" true
      tDistLegacy
      fDist
      fDistRT "F.DIST.RT"
      fDistRT "FDIST"
      fInv "F.INV" false
      fInv "F.INV.RT" true
      fInv "FINV" true
      chisqDist
      chisqDistRT "CHISQ.DIST.RT"
      chisqDistRT "CHIDIST"
      chisqInv "CHISQ.INV" false
      chisqInv "CHISQ.INV.RT" true
      chisqInv "CHIINV" true
      betaDist
      betaDistLegacy
      betaInv "BETA.INV"
      betaInv "BETAINV"
      gammaFunction
      gammaDist "GAMMA.DIST"
      gammaDist "GAMMADIST"
      gammaInv "GAMMA.INV"
      gammaInv "GAMMAINV"
      gammaLn "GAMMALN"
      gammaLn "GAMMALN.PRECISE"
      binomDist "BINOM.DIST"
      binomDist "BINOMDIST"
      binomDistRange
      binomInv "BINOM.INV"
      binomInv "CRITBINOM"
      poissonDist "POISSON.DIST"
      poissonDist "POISSON"
      weibullDist "WEIBULL.DIST"
      weibullDist "WEIBULL"
      lognormDist
      lognormDistLegacy
      lognormInv "LOGNORM.INV"
      lognormInv "LOGINV"
      exponDist "EXPON.DIST"
      exponDist "EXPONDIST"
      hypgeomDist
      hypgeomDistLegacy
      negbinomDist
      negbinomDistLegacy
      confidenceT
      tTest "T.TEST"
      tTest "TTEST"
      fTest "F.TEST"
      fTest "FTEST"
      chisqTest "CHISQ.TEST"
      chisqTest "CHITEST" ]

let install () = registerFunctions moduleName functions
