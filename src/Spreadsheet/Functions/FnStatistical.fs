/// Statistical functions of the function table: moments, the A variants of
/// the deviations, exclusive percentiles, the normal-distribution helpers,
/// the least-squares fits (LINEST, LOGEST, TREND, GROWTH).
module Rbt.Spreadsheet.Functions.FnStatistical

open System
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Functions.Kit

let moduleName = "Statistical"

/// The standard normal distribution of NormalDistribution.fs (NORM.S.DIST's).
let private normalCdf (z: float) = NormalDistribution.cdf z

let private variance (sample: bool) (xs: float[]) : Result<float, string> =
    let n = xs.Length
    if n = 0 || (sample && n < 2) then Result.Error Err.div0
    else
        let m = Array.average xs
        Ok ((xs |> Array.sumBy (fun x -> (x - m) * (x - m))) / float (if sample then n - 1 else n))

let private deviation (name: string) (sample: bool) (root: bool) =
    scalar name 1 -1 "|R" (fun a -> res {
        let! xs = a.NumbersA 0
        let! v = variance sample xs
        return! numOk (if root then Math.Sqrt v else v) })

/// The exclusive percentile: rank k(n+1) among the n sorted values, interpolated.
let private percentileExc (xs: float[]) (k: float) : Result<float, string> =
    let n = xs.Length
    if n = 0 then Result.Error Err.num
    else
        let sorted = Array.sort xs
        let rank = k * float (n + 1)
        if k <= 0.0 || k >= 1.0 || rank < 1.0 || rank > float n then Result.Error Err.num
        else
            let low = int (Math.Floor rank)
            let fraction = rank - float low
            if low >= n then Ok sorted.[n - 1]
            else Ok (sorted.[low - 1] + fraction * (sorted.[low] - sorted.[low - 1]))

/// The value truncated (not rounded) to `digits` decimals.
let private truncateTo (digits: int) (x: float) =
    let scale = Math.Pow (10.0, float digits)
    Math.Floor (x * scale + 1e-9) / scale

/// The values that occur most often (twice at least), in the order they first occur.
let private modes (xs: float[]) : float[] =
    let counts = Collections.Generic.Dictionary<float, int>()
    let order = ResizeArray<float>()
    for x in xs do
        match Portable.Dict.tryFind x counts with
        | ValueSome c -> counts.[x] <- c + 1
        | ValueNone ->
            counts.[x] <- 1
            order.Add x
    let best = if counts.Count = 0 then 0 else Seq.max counts.Values
    if best < 2 then [||] else [| for x in order do if counts.[x] = best then yield x |]

// ----- least squares (LINEST, LOGEST, TREND, GROWTH) ------------------------

/// The numbers of an array argument: anything else (text, a logical, a blank
/// cell, an error) is #VALUE!, as LINEST has it.
let private numericMatrix (rows: CellValue[][]) : Result<float[][], string> =
    let mutable bad = false
    let m =
        rows |> Array.map (Array.map (fun v ->
            match v with
            | CellValue.Number x -> x
            | _ -> bad <- true; 0.0))
    if bad then Result.Error Err.value else Ok m

/// The data of a fit: y (n values), the variables as columns of X (n x k),
/// and how new x values are laid out (`ByRows`: an observation is a row of
/// the new x's, as when y is a column; else a column of them).
type private Fit =
    { Y: float[]; X: float[][]; Vars: int; ByRows: bool; YRows: int; YCols: int }

/// 1, 2, 3, ... in the shape of y (the x's when they are left out).
let private counting (rows: int) (cols: int) =
    Array.init rows (fun r -> Array.init cols (fun c -> float (r * cols + c + 1)))

/// Read known y's and known x's as LINEST lays them out: y a column, each
/// column of x a variable; y a row, each row of x; x of y's shape (or left
/// out) one variable. Shapes that do not fit are #REF!.
let private fitData (a: Args) (logY: bool) : Result<Fit, string> =
    res {
        let! ym = numericMatrix (a.Array 0)
        let yRows = ym.Length
        let yCols = if yRows > 0 then ym.[0].Length else 0
        do! check (yRows > 0 && yCols > 0) Err.value
        do! check (not logY || ym |> Array.forall (Array.forall (fun v -> v > 0.0))) Err.num
        let y = Array.concat ym |> Array.map (fun v -> if logY then Math.Log v else v)
        let n = y.Length
        let! xm = if a.Supplied 1 then numericMatrix (a.Array 1) else Ok (counting yRows yCols)
        let xRows = xm.Length
        let xCols = if xRows > 0 then xm.[0].Length else 0
        if xRows = yRows && xCols = yCols then
            return { Y = y; X = Array.init n (fun i -> [| xm.[i / yCols].[i % yCols] |]); Vars = 1; ByRows = (yCols = 1); YRows = yRows; YCols = yCols }
        elif yCols = 1 && xRows = n then
            return { Y = y; X = xm; Vars = xCols; ByRows = true; YRows = yRows; YCols = yCols }
        elif yRows = 1 && xCols = n then
            return { Y = y; X = Array.init n (fun i -> Array.init xRows (fun j -> xm.[j].[i])); Vars = xRows; ByRows = false; YRows = yRows; YCols = yCols }
        else return! fail Err.ref }

/// A least-squares solution: the coefficients (0 for a column left out), the
/// intercept, which columns were kept, (Xc'Xc)^-1 of the kept ones (in the
/// original column order, 0 elsewhere), the means the columns were centred by
/// and the number of parameters Excel counts: the kept columns and the columns
/// that are 0 throughout (a constant x with an intercept) - only a column
/// collinear with other columns is not counted.
type private Solution = { Coef: float[]; Intercept: float; Kept: bool[]; Inverse: float[][]; Means: float[]; Parameters: int }

/// The sum of f i for i = first .. last.
let private sumFrom (first: int) (last: int) (f: int -> float) =
    let mutable s = 0.0
    for i in first .. last do s <- s + f i
    s

/// The least-squares fit of y on the columns of X (with an intercept when
/// `constant`, by centring), by Householder QR with column pivoting: the column
/// with the largest remaining norm first; a column whose remaining norm is
/// below 1e-13 of its own is collinear with those before it and left out
/// (coefficient 0, as Excel reports it).
let private leastSquares (y: float[]) (x: float[][]) (k: int) (constant: bool) : Solution =
    let n = y.Length
    let means = Array.init k (fun j -> if constant then sumFrom 0 (n - 1) (fun i -> x.[i].[j]) / float n else 0.0)
    let yMean = if constant then Array.average y else 0.0
    let a = Array.init n (fun i -> Array.init k (fun j -> x.[i].[j] - means.[j]))
    let b = Array.init n (fun i -> y.[i] - yMean)
    let original = Array.init k (fun j -> Math.Sqrt (sumFrom 0 (n - 1) (fun i -> a.[i].[j] * a.[i].[j])))
    let order = Array.init k id
    let mutable rank = 0
    let mutable stop = false
    while not stop && rank < min n k do
        let s = rank
        let mutable best = -1
        let mutable bestNorm = 0.0
        for c in s .. k - 1 do
            let r = Math.Sqrt (sumFrom s (n - 1) (fun i -> a.[i].[c] * a.[i].[c]))
            if r > 0.0 && r > 1e-13 * original.[order.[c]] && r > bestNorm then
                best <- c
                bestNorm <- r
        if best < 0 then stop <- true
        else
            if best <> s then
                for i in 0 .. n - 1 do
                    let t = a.[i].[s]
                    a.[i].[s] <- a.[i].[best]
                    a.[i].[best] <- t
                let t = order.[s]
                order.[s] <- order.[best]
                order.[best] <- t
            let alpha = if a.[s].[s] > 0.0 then -bestNorm else bestNorm
            let v = Array.init n (fun i -> if i < s then 0.0 elif i = s then a.[s].[s] - alpha else a.[i].[s])
            let vv = sumFrom s (n - 1) (fun i -> v.[i] * v.[i])
            if vv > 0.0 then
                for j in s + 1 .. k - 1 do
                    let f = 2.0 * sumFrom s (n - 1) (fun i -> v.[i] * a.[i].[j]) / vv
                    for i in s .. n - 1 do a.[i].[j] <- a.[i].[j] - f * v.[i]
                let f = 2.0 * sumFrom s (n - 1) (fun i -> v.[i] * b.[i]) / vv
                for i in s .. n - 1 do b.[i] <- b.[i] - f * v.[i]
            a.[s].[s] <- alpha
            for i in s + 1 .. n - 1 do a.[i].[s] <- 0.0
            rank <- rank + 1
    // R z = Q'y by back substitution; R^-1 for the covariance.
    let z = Array.zeroCreate rank
    for i in rank - 1 .. -1 .. 0 do
        z.[i] <- (b.[i] - sumFrom (i + 1) (rank - 1) (fun j -> a.[i].[j] * z.[j])) / a.[i].[i]
    let rinv : float[][] = Array.init rank (fun _ -> Array.zeroCreate rank)
    for i in rank - 1 .. -1 .. 0 do
        rinv.[i].[i] <- 1.0 / a.[i].[i]
        for j in i + 1 .. rank - 1 do
            rinv.[i].[j] <- -(sumFrom (i + 1) j (fun m -> a.[i].[m] * rinv.[m].[j])) / a.[i].[i]
    let coef = Array.zeroCreate k
    let kept = Array.create k false
    let inverse : float[][] = Array.init k (fun _ -> Array.zeroCreate k)
    for i in 0 .. rank - 1 do
        coef.[order.[i]] <- z.[i]
        kept.[order.[i]] <- true
        for j in 0 .. rank - 1 do
            inverse.[order.[i]].[order.[j]] <- sumFrom (max i j) (rank - 1) (fun m -> rinv.[i].[m] * rinv.[j].[m])
    let intercept = if constant then yMean - sumFrom 0 (k - 1) (fun j -> coef.[j] * means.[j]) else 0.0
    let zero = original |> Array.filter (fun v -> v = 0.0) |> Array.length
    { Coef = coef; Intercept = intercept; Kept = kept; Inverse = inverse; Means = means; Parameters = rank + zero }

let private predict (sol: Solution) (row: float[]) =
    sol.Intercept + sumFrom 0 (row.Length - 1) (fun j -> sol.Coef.[j] * row.[j])

/// LINEST/LOGEST: the coefficients (the last variable's first, the intercept
/// last) and, with `stats`, the four rows of statistics, #N/A where a row has
/// no value. LOGEST fits ln y and shows e^coefficient; its statistics are those
/// of the logarithmic fit.
let private linest (exponential: bool) (a: Args) : Result<CellValue[][], string> =
    res {
        let! fit = fitData a exponential
        let! constant = a.BoolOr 2 true
        let! stats = a.BoolOr 3 false
        let k = fit.Vars
        let n = fit.Y.Length
        let sol = leastSquares fit.Y fit.X k constant
        let shown (v: float) = num (if exponential then Math.Exp v else v)
        let first = Array.append [| for j in k - 1 .. -1 .. 0 -> shown sol.Coef.[j] |] [| shown sol.Intercept |]
        if not stats then return [| first |]
        else
            let kept = sol.Parameters
            // As many parameters as values: an exact fit (Excel's 0 residual, not rounding's).
            let df = max 0 (n - kept - (if constant then 1 else 0))
            let yMean = if constant then Array.average fit.Y else 0.0
            let ssTotal = fit.Y |> Array.sumBy (fun v -> (v - yMean) * (v - yMean))
            let ssResid =
                if df = 0 then 0.0
                else Array.map2 (fun yv row -> let r = yv - predict sol row in r * r) fit.Y fit.X |> Array.sum
            let ssReg = max 0.0 (ssTotal - ssResid)
            let na = err Err.na
            let seY = if ssResid = 0.0 then 0.0 else Math.Sqrt (ssResid / float df)
            let variance = seY * seY
            let seCoef = [| for j in k - 1 .. -1 .. 0 -> if sol.Kept.[j] then num (Math.Sqrt (variance * sol.Inverse.[j].[j])) else CellValue.Number 0.0 |]
            let seIntercept =
                if not constant then na
                else
                    // Var(b) = s^2 (1/n + m' (Xc'Xc)^-1 m), m the means of the columns.
                    let q = sumFrom 0 (k - 1) (fun i -> sumFrom 0 (k - 1) (fun j -> sol.Means.[i] * sol.Inverse.[i].[j] * sol.Means.[j]))
                    num (Math.Sqrt (variance * (1.0 / float n + q)))
            let r2 = if ssTotal = 0.0 then 1.0 else ssReg / ssTotal
            let f =
                if ssResid = 0.0 then err Err.num
                elif kept = 0 then CellValue.Number 0.0
                else num ((ssReg / float kept) / (ssResid / float df))
            let pad (cells: CellValue list) = Array.append (Array.ofList cells) (Array.create (k + 1 - cells.Length) na)
            return
                [| first
                   Array.append seCoef [| seIntercept |]
                   pad [ num r2; num seY ]
                   pad [ f; CellValue.Number (float df) ]
                   pad [ num ssReg; num ssResid ] |] }

/// TREND/GROWTH: the fitted values at the new x's (by default the known x's):
/// of one variable in the new x's shape, of several one per row (y a column)
/// or column (y a row) of them.
let private trend (exponential: bool) (a: Args) : Result<CellValue[][], string> =
    res {
        let! fit = fitData a exponential
        let! constant = a.BoolOr 3 true
        let k = fit.Vars
        let sol = leastSquares fit.Y fit.X k constant
        let shown (v: float) = num (if exponential then Math.Exp v else v)
        let! fresh =
            if a.Supplied 2 then numericMatrix (a.Array 2)
            elif a.Supplied 1 then numericMatrix (a.Array 1)
            else Ok (counting fit.YRows fit.YCols)
        let rows = fresh.Length
        let cols = if rows > 0 then fresh.[0].Length else 0
        if k = 1 then
            return fresh |> Array.map (Array.map (fun v -> shown (predict sol [| v |])))
        elif fit.ByRows then
            do! check (cols = k) Err.ref
            return fresh |> Array.map (fun row -> [| shown (predict sol row) |])
        else
            do! check (rows = k) Err.ref
            return [| Array.init cols (fun c -> shown (predict sol (Array.init k (fun j -> fresh.[j].[c])))) |] }


// ----- FORECAST.ETS and its kin -------------------------------------------------
//
// Additive error, trend and season (AAA) exponential smoothing in the error
// correction form, as Excel 2016+ describes it. Measured in Excel 2019
// (tools/excel-parity/make-ets-probe.py):
//  - Without a season the model is Excel's exactly: the level starts at the
//    first value, the trend at the least-squares slope of the whole series;
//    l = l + b + alpha e, b = b + beta e (beta multiplies the error, not the
//    level's change); the errors from the second point on, the first counted
//    as 0 in MAE, RMSE, SMAPE; MASE against the naive one-step forecast.
//    Given Excel's alpha and beta this reproduces its statistics and
//    forecasts to the last digit (for 8 or more points; 4 and 5 points start
//    the trend otherwise).
//  - How Excel chooses alpha, beta and gamma is not reproduced: its values lie
//    on a coarse lattice (0.001 apart from 1/n-like points: 0.002, 0.126,
//    0.251, 1/6, 0.5, 0.501, 0.75, 0.998 ...) and are not the minimum of the
//    in-sample squared or absolute errors. Here they minimize the squared
//    errors over that range, by a lattice search refined coordinate-wise.
//  - The season's start, the detection of its length (here the strongest
//    autocorrelation of the detrended series) and the confidence interval's
//    spread past one step are this module's, not measured.
//  - The rules of the arguments are Excel's: values and timeline of the same
//    count (#N/A); points without a number left out; the timeline sorted, its
//    step the least difference; repeated dates aggregated (1 AVERAGE, 2 COUNT,
//    3 COUNTA, 4 MAX, 5 MEDIAN, 6 MIN, 7 SUM; else #NUM!); missing steps
//    completed (1 linear interpolation, 0 zeros; else #NUM!); seasonality 0,
//    1 (detect) or a whole number (else #NUM!; 8761 too), a season too long
//    for the data dropped; a target before the timeline #NUM!, on it the
//    series' value; statistic_type 1 to 8, confidence_level 0 <= c < 1.

type private EtsFit =
    { Alpha: float; Beta: float; Gamma: float; Season: int
      Level: float; Trend: float; Seasonals: float[]; Errors: float[]; Fitted: float[]
      Series: float[]; Start: float; Step: float }

/// The values on a regular timeline: (series, first time, step).
let private etsSeries (a: Args) (vi: int) (ti: int) (completion: float) (aggregation: float) : Result<float[] * float * float, string> =
    let vs = a.Flat vi
    let ts = a.Flat ti
    if vs.Length <> ts.Length then Result.Error Err.na
    else
        let pairs =
            [| for k in 0 .. vs.Length - 1 do
                 match ts.[k], vs.[k] with
                 | CellValue.Number t, v -> yield t, v
                 | _ -> () |]
        match pairs |> Array.tryPick (fun (_, v) -> match v with CellValue.Error e -> Some e | _ -> None) with
        | Some e -> Result.Error e
        | None ->
        let groups = pairs |> Array.groupBy fst |> Array.sortBy fst
        if groups.Length < 2 then Result.Error Err.num
        else
            let times = groups |> Array.map fst
            let step = Array.init (times.Length - 1) (fun k -> times.[k + 1] - times.[k]) |> Array.min
            let aggregate (items: CellValue[]) =
                let nums = items |> Array.choose (fun v -> match v with CellValue.Number x -> Some x | _ -> None)
                match int aggregation with
                | 2 -> Some (float nums.Length)
                | 3 -> Some (float (items |> Array.filter (fun v -> match v with CellValue.Empty -> false | _ -> true) |> Array.length))
                | _ when nums.Length = 0 -> None
                | 4 -> Some (Array.max nums)
                | 5 -> let s = Array.sort nums in Some (if s.Length % 2 = 1 then s.[s.Length / 2] else (s.[s.Length / 2 - 1] + s.[s.Length / 2]) / 2.0)
                | 6 -> Some (Array.min nums)
                | 7 -> Some (Array.sum nums)
                | _ -> Some (Array.average nums)
            let count = int (Math.Round ((times.[times.Length - 1] - times.[0]) / step)) + 1
            if count > 100000 then Result.Error Err.num
            else
                let slots : float option[] = Array.create count None
                for (t, items) in groups do
                    let k = int (Math.Round ((t - times.[0]) / step))
                    slots.[k] <- aggregate (items |> Array.map snd)
                // Completion: linear between the known neighbours, or zeros.
                let ys = Array.zeroCreate count
                let known = [| for k in 0 .. count - 1 do if slots.[k].IsSome then yield k |]
                if known.Length < 2 then Result.Error Err.num
                else
                    for k in 0 .. count - 1 do
                        ys.[k] <-
                            match slots.[k] with
                            | Some v -> v
                            | None when completion = 0.0 -> 0.0
                            | None ->
                                let lo = known |> Array.filter (fun j -> j < k)
                                let hi = known |> Array.filter (fun j -> j > k)
                                match lo, hi with
                                | [||], _ -> slots.[hi.[0]].Value
                                | _, [||] -> slots.[lo.[lo.Length - 1]].Value
                                | _ ->
                                    let j0, j1 = lo.[lo.Length - 1], hi.[0]
                                    let v0, v1 = slots.[j0].Value, slots.[j1].Value
                                    v0 + (v1 - v0) * float (k - j0) / float (j1 - j0)
                    Ok (ys, times.[0], step)

let private olsLine (ys: float[]) =
    let n = float ys.Length
    let mx = (n - 1.0) / 2.0
    let my = Array.average ys
    let mutable sxy = 0.0
    let mutable sxx = 0.0
    for i in 0 .. ys.Length - 1 do
        sxy <- sxy + (float i - mx) * (ys.[i] - my)
        sxx <- sxx + (float i - mx) * (float i - mx)
    let slope = if sxx = 0.0 then 0.0 else sxy / sxx
    my - slope * mx, slope

/// The smoothing of ys with the parameters: the final states and the one-step errors.
let private etsRun (ys: float[]) (m: int) (alpha: float) (beta: float) (gamma: float) =
    let n = ys.Length
    let c, slope = olsLine ys
    let errors = Array.zeroCreate n
    let fitted = Array.copy ys
    if m = 0 then
        let mutable l = ys.[0]
        let mutable b = slope
        for i in 1 .. n - 1 do
            let f = l + b
            let e = ys.[i] - f
            fitted.[i] <- f
            errors.[i] <- e
            l <- l + b + alpha * e
            b <- b + beta * e
        l, b, [||], errors, fitted
    else
        // The first season from the detrended series, averaged over the whole seasons.
        let full = max 1 (n / m)
        let seasonal = Array.init m (fun p ->
            let vals = [| for j in 0 .. full - 1 do let i = j * m + p in if i < n then yield ys.[i] - (c + slope * float i) |]
            Array.average vals)
        let mean = Array.average seasonal
        let s = ResizeArray(seasonal |> Array.map (fun v -> v - mean))
        let mutable l = c + slope * float (m - 1)
        let mutable b = slope
        for i in m .. n - 1 do
            let f = l + b + s.[i - m]
            let e = ys.[i] - f
            fitted.[i] <- f
            errors.[i] <- e
            l <- l + b + alpha * e
            b <- b + beta * e
            s.Add(s.[i - m] + gamma * e)
        l, b, s.ToArray(), errors, fitted

let private sse (errors: float[]) = errors |> Array.sumBy (fun e -> e * e)

/// The parameters with the least squared errors: a lattice, then coordinate refinement.
let private etsFit (ys: float[]) (m: int) (start: float) (step: float) : EtsFit =
    let lattice = [| 0.001; 0.01; 0.05; 0.1; 0.2; 0.3; 0.4; 0.5; 0.6; 0.7; 0.8; 0.9; 0.999 |]
    let cost (p: float[]) = let _, _, _, e, _ = etsRun ys m p.[0] p.[1] p.[2] in sse e
    let gammas = if m = 0 then [| 0.0 |] else lattice
    let mutable best = [| 0.5; 0.001; (if m = 0 then 0.0 else 0.001) |]
    let mutable bestCost = cost best
    for a in lattice do
        for b in lattice do
            for g in gammas do
                let p = [| a; b; g |]
                let c = cost p
                if c < bestCost then
                    best <- p
                    bestCost <- c
    let dims = if m = 0 then 2 else 3
    let mutable delta = 0.05
    while delta >= 0.0005 do
        let mutable improved = true
        while improved do
            improved <- false
            for d in 0 .. dims - 1 do
                for sign in [ -1.0; 1.0 ] do
                    let p = Array.copy best
                    p.[d] <- max 0.001 (min 0.999 (p.[d] + sign * delta))
                    let c = cost p
                    if c < bestCost - 1e-12 * (1.0 + bestCost) then
                        best <- p
                        bestCost <- c
                        improved <- true
        delta <- delta / 2.0
    let l, b, s, e, f = etsRun ys m best.[0] best.[1] best.[2]
    { Alpha = best.[0]; Beta = best.[1]; Gamma = (if m = 0 then 2.220446049250313e-16 else best.[2]); Season = m
      Level = l; Trend = b; Seasonals = s; Errors = e; Fitted = f; Series = ys; Start = start; Step = step }

/// The season length found in ys: the lag of the strongest autocorrelation of
/// the detrended series, when it is clear (above 0.5 and a peak), else 0.
let private detectSeason (ys: float[]) =
    let n = ys.Length
    let c, slope = olsLine ys
    let r = Array.init n (fun i -> ys.[i] - (c + slope * float i))
    let var = r |> Array.sumBy (fun v -> v * v)
    if var = 0.0 || n < 6 then 0
    else
        let acf k = (Array.init (n - k) (fun i -> r.[i] * r.[i + k]) |> Array.sum) / var
        let maxLag = min 8760 (n / 2)
        let mutable best = 0
        let mutable bestValue = 0.5
        for k in 2 .. maxLag do
            let v = acf k
            if v > bestValue && v >= acf (k - 1) && (k + 1 > n - 1 || v >= acf (k + 1)) then
                best <- k
                bestValue <- v
        best

/// The arguments shared by the four functions, from index `first`:
/// seasonality (when `withSeason`), data_completion, aggregation; the fit.
let private etsArgs (a: Args) (vi: int) (ti: int) (first: int) (withSeason: bool) : Result<EtsFit, string> = res {
    let i = if withSeason then first + 1 else first
    let! seasonality = if withSeason then a.NumberOr first 1.0 else Ok 1.0
    let! completion = a.NumberOr i 1.0
    let! aggregation = a.NumberOr (i + 1) 1.0
    do! check (seasonality >= 0.0 && seasonality = Math.Floor seasonality) Err.num
    do! check (completion = 0.0 || completion = 1.0) Err.num
    do! check (aggregation >= 1.0 && aggregation <= 7.0 && aggregation = Math.Floor aggregation) Err.num
    let! ys, start, step = etsSeries a vi ti completion aggregation
    let m =
        match int seasonality with
        | 1 -> detectSeason ys
        | k -> k
    // A season needs two of it in the data.
    let m = if m >= 2 && 2 * m <= ys.Length then m else 0
    return etsFit ys m start step }

/// The forecast h steps (a real number) after the last point.
let private etsForecast (fit: EtsFit) (target: float) : Result<float, string> =
    let n = fit.Series.Length
    let pos = (target - fit.Start) / fit.Step
    if pos < 0.0 then Result.Error Err.num
    elif pos <= float (n - 1) then
        let k = int (Math.Floor pos)
        if k >= n - 1 then Ok fit.Series.[n - 1]
        else Ok (fit.Series.[k] + (fit.Series.[k + 1] - fit.Series.[k]) * (pos - float k))
    else
        let h = pos - float (n - 1)
        let seasonal =
            if fit.Season = 0 then 0.0
            else
                let s = fit.Seasonals
                let k = int (Math.Ceiling h)
                let idx = s.Length - fit.Season + ((k - 1) % fit.Season)
                s.[idx]
        Ok (fit.Level + h * fit.Trend + seasonal)

let private etsStatistic (fit: EtsFit) (kind: int) =
    let n = float fit.Series.Length
    let errs = fit.Errors
    match kind with
    | 1 -> fit.Alpha
    | 2 -> fit.Beta
    | 3 -> fit.Gamma
    | 4 ->
        let ys = fit.Series
        let naive = (Array.init (ys.Length - 1) (fun i -> abs (ys.[i + 1] - ys.[i])) |> Array.sum) / (n - 1.0)
        let mae = (errs |> Array.sumBy abs) / n
        if naive = 0.0 then 0.0 else mae / naive
    | 5 ->
        (Array.init errs.Length (fun i ->
            let d = (abs fit.Series.[i] + abs fit.Fitted.[i]) / 2.0
            if d = 0.0 then 0.0 else abs errs.[i] / d) |> Array.sum) / n
    | 6 -> (errs |> Array.sumBy abs) / n
    | 7 -> Math.Sqrt (sse errs / n)
    | _ -> fit.Step

let private etsFunctions : FunctionDef list =
    [ scalar "FORECAST.ETS" 3 6 "SRRSSS" (fun a -> res {
        let! target = a.Number 0
        let! fit = etsArgs a 1 2 3 true
        let! v = etsForecast fit target
        return! numOk v })
      scalar "FORECAST.ETS.CONFINT" 3 7 "SRRSSSS" (fun a -> res {
        let! target = a.Number 0
        let! level = a.NumberOr 3 0.95
        do! check (level >= 0.0 && level < 1.0) Err.num
        let! fit = etsArgs a 1 2 4 true
        let! _ = etsForecast fit target
        let n = fit.Series.Length
        let h = max 1.0 (Math.Ceiling ((target - fit.Start) / fit.Step - float (n - 1)))
        let rmse = Math.Sqrt (sse fit.Errors / float n)
        let al, be = fit.Alpha, fit.Beta
        // The spread of the h-step error of the AAN model (Hyndman et al.).
        let factor = 1.0 + (h - 1.0) * (al * al + al * be * h + be * be * h * (2.0 * h - 1.0) / 6.0)
        let z = NormalDistribution.quantile (0.5 + level / 2.0)
        return! numOk (if level = 0.0 then 0.0 else z * rmse * Math.Sqrt factor) })
      scalar "FORECAST.ETS.SEASONALITY" 2 4 "RRSS" (fun a -> res {
        let! fit = etsArgs a 0 1 2 false
        return! numOk (float fit.Season) })
      scalar "FORECAST.ETS.STAT" 3 6 "RRSSSS" (fun a -> res {
        let! kind = a.Number 2
        let kind = truncate kind
        do! check (kind >= 1.0 && kind <= 8.0) Err.num
        let! fit = etsArgs a 0 1 3 true
        return! numOk (etsStatistic fit (int kind)) }) ]

let functions : FunctionDef list =
    etsFunctions @
    [ arrayed "LINEST" 1 4 "" (linest false)
      arrayed "LOGEST" 1 4 "" (linest true)
      arrayed "TREND" 1 4 "" (trend false)
      arrayed "GROWTH" 1 4 "" (trend true)
      scalar "KURT" 1 -1 "|R" (fun a -> res {
        let! xs = a.Numbers 0
        let n = float xs.Length
        do! check (xs.Length >= 4) Err.div0
        let m = Array.average xs
        let s = Math.Sqrt ((xs |> Array.sumBy (fun x -> (x - m) * (x - m))) / (n - 1.0))
        do! check (s > 0.0) Err.div0
        let fourth = xs |> Array.sumBy (fun x -> Math.Pow ((x - m) / s, 4.0))
        return! numOk (n * (n + 1.0) / ((n - 1.0) * (n - 2.0) * (n - 3.0)) * fourth - 3.0 * (n - 1.0) * (n - 1.0) / ((n - 2.0) * (n - 3.0))) })
      scalar "SKEW" 1 -1 "|R" (fun a -> res {
        let! xs = a.Numbers 0
        let n = float xs.Length
        do! check (xs.Length >= 3) Err.div0
        let m = Array.average xs
        let s = Math.Sqrt ((xs |> Array.sumBy (fun x -> (x - m) * (x - m))) / (n - 1.0))
        do! check (s > 0.0) Err.div0
        return! numOk (n / ((n - 1.0) * (n - 2.0)) * (xs |> Array.sumBy (fun x -> Math.Pow ((x - m) / s, 3.0)))) })
      scalar "SKEW.P" 1 -1 "|R" (fun a -> res {
        let! xs = a.Numbers 0
        let n = float xs.Length
        do! check (xs.Length >= 1) Err.div0
        let m = Array.average xs
        let s = Math.Sqrt ((xs |> Array.sumBy (fun x -> (x - m) * (x - m))) / n)
        do! check (s > 0.0) Err.div0
        return! numOk ((xs |> Array.sumBy (fun x -> Math.Pow ((x - m) / s, 3.0))) / n) })
      deviation "STDEVA" true true
      deviation "STDEVPA" false true
      deviation "VARA" true false
      deviation "VARPA" false false
      scalar "TRIMMEAN" 2 2 "RS" (fun a -> res {
        let! xs = a.NumbersBetween 0 0
        let! percent = a.Number 1
        do! check (percent >= 0.0 && percent < 1.0) Err.num
        do! check (xs.Length > 0) Err.num
        // The number left out, rounded down to an even one, half from each end.
        let excluded = int (Math.Floor (float xs.Length * percent / 2.0)) * 2
        let kept = Array.sort xs |> Array.skip (excluded / 2) |> Array.take (xs.Length - excluded)
        return! numOk (Array.average kept) })
      scalar "PERCENTILE.EXC" 2 2 "RS" (fun a -> res {
        let! xs = a.NumbersBetween 0 0
        let! k = a.Number 1
        let! p = percentileExc xs k
        return! numOk p })
      scalar "QUARTILE.EXC" 2 2 "RS" (fun a -> res {
        let! xs = a.NumbersBetween 0 0
        let! q = a.Number 1
        let q = truncate q
        do! check (q >= 1.0 && q <= 3.0) Err.num
        let! p = percentileExc xs (q / 4.0)
        return! numOk p })
      scalar "PERCENTRANK.EXC" 2 3 "RSS" (fun a -> res {
        let! xs = a.NumbersBetween 0 0
        let! x = a.Number 1
        let! significance = a.NumberOr 2 3.0
        let digits = int (truncate significance)
        do! check (digits >= 1) Err.num
        do! check (xs.Length > 0) Err.na
        let sorted = Array.sort xs
        let n = sorted.Length
        do! check (x >= sorted.[0] && x <= sorted.[n - 1]) Err.na
        let rankOf (v: float) = float ((sorted |> Array.filter (fun y -> y < v) |> Array.length) + 1) / float (n + 1)
        let rank =
            if sorted |> Array.contains x then rankOf x
            else
                let lower = sorted |> Array.filter (fun y -> y < x) |> Array.max
                let upper = sorted |> Array.filter (fun y -> y > x) |> Array.min
                let r0 = rankOf lower
                r0 + (x - lower) / (upper - lower) * (rankOf upper - r0)
        return! numOk (truncateTo digits rank) })
      arrayed "MODE.MULT" 1 -1 "" (fun a -> res {
        // A text or logical written as an argument is #VALUE!, as for MODE.SNGL.
        do! check (a.Exprs |> Array.forall (function Expr.Str _ | Expr.Bool _ -> false | _ -> true)) Err.value
        let! xs = a.Numbers 0
        let found = modes xs
        do! check (found.Length > 0) Err.na
        return found |> Array.map (fun x -> [| CellValue.Number x |]) })
      numeric2 "PERMUT" (fun n k ->
        let n = truncate n
        let k = truncate k
        if n <= 0.0 && not (n = 0.0 && k = 0.0) || k < 0.0 || n < k then fail Err.num
        else
            let mutable product = 1.0
            let mutable i = 0.0
            while i < k do
                product <- product * (n - i)
                i <- i + 1.0
            numOk product)
      numeric2 "PERMUTATIONA" (fun n k ->
        let n = truncate n
        let k = truncate k
        if n < 0.0 || k < 0.0 then fail Err.num else numOk (Math.Pow (n, k)))
      scalar "PROB" 3 4 "RRSS" (fun a -> res {
        let xs = a.Flat 0
        let ps = a.Flat 1
        do! check (xs.Length = ps.Length) Err.na
        let pairs =
            Array.zip xs ps |> Array.choose (fun (x, p) ->
                match x, p with
                | CellValue.Number x, CellValue.Number p -> Some (Ok (x, p))
                | CellValue.Error e, _ | _, CellValue.Error e -> Some (Result.Error e)
                | _ -> None)
        match pairs |> Array.tryPick (fun r -> match r with Result.Error e -> Some e | _ -> None) with
        | Some e -> return! fail e
        | None ->
            let pairs = pairs |> Array.choose (fun r -> match r with Ok p -> Some p | _ -> None)
            do! check (pairs |> Array.forall (fun (_, p) -> p > 0.0 && p <= 1.0)) Err.num
            do! check (abs ((pairs |> Array.sumBy snd) - 1.0) < 1e-7) Err.num
            let! lower = a.Number 2
            let! upper = a.NumberOr 3 lower
            return! numOk (pairs |> Array.sumBy (fun (x, p) -> if x >= lower && x <= upper then p else 0.0)) })
      numeric1 "FISHER" (fun x -> if x <= -1.0 || x >= 1.0 then fail Err.num else numOk (0.5 * Math.Log ((1.0 + x) / (1.0 - x))))
      numeric1 "FISHERINV" (fun y -> let e = Math.Exp (2.0 * y) in numOk ((e - 1.0) / (e + 1.0)))
      scalar "STANDARDIZE" 3 3 "|S" (fun a -> res {
        let! x = a.Number 0
        let! m = a.Number 1
        let! sd = a.Number 2
        do! check (sd > 0.0) Err.num
        return! numOk ((x - m) / sd) })
      numeric1 "PHI" (fun x -> numOk (NormalDistribution.pdf x))
      scalar "GAUSS" 1 1 "|S" (fun a -> res {
        let! z = a.Number 0
        return! numOk (normalCdf z - 0.5) })
      scalar "CONFIDENCE.NORM" 3 3 "|S" (fun a -> res {
        let! alpha = a.Number 0
        let! sd = a.Number 1
        let! size = a.Number 2
        let size = truncate size
        do! check (alpha > 0.0 && alpha < 1.0 && sd > 0.0 && size >= 1.0) Err.num
        let z = NormalDistribution.quantile (1.0 - alpha / 2.0)
        return! numOk (z * sd / Math.Sqrt size) })
      scalar "Z.TEST" 2 3 "RSS" (fun a -> res {
        let! xs = a.NumbersBetween 0 0
        let! x = a.Number 1
        do! check (xs.Length > 0) Err.na
        // The sample deviation is computed (and #DIV/0! for one value) even when sigma is given.
        let! sd = variance true xs |> Result.map Math.Sqrt
        let! sigma = if a.Supplied 2 then a.Number 2 else Ok sd
        do! check (a.Supplied 2 || sigma > 0.0) Err.div0
        do! check (sigma > 0.0) Err.num
        // The upper tail as the lower one of the mirrored statistic: Excel keeps
        // the small probabilities (7.2e-73), which 1 - P would lose.
        return! numOk (normalCdf ((x - Array.average xs) / (sigma / Math.Sqrt (float xs.Length)))) }) ]
    |> fun defs ->
        let byName name = defs |> List.find (fun d -> d.Name = name)
        defs @ [ alias "CONFIDENCE" (byName "CONFIDENCE.NORM"); alias "ZTEST" (byName "Z.TEST") ]

let install () = registerFunctions moduleName functions
