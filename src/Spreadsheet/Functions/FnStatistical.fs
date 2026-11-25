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


let functions : FunctionDef list =
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
