/// Statistical functions of the function table: moments, the A variants of
/// the deviations, exclusive percentiles, the normal-distribution helpers.
module Rbt.Spreadsheet.Functions.FnStatistical

open System
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Functions.Kit

let moduleName = "Statistical"

/// The standard normal distribution function, Formula.fs's (NORM.S.DIST).
let private normalCdf (a: Args) (z: float) =
    coerceNumber (a.Ev.Call "NORM.S.DIST" [ Number z; Bool true ] a.Ctx)

let private normalInverse (a: Args) (p: float) =
    coerceNumber (a.Ev.Call "NORM.S.INV" [ Number p ] a.Ctx)

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
        match counts.TryGetValue x with
        | true, c -> counts.[x] <- c + 1
        | false, _ ->
            counts.[x] <- 1
            order.Add x
    let best = if counts.Count = 0 then 0 else Seq.max counts.Values
    if best < 2 then [||] else [| for x in order do if counts.[x] = best then yield x |]

let functions : FunctionDef list =
    [ scalar "KURT" 1 -1 "|R" (fun a -> res {
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
      numeric1 "PHI" (fun x -> numOk (Math.Exp (-x * x / 2.0) / Math.Sqrt (2.0 * Math.PI)))
      scalar "GAUSS" 1 1 "|S" (fun a -> res {
        let! z = a.Number 0
        let! p = normalCdf a z
        return! numOk (p - 0.5) })
      scalar "CONFIDENCE.NORM" 3 3 "|S" (fun a -> res {
        let! alpha = a.Number 0
        let! sd = a.Number 1
        let! size = a.Number 2
        let size = truncate size
        do! check (alpha > 0.0 && alpha < 1.0 && sd > 0.0 && size >= 1.0) Err.num
        let! z = normalInverse a (1.0 - alpha / 2.0)
        return! numOk (z * sd / Math.Sqrt size) })
      scalar "Z.TEST" 2 3 "RSS" (fun a -> res {
        let! xs = a.NumbersBetween 0 0
        let! x = a.Number 1
        do! check (xs.Length > 0) Err.na
        let! sigma =
            if a.Supplied 2 then a.Number 2
            else variance true xs |> Result.map Math.Sqrt
        do! check (sigma > 0.0) Err.div0
        let! p = normalCdf a ((Array.average xs - x) / (sigma / Math.Sqrt (float xs.Length)))
        return! numOk (1.0 - p) }) ]
    |> fun defs ->
        let byName name = defs |> List.find (fun d -> d.Name = name)
        defs @ [ alias "CONFIDENCE" (byName "CONFIDENCE.NORM"); alias "ZTEST" (byName "Z.TEST") ]

let install () = registerFunctions moduleName functions
