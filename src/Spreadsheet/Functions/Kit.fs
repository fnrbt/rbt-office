/// What the function modules of this folder share: reading the arguments of a
/// call with Excel's coercions, and making a `FunctionDef`. Only the modules
/// import this file (the grid does not), so it is loaded with the first of them.
module Rbt.Spreadsheet.Functions.Kit

open System
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula

let err (code: string) : CellValue = CellValue.Error code

/// A number result: NaN and the infinities are #NUM!, as Excel has no such numbers.
let num (x: float) : CellValue =
    if Double.IsNaN x || Double.IsInfinity x then err Err.num else CellValue.Number x

let numOk (x: float) : Result<CellValue, string> =
    if Double.IsNaN x || Double.IsInfinity x then Result.Error Err.num else Ok (CellValue.Number x)

type ResultBuilder() =
    member _.Bind(r: Result<'a, string>, f: 'a -> Result<'b, string>) = Result.bind f r
    member _.Return(x: 'a) : Result<'a, string> = Ok x
    member _.ReturnFrom(r: Result<'a, string>) = r
    member _.Zero() : Result<unit, string> = Ok ()

/// `res { let! x = ... }`: the first error is the result.
let res = ResultBuilder()

let fail (code: string) : Result<'a, string> = Result.Error code

/// Fail with `code` unless `condition`.
let check (condition: bool) (code: string) : Result<unit, string> = if condition then Ok () else Result.Error code

/// A value as a number for an argument that is one: an error is that error.
let numberOf (value: CellValue) : Result<float, string> = coerceNumber value

let truncate (x: float) = Math.Truncate x

let private isMissing (e: Expr) = match e with MissingArg -> true | _ -> false

/// The arguments of one call.
type Args(ev: FunctionEvaluator, args: Expr list, ctx: EvalContext) =
    let items = List.toArray args
    member _.Count = items.Length
    member _.Ev = ev
    member _.Ctx = ctx
    member _.Exprs = items
    /// Was argument `i` given (and not left empty, `F(1,)`)?
    member _.Supplied (i: int) = i < items.Length && not (isMissing items.[i])
    member _.Expr (i: int) = items.[i]
    /// The value of argument `i` (an omitted one is Empty).
    member _.Value (i: int) : CellValue =
        if i >= items.Length || isMissing items.[i] then CellValue.Empty
        else unwrapValue (ev.Eval items.[i] ctx)
    member this.Number (i: int) : Result<float, string> = coerceNumber (this.Value i)
    /// Argument `i` as a number, `fallback` when it is not given.
    member this.NumberOr (i: int) (fallback: float) : Result<float, string> =
        if this.Supplied i then this.Number i else Ok fallback
    member this.Text (i: int) : Result<string, string> = coerceText (this.Value i)
    member this.Bool (i: int) : Result<bool, string> = coerceBool (this.Value i)
    member this.BoolOr (i: int) (fallback: bool) : Result<bool, string> =
        if this.Supplied i then this.Bool i else Ok fallback
    /// Argument `i` as rows of values (references' blank cells Empty).
    member _.Array (i: int) : CellValue[][] =
        if i >= items.Length || isMissing items.[i] then [| [| CellValue.Empty |] |]
        else ev.EvalArray items.[i] ctx |> Array.map (Array.map unwrapValue)
    /// Is argument `i` a reference or an array (rather than one value written in)?
    member _.IsRangeLike (i: int) =
        i < items.Length
        && (match items.[i] with
            | CellRef _ | RangeRef _ | TableRef _ | SpillRef _ | ArrayLit _ -> true
            | MissingArg | Number _ | Str _ | Bool _ | ErrorLit _ -> false
            | e -> ev.IsArray e || (ev.Reference e ctx).IsSome)
    /// The reference argument `i` stands for, when it is one.
    member _.Reference (i: int) : Expr option =
        if i >= items.Length then None else ev.Reference items.[i] ctx

    /// The numbers of the arguments from `first` on, as STDEV or KURT take
    /// them: of a reference or an array its numbers (text, logicals and blank
    /// cells are skipped); a value written in is coerced (TRUE 1, "3" 3, other
    /// text #VALUE!). An error anywhere is the result.
    member this.Numbers (first: int) : Result<float[], string> = this.NumbersBetween first (items.Length - 1)

    /// `Numbers` of the arguments `first` to `last`.
    member this.NumbersBetween (first: int) (last: int) : Result<float[], string> =
        let found = ResizeArray<float>()
        let mutable failure = None
        for i in first .. min last (items.Length - 1) do
            if failure.IsNone && not (isMissing items.[i]) then
                if this.IsRangeLike i then
                    for row in this.Array i do
                        for v in row do
                            match v with
                            | CellValue.Number x -> if failure.IsNone then found.Add x
                            | CellValue.Error e -> if failure.IsNone then failure <- Some e
                            | _ -> ()
                else
                    match coerceNumber (this.Value i) with
                    | Ok x -> found.Add x
                    | Result.Error e -> failure <- Some e
        match failure with
        | Some e -> Result.Error e
        | None -> Ok (found.ToArray())

    /// The numbers as the A variants (STDEVA, VARA) take them: in a reference
    /// or array text is 0, TRUE 1, FALSE 0 and blank cells are skipped; a value
    /// written in is coerced (text that is no number is #VALUE!).
    member this.NumbersA (first: int) : Result<float[], string> =
        let found = ResizeArray<float>()
        let mutable failure = None
        for i in first .. items.Length - 1 do
            if failure.IsNone && not (isMissing items.[i]) then
                if this.IsRangeLike i then
                    for row in this.Array i do
                        for v in row do
                            if failure.IsNone then
                                match v with
                                | CellValue.Number x -> found.Add x
                                | CellValue.Boolean b -> found.Add (if b then 1.0 else 0.0)
                                | CellValue.String _ -> found.Add 0.0
                                | CellValue.Error e -> failure <- Some e
                                | _ -> ()
                else
                    match coerceNumber (this.Value i) with
                    | Ok x -> found.Add x
                    | Result.Error e -> failure <- Some e
        match failure with
        | Some e -> Result.Error e
        | None -> Ok (found.ToArray())

    /// Every value of argument `i`, row by row.
    member this.Flat (i: int) : CellValue[] = this.Array i |> Array.concat

let private resultValue (r: Result<CellValue, string>) =
    match r with
    | Ok v -> v
    | Result.Error e -> err e

/// A function of one value.
let scalar (name: string) (minArgs: int) (maxArgs: int) (signature: string) (f: Args -> Result<CellValue, string>) : FunctionDef =
    { Name = name; MinArgs = minArgs; MaxArgs = maxArgs; Signature = signature; Volatile = false
      Impl = ScalarFunction (fun ev args ctx -> resultValue (f (Args(ev, args, ctx)))) }

/// A function that gives an array.
let arrayed (name: string) (minArgs: int) (maxArgs: int) (signature: string) (f: Args -> Result<CellValue[][], string>) : FunctionDef =
    { Name = name; MinArgs = minArgs; MaxArgs = maxArgs; Signature = signature; Volatile = false
      Impl = ArrayFunction (fun ev args ctx ->
        match f (Args(ev, args, ctx)) with
        | Ok rows when rows.Length > 0 && rows.[0].Length > 0 -> rows
        | Ok _ -> [| [| err Err.calc |] |]
        | Result.Error e -> [| [| err e |] |]) }

/// The same function under another name (a compatibility name: CONFIDENCE for CONFIDENCE.NORM).
let alias (name: string) (def: FunctionDef) : FunctionDef = { def with Name = name }

let volatileFn (def: FunctionDef) : FunctionDef = { def with Volatile = true }

/// A function of numbers given as scalars (lifted over arrays): `count` of them.
let numeric1 (name: string) (f: float -> Result<CellValue, string>) =
    scalar name 1 1 "|S" (fun a -> res { let! x = a.Number 0 in return! f x })

let numeric2 (name: string) (f: float -> float -> Result<CellValue, string>) =
    scalar name 2 2 "|S" (fun a -> res {
        let! x = a.Number 0
        let! y = a.Number 1
        return! f x y })

/// Two arrays of the same number of values, paired (SUMXMY2 and its kin):
/// pairs where either is not a number are skipped; an error is the result;
/// arrays of different sizes are #N/A.
let pairedNumbers (a: Args) (i: int) (j: int) : Result<(float * float)[], string> =
    let xs = a.Flat i
    let ys = a.Flat j
    if xs.Length <> ys.Length then Result.Error Err.na
    else
        let pairs = ResizeArray<float * float>()
        let mutable failure = None
        for k in 0 .. xs.Length - 1 do
            match xs.[k], ys.[k] with
            | CellValue.Error e, _ | _, CellValue.Error e -> if failure.IsNone then failure <- Some e
            | CellValue.Number x, CellValue.Number y -> pairs.Add((x, y))
            | _ -> ()
        match failure with
        | Some e -> Result.Error e
        | None -> Ok (pairs.ToArray())

/// The value at row r, column c of an array, Excel's broadcasting of a 1-wide or 1-high one.
let cellOf (arr: CellValue[][]) (r: int) (c: int) =
    let rows = arr.Length
    let cols = if rows > 0 then arr.[0].Length else 0
    if rows = 0 || cols = 0 then CellValue.Error Err.value
    else arr.[(if rows = 1 then 0 else r)].[(if cols = 1 then 0 else c)]

/// Sample mean and the sums of the powers of the deviations.
let mean (xs: float[]) = Array.average xs
