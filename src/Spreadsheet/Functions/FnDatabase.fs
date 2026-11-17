/// Database functions of the function table (the rest - DSUM, DGET ... - are Formula.fs's).
module Rbt.Spreadsheet.Functions.FnDatabase

open System
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Functions.Kit

let moduleName = "Database"

/// The values of the field of the records that meet the criteria, as DSUM
/// reads them: the first row of the database and of the criteria are column
/// labels (matched without case); each criteria row is an AND of its non-blank
/// cells, and the rows are OR-ed. The field is a label or a 1-based column.
let databaseValues (a: Args) : Result<CellValue[], string> =
    let database = a.Array 0
    let criteria = a.Array 2
    if database.Length < 1 || criteria.Length < 2 then Result.Error Err.value
    else
        let header (v: CellValue) = match coerceText v with Ok t -> t.Trim() | _ -> ""
        let columns = database.[0] |> Array.map header
        let columnIndex (name: string) =
            columns |> Array.tryFindIndex (fun c -> String.Equals(c, name.Trim(), StringComparison.OrdinalIgnoreCase))
        let field =
            match a.Value 1 with
            | CellValue.Number index when index >= 1.0 && int index <= columns.Length -> Ok (int index - 1)
            | CellValue.String name -> (match columnIndex name with Some i -> Ok i | None -> Result.Error Err.value)
            | CellValue.Error e -> Result.Error e
            | _ -> Result.Error Err.value
        match field with
        | Result.Error e -> Result.Error e
        | Ok fieldIndex ->
            let criteriaColumns = criteria.[0] |> Array.map (fun v -> columnIndex (header v))
            let tests =
                [| for r in 1 .. criteria.Length - 1 ->
                    [| for c in 0 .. criteria.[r].Length - 1 do
                        match criteria.[r].[c], criteriaColumns.[c] with
                        | CellValue.Empty, _ -> ()
                        | CellValue.String "", _ -> ()
                        | condition, Some column -> yield column, a.Ev.Criteria condition
                        | _, None -> yield -1, (fun _ -> false) |] |]
            let matches (record: CellValue[]) =
                tests |> Array.exists (fun row ->
                    row |> Array.forall (fun (column, test) -> column >= 0 && column < record.Length && test record.[column]))
            Ok [| for r in 1 .. database.Length - 1 do
                    if matches database.[r] && fieldIndex < database.[r].Length then yield database.[r].[fieldIndex] |]

/// The numbers among the values (an error among them is the result).
let databaseNumbers (a: Args) : Result<float[], string> =
    match databaseValues a with
    | Result.Error e -> Result.Error e
    | Ok values ->
        match values |> Array.tryPick (fun v -> match v with CellValue.Error e -> Some e | _ -> None) with
        | Some e -> Result.Error e
        | None -> Ok (values |> Array.choose (fun v -> match v with CellValue.Number x -> Some x | _ -> None))

/// The variance of numbers, of a sample or of the population.
let variance (sample: bool) (xs: float[]) : Result<float, string> =
    let n = xs.Length
    if n = 0 || (sample && n < 2) then Result.Error Err.div0
    else
        let m = Array.average xs
        let squares = xs |> Array.sumBy (fun x -> (x - m) * (x - m))
        Ok (squares / float (if sample then n - 1 else n))

let private databaseStatistic (name: string) (sample: bool) (root: bool) =
    scalar name 3 3 "RSR" (fun a -> res {
        let! numbers = databaseNumbers a
        let! v = variance sample numbers
        return! numOk (if root then Math.Sqrt v else v) })

let functions : FunctionDef list =
    [ databaseStatistic "DSTDEV" true true
      databaseStatistic "DSTDEVP" false true
      databaseStatistic "DVAR" true false
      databaseStatistic "DVARP" false false ]

let install () = registerFunctions moduleName functions
