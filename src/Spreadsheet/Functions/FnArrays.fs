/// Dynamic-array functions of the function table that reshape an array.
module Rbt.Spreadsheet.Functions.FnArrays

open System
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Functions.Kit

let moduleName = "Arrays"

/// What fills the places an array does not cover: the pad argument, #N/A when omitted.
let private padArg (a: Args) (i: int) = if a.Supplied i then a.Value i else CellValue.Error Err.na

/// An array argument as rows (an error value standing for the whole array is the result).
let private arrayArg (a: Args) (i: int) : Result<CellValue[][], string> =
    let rows = a.Array i
    match rows with
    | [| [| CellValue.Error e |] |] -> Result.Error e
    | _ -> Ok rows

/// WRAPROWS / WRAPCOLS: a row or column cut into rows (columns) of `count` values.
let private wrap (name: string) (byRows: bool) =
    arrayed name 2 3 "" (fun a -> res {
        let! rows = arrayArg a 0
        let! count = a.Number 1
        let count = int (truncate count)
        let width = if rows.Length > 0 then rows.[0].Length else 0
        do! check (rows.Length = 1 || width = 1) Err.value
        do! check (count >= 1) Err.num
        let values = Array.concat rows
        let pad = padArg a 2
        let lines = (values.Length + count - 1) / count
        let at k = if k < values.Length then values.[k] else pad
        if byRows then return [| for r in 0 .. lines - 1 -> [| for c in 0 .. count - 1 -> at (r * count + c) |] |]
        else return [| for r in 0 .. count - 1 -> [| for c in 0 .. lines - 1 -> at (c * count + r) |] |] })

let functions : FunctionDef list =
    [ arrayed "EXPAND" 2 4 "" (fun a -> res {
        let! rows = arrayArg a 0
        let height = rows.Length
        let width = if height > 0 then rows.[0].Length else 0
        let! wantRows = if a.Supplied 1 then a.Number 1 |> Result.map (truncate >> int) else Ok height
        let! wantCols = if a.Supplied 2 then a.Number 2 |> Result.map (truncate >> int) else Ok width
        do! check (wantRows >= height && wantCols >= width) Err.value
        do! check (wantRows <= 1048576 && wantCols <= 16384) Err.num
        let pad = padArg a 3
        return [| for r in 0 .. wantRows - 1 ->
                    [| for c in 0 .. wantCols - 1 -> if r < height && c < width then rows.[r].[c] else pad |] |] })
      wrap "WRAPROWS" true
      wrap "WRAPCOLS" false ]

let install () = registerFunctions moduleName functions
