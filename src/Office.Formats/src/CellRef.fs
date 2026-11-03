/// References to the cells of a sheet as the file format writes them: a cell ("B5", "$B$5"), a range
/// ("B5:F10"), a list of ranges ("A1:B2 D4"), each read into rows and columns (from 0) and written back.
module CellRef

/// "AB12" (or "$AB$12") as (row, column) from 0; None for what is not a cell reference.
let parse (reference: string) : (int * int) option =
    let mutable col = 0
    let mutable row = 0
    let mutable sound = reference.Length > 0
    let mutable inDigits = false
    for ch in reference do
        if ch >= 'A' && ch <= 'Z' && not inDigits then col <- col * 26 + (int ch - int 'A' + 1)
        elif ch >= '0' && ch <= '9' then
            inDigits <- true
            row <- row * 10 + (int ch - int '0')
        elif ch <> '$' then sound <- false
    if sound && row >= 1 && col >= 1 then Some (row - 1, col - 1) else None

/// The cells that name the corners of a range, as written ("B5:F10" -> [ "B5"; "F10" ]; "B5" -> [ "B5" ]).
let corners (range: string) : string list = range.Split(':') |> Array.toList

/// The first and last corner of a range ("B5:F10", or one cell), as (row, column) from 0.
let parseRange (range: string) : ((int * int) * (int * int)) option =
    match corners range with
    | [] -> None
    | first :: _ as all ->
        match parse first, parse (List.last all) with
        | Some a, Some b -> Some (a, b)
        | _ -> None

/// The first cell of a range reference ("B5:F10" -> "B5"), without "$", in capitals.
let firstCell (range: string) : string =
    (List.head (corners range)).Replace("$", "").ToUpperInvariant()

/// The ranges of a list of ranges (`sqref`: "A1:B2 D4" -> [ "A1:B2"; "D4" ]).
let ranges (sqref: string) : string list = sqref.Split(' ') |> Array.toList |> List.filter (fun part -> part <> "")

/// A list of ranges.
let ofRanges (ranges: string list) : string = String.concat " " ranges

/// A range of corners.
let ofCorners (corners: string list) : string = String.concat ":" corners

/// A corner made absolute ("B5" -> "$B$5", "B" -> "$B", "5" -> "$5").
let absolute (corner: string) : string =
    let letters = corner |> Seq.takeWhile (fun c -> c < '0' || c > '9') |> Seq.filter (fun c -> c <> '$') |> Seq.toArray |> System.String
    let digits = corner |> Seq.filter (fun c -> c >= '0' && c <= '9') |> Seq.toArray |> System.String
    (if letters = "" then "" else "$" + letters) + (if digits = "" then "" else "$" + digits)

/// The areas of a reference of several (a print area: "A1:B2,D1:E5" -> [ "A1:B2"; "D1:E5" ]).
let areas (reference: string) : string list = reference.Split(',') |> Array.toList |> List.map (fun area -> area.Trim())
