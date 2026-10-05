#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
module Ops = Rbt.Spreadsheet.SheetOps
module Formula = Rbt.Spreadsheet.Formula

// Find and replace: what matches (part of a cell or the whole cell, case,
// regular expressions, what cells show or what was entered), in which order,
// and what a replacement makes of a cell.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-44s expected %A  got %A" label expected got)
let throws label (run: unit -> unit) (part: string) =
    try
        run ()
        fail <- fail + 1
        printfn "FAIL %-44s did not throw" label
    with error ->
        if error.Message.Contains part then pass <- pass + 1
        else (fail <- fail + 1; printfn "FAIL %-44s threw %s" label error.Message)

Formula.formulaHooksInstalled |> ignore

let at row col : CellAddress = { Row = row; Col = col }
let cell (value: CellValue) = { Value = value; Format = defaultFormat; Comment = None }
let sheetOf cells = cells |> List.fold (fun m (addr, data) -> setCell addr data m) (createEmpty 60 10)
let a1 (addr: CellAddress) = formatCellRef addr.Row addr.Col
let options = Ops.defaultFindOptions
let find text opts range sheet = Ops.findInSheet (Ops.finderOf "Find" text opts) opts range sheet |> List.map (fun (addr, text) -> a1 addr, text)

let sheet =
    sheetOf [
        at 0 0, cell (CellValue.String "Apple pie")
        at 0 1, cell (CellValue.String "apple")
        at 1 0, cell (CellValue.Number 1234.5)
        at 1 1, { cell (CellValue.Number 1234.5) with Format = { defaultFormat with NumberFormat = Some "#,##0.00" } }
        at 2 0, cell (CellValue.Formula ("SUM(B2,1)", CellValue.Number 1235.5))
        at 2 2, cell (CellValue.String "pineapple")
        at 3 0, cell (CellValue.Boolean true)
        at 4 0, cell CellValue.Empty
    ]

// ---- what matches ---------------------------------------------------------------
eq "a part of a cell, whatever the case" (find "apple" options None sheet) [ "A1", "Apple pie"; "B1", "apple"; "C3", "pineapple" ]
eq "matching case" (find "apple" { options with MatchCase = true } None sheet) [ "B1", "apple"; "C3", "pineapple" ]
eq "the whole cell" (find "APPLE" { options with WholeCell = true } None sheet) [ "B1", "apple" ]
eq "the whole cell, matching case" (find "APPLE" { options with WholeCell = true; MatchCase = true } None sheet) []
eq "values are what the cells show" (find "1,234" options None sheet) [ "B2", "1,234.50" ]
eq "a number as it is shown" (find "1234.5" options None sheet |> List.map fst) [ "A2" ]
eq "a formula cell shows its value" (find "1235" options None sheet) [ "A3", "1235.5" ]
eq "its formula is not searched in values" (find "SUM" options None sheet) []
eq "formulas: the formula text" (find "sum(" { options with InFormulas = true } None sheet) [ "A3", "=SUM(B2,1)" ]
eq "formulas: other cells as entered" (find "1234.5" { options with InFormulas = true } None sheet |> List.map fst) [ "A2"; "B2" ]
eq "a boolean" (find "true" options None sheet |> List.map fst) [ "A4" ]
eq "row by row" (find "p" options None sheet |> List.map fst) [ "A1"; "B1"; "C3" ]
eq "within a range" (find "apple" options (Some { From = at 0 1; To = at 2 2 }) sheet |> List.map fst) [ "B1"; "C3" ]
eq "nothing found" (find "zebra" options None sheet) []
eq "an empty cell never matches" (find "." { options with Regex = true } None sheet |> List.length) 7

// ---- regular expressions ---------------------------------------------------------
let regex = { options with Regex = true }
eq "a regular expression" (find "^a.*e$" regex None sheet |> List.map fst) [ "A1"; "B1" ]
eq "a regular expression, matching case" (find "^a.*e$" { regex with MatchCase = true } None sheet |> List.map fst) [ "B1" ]
eq "a regular expression on the whole cell" (find "a..le" { regex with WholeCell = true } None sheet |> List.map fst) [ "B1" ]
eq "an alternation on the whole cell" (find "apple|pine" { regex with WholeCell = true } None sheet |> List.map fst) [ "B1" ]
eq "digits" (find "\\d{4}" regex None sheet |> List.map fst) [ "A2"; "A3" ]
throws "an unreadable regular expression" (fun () -> Ops.finderOf "Find" "(" regex |> ignore) "is not a regular expression"
throws "an empty text" (fun () -> Ops.finderOf "Find" "" options |> ignore) "no text to search"
eq "plain text is not a pattern" (find "(" options None (sheetOf [ at 0 0, cell (CellValue.String "f(x)") ]) |> List.map fst) [ "A1" ]

// ---- replacing --------------------------------------------------------------------
let replace text replacement opts (data: CellData) = Ops.replaceInCell (Ops.finderOf "ReplaceMatch" text opts) opts replacement 2024 data
let valueOf (data: CellData option) = data |> Option.map (fun d -> d.Value)
eq "every occurrence in the cell" (replace "p" "b" options (cell (CellValue.String "Apple pie")) |> valueOf) (Some (CellValue.String "Abble bie"))
eq "whatever the case" (replace "APPLE" "Cherry" options (cell (CellValue.String "Apple pie")) |> valueOf) (Some (CellValue.String "Cherry pie"))
eq "matching case leaves the others" (replace "a" "o" { options with MatchCase = true } (cell (CellValue.String "A banana")) |> valueOf) (Some (CellValue.String "A bonono"))
eq "the whole cell is replaced" (replace "apple" "pear" { options with WholeCell = true } (cell (CellValue.String "Apple")) |> valueOf) (Some (CellValue.String "pear"))
eq "the whole cell: a part does not match" (replace "apple" "pear" { options with WholeCell = true } (cell (CellValue.String "Apple pie")) |> valueOf) None
eq "no match, no change" (replace "x" "y" options (cell (CellValue.String "Apple")) |> valueOf) None
eq "the result is read as input (a number)" (replace "n" "" options (cell (CellValue.String "n42")) |> valueOf) (Some (CellValue.Number 42.0))
eq "a number changes as it is shown" (replace "5" "7" options (cell (CellValue.Number 1525.0)) |> valueOf) (Some (CellValue.Number 1727.0))
eq "replaced away: the cell is empty" (replace "apple" "" { options with WholeCell = true } (cell (CellValue.String "apple")) |> valueOf) (Some CellValue.Empty)
eq "a dollar sign is a dollar sign" (replace "USD" "$1" options (cell (CellValue.String "USD x")) |> valueOf) (Some (CellValue.String "$1 x"))
let bold = { cell (CellValue.String "apple") with Format = { defaultFormat with Bold = true }; Comment = Some { Text = "note"; Author = None } }
eq "the cell keeps its format and note" (replace "apple" "pear" options bold |> Option.map (fun d -> d.Format.Bold, d.Comment.IsSome)) (Some (true, true))
let formula = cell (CellValue.Formula ("SUM(B2,1)", CellValue.Number 1235.5))
eq "values: a formula cell is left alone" (replace "1235" "9" options formula |> valueOf) None
let inFormulas = { options with InFormulas = true }
eq "formulas: the formula text changes"
    (replace "B2" "C7" inFormulas formula |> Option.map (fun d -> match d.Value with CellValue.Formula (text, _) -> formulaText text | other -> sprintf "%A" other))
    (Some "=SUM(C7,1)")
eq "formulas: a formula can become a value" (replace "=SUM(B2,1)" "12" { inFormulas with WholeCell = true } formula |> valueOf) (Some (CellValue.Number 12.0))
eq "a regular expression with groups" (replace "(\\w+) (\\w+)" "$2 $1" regex (cell (CellValue.String "Apple pie")) |> valueOf) (Some (CellValue.String "pie Apple"))
eq "a regular expression on the whole cell" (replace "a(.*)e" "<$1>" { regex with WholeCell = true } (cell (CellValue.String "apple")) |> valueOf) (Some (CellValue.String "<ppl>"))

// Replace all: the cells that change, and never the cells of a spilled array.
let spilled = { sheetOf [ at 0 0, cell (CellValue.String "cat"); at 1 0, cell (CellValue.String "cat"); at 2 0, cell (CellValue.String "dog") ] with Spill = Map.ofList [ at 1 0, at 0 5 ] }
let all = Ops.replaceInSheet (Ops.finderOf "ReplaceAll" "cat" options) options "cow" 2024 None spilled
eq "replace all changes the matching cells" (all |> List.map (fun (addr, before, after) -> a1 addr, before.Value, after.Value)) [ "A1", CellValue.String "cat", CellValue.String "cow" ]
eq "replace all in a range" (Ops.replaceInSheet (Ops.finderOf "ReplaceAll" "apple" options) options "kiwi" 2024 (Some { From = at 0 1; To = at 0 1 }) sheet |> List.map (fun (addr, _, after) -> a1 addr, after.Value)) [ "B1", CellValue.String "kiwi" ]

// ---- stepping ----------------------------------------------------------------------
let places = [| (0, 0, 0); (0, 0, 1); (0, 2, 2); (2, 1, 0) |]
eq "next: the first match after the place" (Ops.stepMatch true (0, 0, 0) places) (Some 1)
eq "next: from a cell between matches" (Ops.stepMatch true (0, 1, 5) places) (Some 2)
eq "next: on to a following sheet" (Ops.stepMatch true (0, 2, 2) places) (Some 3)
eq "next: from a sheet without matches" (Ops.stepMatch true (1, 9, 9) places) (Some 3)
eq "next: starts over after the last" (Ops.stepMatch true (2, 1, 0) places) (Some 0)
eq "previous: the last match before the place" (Ops.stepMatch false (0, 2, 2) places) (Some 1)
eq "previous: back to a sheet before" (Ops.stepMatch false (2, 1, 0) places) (Some 2)
eq "previous: the last one before the first" (Ops.stepMatch false (0, 0, 0) places) (Some 3)
eq "one match is its own next" (Ops.stepMatch true (0, 0, 0) [| (0, 0, 0) |]) (Some 0)
eq "no match, no step" (Ops.stepMatch true (0, 0, 0) [||]) None

printfn "Find and replace tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
