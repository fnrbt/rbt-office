#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model

// Cell borders: every Excel border style has a complete CSS declaration
// (width, line style, colour), and ApplyBorder's kinds ("all", "outside",
// "inside", "none", one side) change the right sides of the right cells.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-34s expected %A  got %A" label expected got)

let css style = borderCss { Style = style; Color = Some "#112233" }
eq "thin" (css "thin") "1px solid #112233"
eq "medium" (css "medium") "2px solid #112233"
eq "thick" (css "thick") "3px solid #112233"
eq "dashed" (css "dashed") "1px dashed #112233"
eq "dotted" (css "dotted") "1px dotted #112233"
eq "double" (css "double") "3px double #112233"
eq "hair" (css "hair") "1px dotted #112233"
eq "mediumDashed" (css "mediumDashed") "2px dashed #112233"
eq "dashDot" (css "dashDot") "1px dashed #112233"
eq "mediumDashDot" (css "mediumDashDot") "2px dashed #112233"
eq "dashDotDot" (css "dashDotDot") "1px dashed #112233"
eq "mediumDashDotDot" (css "mediumDashDotDot") "2px dashed #112233"
eq "slantDashDot" (css "slantDashDot") "2px dashed #112233"
eq "no colour is black" (borderCss { Style = "thin"; Color = None }) "1px solid #000000"
eq "none" (css "none") "none"
eq "every style has a line style"
    ([ "thin"; "medium"; "thick"; "dashed"; "dotted"; "double"; "hair"; "mediumDashed"; "dashDot"; "mediumDashDot"; "dashDotDot"; "mediumDashDotDot"; "slantDashDot" ]
     |> List.forall (fun style -> match borderCssParts style with Some (width, line) -> width >= 1 && line <> "" | None -> false)) true

// Sides of the cells of the block B2:D4 (rows 1-3, columns 1-3).
let block = 1, 1, 3, 3
let sides (format: CellFormat) =
    [ if format.BorderTop.IsSome then "T"
      if format.BorderBottom.IsSome then "B"
      if format.BorderLeft.IsSome then "L"
      if format.BorderRight.IsSome then "R" ] |> String.concat ""
let at side style row col = borderFormatInBlock block side style "#000000" row col defaultFormat |> sides
eq "all" (at "all" "thin" 2 2) "TBLR"
eq "outside corner" (at "outside" "thin" 1 1) "TL"
eq "outside top edge" (at "outside" "thin" 1 2) "T"
eq "outside middle" (at "outside" "thin" 2 2) ""
eq "outside far corner" (at "outside" "thin" 3 3) "BR"
eq "inside corner" (at "inside" "thin" 1 1) "BR"
eq "inside middle" (at "inside" "thin" 2 2) "TBLR"
eq "inside far corner" (at "inside" "thin" 3 3) "TL"
eq "one side" (at "left" "medium" 2 2) "L"
let boxed = borderFormatInBlock block "all" "thick" "#ff0000" 2 2 defaultFormat
eq "style and colour stored" boxed.BorderTop (Some { Style = "thick"; Color = Some "#ff0000" })
eq "none clears" (borderFormatInBlock block "none" "thin" "#000000" 2 2 boxed |> sides) ""
eq "a style of none clears that side" (borderFormatInBlock block "top" "none" "#000000" 2 2 boxed |> sides) "BLR"
eq "nothing to do keeps the format" (obj.ReferenceEquals(borderFormatInBlock block "outside" "thin" "#000000" 2 2 boxed, boxed)) true
eq "one cell: outside is its box" (borderFormatInBlock (4, 4, 4, 4) "outside" "thin" "#000000" 4 4 defaultFormat |> sides) "TBLR"
eq "one cell: inside is nothing" (borderFormatInBlock (4, 4, 4, 4) "inside" "thin" "#000000" 4 4 defaultFormat |> sides) ""

printfn "stage29-borders: %d passed, %d failed" pass fail
if fail > 0 then exit 1
