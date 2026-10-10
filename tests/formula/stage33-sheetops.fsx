#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
module Ops = Rbt.Spreadsheet.SheetOps
module Field = Rbt.Spreadsheet.SheetOps.FormatField

// What the component's host API stands on: A1 ranges as hosts write them
// (cells, blocks, whole rows and columns, `$`, a sheet in front), the format
// a cell shows with, and what the cells of a selection have in common.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-40s expected %A  got %A" label expected got)

let at row col : CellAddress = { Row = row; Col = col }
let block r0 c0 r1 c1 : CellRange = { From = at r0 c0; To = at r1 c1 }
let cell value format = { Value = value; Format = format; Comment = None }

// ---- A1 ranges ------------------------------------------------------------------
eq "a cell" (Ops.tryParseRange "B2") (Some (block 1 1 1 1))
eq "a block" (Ops.tryParseRange "B2:D5") (Some (block 1 1 4 3))
eq "corners in any order" (Ops.tryParseRange "D5:B2") (Some (block 1 1 4 3))
eq "mixed corners" (Ops.tryParseRange "D2:B5") (Some (block 1 1 4 3))
eq "dollars and lower case" (Ops.tryParseRange " $b$2:d$5 ") (Some (block 1 1 4 3))
eq "whole columns" (Ops.tryParseRange "A:C") (Some (block 0 0 Ops.maxSheetRow 2))
eq "one whole column" (Ops.tryParseRange "$C:$C") (Some (block 0 2 Ops.maxSheetRow 2))
eq "whole rows" (Ops.tryParseRange "3:5") (Some (block 2 0 4 Ops.maxSheetCol))
eq "the last cell" (Ops.tryParseRange "XFD1048576") (Some (block Ops.maxSheetRow Ops.maxSheetCol Ops.maxSheetRow Ops.maxSheetCol))
for bad in [ ""; "A"; "5"; "A:5"; "A1:B"; "1A"; "A0"; "XFE1"; "A1048577"; "A1:B2:C3"; "A-1"; "$$A1"; "A$"; "Sheet1!A1"; "A 1"; "A1:"; ":A1"; "R1C1:X" ] do
    eq (sprintf "not a range: %s" bad) (Ops.tryParseRange bad) None
eq "null is not a range" (Ops.tryParseRange null) None
eq "a cell reference" (Ops.tryParseCell "$C$7") (Some (at 6 2))
eq "a block is not a cell" (Ops.tryParseCell "C7:C8") None
eq "parseRange names the member" (try Ops.parseRange "GetRange" "nope!" |> ignore; "" with e -> e.Message.Substring(0, 26)) "GetRange: \"nope!\" is not a"

for text in [ "B2"; "B2:D5"; "A:C"; "C:C"; "3:5"; "7:7"; "AA10:AB11" ] do
    eq (sprintf "text of %s" text) (Ops.tryParseRange text |> Option.map Ops.rangeText) (Some text)
eq "whole sheet" (Ops.rangeText (block 0 0 Ops.maxSheetRow Ops.maxSheetCol)) "1:1048576"
eq "whole columns are told apart" ((Ops.isWholeColumns (block 0 1 Ops.maxSheetRow 1)), (Ops.isWholeRows (block 0 1 Ops.maxSheetRow 1))) (true, false)

eq "no sheet in front" (Ops.splitSheetPrefix "A1:B2") (None, "A1:B2")
eq "a sheet in front" (Ops.splitSheetPrefix "Data!A1") (Some "Data", "A1")
eq "a quoted sheet" (Ops.splitSheetPrefix "'My sheet'!A:A") (Some "My sheet", "A:A")
eq "a quote in the name" (Ops.splitSheetPrefix "'Bob''s!'!B2") (Some "Bob's!", "B2")
eq "names that need quotes" ([ "Data"; "My sheet"; "Bob's"; "2024"; "A1"; "Q1.Plan"; "a-b" ] |> List.map Ops.quoteSheetName) [ "Data"; "'My sheet'"; "'Bob''s'"; "'2024'"; "'A1'"; "Q1.Plan"; "'a-b'" ]
eq "quoting round trips" (Ops.splitSheetPrefix (Ops.quoteSheetName "It's 2024!" + "!C3")) (Some "It's 2024!", "C3")

eq "clipped to the sheet" (Ops.clipRange 100 10 (block 0 2 Ops.maxSheetRow 2)) (Some (block 0 2 99 2))
eq "outside the sheet" (Ops.clipRange 100 10 (block 100 0 120 3)) None

// ---- used range -------------------------------------------------------------------
let bold = { defaultFormat with Bold = true }
let sheet =
    createEmpty 1000 26
    |> setCell (at 1 1) (cell (CellValue.Number 1.0) defaultFormat)
    |> setCell (at 4 3) (cell (CellValue.String "x") bold)
eq "an empty sheet has no used range" (Ops.usedRange (createEmpty 10 10)) None
eq "the used range starts at A1" (Ops.usedRange sheet) (Some (block 0 0 4 3))
eq "a merged cell is used" (Ops.usedRange (addMergedCell 6 0 7 5 sheet)) (Some (block 0 0 7 5))
eq "a whole column cut to the used rows" (Ops.clipToUsed sheet (block 0 1 Ops.maxSheetRow 1)) (block 0 1 4 1)
eq "whole rows cut to the used columns" (Ops.clipToUsed sheet (block 1 0 2 Ops.maxSheetCol)) (block 1 0 2 3)
eq "a range beyond the used cells keeps its first cell" (Ops.clipToUsed sheet (block 0 9 Ops.maxSheetRow 9)) (block 0 9 4 9)

// ---- the format a cell shows with ----------------------------------------------------
let italic = { defaultFormat with Italic = true }
let styled = { sheet with ColumnStyles = [ { First = 3; Last = 5; Format = italic; Source = -1 } ] }
eq "a cell's own format" (Ops.effectiveFormat styled (at 4 3)) bold
eq "an empty cell of a styled column" (Ops.effectiveFormat styled (at 9 4)) italic
eq "an empty cell elsewhere" (Ops.effectiveFormat styled (at 9 0)) defaultFormat

// ---- what a selection has in common ---------------------------------------------------
eq "the same format: nothing differs" (Ops.formatDifference bold { defaultFormat with Bold = true }) 0
eq "General is no number format" (Ops.formatDifference { defaultFormat with NumberFormat = Some "General" } defaultFormat) 0
eq "center is middle" (Ops.formatDifference { defaultFormat with VerticalAlign = Some "center" } { defaultFormat with VerticalAlign = Some "middle" }) 0
eq "two fields differ"
    (Ops.formatDifference { defaultFormat with Bold = true; BgColor = Some "#FF0000" } defaultFormat) (Field.Bold ||| Field.Fill)
let thin = Some { Style = "thin"; Color = Some "#000000" }
eq "a border side" (Ops.formatDifference { defaultFormat with BorderTop = thin } defaultFormat) Field.BorderTop

let grid =
    [ for row in 0 .. 9 do
        for col in 0 .. 3 ->
            at row col, cell (CellValue.Number (float row)) { defaultFormat with Bold = (col = 0); FontSize = Some (if row < 5 then 11 else 14); NumberFormat = Some "0.00" } ]
    |> List.fold (fun m (addr, data) -> setCell addr data m) (createEmpty 200000 30)
let summary = Ops.summarizeFormats grid (0, 0, 9, 3)
eq "bold and size are mixed, the number format is not" summary.Mixed (Field.Bold ||| Field.FontSize)
eq "the summary starts from the first cell" (summary.Format.Bold, summary.Format.NumberFormat) (true, Some "0.00")
eq "one column, the top rows: nothing mixed" (Ops.summarizeFormats grid (0, 1, 4, 3)).Mixed 0
eq "one cell" (Ops.summarizeFormats grid (7, 0, 7, 0)) { Format = (Map.find (at 7 0) grid.Cells).Format; Mixed = 0 }
// Empty cells count with the default format.
eq "a block that runs into empty cells" (Ops.summarizeFormats grid (0, 1, 4, 5)).Mixed (Field.FontSize ||| Field.NumberFormat)
// A selection too large to walk cell by cell gives what the walk would.
let whole = Ops.summarizeFormats grid (0, 0, 199999, 29)
eq "the whole sheet" whole.Mixed (Field.Bold ||| Field.FontSize ||| Field.NumberFormat)
let filled = Ops.summarizeFormats grid (0, 1, 199999, 3)
eq "whole columns with empty cells below" filled.Mixed (Field.FontSize ||| Field.NumberFormat)
// A column style stands for the cells of its columns that do not exist.
let withStyle = { createEmpty 200000 30 with ColumnStyles = [ { First = 0; Last = 29; Format = italic; Source = -1 } ] }
eq "every column styled: the style is the common format" (Ops.summarizeFormats withStyle (0, 0, 199999, 29)) { Format = italic; Mixed = 0 }
let partly = { createEmpty 200000 30 with ColumnStyles = [ { First = 0; Last = 3; Format = italic; Source = -1 } ] }
eq "some columns styled: italic is mixed" (Ops.summarizeFormats partly (0, 0, 199999, 29)).Mixed Field.Italic
eq "small blocks see column styles too" (Ops.summarizeFormats partly (0, 2, 3, 6)).Mixed Field.Italic

// ---- merges and hyperlinks ---------------------------------------------------------------
let merged = addMergedCell 2 1 3 4 sheet
eq "a cell of a merged cell" (Ops.mergeRangeAt merged (at 3 3) |> Option.map Ops.rangeText) (Some "B3:E4")
eq "a cell outside it" (Ops.mergeRangeAt merged (at 4 3)) None
let linked = { sheet with Hyperlinks = [ { Range = block 0 0 2 2; Target = "https://example.com/"; IsInternal = false; Tooltip = None; Display = None }
                                         { Range = block 1 1 1 1; Target = "Sheet2!A1"; IsInternal = true; Tooltip = None; Display = None } ] }
eq "the last link on a cell wins" (Ops.hyperlinkAt linked (at 1 1) |> Option.map (fun link -> link.Target)) (Some "Sheet2!A1")
eq "a link on a range" (Ops.hyperlinkAt linked (at 2 0) |> Option.map (fun link -> link.Target)) (Some "https://example.com/")
eq "no link" (Ops.hyperlinkAt linked (at 5 5)) None

// ---- number format codes: decimals, validity ----------------------------------------
let num x = CellValue.Number x
eq "more decimals: General starts from the value" (Ops.changeDecimals 1 None (num 12.5)) (Some "0.00")
eq "fewer decimals: General" (Ops.changeDecimals -1 None (num 12.5)) (Some "0")
eq "a whole number in General gains one" (Ops.changeDecimals 1 None (num 7.0)) (Some "0.0")
eq "a whole number in General has none to lose" (Ops.changeDecimals -1 None (num 7.0)) None
eq "the result of a formula counts" (Ops.changeDecimals 1 (Some "General") (CellValue.Formula ("=1/4", num 0.25))) (Some "0.000")
eq "text has no decimals" (Ops.changeDecimals 1 None (CellValue.String "abc")) None
eq "more decimals of a code" (Ops.changeDecimals 1 (Some "0.00") (num 1.0)) (Some "0.000")
eq "fewer decimals of a code" (Ops.changeDecimals -1 (Some "0.00") (num 1.0)) (Some "0.0")
eq "the last decimal takes the point with it" (Ops.changeDecimals -1 (Some "#,##0.0") (num 1.0)) (Some "#,##0")
eq "none to lose" (Ops.changeDecimals -1 (Some "#,##0") (num 1.0)) None
eq "a first decimal" (Ops.changeDecimals 1 (Some "#,##0") (num 1.0)) (Some "#,##0.0")
eq "a percent" (Ops.changeDecimals 1 (Some "0%") (num 0.5)) (Some "0.0%")
eq "every section" (Ops.changeDecimals 1 (Some "#,##0.00;[Red]-#,##0.00;\"-\";@") (num 1.0)) (Some "#,##0.000;[Red]-#,##0.000;\"-\";@")
eq "text in quotes and escapes is not the number" (Ops.changeDecimals 1 (Some "\"No. 0.0\" 0 \\0") (num 1.0)) (Some "\"No. 0.0\" 0.0 \\0")
eq "an exponent keeps its digits" (Ops.changeDecimals 1 (Some "0.00E+00") (num 1.0)) (Some "0.000E+00")
eq "a currency with a locale tag" (Ops.changeDecimals -1 (Some "[$€-407] #,##0.00") (num 1.0)) (Some "[$€-407] #,##0.0")
eq "a date format is left alone" (Ops.changeDecimals 1 (Some "yyyy-mm-dd") (num 45000.0)) None
eq "a fraction is left alone" (Ops.changeDecimals 1 (Some "# ?/?") (num 1.5)) None
eq "the editor's named formats become codes" (Ops.changeDecimals 1 (Some "currency") (num 1.0)) (Some "$#,##0.000")
eq "at most thirty decimals" (Ops.changeDecimals 1 (Some ("0." + System.String('0', 30))) (num 1.0)) None
eq "a good code" (Ops.numberFormatProblem "#,##0.00;[Red]-#,##0.00") None
eq "a date code" (Ops.numberFormatProblem "dd/mm/yyyy hh:mm") None
eq "an empty code" (Ops.numberFormatProblem "  ").IsSome true
eq "an open quote" (Ops.numberFormatProblem "0.0 \"units").IsSome true
eq "an open bracket" (Ops.numberFormatProblem "[Red 0.0").IsSome true
eq "a dangling escape" (Ops.numberFormatProblem "0.0\\").IsSome true
eq "five sections" (Ops.numberFormatProblem "0;0;0;@;0").IsSome true

// ---- clearing cells ----------------------------------------------------------------
let boldFormat = { defaultFormat with Bold = true }
let toClear =
    createEmpty 100 10
    |> setCell (at 0 0) (cell (num 1.0) boldFormat)
    |> setCell (at 1 0) (cell (num 2.0) defaultFormat)
    |> setCell (at 2 0) { Value = CellValue.String "noted"; Format = boldFormat; Comment = Some { Text = "why"; Author = None; Runs = [] } }
    |> setCell (at 5 5) (cell (num 9.0) boldFormat)
let shape (m: SpreadsheetModel) = m.Cells |> Map.toList |> List.map (fun (a, c) -> (a.Row, a.Col), c.Value, c.Format.Bold, c.Comment.IsSome)
eq "clear contents keeps formats and comments" (shape (Ops.clearCells Ops.ClearContents (0, 0, 2, 0) toClear))
    [ (0, 0), CellValue.Empty, true, false; (2, 0), CellValue.Empty, true, true; (5, 5), num 9.0, true, false ]
eq "clear formats keeps values and comments" (shape (Ops.clearCells Ops.ClearFormats (0, 0, 2, 0) toClear))
    [ (0, 0), num 1.0, false, false; (1, 0), num 2.0, false, false; (2, 0), CellValue.String "noted", false, true; (5, 5), num 9.0, true, false ]
eq "clear all leaves empty cells" (shape (Ops.clearCells Ops.ClearEverything (0, 0, 2, 0) toClear)) [ (5, 5), num 9.0, true, false ]
eq "a whole column costs the cells that exist" (shape (Ops.clearCells Ops.ClearContents (0, 0, Ops.maxSheetRow, 0) toClear)) (shape (Ops.clearCells Ops.ClearContents (0, 0, 2, 0) toClear))
eq "nothing to clear: the same model" (obj.ReferenceEquals(Ops.clearCells Ops.ClearFormats (10, 0, 20, 3) toClear, toClear)) true
eq "rich text goes with the contents" (Ops.clearCells Ops.ClearContents (0, 0, 0, 0) { toClear with RichRuns = Map.ofList [ at 0 0, [ "1", emptyRunOverride ] ] }).RichRuns.IsEmpty true

// ---- what a cell shows under the sheet's view options --------------------------------
let viewSheet = createEmpty 10 5
let shownAs (m: SpreadsheetModel) (v: CellValue) = viewedValue m v
let sum = CellValue.Formula ("=A1-5", CellValue.Number 0.0)
eq "the default view shows values as they are" (shownAs viewSheet sum, shownAs viewSheet (CellValue.Number 0.0)) (sum, CellValue.Number 0.0)
eq "formulas shown: the formula text, with its =" (shownAs { viewSheet with ShowFormulas = true } sum, shownAs { viewSheet with ShowFormulas = true } (CellValue.Formula ("A1*2", CellValue.Number 4.0)))
    (CellValue.String "=A1-5", CellValue.String "=A1*2")
eq "formulas shown: other cells are unchanged" (shownAs { viewSheet with ShowFormulas = true } (CellValue.Number 3.0)) (CellValue.Number 3.0)
let noZeros = { viewSheet with ShowZeros = false }
eq "zeros hidden: a typed and a calculated zero" (shownAs noZeros (CellValue.Number 0.0), shownAs noZeros sum) (CellValue.Empty, CellValue.Empty)
eq "zeros hidden: everything else stays" (shownAs noZeros (CellValue.Number 0.5), shownAs noZeros (CellValue.String "0"), shownAs noZeros (CellValue.Formula ("=1", CellValue.Number 1.0)))
    (CellValue.Number 0.5, CellValue.String "0", CellValue.Formula ("=1", CellValue.Number 1.0))
eq "both: formulas win over hidden zeros" (shownAs { noZeros with ShowFormulas = true } sum) (CellValue.String "=A1-5")
eq "a new sheet: 100 %, headers, values, zeros" (viewSheet.Zoom, viewSheet.ShowHeaders, viewSheet.ShowFormulas, viewSheet.ShowZeros) (1.0, true, false, true)

printfn "stage33-sheetops: %d passed, %d failed" pass fail
if fail > 0 then exit 1
