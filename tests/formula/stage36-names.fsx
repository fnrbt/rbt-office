#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
module Wb = Rbt.Spreadsheet.Workbook
module Ops = Rbt.Spreadsheet.SheetOps
module Recalc = Rbt.Spreadsheet.Recalc

// Defined names as the host API manages them: which names Excel accepts,
// what a refers-to text may be, set / replace / delete by scope with the
// spelling and comment kept, and the formulas that use a name following it
// (re-pointed, deleted -> #NAME?, put back).

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-40s expected %A  got %A" label expected got)

let at row col : CellAddress = { Row = row; Col = col }
let num (n: float) = { Value = CellValue.Number n; Format = defaultFormat; Comment = None }
let fml (text: string) = { Value = CellValue.Formula (text, CellValue.Empty); Format = defaultFormat; Comment = None }
let sheetOf cells = cells |> List.fold (fun m (addr, cell) -> setCell addr cell m) (createEmpty 50 10)

// ---- which names are names ---------------------------------------------------------------
let accepted name = (Wb.definedNameProblem name).IsNone
eq "plain names" ([ "Rate"; "_total"; "\\x"; "Tax.Rate_2"; "Really?"; "Größe"; "A1B"; "ABCD1"; "RC1X"; "Rate1" ] |> List.map accepted) (List.replicate 10 true)
eq "not names" ([ ""; "1st"; "My name"; "A1"; "xfd1048576"; "R1C1"; "r"; "C"; "R5"; "C12"; "RC"; "TRUE"; "false"; "a+b"; "$A$1"; String.replicate 256 "x" ] |> List.map accepted) (List.replicate 16 false)
eq "the longest name" (accepted (String.replicate 255 "x")) true
eq "past the last column is a name" (accepted "XFE1", accepted "A1048577") (true, true)
eq "the reason is said" (Wb.definedNameProblem "My name") (Some "the name \"My name\" may hold letters, digits, underscores, periods, backslashes and question marks only")

// ---- what a name may refer to ------------------------------------------------------------
let refers text = (Ops.refersToProblem text).IsNone
eq "references and formulas" ([ "Sheet1!$A$1:$B$9"; "=Sheet1!$A$1"; "0.19"; "\"text\""; "SUM(Data!$A$1:$A$3)*2"; "'My sheet'!$A$1"; "OtherName" ] |> List.map refers) (List.replicate 7 true)
eq "not formulas" ([ ""; "="; "  "; "SUM("; "A1:" ] |> List.map refers) (List.replicate 5 false)
eq "areas of a reference list" (Ops.splitAreas "Data!$A$1:$A$3,'a,b'!$C$1, SUM(1,2)") [ "Data!$A$1:$A$3"; "'a,b'!$C$1"; "SUM(1,2)" ]

// ---- set, replace, delete ------------------------------------------------------------------
let data = sheetOf [ at 0 0, num 1.0; at 1 0, num 2.0; at 2 0, num 3.0; at 0 1, num 10.0; at 1 1, num 20.0
                     at 4 0, fml "=SUM(Totals)"; at 5 0, fml "=Rate*100" ]
let other = sheetOf [ at 0 0, fml "=SUM(Totals)"; at 1 0, fml "=Rate*100"; at 0 1, num 7.0 ]
let wb0 : Wb.Workbook =
    { Sheets = Map.ofList [ "sheet1", ({ Id = "sheet1"; Name = "Data"; Model = data; Visible = true } : Wb.Sheet)
                            "sheet2", ({ Id = "sheet2"; Name = "Other"; Model = other; Visible = true } : Wb.Sheet) ]
      SheetOrder = [ "sheet1"; "sheet2" ]; ActiveSheet = "sheet1"; Tables = Map.empty; Names = Map.empty; NameInfo = Map.empty }
let valueOf sheet row col (wb: Wb.Workbook) = getCellValue (at row col) wb.Sheets.[sheet].Model |> unwrapFormulaValue
let calc = recalculateWorkbook

let none = calc wb0
eq "without the names: #NAME?" (valueOf "sheet1" 4 0 none, valueOf "sheet2" 1 0 none) (CellValue.Error "NAME?", CellValue.Error "NAME?")

let wb1 =
    wb0
    |> Wb.setDefinedName None "Totals" "=Data!$A$1:$A$3" (Some "the first column")
    |> Wb.setDefinedName None "Rate" "0.19" None
eq "the text is kept without =" (Wb.getName None "totals" wb1, Wb.getName None "RATE" wb1) (Some "Data!$A$1:$A$3", Some "0.19")
eq "listed with spelling and comment" (Wb.definedNames wb1 |> List.map (fun n -> n.Scope, n.Name, n.RefersTo, n.Comment))
    [ None, "Rate", "0.19", None; None, "Totals", "Data!$A$1:$A$3", Some "the first column" ]
let c1 = calc wb1
eq "formulas on both sheets read the names" (valueOf "sheet1" 4 0 c1, valueOf "sheet1" 5 0 c1, valueOf "sheet2" 0 0 c1, valueOf "sheet2" 1 0 c1)
    (CellValue.Number 6.0, CellValue.Number 19.0, CellValue.Number 6.0, CellValue.Number 19.0)

// A sheet-level name shadows the workbook's on its sheet only.
let wb2 = wb1 |> Wb.setDefinedName (Some "sheet2") "rate" "Other!$B$1" None
let c2 = calc wb2
eq "a sheet's own name wins on that sheet" (valueOf "sheet1" 5 0 c2, valueOf "sheet2" 1 0 c2) (CellValue.Number 19.0, CellValue.Number 700.0)
eq "found by scope, whatever the case" ((Wb.tryFindDefinedName (Some "sheet2") "RATE" wb2) |> Option.map (fun n -> n.Name, n.RefersTo), (Wb.tryFindDefinedName (Some "sheet1") "Rate" wb2).IsNone)
    (Some ("rate", "Other!$B$1"), true)

// Replacing re-points every formula that uses the name; the spelling and comment are the new ones.
let wb3 = wb2 |> Wb.setDefinedName None "TOTALS" "Data!$B$1:$B$2" None
let c3 = calc wb3
eq "a re-pointed name" (valueOf "sheet1" 4 0 c3, valueOf "sheet2" 0 0 c3) (CellValue.Number 30.0, CellValue.Number 30.0)
eq "one name, the new spelling, no comment" (Wb.definedNames wb3 |> List.filter (fun n -> n.Name.ToUpperInvariant() = "TOTALS") |> List.map (fun n -> n.Name, n.Comment)) [ "TOTALS", None ]

// The incremental engine does not go on over changed names: the edit is declined to a full pass.
let engine = Recalc.build c2
eq "a changed name is not an edit of cells" ((Recalc.recalculate engine wb3 None).IsNone) true

// Deleting: #NAME?; the sheet-level one gone, the workbook's shows again.
let wb4 = (Wb.deleteDefinedName None "totals" wb3).Value
let c4 = calc wb4
eq "a deleted name" (valueOf "sheet1" 4 0 c4, valueOf "sheet2" 0 0 c4) (CellValue.Error "NAME?", CellValue.Error "NAME?")
eq "its entry and its info are gone" (wb4.Names.ContainsKey (None, "TOTALS"), wb4.NameInfo.ContainsKey (None, "TOTALS")) (false, false)
eq "deleting what is not there" ((Wb.deleteDefinedName None "Nope" wb4).IsNone, (Wb.deleteDefinedName (Some "sheet1") "Rate" wb4).IsNone) (true, true)
let wb5 = (Wb.deleteDefinedName (Some "sheet2") "Rate" wb4).Value
eq "without its own name the sheet reads the workbook's" (valueOf "sheet2" 1 0 (calc wb5)) (CellValue.Number 19.0)

// Undo / redo of one name: put back as it was.
let before = Wb.tryFindDefinedName None "Totals" wb3
let undone = Wb.restoreDefinedName None "Totals" before wb4
eq "restored with its spelling" (Wb.definedNames undone = Wb.definedNames wb3, valueOf "sheet1" 4 0 (calc undone)) (true, CellValue.Number 30.0)
let unmade = Wb.restoreDefinedName None "Totals" None undone
eq "restored to nothing" (unmade.Names.ContainsKey (None, "TOTALS")) false
// A replaced name keeps the file's hidden flag.
let hidden = { wb1 with NameInfo = wb1.NameInfo |> Map.add (None, "RATE") { Name = "Rate"; Comment = None; Hidden = true } } |> Wb.setDefinedName None "Rate" "0.2" (Some "new")
eq "hidden stays, comment is new" (Wb.tryFindDefinedName None "rate" hidden |> Option.map (fun n -> n.Hidden, n.Comment, n.RefersTo)) (Some (true, Some "new", "0.2"))

// Names follow rows inserted above what they refer to.
let moved, _ = Wb.applyStructureEdit "sheet1" (InsertRows (0, 2)) wb1
eq "a name follows inserted rows, its info stays" (Wb.tryFindDefinedName None "Totals" moved |> Option.map (fun n -> n.Name, n.RefersTo, n.Comment)) (Some ("Totals", "Data!$A$3:$A$5", Some "the first column"))

printfn "stage36-names: %d passed, %d failed" pass fail
if fail > 0 then exit 1
