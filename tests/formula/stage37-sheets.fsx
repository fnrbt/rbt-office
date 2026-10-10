#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
module Wb = Rbt.Spreadsheet.Workbook

// The list of sheets as the host API changes it: which names a sheet may
// have, renaming a sheet with every formula, name, rule and link that names
// it following (quoted where needed), deleting one (#REF!, and back),
// moving, copying, hiding, and the ids that stay with the sheets throughout.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-40s expected %A  got %A" label expected got)

let at row col : CellAddress = { Row = row; Col = col }
let num (n: float) = { Value = CellValue.Number n; Format = defaultFormat; Comment = None }
let fml (text: string) = { Value = CellValue.Formula (text, CellValue.Empty); Format = defaultFormat; Comment = None }
let sheetOf cells = cells |> List.fold (fun m (addr, cell) -> setCell addr cell m) (createEmpty 50 10)
let range r0 c0 r1 c1 : CellRange = { From = at r0 c0; To = at r1 c1 }

// ---- names ---------------------------------------------------------------------------------
let fine name = (Wb.sheetNameProblem name).IsNone
eq "names a sheet may have" ([ "Data"; "My Sheet"; "2024"; "a'b"; "Größe (2)"; String.replicate 31 "x"; "history2" ] |> List.map fine) (List.replicate 7 true)
eq "names it may not have" ([ ""; "   "; String.replicate 32 "x"; "a:b"; "a\\b"; "a/b"; "a?b"; "a*b"; "a[b"; "a]b"; "'a"; "a'"; "History"; "HISTORY" ] |> List.map fine) (List.replicate 14 false)
eq "the reason is said" (Wb.sheetNameProblem "a/b") (Some "a sheet name cannot contain any of : \\ / ? * [ ]")

// ---- a workbook ----------------------------------------------------------------------------
let data =
    { sheetOf [ at 0 0, num 1.0; at 1 0, num 2.0; at 2 0, num 3.0; at 3 0, fml "=SUM(A1:A3)"; at 4 0, fml "=Data!A1+Summary!B1" ] with
        ConditionalFormats = [ { Id = 1; Ranges = [ range 0 0 2 0 ]; Format = { Rule = Expression "Summary!$B$1>Data!$A$1"; Style = defaultFormat }; Priority = 1; StopIfTrue = false } ]
        Validations = [ { Id = 1; Ranges = [ range 0 2 5 2 ]; Validation = validationOf (ListSource "Summary!$B$1:$B$3") } ]
        Hyperlinks = [ { Range = range 6 0 6 0; Target = "Summary!A1"; IsInternal = true; Tooltip = None; Display = None }; { Range = range 7 0 7 0; Target = "https://example.com/Summary!A1"; IsInternal = false; Tooltip = None; Display = None } ] }
let summary =
    sheetOf [ at 0 0, fml "=SUM(Data!A1:A3)"; at 1 0, fml "=Data!$A$4*2"; at 2 0, fml "=COUNT(data!A:A)+Rate"; at 3 0, fml "=\"Data!A1\"&Other!A1"
              at 0 1, num 10.0; at 4 0, fml "=SUM(DataTotal)"; at 5 0, fml "=Local" ]
let other = sheetOf [ at 0 0, num 5.0; at 1 0, fml "=Summary!A1+1" ]
let sheetEntry id name model : Wb.SheetId * Wb.Sheet = id, { Id = id; Name = name; Model = model; Visible = true }
let wbBare : Wb.Workbook =
    { Sheets = Map.ofList [ sheetEntry "sheet1" "Data" data; sheetEntry "sheet2" "Summary" summary; sheetEntry "sheet3" "Other" other ]
      SheetOrder = [ "sheet1"; "sheet2"; "sheet3" ]; ActiveSheet = "sheet1"; Tables = Map.empty; Names = Map.empty; NameInfo = Map.empty; Iteration = None }
let wb0 : Wb.Workbook =
    wbBare
    |> Wb.setDefinedName None "DataTotal" "Data!$A$1:$A$3" None
    |> Wb.setDefinedName None "Rate" "0.5" None
    |> Wb.setDefinedName (Some "sheet1") "Local" "Data!$A$1" (Some "of the data sheet")
let formulaAt sheet row col (wb: Wb.Workbook) =
    match getCellValue (at row col) wb.Sheets.[sheet].Model with
    | CellValue.Formula (text, _) -> text
    | value -> sprintf "%A" value
let valueAt sheet row col (wb: Wb.Workbook) = getCellValue (at row col) wb.Sheets.[sheet].Model |> unwrapFormulaValue
let calc = recalculateWorkbook
let names (wb: Wb.Workbook) = wb.SheetOrder |> List.map (fun id -> wb.Sheets.[id].Name)

eq "taken whatever the case" (Wb.sheetNameTaken None "data" wb0, Wb.sheetNameTaken (Some "sheet1") "DATA" wb0, Wb.sheetNameTaken None "Else" wb0) (true, false, false)
eq "the next free SheetN" (Wb.nextSheetName wb0, Wb.nextSheetName (Wb.renameSheetEverywhere "sheet3" "Sheet1" wb0)) ("Sheet1", "Sheet2")
eq "the name of a copy" (Wb.copySheetName "Data" wb0, Wb.copySheetName "Data" (Wb.renameSheetEverywhere "sheet3" "Data (2)" wb0)) ("Data (2)", "Data (3)")
eq "... cut to 31 characters" (Wb.copySheetName (String.replicate 31 "x") wb0) (String.replicate 27 "x" + " (2)")
eq "a fresh id is no sheet's and none used before" (Wb.freshSheetId (fun id -> id = "sheet4") wb0) "sheet5"

// ---- rename -------------------------------------------------------------------------------
let renamed = Wb.renameSheetEverywhere "sheet1" "My Data" wb0
eq "rename: the sheet" (names renamed, renamed.SheetOrder) ([ "My Data"; "Summary"; "Other" ], [ "sheet1"; "sheet2"; "sheet3" ])
eq "rename: formulas on other sheets, quoted" (formulaAt "sheet2" 0 0 renamed, formulaAt "sheet2" 1 0 renamed) ("=SUM('My Data'!A1:A3)", "='My Data'!$A$4*2")
eq "rename: whatever the case it was written in" (formulaAt "sheet2" 2 0 renamed) "=COUNT('My Data'!A:A)+Rate"
eq "rename: text that looks like a reference stays" (formulaAt "sheet2" 3 0 renamed) "=\"Data!A1\"&Other!A1"
eq "rename: the sheet's own qualified references" (formulaAt "sheet1" 4 0 renamed, formulaAt "sheet1" 3 0 renamed) ("='My Data'!A1+Summary!B1", "=SUM(A1:A3)")
eq "rename: untouched formulas are the same cells" (obj.ReferenceEquals(renamed.Sheets.["sheet3"].Model, wb0.Sheets.["sheet3"].Model)) true
eq "rename: defined names" (Wb.getName None "DataTotal" renamed, Wb.getName (Some "sheet1") "Local" renamed, Wb.getName None "Rate" renamed) (Some "'My Data'!$A$1:$A$3", Some "'My Data'!$A$1", Some "0.5")
eq "rename: a conditional format's formula" (renamed.Sheets.["sheet1"].Model.ConditionalFormats |> List.map (fun f -> f.Format.Rule)) [ Expression "Summary!$B$1>'My Data'!$A$1" ]
let renamedSummary = Wb.renameSheetEverywhere "sheet2" "It's" wb0
eq "rename: an apostrophe is doubled" (formulaAt "sheet1" 4 0 renamedSummary, formulaAt "sheet3" 1 0 renamedSummary) ("=Data!A1+'It''s'!B1", "='It''s'!A1+1")
eq "rename: a validation's list source" (renamedSummary.Sheets.["sheet1"].Model.Validations |> List.map (fun v -> v.Validation.Rule)) [ ListSource "'It''s'!$B$1:$B$3" ]
eq "rename: hyperlinks into the sheet, not addresses" (renamedSummary.Sheets.["sheet1"].Model.Hyperlinks |> List.map (fun l -> l.Target)) [ "'It''s'!A1"; "https://example.com/Summary!A1" ]
let back = Wb.renameSheetEverywhere "sheet1" "Data" renamed
eq "rename back: as it was" (formulaAt "sheet2" 0 0 back, formulaAt "sheet2" 2 0 back, Wb.getName None "DataTotal" back, names back) ("=SUM(Data!A1:A3)", "=COUNT(Data!A:A)+Rate", Some "Data!$A$1:$A$3", names wb0)
eq "rename: values stand" (valueAt "sheet2" 0 0 (calc renamed), valueAt "sheet2" 4 0 (calc renamed), valueAt "sheet2" 5 0 (calc renamed)) (CellValue.Number 6.0, CellValue.Number 6.0, CellValue.Error "NAME?")
eq "rename: plain names need no quotes" (formulaAt "sheet2" 0 0 (Wb.renameSheetEverywhere "sheet1" "Facts_2" wb0)) "=SUM(Facts_2!A1:A3)"
eq "rename: a name like a cell reference is quoted" (formulaAt "sheet2" 0 0 (Wb.renameSheetEverywhere "sheet1" "A1" wb0)) "=SUM('A1'!A1:A3)"
eq "rename: to the same name is nothing" (obj.ReferenceEquals(Wb.renameSheetEverywhere "sheet1" "Data" wb0, wb0)) true

// ---- delete ---------------------------------------------------------------------------------
eq "the only visible sheet stays" (Wb.lastVisibleSheetProblem "sheet1" (wb0 |> Wb.setSheetVisibility "sheet2" false false |> Wb.setSheetVisibility "sheet3" false true)) (Some "a workbook must keep at least one visible sheet")
eq "others may go" (Wb.lastVisibleSheetProblem "sheet1" wb0, Wb.lastVisibleSheetProblem "sheet9" wb0) (None, Some "no such sheet")
let without, removal = (Wb.removeSheet "sheet1" wb0).Value
eq "delete: the sheets left, the next one active" (names without, without.ActiveSheet) ([ "Summary"; "Other" ], "sheet2")
eq "delete: references are #REF!" (formulaAt "sheet2" 0 0 without, formulaAt "sheet2" 1 0 without, formulaAt "sheet2" 2 0 without) ("=SUM(#REF!)", "=#REF!*2", "=COUNT(#REF!)+Rate")
eq "delete: other formulas stay" (formulaAt "sheet3" 1 0 without, formulaAt "sheet2" 3 0 without) ("=Summary!A1+1", "=\"Data!A1\"&Other!A1")
eq "delete: names that referred to it" (Wb.getName None "DataTotal" without, Wb.getName None "Rate" without) (Some "#REF!", Some "0.5")
eq "delete: the names of its scope are gone" (Wb.getName (Some "sheet1") "Local" without, without.NameInfo |> Map.containsKey (Some "sheet1", "LOCAL")) (None, false)
let withoutCalc = calc without
eq "delete: the values" (valueAt "sheet2" 0 0 withoutCalc, valueAt "sheet2" 4 0 withoutCalc) (CellValue.Error "REF!", CellValue.Error "REF!")
// A link into the deleted sheet leads nowhere: its target says #REF! (other links stay).
eq "delete: links into the sheet" (((Wb.removeSheet "sheet2" wb0).Value |> fst).Sheets.["sheet1"].Model.Hyperlinks |> List.map (fun l -> l.Target, l.IsInternal)) [ "#REF!", true; "https://example.com/Summary!A1", false ]
eq "undo: the links are back" ((Wb.restoreSheet (snd (Wb.removeSheet "sheet2" wb0).Value) (fst (Wb.removeSheet "sheet2" wb0).Value)).Sheets.["sheet1"].Model.Hyperlinks |> List.map (fun l -> l.Target)) (wb0.Sheets.["sheet1"].Model.Hyperlinks |> List.map (fun l -> l.Target))
let restored = Wb.restoreSheet removal without
eq "undo: the sheet is back in its place and active" (names restored, restored.SheetOrder, restored.ActiveSheet) (names wb0, wb0.SheetOrder, "sheet1")
eq "undo: the references are back" (formulaAt "sheet2" 0 0 restored, formulaAt "sheet2" 1 0 restored, formulaAt "sheet2" 2 0 restored) ("=SUM(Data!A1:A3)", "=Data!$A$4*2", "=COUNT(data!A:A)+Rate")
eq "undo: the names are back" (Wb.getName None "DataTotal" restored, Wb.getName (Some "sheet1") "Local" restored, (Wb.definedNames restored |> List.find (fun n -> n.Name = "Local")).Comment) (Some "Data!$A$1:$A$3", Some "Data!$A$1", Some "of the data sheet")
eq "undo: the values" (valueAt "sheet2" 0 0 (calc restored), valueAt "sheet2" 5 0 (calc restored)) (CellValue.Number 6.0, CellValue.Error "NAME?")
// A formula edited after the delete keeps the edit.
let edited = { without with Sheets = without.Sheets |> Map.add "sheet2" { without.Sheets.["sheet2"] with Model = setCell (at 0 0) (fml "=1+1") without.Sheets.["sheet2"].Model } }
eq "undo: a formula changed meanwhile is left" (formulaAt "sheet2" 0 0 (Wb.restoreSheet removal edited), formulaAt "sheet2" 1 0 (Wb.restoreSheet removal edited)) ("=1+1", "=Data!$A$4*2")
let middle, middleRemoval = (Wb.removeSheet "sheet2" wb0).Value
eq "delete: rules and links of other sheets" (middle.Sheets.["sheet1"].Model.ConditionalFormats |> List.map (fun f -> f.Format.Rule), middle.Sheets.["sheet1"].Model.Validations |> List.map (fun v -> v.Validation.Rule)) ([ Expression "#REF!>Data!$A$1" ], [ ListSource "#REF!" ])
eq "delete: a sheet that was not active leaves the active one" (middle.ActiveSheet, names middle) ("sheet1", [ "Data"; "Other" ])
let middleBack = Wb.restoreSheet middleRemoval middle
eq "undo: rules are back" (middleBack.Sheets.["sheet1"].Model.ConditionalFormats = data.ConditionalFormats, middleBack.Sheets.["sheet1"].Model.Validations = data.Validations, names middleBack) (true, true, names wb0)
eq "the last sheet is never removed" ((Wb.removeSheet "sheet2" (fst (Wb.removeSheet "sheet1" (fst (Wb.removeSheet "sheet3" wb0).Value)).Value)).IsNone) true
// Deleting the active last sheet shows the one before it.
eq "delete: the last sheet, active" ((fst (Wb.removeSheet "sheet3" { wb0 with ActiveSheet = "sheet3" }).Value).ActiveSheet) "sheet2"
// ... and a hidden neighbour is passed over.
eq "delete: the nearest visible sheet becomes active" ((fst (Wb.removeSheet "sheet1" (Wb.setSheetVisibility "sheet2" false false wb0)).Value).ActiveSheet) "sheet3"

// ---- add, move --------------------------------------------------------------------------------
let added = Wb.insertSheet { Id = "sheet4"; Name = "New"; Model = createEmpty 10 5; Visible = true } 1 wb0
eq "add: in its place and active" (names added, added.ActiveSheet) ([ "Data"; "New"; "Summary"; "Other" ], "sheet4")
eq "add: past the end is the end" (names (Wb.insertSheet { Id = "sheet4"; Name = "New"; Model = createEmpty 10 5; Visible = true } 99 wb0)) [ "Data"; "Summary"; "Other"; "New" ]
eq "move: to the end, to the front" (names (Wb.moveSheet "sheet1" 2 wb0), names (Wb.moveSheet "sheet3" 0 wb0), (Wb.moveSheet "sheet1" 2 wb0).ActiveSheet) ([ "Summary"; "Other"; "Data" ], [ "Other"; "Data"; "Summary" ], "sheet1")
eq "move: to where it is" (obj.ReferenceEquals(Wb.moveSheet "sheet2" 1 wb0, wb0)) true
eq "move: formulas read the same" (valueAt "sheet2" 0 0 (calc (Wb.moveSheet "sheet1" 2 wb0))) (CellValue.Number 6.0)

// ---- visibility -------------------------------------------------------------------------------
let hidden = Wb.setSheetVisibility "sheet1" false false wb0
eq "hide: the active sheet gives way" (hidden.Sheets.["sheet1"].Visible, hidden.ActiveSheet, Wb.sheetVisibility hidden.Sheets.["sheet1"]) (false, "sheet2", "hidden")
let veryHidden = Wb.setSheetVisibility "sheet3" false true wb0
eq "very hidden" (Wb.sheetVisibility veryHidden.Sheets.["sheet3"], veryHidden.ActiveSheet, veryHidden.Sheets.["sheet3"].Model.VeryHidden) ("veryHidden", "sheet1", true)
eq "shown again" (Wb.sheetVisibility (Wb.setSheetVisibility "sheet3" true true veryHidden).Sheets.["sheet3"], (Wb.setSheetVisibility "sheet3" true false veryHidden).Sheets.["sheet3"].Model.VeryHidden) ("visible", false)
eq "a hidden sheet is still read" (valueAt "sheet2" 0 0 (calc hidden)) (CellValue.Number 6.0)

// ---- copy -------------------------------------------------------------------------------------
let tableModel : Table = { Name = "Sales"; StartRow = 1; StartCol = 5; EndRow = 3; EndCol = 6; HeaderRowCount = 1; Style = None; AutoFilter = None; HasTotals = false }
let pictured : SheetDrawing =
    { Id = 1; Anchor = { Kind = "oneCell"; FromRow = 1; FromCol = 1; FromOffsetX = 0.0; FromOffsetY = 0.0; ToRow = None; ToCol = None; ToOffsetX = 0.0; ToOffsetY = 0.0; X = 0.0; Y = 0.0; Width = Some 10.0; Height = Some 10.0; EditAs = None }
      Name = Some "p"; Description = None; Content = Picture ("image/png", "png", [| 1uy; 2uy |]) }
let kept = { pictured with Name = Some "chart"; Content = KeptObject ("chart", "<xdr:graphicFrame/>") }
let text (value: string) = { num 0.0 with Value = CellValue.String value }
let extra = [ at 0 4, text "Amount"; at 0 5, text "Tax"; at 1 4, num 100.0; at 2 4, num 50.0; at 5 0, fml "=SUM(Sales[Amount])+Data!A1"; at 6 0, fml "=Summary!B1+Local" ]
let source =
    { (extra |> List.fold (fun m (addr, cell) -> setCell addr cell m) data) with
        Tables = [ tableModel ]; Drawings = [ pictured; kept ]; KeptRanges = [ { Key = "xl/pivotTables/pivotTable1.xml"; Range = range 10 0 12 2 } ]
        SourceSheet = Some { Book = 1; Sheet = 0 } }
let wbCopy0 =
    { wb0 with
        Sheets = wb0.Sheets |> Map.add "sheet1" { wb0.Sheets.["sheet1"] with Model = source }
        Tables = Map.ofList [ "Sales", Wb.makeTableFromRegion "sheet1" "Sales" 1 5 3 6 1 0 [] source ] }
let copied = Wb.duplicateSheet "sheet1" "sheet4" "Data (2)" 1 wbCopy0
let copy = copied.Sheets.["sheet4"].Model
eq "copy: beside the original and active" (names copied, copied.ActiveSheet) ([ "Data"; "Data (2)"; "Summary"; "Other" ], "sheet4")
eq "copy: its formulas name the copy, and the copy of the table" (formulaAt "sheet4" 5 0 copied, formulaAt "sheet4" 4 0 copied, formulaAt "sheet4" 6 0 copied) ("=SUM(Sales_2[Amount])+'Data (2)'!A1", "='Data (2)'!A1+Summary!B1", "=Summary!B1+Local")
eq "copy: the original is as it was" (obj.ReferenceEquals(copied.Sheets.["sheet1"].Model, source), formulaAt "sheet1" 5 0 copied) (true, "=SUM(Sales[Amount])+Data!A1")
eq "copy: tables" (copy.Tables |> List.map (fun t -> t.Name), copied.Tables |> Map.toList |> List.map (fun (key, t) -> key, t.Name, t.SheetId, t.Columns)) ([ "Sales_2" ], [ "Sales", "Sales", "sheet1", [ "Amount"; "Tax" ]; "Sales_2", "Sales_2", "sheet4", [ "Amount"; "Tax" ] ])
eq "copy: the names of the sheet's scope" (Wb.getName (Some "sheet4") "Local" copied, Wb.getName (Some "sheet1") "Local" copied, (Wb.definedNames copied |> List.filter (fun n -> n.Name = "Local") |> List.map (fun n -> n.Scope, n.Comment))) (Some "'Data (2)'!$A$1", Some "Data!$A$1", [ Some "sheet1", Some "of the data sheet"; Some "sheet4", Some "of the data sheet" ])
// (The kept drawing objects - charts - come along: written to XLSX, their chart parts are copied for the copy, and so are
// the pivot table parts, each where the copy has its cells; the range a pivot cache reads stays the source's alone.)
eq "copy: pictures, kept objects and where the pivot tables are come along" (copy.Drawings |> List.map (fun d -> d.Name), copy.KeptRanges, copy.SourceSheet, copy.KeptFrom)
    ([ Some "p"; Some "chart" ], source.KeptRanges |> List.filter (fun kept -> kept.Key.StartsWith "xl/pivotTables/"), None, Some { Book = 1; Sheet = 0 })
eq "copy: the source has a pivot table to copy" (source.KeptRanges |> List.exists (fun kept -> kept.Key.StartsWith "xl/pivotTables/")) true
eq "copy of a copy: the kept objects are still the file's" ((Wb.duplicateSheet "sheet4" "sheet5" "Data (3)" 2 copied).Sheets.["sheet5"].Model.KeptFrom) (Some { Book = 1; Sheet = 0 })
eq "copy: rules name the copy" (copy.ConditionalFormats |> List.map (fun f -> f.Format.Rule)) [ Expression "Summary!$B$1>'Data (2)'!$A$1" ]
let copiedCalc = calc copied
eq "copy: the values" (valueAt "sheet4" 5 0 copiedCalc, valueAt "sheet1" 5 0 copiedCalc, valueAt "sheet4" 6 0 copiedCalc) (CellValue.Number 151.0, CellValue.Number 151.0, CellValue.Number 11.0)
let uncopied, copyRemoval = (Wb.removeSheet "sheet4" copied).Value
eq "copy undone: the table and the names of the copy are gone" (uncopied.Tables |> Map.toList |> List.map fst, Wb.getName (Some "sheet4") "Local" uncopied, names uncopied) ([ "Sales" ], None, names wb0)
let recopied = Wb.restoreSheet copyRemoval uncopied
eq "copy redone" (names recopied, recopied.Tables |> Map.toList |> List.map fst, Wb.getName (Some "sheet4") "Local" recopied) (names copied, [ "Sales"; "Sales_2" ], Some "'Data (2)'!$A$1")

// ---- 3-D references follow a rename of either end -----------------------------------------------
let threeD = { wb0 with Sheets = wb0.Sheets |> Map.add "sheet3" { wb0.Sheets.["sheet3"] with Model = sheetOf [ at 0 0, fml "=SUM(Data:Summary!B1)" ] } }
let threeDRenamed = Wb.renameSheetEverywhere "sheet2" "Sums" threeD
eq "3-D: an end renamed (written quoted, as Excel also reads it)" (formulaAt "sheet3" 0 0 threeDRenamed) "=SUM('Data:Sums'!B1)"
eq "3-D: it still reads both sheets" (valueAt "sheet3" 0 0 (calc threeDRenamed), valueAt "sheet3" 0 0 (calc threeD)) (CellValue.Number 10.0, CellValue.Number 10.0)


// ---- 3-D references and the sheets of their span -------------------------------------------------
// Four sheets in tab order: Jan, Feb, Mar, Total. B1 holds 1, 10, 100.
do
    let month value = sheetOf [ at 0 1, num value; at 4 1, num (value * 2.0) ]
    let total = sheetOf [ at 0 0, fml "=SUM(Jan:Mar!B1)"; at 1 0, fml "=SUM(Jan:Feb!B1)+SUM(Feb:Mar!B1)"; at 2 0, fml "=SUM(Mar:Jan!B1:B5)"; at 3 0, fml "=SUM(Feb:Feb!B1)+Feb!B1" ]
    let book : Wb.Workbook =
        { wbBare with
            Sheets = Map.ofList [ sheetEntry "sheet1" "Jan" (month 1.0); sheetEntry "sheet2" "Feb" (month 10.0); sheetEntry "sheet3" "Mar" (month 100.0); sheetEntry "sheet4" "Total" total ]
            SheetOrder = [ "sheet1"; "sheet2"; "sheet3"; "sheet4" ] }
        |> Wb.setDefinedName None "Quarter" "SUM(Jan:Mar!$B$1)" None
    let totals (wb: Wb.Workbook) = [ for row in 0 .. 3 -> valueAt "sheet4" row 0 (calc wb) ]
    let formulas (wb: Wb.Workbook) = [ for row in 0 .. 3 -> formulaAt "sheet4" row 0 wb ]
    eq "3-D: calculated" (totals book) [ CellValue.Number 111.0; CellValue.Number 121.0; CellValue.Number 333.0; CellValue.Number 20.0 ]
    // The last sheet of a span deleted: the span ends at the sheet before it. A span of two is left with one sheet.
    let noMar = fst (Wb.removeSheet "sheet3" book).Value
    eq "3-D: the last sheet of the span deleted" (formulas noMar) [ "=SUM('Jan:Feb'!B1)"; "=SUM('Jan:Feb'!B1)+SUM(Feb!B1)"; "=SUM('Feb:Jan'!B1:B5)"; "=SUM(Feb:Feb!B1)+Feb!B1" ]
    eq "3-D: ... and what it reads" (totals noMar) [ CellValue.Number 11.0; CellValue.Number 21.0; CellValue.Number 33.0; CellValue.Number 20.0 ]
    eq "3-D: ... in a defined name too" (Wb.getName None "Quarter" noMar) (Some "SUM('Jan:Feb'!$B$1)")
    // The first sheet deleted: the span starts at the sheet after it.
    let noJan = fst (Wb.removeSheet "sheet1" book).Value
    eq "3-D: the first sheet of the span deleted" (formulas noJan) [ "=SUM('Feb:Mar'!B1)"; "=SUM(Feb!B1)+SUM('Feb:Mar'!B1)"; "=SUM('Mar:Feb'!B1:B5)"; "=SUM(Feb:Feb!B1)+Feb!B1" ]
    eq "3-D: ... and what it reads" (totals noJan) [ CellValue.Number 110.0; CellValue.Number 120.0; CellValue.Number 330.0; CellValue.Number 20.0 ]
    // A sheet from the middle: the reference is written as before and reads one sheet less; a span that was that sheet alone is #REF!.
    let noFeb = fst (Wb.removeSheet "sheet2" book).Value
    eq "3-D: a sheet deleted from the middle; a span of that sheet alone" (formulas noFeb) [ "=SUM(Jan:Mar!B1)"; "=SUM(Jan!B1)+SUM(Mar!B1)"; "=SUM(Mar:Jan!B1:B5)"; "=SUM(#REF!)+#REF!" ]
    eq "3-D: ... and what it reads" (totals noFeb |> List.truncate 3) [ CellValue.Number 101.0; CellValue.Number 101.0; CellValue.Number 303.0 ]
    // Undo of the deletion puts the formulas back.
    eq "3-D: the deletion undone" (formulas (Wb.restoreSheet (snd (Wb.removeSheet "sheet3" book).Value) noMar)) (formulas book)
    // Sheets moved: the reference is written as before; the sheets between its two ends are what it reads.
    let febLast = Wb.moveSheet "sheet2" 3 book          // Jan, Mar, Total, Feb
    eq "3-D: a sheet moved out of the span" (formulas febLast, (totals febLast).[0]) (formulas book, CellValue.Number 101.0)
    let totalInside = Wb.moveSheet "sheet4" 1 book       // Jan, Total, Feb, Mar
    eq "3-D: the formulas' own sheet moved into the span (its B1 is empty)" ((totals totalInside).[0]) (CellValue.Number 111.0)
    let marFirst = Wb.moveSheet "sheet3" 0 book          // Mar, Jan, Feb, Total: Jan:Mar is Mar and Jan
    eq "3-D: an end moved to the other side of the other end" ((totals marFirst).[0], (totals marFirst).[1]) (CellValue.Number 101.0, CellValue.Number 122.0)
    // A sheet added inside the span is read.
    let added = Wb.insertSheet { Id = "sheet9"; Name = "Extra"; Model = month 1000.0; Visible = true } 1 book
    eq "3-D: a sheet added inside the span" ((totals added).[0]) (CellValue.Number 1111.0)
    // Rows inserted on one sheet of a span: the reference stays as it is written, as in Excel (it is
    // the same cells on every sheet of the span; only an insertion on all of them at once - sheets
    // grouped - would move it). It reads that sheet's cell as it now is.
    let inserted, undo = Wb.applyStructureEdit "sheet2" (InsertRows (0, 2)) book
    eq "3-D: rows inserted on a sheet of the span: the reference is not re-pointed" (formulas inserted |> List.truncate 3) (formulas book |> List.truncate 3)
    eq "3-D: ... it reads what is at B1 on each sheet now" (totals inserted) [ CellValue.Number 101.0; CellValue.Number 101.0; CellValue.Number 313.0; CellValue.Number 10.0 ]
    eq "3-D: ... a plain reference to that sheet beside it is re-pointed" (formulaAt "sheet4" 3 0 inserted) "=SUM('Feb:Feb'!B1)+Feb!B3"
    eq "3-D: ... undone" (formulas (Wb.undoStructureEdit undo inserted), totals (Wb.undoStructureEdit undo inserted)) (formulas book, totals book)
    let cols, _ = Wb.applyStructureEdit "sheet4" (InsertCols (0, 1)) book
    eq "3-D: columns inserted on the formulas' own sheet, outside the span" (formulaAt "sheet4" 0 1 cols, valueAt "sheet4" 0 1 (calc cols)) ("=SUM(Jan:Mar!B1)", CellValue.Number 111.0)

printfn "\nSheet tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
