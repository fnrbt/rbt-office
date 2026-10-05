#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.ConditionalFormatting
module Ops = Rbt.Spreadsheet.SheetOps

// Data validations stored by range, and what the validation part of the host
// API stands on: a whole column without a cell, set / clear by range (what a
// cut validation keeps), ranges and formulas through row / column edits, a
// validation as plain data and back, what a description is refused for,
// lists read from cells (the sheet, another sheet, a name), and the scan for
// invalid cells.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-40s expected %A  got %A" label expected got)
let throws label (run: unit -> unit) (start: string) =
    let message = try run (); "(nothing thrown)" with e -> e.Message
    if message.StartsWith start then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-40s expected an error starting \"%s\"  got \"%s\"" label start message)

let at row col : CellAddress = { Row = row; Col = col }
let block r0 c0 r1 c1 : CellRange = { From = at r0 c0; To = at r1 c1 }
let column c = block 0 c Ops.maxSheetRow c
let num (n: float) = { Value = CellValue.Number n; Format = defaultFormat; Comment = None }
let str (s: string) = { Value = CellValue.String s; Format = defaultFormat; Comment = None }
let empty = createEmpty 100 26
let whole lo hi = validationOf (WholeNumber (Some lo, Some hi))
let invalid outcome = match outcome with EntryInvalid _ -> true | EntryValid -> false

// ---- a whole column: no cell is made, every row is covered ---------------------------
let col = empty |> setValidationRanges [ column 0 ] (whole 1.0 10.0)
eq "a whole column makes no cell" (col.Cells.Count, col.Validations.Length) (0, 1)
eq "its first and its last row are covered" ((getValidation (at 0 0) col).IsSome, (getValidation (at Ops.maxSheetRow 0) col).IsSome, (getValidation (at 5 1) col).IsSome) (true, true, false)
eq "an entry far down is checked" (invalid (checkEntry (at 50000 0) (CellValue.Number 99.0) col), invalid (checkEntry (at 50000 0) (CellValue.Number 5.0) col)) (true, false)
eq "the formulas are written for the first cell" (getValidation (at 7 0) col).Value.Anchor (Some (at 0 0))
eq "as stored the anchor is left out" col.Validations.Head.Validation.Anchor None

// ---- one validation per cell: a new one replaces what its cells had -----------------------
let two = col |> setValidationRanges [ block 2 0 4 0 ] (validationOf (ListValues [ "a"; "b" ]))
eq "the ids go up" (two.Validations |> List.map (fun v -> v.Id)) [ 1; 2 ]
eq "the first one is cut around the second" (two.Validations.Head.Ranges) [ block 0 0 1 0; block 5 0 Ops.maxSheetRow 0 ]
eq "the cut one keeps its anchor" two.Validations.Head.Validation.Anchor (Some (at 0 0))
eq "each cell has one" ((getValidation (at 3 0) two).Value.Rule, (getValidation (at 9 0) two).Value.Rule) (ListValues [ "a"; "b" ], WholeNumber (Some 1.0, Some 10.0))
eq "a list offers its dropdown" (showsListDropdown (at 3 0) two, showsListDropdown (at 9 0) two) (true, false)
let covered = two |> setValidationRanges [ column 0; block 0 2 0 2 ] (validationOf AnyValue)
eq "covering both removes them" (covered.Validations |> List.map (fun v -> v.Id, v.Ranges)) [ 3, [ column 0; block 0 2 0 2 ] ]
let cleared = two |> clearValidationRanges [ block 0 0 2 0 ]
eq "clear cuts by range" (cleared.Validations |> List.map (fun v -> v.Id, v.Ranges)) [ 1, [ block 5 0 Ops.maxSheetRow 0 ]; 2, [ block 3 0 4 0 ] ]
eq "the cut list has its anchor where it was" (cleared.Validations.[1].Validation.Anchor) (Some (at 2 0))
eq "clear of nothing changes nothing" (obj.ReferenceEquals(clearValidationRanges [ block 0 5 9 9 ] two, two)) true
eq "clear of everything" (clearValidationRanges [ block 0 0 Ops.maxSheetRow Ops.maxSheetCol ] two).Validations []
eq "the legacy per-cell call" ((empty |> setValidation (at 3 3) (whole 1.0 2.0)).Validations |> List.map (fun v -> v.Ranges)) [ [ block 3 3 3 3 ] ]

// A relative formula keeps its meaning for the cells left when the first is taken away.
let relative =
    empty
    |> setValidationRanges [ block 0 1 9 1 ] (validationOf (CustomFormula "=A1>0"))
    |> clearValidationRanges [ block 0 1 2 1 ]
let withA = { relative with Cells = Map.ofList [ at 5 0, num 3.0; at 6 0, num -3.0 ] }
eq "relative formula after a cut" (invalid (checkEntry (at 5 1) (CellValue.Number 1.0) withA), invalid (checkEntry (at 6 1) (CellValue.Number 1.0) withA)) (false, true)
eq "the file shape shifts the formula to the first cell left"
   (validationsToOoxml relative.Validations |> List.map (fun (sqref, v) -> sqref, v.Formula1)) [ "B4:B10", Some "A4>0" ]

// ---- rows and columns --------------------------------------------------------------
let shaped =
    empty
    |> setValidationRanges [ column 0 ] (whole 1.0 10.0)
    |> setValidationRanges [ block 4 2 8 3 ] (validationOf (Comparison (DecimalKind, OpGreaterThan, "=$F$5", None)))
    |> setValidationRanges [ block 20 5 20 5 ] (validationOf (ListSource "$H$2:$H$6"))
let inserted = insertRow 2 shaped
eq "insert rows: a whole column stays one, a block moves, formulas follow"
   (inserted.Validations |> List.map (fun v -> v.Ranges, v.Validation.Rule))
   [ [ column 0 ], WholeNumber (Some 1.0, Some 10.0)
     [ block 5 2 9 3 ], Comparison (DecimalKind, OpGreaterThan, "=$F$6", None)
     [ block 21 5 21 5 ], ListSource "$H$2:$H$7" ]
let grown = insertRow 6 shaped
eq "a row inserted inside a block grows it" (grown.Validations.[1].Ranges) [ block 4 2 9 3 ]
let deleted = shaped |> deleteRow 20
eq "a validation whose only row goes is gone" (deleted.Validations |> List.map (fun v -> v.Id)) [ 1; 2 ]
let colDeleted = shaped |> deleteColumn 2
eq "delete a column of a block" (colDeleted.Validations.[1].Ranges, colDeleted.Validations.[2].Ranges) ([ block 4 2 8 2 ], [ block 20 4 20 4 ])
eq "delete the whole-column's column" ((shaped |> deleteColumn 0).Validations |> List.map (fun v -> v.Id)) [ 2; 3 ]
eq "ids stay through edits" (inserted.Validations |> List.map (fun v -> v.Id)) [ 1; 2; 3 ]

// ---- a validation as plain data ---------------------------------------------------
let d kind = Ops.emptyValidationDescription kind
let build description = Ops.validationOfDescription "Set" description
let roundTrip (validation: DataValidation) = build (Ops.describeValidation validation)
let kinds =
    [ validationOf AnyValue
      validationOf (ListValues [ "Yes"; "No" ])
      { validationOf (ListSource "Lists!$A$1:$A$9") with ShowDropdown = false }
      validationOf (WholeNumber (Some 1.0, Some 10.0))
      validationOf (WholeNumber (Some 5.0, None))
      validationOf (DecimalRange (None, Some 2.5))
      validationOf (TextLength (Some 3.0, Some 3.0))
      validationOf (DateRange (Some 45000.0, Some 45100.0))
      validationOf (CustomFormula "ISNUMBER(A1)")
      { validationOf (Comparison (TextLengthKind, OpLessThan, "8", None)) with
          AllowBlank = false; ShowInput = false; InputTitle = Some "T"; InputMessage = Some "M"
          ShowError = false; Style = WarningStyle; ErrorTitle = Some "ET"; ErrorMessage = Some "EM" } ]
    @ [ for op in [ OpGreaterThan; OpLessThan; OpGreaterOrEqual; OpLessOrEqual; OpEqual; OpNotEqual; OpBetween; OpNotBetween ] do
          for kind in [ WholeKind; DecimalKind; DateKind; TimeKind; TextLengthKind ] ->
              validationOf (Comparison (kind, op, "$B$1", (if op = OpBetween || op = OpNotBetween then Some "$B$2" else None))) ]
for validation in kinds do
    eq (sprintf "as data and back: %A" validation.Rule) (roundTrip validation) validation
eq "describe a list" (let v = Ops.describeValidation (validationOf (ListValues [ "a"; "b" ])) in v.Type, v.Values, v.Source, v.Dropdown) ("list", Some [ "a"; "b" ], None, true)
eq "describe a source" (let v = Ops.describeValidation (validationOf (ListSource "=Names")) in v.Type, v.Values, v.Source) ("list", None, Some "Names")
eq "describe a range of numbers" (let v = Ops.describeValidation (whole 1.0 10.0) in v.Type, v.Operator, v.Formula1, v.Formula2) ("whole", Some "between", Some "1", Some "10")
eq "describe any" (Ops.describeValidation (validationOf AnyValue)).Type "any"
eq "describe a custom formula" (let v = Ops.describeValidation (validationOf (CustomFormula "=A1>0")) in v.Type, v.Formula1) ("custom", Some "A1>0")
eq "numbers give the simple rules" (build { d "decimal" with Operator = Some "between"; Formula1 = Some "1.5"; Formula2 = Some "9" }).Rule (DecimalRange (Some 1.5, Some 9.0))
eq "a leading = is dropped" (build { d "whole" with Operator = Some "greaterThan"; Formula1 = Some "=$C$1" }).Rule (Comparison (WholeKind, OpGreaterThan, "$C$1", None))
eq "a date as text" (build { d "date" with Operator = Some "greaterThanOrEqual"; Formula1 = Some "2024-01-01" }).Rule (DateRange (Some 45292.0, None))
eq "a date and a time" (build { d "date" with Operator = Some "lessThan"; Formula1 = Some "2024-01-01T12:00" }).Rule (Comparison (DateKind, OpLessThan, "45292.5", None))
eq "a time as text" (build { d "time" with Operator = Some "between"; Formula1 = Some "06:00"; Formula2 = Some "18:00:00" }).Rule (Comparison (TimeKind, OpBetween, "0.25", Some "0.75"))
eq "list values are trimmed" (build { d "list" with Values = Some [ " a "; ""; "b" ] }).Rule (ListValues [ "a"; "b" ])
eq "formula2 is dropped where it does not count" (build { d "whole" with Operator = Some "notEqual"; Formula1 = Some "$A$1"; Formula2 = Some "9" }).Rule (Comparison (WholeKind, OpNotEqual, "$A$1", None))
throws "unknown type" (fun () -> build (d "sparkle") |> ignore) "Set: unknown validation type \"sparkle\""
throws "no operator" (fun () -> build { d "whole" with Formula1 = Some "1" } |> ignore) "Set: a \"whole\" rule needs an operator"
throws "bad operator" (fun () -> build { d "whole" with Operator = Some ">"; Formula1 = Some "1" } |> ignore) "Set: unknown operator \">\""
throws "no operand" (fun () -> build { d "decimal" with Operator = Some "equal" } |> ignore) "Set: a \"decimal\" rule needs formula1"
throws "between needs two" (fun () -> build { d "date" with Operator = Some "between"; Formula1 = Some "1" } |> ignore) "Set: a \"date\" rule needs formula2"
throws "list without values" (fun () -> build (d "list") |> ignore) "Set: a list rule needs values"
throws "list with no value left" (fun () -> build { d "list" with Values = Some [ " " ] } |> ignore) "Set: a list rule needs at least one value"
throws "list with both" (fun () -> build { d "list" with Values = Some [ "a" ]; Source = Some "A1:A3" } |> ignore) "Set: a list rule takes either values or a source"
throws "a comma in a value" (fun () -> build { d "list" with Values = Some [ "a,b" ] } |> ignore) "Set: the list value \"a,b\" holds a comma"
throws "custom without formula" (fun () -> build (d "custom") |> ignore) "Set: a \"custom\" rule needs a formula"
throws "bad error style" (fun () -> build { d "any" with ErrorStyle = "loud" } |> ignore) "Set: unknown errorStyle \"loud\""

// ---- lists read from cells ---------------------------------------------------------
let fruit = { empty with Cells = Map.ofList [ at 0 7, str "Apple"; at 1 7, str "Pear"; at 3 7, num 7.0; at 0 8, str "Red"; at 1 8, str "Green" ] }
let other = { empty with Cells = Map.ofList [ at 0 0, str "North"; at 1 0, str "South" ] }
let reader =
    Ops.listSourceReader
        (fun name -> if name.ToLowerInvariant() = "other" then Some other elif name = "My sheet" then Some fruit else None)
        (fun name -> match name.ToUpperInvariant() with "REGIONS" -> Some "=Other!$A$1:$A$5" | "FRUIT" -> Some "$H$1:$H$9" | "LOOP" -> Some "Loop" | _ -> None)
let listed source addr = fruit |> setValidationRanges [ block 0 0 9 1 ] (validationOf (ListSource source)) |> Ops.listValuesAt reader addr
eq "a range of the sheet (blanks left out, numbers as texts)" (listed "$H$1:$H$9" (at 0 0)) (Some [ "Apple"; "Pear"; "7" ])
eq "a whole column" (listed "$H:$H" (at 0 0)) (Some [ "Apple"; "Pear"; "7" ])
eq "a relative range moves with the cell" (listed "H$1:H$9" (at 4 1)) (Some [ "Red"; "Green" ])
eq "another sheet" (listed "Other!$A$1:$A$5" (at 0 0)) (Some [ "North"; "South" ])
eq "a quoted sheet" (listed "'My sheet'!$I$1:$I$2" (at 0 0)) (Some [ "Red"; "Green" ])
eq "a name" (listed "Regions" (at 0 0), listed "fruit" (at 0 0)) (Some [ "North"; "South" ], Some [ "Apple"; "Pear"; "7" ])
eq "a name that leads nowhere, an unknown sheet" (listed "Loop" (at 0 0), listed "Gone!A1:A2" (at 0 0), listed "SUM(1)" (at 0 0)) (None, None, None)
eq "a written-out list" (Ops.listValuesAt reader (at 3 0) two) (Some [ "a"; "b" ])
eq "no list on the cell" (Ops.listValuesAt reader (at 9 0) two, Ops.listValuesAt reader (at 9 9) two) (None, None)
let sourced = fruit |> setValidationRanges [ block 0 0 9 0 ] (validationOf (ListSource "$H$1:$H$9"))
eq "an entry is checked against the cells" (invalid (checkEntry (at 2 0) (CellValue.String "pear") sourced), invalid (checkEntry (at 2 0) (CellValue.String "Plum") sourced), invalid (checkEntry (at 2 0) (CellValue.Number 7.0) sourced)) (false, true, false)
let elsewhere = fruit |> setValidationRanges [ block 0 0 9 0 ] (validationOf (ListSource "Regions"))
eq "with the workbook's reader: a name" (invalid (checkEntryIn reader evaluateFormulaAt (at 2 0) (CellValue.String "South") elsewhere), invalid (checkEntryIn reader evaluateFormulaAt (at 2 0) (CellValue.String "East") elsewhere)) (false, true)
eq "a source that cannot be read locks nothing" (invalid (checkEntry (at 2 0) (CellValue.String "East") elsewhere)) false

// ---- the invalid cells of a sheet ----------------------------------------------------
let filled =
    { empty with Cells = Map.ofList [ at 0 0, num 5.0; at 1 0, num 50.0; at 2 0, str "x"; at 3 0, num 2.5; at 4 0, { Value = CellValue.Empty; Format = defaultFormat; Comment = None }
                                      at 0 1, str "Pear"; at 1 1, str "Plum"; at 0 2, num 1.0; at 1 2, num -1.0; at 5 5, num 99.0 ] }
    |> setValidationRanges [ column 0 ] { whole 1.0 10.0 with ErrorMessage = Some "1 to 10" }
    |> setValidationRanges [ block 0 1 9 1 ] { validationOf (ListValues [ "Apple"; "Pear" ]) with ShowError = false }
    |> setValidationRanges [ block 0 2 9 2 ] { validationOf (CustomFormula "=C1>0") with Style = WarningStyle }
let found = invalidCells sameSheetListValues evaluateFormulaAt filled
eq "the invalid cells, in row order" (found |> List.map fst) [ at 1 0; at 1 1; at 1 2; at 2 0; at 3 0 ]
eq "each with its alert" (found |> List.map snd |> List.take 3)
   [ EntryInvalid (StopStyle, None, "1 to 10"); EntryInvalid (StopStyle, None, "Value must be one of: Apple, Pear"); EntryInvalid (WarningStyle, None, "This value does not meet the validation rule set for this cell.") ]
eq "an entry is not refused where the alert is off" (invalid (checkEntry (at 1 1) (CellValue.String "Plum") filled)) false
eq "a sheet without validations" (invalidCells sameSheetListValues evaluateFormulaAt empty) []

// ---- many validations: a lookup does not walk a list per cell -------------------------------
let many =
    [ for i in 0 .. 1999 -> { Id = i + 1; Ranges = [ block i 3 i 3 ]; Validation = whole 1.0 (float (i + 2)) } ]
    |> fun validations -> { empty with Validations = validations }
let clock = System.Diagnostics.Stopwatch.StartNew()
let mutable hits = 0
for _ in 1 .. 50 do
    for row in 0 .. 1999 do
        if (validationEntryAt (at row 3) many).IsSome then hits <- hits + 1
        if (validationEntryAt (at row 9) many).IsSome then hits <- hits + 1
eq "every lookup found its validation" hits 100000
eq "200 000 lookups among 2 000 validations take well under a second" (clock.ElapsedMilliseconds < 3000L) true

printfn "stage35-validations: %d passed, %d failed" pass fail
if fail > 0 then exit 1
