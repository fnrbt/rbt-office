#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// Data validation: rule enforcement (list / whole / decimal / textLength /
// date), plus setValidation flagging the dropdown.

open Rbt.Spreadsheet.Model

let mutable pass = 0
let mutable fail = 0
let check label (got: bool) (expected: bool) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-26s expected %b got %b" label expected got
let ok r = match r with Result.Ok _ -> true | _ -> false
let S = CellValue.String
let N = CellValue.Number

// ── list ──
let lst = ListValues [ "Apple"; "Banana" ]
check "list match"     (ok (validateValue lst (S "Apple"))) true
check "list case-insens" (ok (validateValue lst (S "banana"))) true
check "list reject"    (ok (validateValue lst (S "Cherry"))) false
check "list blank-ok"  (ok (validateValue lst CellValue.Empty)) true

// ── whole number 1..10 ──
let wn = WholeNumber (Some 1.0, Some 10.0)
check "whole ok"       (ok (validateValue wn (N 5.0))) true
check "whole decimal"  (ok (validateValue wn (N 5.5))) false
check "whole over"     (ok (validateValue wn (N 11.0))) false
check "whole under"    (ok (validateValue wn (N 0.0))) false
check "whole text"     (ok (validateValue wn (S "x"))) false

// ── decimal 0..1 ──
let dec = DecimalRange (Some 0.0, Some 1.0)
check "decimal ok"     (ok (validateValue dec (N 0.5))) true
check "decimal over"   (ok (validateValue dec (N 2.0))) false

// ── text length <= 3 ──
let tl = TextLength (None, Some 3.0)
check "textlen ok"     (ok (validateValue tl (S "ab"))) true
check "textlen over"   (ok (validateValue tl (S "abcd"))) false

// ── date >= 43831 (2020-01-01) ──
let dr = DateRange (Some 43831.0, None)
check "date ok"        (ok (validateValue dr (N 43831.0))) true
check "date under"     (ok (validateValue dr (N 40000.0))) false

// ── setValidation flags the dropdown; validateEdit uses the rule + message ──
let model =
    { createEmpty 20 8 with Cells = Map.ofList [ { Row=0; Col=0 }, { Value = N 5.0; Format = defaultFormat; Comment = None } ] }
    |> setValidation { Row=0; Col=0 } { Rule = wn; ErrorMessage = Some "Only 1-10!"; InputMessage = None }
    |> setValidation { Row=1; Col=0 } { Rule = lst; ErrorMessage = None; InputMessage = None }
check "validateEdit ok"    (ok (validateEdit { Row=0; Col=0 } (N 7.0) model)) true
check "validateEdit reject" (ok (validateEdit { Row=0; Col=0 } (N 99.0) model)) false
check "validateEdit message" (match validateEdit { Row=0; Col=0 } (N 99.0) model with Result.Error m -> m = "Only 1-10!" | _ -> false) true
check "list sets dropdown" (match getCell { Row=1; Col=0 } model with Some cd -> cd.Format.HasListDropdown | None -> false) true
check "no rule → ok"       (ok (validateEdit { Row=5; Col=5 } (S "anything") model)) true

printfn ""
printfn "Stage 15 validation tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
