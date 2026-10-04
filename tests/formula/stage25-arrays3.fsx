#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Workbook
open Rbt.Spreadsheet.Formula

// Array semantics, part 3: scalar functions and operators lift over arrays by
// default (driven by the per-function parameter signatures), lookups accept
// arrays as their table, multi-cell references spill, the reference operators
// (range, intersection, union, 3-D, sheet-qualified spill, implicit
// intersection), INDIRECT with R1C1 text, strict lexing, defined-name cycles.

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label detail)
let eq label (got: 'a) (expected: 'a) =
    ok label (got = expected) (sprintf "expected %A  got %A" expected got)

let cell v = { Value = v; Format = defaultFormat; Comment = None }
let num x = cell (CellValue.Number x)
let str x = cell (CellValue.String x)
let fml f = cell (CellValue.Formula (f, CellValue.Empty))
let at r c = { Row = r; Col = c }
let N x = CellValue.Number x
let S x = CellValue.String x
let B x = CellValue.Boolean x
let E x = CellValue.Error x
let modelOf cells = { createEmpty 100 26 with Cells = Map.ofList cells }
let valueAt r c (m: SpreadsheetModel) = getCellValue (at r c) m |> unwrapFormulaValue

// A1:A5 = 3,1,4,1,5   B1:B5 = "c","a","d","a","e"   C1:C3 = 10,20,30   D1 = "apple pie"
let model =
    modelOf [ at 0 0, num 3.0; at 1 0, num 1.0; at 2 0, num 4.0; at 3 0, num 1.0; at 4 0, num 5.0
              at 0 1, str "c"; at 1 1, str "a"; at 2 1, str "d"; at 3 1, str "a"; at 4 1, str "e"
              at 0 2, num 10.0; at 1 2, num 20.0; at 2 2, num 30.0
              at 0 3, str "apple pie" ]
let check label formula expected = eq label (evaluateFormula formula model None) expected

/// Put `formula` in F1 of the model, recalculate, and read the cells it filled.
let spillOf formula =
    let m = model |> setCell (at 0 5) (fml formula) |> recalcAndSpill None None
    [ for r in 0 .. 5 do
        let row = [ for c in 5 .. 8 do match valueAt r c m with CellValue.Empty -> () | v -> yield v ]
        if not row.IsEmpty then yield row ]

// --- 4. Lifting is the default for scalar parameters ------------------------
check "lift-sum-len"        "=SUM(LEN(B1:B5))"                 (N 5.0)
check "lift-date-parts"     "=SUM(DATE(2024,A1:A3,1)-DATE(2024,1,1))" (N (60.0 + 0.0 + 91.0))
check "lift-newly-covered"  "=SUM(CEILING(A1:A5/2,1))"         (N 9.0)
check "lift-two-arrays"     "=SUM(MAX(A1,0)*POWER(A1:A3,C1:C3/10))" (N (3.0 * (3.0 + 1.0 + 64.0)))
check "lift-concat"         "=TEXTJOIN(\"\",TRUE,REPT(B1:B3,A1:A3))" (S "cccadddd")
check "lift-nested"         "=SUM(ROUND(SQRT(A1:A5),0))"       (N 8.0)
check "lift-choose-index"   "=SUM(CHOOSE({1,2},10,20))"        (N 30.0)
check "lift-if-lazy"        "=SUM(IF(A1:A5>2,A1:A5,0))"        (N 12.0)
check "lift-iferror"        "=SUM(IFERROR(1/(A1:A5-1),0))"     (N (0.5 + 1.0 / 3.0 + 0.25))
check "lift-broadcast-col-row" "=SUM(A1:A3*{1,10})"            (N 88.0)
check "lift-mismatch-na"    "=COUNT(IFERROR(A1:A3+{1;2},\"x\"))" (N 2.0)
check "lift-text-fn"        "=CONCAT(UPPER(B1:B3))"            (S "CAD")
check "lift-constant"       "=SUM(ABS({-1,-2;3,-4}))"          (N 10.0)
// Range parameters are still taken whole, never lifted.
check "whole-countif"       "=SUM(COUNTIF(A1:A5,{1,5}))"       (N 3.0)
check "whole-rank"          "=SUM(RANK(A1:A3,A1:A5))"          (N (3.0 + 4.0 + 2.0))
check "whole-vlookup-key"   "=SUM(VLOOKUP({1,4},A1:C3,3,FALSE))" (N 50.0)
check "whole-correl"        "=ROUND(CORREL(A1:A3,C1:C3),4)"    (N 0.3273)
check "whole-large"         "=SUM(LARGE(A1:A5,{1,2}))"         (N 9.0)
check "whole-index-scalar"  "=INDEX(A1:A5,3)*2"                (N 8.0)
check "whole-sumproduct"    "=SUMPRODUCT((A1:A5>2)*A1:A5)"     (N 12.0)
// A user-defined lambda is never lifted: it gets its argument as written.
check "lambda-not-lifted"   "=LAMBDA(r,SUM(r))(A1:A5)"         (N 14.0)
check "let-not-lifted"      "=LET(r,A1:A5,SUM(r)*2)"           (N 28.0)

eq "sig-scalar"   (functionSignature "len" |> ValueOption.map (fun s -> paramKindAt s 0)) (ValueSome ScalarParam)
eq "sig-range"    (functionSignature "SUM" |> ValueOption.map (fun s -> paramKindAt s 7)) (ValueSome RangeParam)
eq "sig-cycle"    (functionSignature "SUMIFS" |> ValueOption.map (fun s -> [ for i in 0 .. 4 -> paramKindAt s i ]))
                  (ValueSome [ RangeParam; RangeParam; ScalarParam; RangeParam; ScalarParam ])
eq "sig-array"    (functionSignature "VLOOKUP" |> ValueOption.map (fun s -> paramKindAt s 1)) (ValueSome ArrayParam)
eq "sig-unknown"  (functionSignature "LET") ValueNone

eq "lift-spills"        (spillOf "=LEN(B1:B3)&\"!\"")   [ [ S "1!" ]; [ S "1!" ]; [ S "1!" ] ]
eq "lift-spills-2d"     (spillOf "=A1:A2*{1,10}")       [ [ N 3.0; N 30.0 ]; [ N 1.0; N 10.0 ] ]
eq "lift-fn-spills"     (spillOf "=ROUND(A1:A3/2,0)")   [ [ N 2.0 ]; [ N 1.0 ]; [ N 2.0 ] ]
eq "aggregate-no-spill" (spillOf "=SUM(A1:A5)")         [ [ N 14.0 ] ]

// --- 5. Lookups take an array constant or an array result -------------------
check "vlookup-array-const" "=VLOOKUP(2,{1,\"a\";2,\"b\"},2,0)"     (S "b")
check "vlookup-array-approx" "=VLOOKUP(2.5,{1,\"a\";2,\"b\";3,\"c\"},2)" (S "b")
check "hlookup-array-const" "=HLOOKUP(\"y\",{\"x\",\"y\";1,2},2,FALSE)" (N 2.0)
check "index-sort"          "=INDEX(SORT(A1:A5),1)"                  (N 1.0)
check "index-sort-last"     "=INDEX(SORT(A1:A5),5)"                  (N 5.0)
check "match-computed"      "=MATCH(8,A1:A5*2,0)"                    (N 3.0)
check "match-array-const"   "=MATCH(\"b\",{\"a\",\"b\",\"c\"},0)"    (N 2.0)
check "match-missing"       "=MATCH(99,A1:A5*2,0)"                   (E "N/A")
check "xlookup-arrays"      "=XLOOKUP(4,A1:A5,UPPER(B1:B5))"         (S "D")
check "xlookup-both-arrays" "=XLOOKUP(10,A1:A5*2,B1:B5&\"!\")"       (S "e!")
check "xmatch-array"        "=XMATCH(5,SORT(A1:A5))"                 (N 5.0)
check "lookup-array"        "=LOOKUP(2,{1,2,3},{\"a\",\"b\",\"c\"})" (S "b")
check "vlookup-hstack"      "=VLOOKUP(4,HSTACK(A1:A5,B1:B5),2,FALSE)" (S "d")
check "vlookup-range-still" "=VLOOKUP(4,A1:B5,2,FALSE)"              (S "d")
check "index-array-row"     "=SUM(INDEX({1,2;3,4},2,0))"             (N 7.0)
check "index-array-col"     "=SUM(INDEX({1,2;3,4},0,2))"             (N 6.0)

// --- 6. References that are computed ---------------------------------------
eq "offset-spills"       (spillOf "=OFFSET(A1,1,0,3,1)")      [ [ N 1.0 ]; [ N 4.0 ]; [ N 1.0 ] ]
eq "indirect-spills"     (spillOf "=INDIRECT(\"C1:C3\")")     [ [ N 10.0 ]; [ N 20.0 ]; [ N 30.0 ] ]
eq "range-spills"        (spillOf "=A1:B2")                   [ [ N 3.0; S "c" ]; [ N 1.0; S "a" ] ]
eq "index-col-spills"    (spillOf "=INDEX(A1:C3,0,3)")        [ [ N 10.0 ]; [ N 20.0 ]; [ N 30.0 ] ]
eq "index-row-spills"    (spillOf "=INDEX(A1:C3,2,0)")        [ [ N 1.0; S "a"; N 20.0 ] ]
eq "index-zero-spills"   (spillOf "=INDEX(C1:C3,0)")          [ [ N 10.0 ]; [ N 20.0 ]; [ N 30.0 ] ]
eq "offset-single-cell"  (spillOf "=OFFSET(A1,2,0)")          [ [ N 4.0 ] ]
eq "array-const-spills"  (spillOf "={1,2;3,4}")               [ [ N 1.0; N 2.0 ]; [ N 3.0; N 4.0 ] ]
// A spill that would land on an occupied cell is blocked.
let blocked = model |> setCell (at 1 5) (num 99.0) |> setCell (at 0 5) (fml "=OFFSET(A1,0,0,3,1)") |> recalcAndSpill None None
eq "offset-spill-blocked" (valueAt 0 5 blocked, valueAt 1 5 blocked) (E "SPILL!", N 99.0)
// A whole-column reference in a scalar position is implicit intersection.
let intersected = model |> setCell (at 2 5) (fml "=A:A*2") |> setCell (at 3 5) (fml "=LEN(B:B)") |> recalcAndSpill None None
eq "implicit-intersection" (valueAt 2 5 intersected, valueAt 3 5 intersected, valueAt 4 5 intersected) (N 8.0, N 1.0, CellValue.Empty)
let single = model |> setCell (at 4 5) (fml "=@A1:A5") |> setCell (at 0 6) (fml "=@SORT(A1:A5)") |> recalcAndSpill None None
eq "at-operator"           (valueAt 4 5 single, valueAt 0 6 single, valueAt 1 6 single) (N 5.0, N 1.0, CellValue.Empty)

check "indirect-r1c1"       "=INDIRECT(\"R3C1\",FALSE)"            (N 4.0)
check "indirect-r1c1-range" "=SUM(INDIRECT(\"R1C1:R3C1\",FALSE))"  (N 8.0)
check "indirect-r1c1-col"   "=SUM(INDIRECT(\"C3\",FALSE))"         (N 60.0)
check "indirect-r1c1-row"   "=COUNTA(INDIRECT(\"R1\",FALSE))"      (N 4.0)
check "indirect-a1-true"    "=INDIRECT(\"A3\",TRUE)"               (N 4.0)
check "indirect-r1c1-bad"   "=INDIRECT(\"A3\",FALSE)"              (E "REF!")
let relative = model |> setCell (at 2 5) (fml "=INDIRECT(\"R[-1]C[-5]\",FALSE)+INDIRECT(\"RC3\",FALSE)") |> recalcAndSpill None None
eq "indirect-r1c1-relative" (valueAt 2 5 relative) (N 31.0)

check "index-as-reference"  "=SUM(A1:INDEX(A:A,3))"                (N 8.0)
check "index-ref-both-ends" "=SUM(INDEX(A1:A5,2):INDEX(A1:A5,4))"  (N 6.0)
check "index-whole-col"     "=SUM(INDEX(A1:C3,0,3))"               (N 60.0)
check "index-whole-row"     "=COUNTA(INDEX(A1:C3,2,0))"            (N 3.0)
check "index-ref-rows"      "=ROWS(INDEX(A1:C5,0,1))"              (N 5.0)
check "range-of-offsets"    "=SUM(OFFSET(A1,0,0):OFFSET(A1,2,0))"  (N 8.0)
check "intersection"        "=SUM(A1:C2 B1:C5)"                    (N 30.0)
check "intersection-cell"   "=A1:A5 A3:C3"                         (N 4.0)
check "intersection-empty"  "=A1:A2 C1:C2"                         (E "NULL!")
check "union-sum"           "=SUM((A1:A2,C1:C2))"                  (N 34.0)
check "union-count"         "=COUNT((A1:A5,C1:C3,A1))"             (N 9.0)
check "union-max"           "=MAX((A1,C3,A5))"                     (N 30.0)
check "args-not-union"      "=SUM(A1,C3)"                          (N 33.0)
check "paren-scalar"        "=(A1+C1)*2"                           (N 26.0)
check "spaces-are-layout"   "= SUM( A1 : A2 , C1 ) + 1"            (N 15.0)

let printed f = match parseFormula f with Ok e -> exprToFormula e | Result.Error m -> "ERR " + m
eq "print-intersection" (printed "=SUM(A1:C2 B1:C5)")              "SUM(A1:C2 B1:C5)"
eq "print-union"        (printed "=SUM((A1:A2,C1:C2))")            "SUM((A1:A2,C1:C2))"
eq "print-range-op"     (printed "=SUM(A1:INDEX(A:A,3))")          "SUM(A1:INDEX(A:A,3))"
eq "print-3d"           (printed "=SUM(Jan:Mar!B2)")               "SUM('Jan:Mar'!B2)"
eq "print-3d-quoted"    (printed "=SUM('Jan:Mar'!B2:C3)")          "SUM('Jan:Mar'!B2:C3)"
eq "print-at"           (printed "=@A1:A5+1")                      "@A1:A5+1"
eq "shift-range-op"     (shiftFormulaText 1 0 "=SUM(A1:INDEX(B:B,3) C1:D4)") "=SUM(A2:INDEX(B:B,3) C2:D5)"

// 3-D references and sheet-qualified spill references need a workbook.
let sheetOf cells = { createEmpty 50 10 with Cells = Map.ofList cells }
let wb3 : Workbook =
    let sheets =
        [ "s1", "Jan", sheetOf [ at 0 0, num 1.0; at 1 0, num 2.0 ]
          "s2", "Feb", sheetOf [ at 0 0, num 10.0; at 1 0, num 20.0 ]
          "s3", "Mar", sheetOf [ at 0 0, num 100.0; at 1 0, num 200.0 ]
          "s4", "Sum", sheetOf [ at 0 0, fml "=SUM(Jan:Mar!A1)"; at 1 0, fml "=SUM(Jan:Mar!A1:A2)"
                                 at 2 0, fml "=COUNT('Jan:Mar'!A:A)"; at 3 0, fml "=MAX(Mar:Jan!A2)"
                                 at 4 0, fml "=AVERAGE(Feb:Mar!A1)"; at 5 0, fml "=SUM(Jan:Nope!A1)"
                                 at 0 2, fml "=SEQUENCE(3)"; at 0 4, fml "=SUM(Sum!C1#)" ]
          "s5", "Reader", sheetOf [ at 0 0, fml "=SUM(Sum!C1#)"; at 1 0, fml "=ROWS(Sum!C1#)" ] ]
    { Sheets = sheets |> List.map (fun (id, name, m) -> id, { Id = id; Name = name; Model = m; Visible = true }) |> Map.ofList
      SheetOrder = sheets |> List.map (fun (id, _, _) -> id)
      ActiveSheet = "s4"; Tables = Map.empty
      Names = namesOfTriples [ None, "Ping", "Pong+1"; None, "Pong", "Ping+1"; None, "Selfish", "Selfish*2"
                               None, "Three", "1+2"; None, "Chain", "Three*2" ] }
let wb3c = recalculateWorkbook wb3
let sumAt r c = valueAt r c wb3c.Sheets.["s4"].Model
eq "3d-cell"          (sumAt 0 0) (N 111.0)
eq "3d-range"         (sumAt 1 0) (N 333.0)
eq "3d-whole-column"  (sumAt 2 0) (N 6.0)
eq "3d-reversed"      (sumAt 3 0) (N 200.0)
eq "3d-partial"       (sumAt 4 0) (N 55.0)
eq "3d-unknown-sheet" (sumAt 5 0) (E "REF!")
eq "spill-ref-own-sheet"   (sumAt 0 4) (N 6.0)
eq "spill-ref-other-sheet" (valueAt 0 0 wb3c.Sheets.["s5"].Model, valueAt 1 0 wb3c.Sheets.["s5"].Model) (N 6.0, N 3.0)

// --- 7. The lexer rejects what is not formula syntax -----------------------
check "lex-unknown-char"     "=1 ? 2"            (E "NAME?")
check "lex-backslash"        "=A1\\2"            (E "NAME?")
check "lex-tilde"            "=~A1"              (E "NAME?")
check "lex-unterminated-str" "=\"abc"            (E "NAME?")
check "lex-unterminated-in-call" "=LEN(\"abc)"   (E "NAME?")
check "lex-unterminated-sheet" "='My Sheet!A1"   (E "NAME?")
check "lex-unterminated-bracket" "=Table1[Col"   (E "NAME?")
check "lex-valid-still"      "=\"a\"\"b\"&\"?\"" (S "a\"b?")
ok "lex-error-message"
   (match parseFormula "=1 ? 2" with Result.Error m -> m.Contains "Unexpected character" | Ok _ -> false) "no parse error"
ok "lex-unterminated-message"
   (match parseFormula "=\"abc" with Result.Error m -> m.Contains "Unterminated" | Ok _ -> false) "no parse error"
// Recalculation shows the error too, rather than evaluating a trimmed formula.
let strict = model |> setCell (at 0 5) (fml "=A1 § 2") |> recalcAndSpill None None
eq "lex-recalc-error" (valueAt 0 5 strict) (E "NAME?")

// --- 8. Defined names that refer to each other ------------------------------
let nameEval f = evaluateFormula f wb3c.Sheets.["s4"].Model (Some wb3c)
eq "name-cycle-two"   (nameEval "=Ping")        (E "CYCLE")
eq "name-cycle-other" (nameEval "=Pong*2")      (E "CYCLE")
eq "name-cycle-self"  (nameEval "=Selfish")     (E "CYCLE")
eq "name-cycle-in-fn" (nameEval "=SUM(Ping,1)") (E "CYCLE")
eq "name-cycle-iferror" (nameEval "=IFERROR(Ping,7)") (N 7.0)
eq "name-chain-ok"    (nameEval "=Chain")       (N 6.0)
eq "name-after-cycle" (nameEval "=Three+Chain") (N 9.0)
let cyc = recalculateWorkbook { wb3 with Sheets = wb3.Sheets |> Map.add "s4" { wb3.Sheets.["s4"] with Model = sheetOf [ at 0 0, fml "=Ping+1"; at 1 0, fml "=Three" ] } }
eq "name-cycle-recalc" (valueAt 0 0 cyc.Sheets.["s4"].Model, valueAt 1 0 cyc.Sheets.["s4"].Model) (E "CYCLE", N 3.0)

printfn ""
printfn "Stage 25 array/reference tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
