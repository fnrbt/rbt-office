#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Workbook
open Rbt.Spreadsheet.Calc
open Rbt.Spreadsheet.Formula

// Parser coverage (whole-column refs, empty arguments, array constants),
// reference-returning functions, dependency edges, and Excel's coercion rules.
//
//   A1..A5 = 10,20,30,40,50      B1..B5 = 1,2,3,4,5
//   C1..C5 = 1,5,2,8,4           D1..D5 = apple,fig,kiwi,fig,pear
//   E1..E5 = 1,"2",TRUE,blank,0  G1..G3 = 2,blank,4
let cell v = { Value = v; Format = defaultFormat; Comment = None }
let n x = CellValue.Number x
let s x = CellValue.String x
let b x = CellValue.Boolean x
let e x = CellValue.Error x
let at r c = { Row = r; Col = c }
let column c values = values |> List.mapi (fun r v -> at r c, cell v)
let cells =
    column 0 [ n 10.0; n 20.0; n 30.0; n 40.0; n 50.0 ]
    @ column 1 [ n 1.0; n 2.0; n 3.0; n 4.0; n 5.0 ]
    @ column 2 [ n 1.0; n 5.0; n 2.0; n 8.0; n 4.0 ]
    @ column 3 [ s "apple"; s "fig"; s "kiwi"; s "fig"; s "pear" ]
    @ [ at 0 4, cell (n 1.0); at 1 4, cell (s "2"); at 2 4, cell (b true); at 4 4, cell (n 0.0) ]
    @ [ at 0 6, cell (n 2.0); at 2 6, cell (n 4.0) ]
    |> Map.ofList
let model = { createEmpty 100 26 with Cells = cells }

let mutable pass = 0
let mutable fail = 0
let check label formula (expected: CellValue) =
    let got = evaluateFormula formula model None
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-24s  %-40s  expected %A  got %A" label formula expected got
let checkApprox label formula (expected: float) =
    match evaluateFormula formula model None with
    | CellValue.Number got when abs (got - expected) < 1e-9 -> pass <- pass + 1
    | got ->
        fail <- fail + 1
        printfn "FAIL %-24s  %-40s  expected ~%g  got %A" label formula expected got
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-24s  %s" label detail)

// ----- item 6: parser ---------------------------------------------------------
check "whole-col"         "=SUM(A:A)"              (n 150.0)
check "whole-col-abs"     "=SUM($A:$B)"            (n 165.0)
check "whole-col-span"    "=COUNT(A:C)"            (n 15.0)
check "whole-row"         "=SUM(1:1)"              (n 15.0)
check "whole-row-abs"     "=SUM($1:$2)"            (n 42.0)
check "whole-col-lookup"  "=VLOOKUP(3,B:C,2,FALSE)" (n 2.0)
check "whole-col-match"   "=MATCH(30,A:A,0)"       (n 3.0)
check "whole-col-countif" "=COUNTIF(D:D,\"fig\")"  (n 2.0)
check "whole-col-array"   "=SUM(A:A*2)"            (n 300.0)
check "empty-arg-if"      "=IF(1,,2)"              (n 0.0)
check "empty-arg-if-else" "=IF(0,1,)"              (n 0.0)
check "empty-arg-offset"  "=OFFSET(A1,,1)"         (n 1.0)
check "empty-arg-xlookup" "=XLOOKUP(99,B1:B5,A1:A5,,-1)" (n 50.0)
check "empty-arg-xlookup-na" "=XLOOKUP(99,B1:B5,A1:A5,,0)" (e "N/A")
check "empty-arg-sum"     "=SUM(1,,2)"             (n 3.0)
check "array-index"       "=INDEX({1,2,3},2)"      (n 2.0)
check "array-2d"          "=INDEX({1,2;3,4},2,1)"  (n 3.0)
check "array-sum"         "=SUM({1,2;3,4})"        (n 10.0)
check "array-strings"     "=INDEX({\"a\",\"b\"},2)" (s "b")
check "array-negative"    "=SUM({-1,2.5})"         (n 1.5)
check "array-scalar"      "={5,6}"                 (n 5.0)
check "array-vlookup-sum" "=SUM(A1:A2*{1;2})"      (n 50.0)
check "legacy-array-braces" "{=SUM(A1:A2*B1:B2)}"  (n 50.0)
check "lit-spill"         "=ISERROR(#SPILL!)"      (b true)
check "lit-calc"          "=IFERROR(#CALC!,7)"     (n 7.0)
check "unterminated-str"  "=\"abc"                 (e "NAME?")
check "parse-error-value" "=SUM(A1:A2"             (e "NAME?")
check "parse-error-ragged" "={1,2;3}"              (e "NAME?")
ok "parse-error-is-error"
   (match evaluateFormula "=1+*" model None with CellValue.Error code -> not (code.Contains " ") | _ -> false)
   "a parse error must be an error code, not a message"

// Whole-column references survive the printer and structure edits.
let printed f = match parseFormula f with Ok expr -> exprToFormula expr | Result.Error m -> "ERR " + m
ok "print-whole-col"  (printed "=SUM(A:A)+SUM($B:$C)" = "SUM(A:A)+SUM($B:$C)") (printed "=SUM(A:A)+SUM($B:$C)")
ok "print-whole-row"  (printed "=SUM(1:1)+SUM($2:$3)+Sheet2!4:5" = "SUM(1:1)+SUM($2:$3)+Sheet2!4:5") (printed "=SUM(1:1)+SUM($2:$3)+Sheet2!4:5")
ok "print-sheet-col"  (printed "='My Data'!A:B" = "'My Data'!A:B") (printed "='My Data'!A:B")
ok "print-empty-arg"  (printed "=IF(A1,,2)" = "IF(A1,,2)") (printed "=IF(A1,,2)")
ok "print-array"      (printed "={1,-2;\"a\",TRUE}" = "{1,-2;\"a\",TRUE}") (printed "={1,-2;\"a\",TRUE}")
ok "shift-whole-col"  (shiftFormulaText 3 1 "=SUM(A:A)" = "=SUM(B:B)") (shiftFormulaText 3 1 "=SUM(A:A)")
ok "shift-whole-row"  (shiftFormulaText 3 1 "=SUM(2:2)" = "=SUM(5:5)") (shiftFormulaText 3 1 "=SUM(2:2)")
ok "insert-row-whole-col" (adjustFormulaTextForEdit (InsertRows (1, 1)) Option.isNone "=SUM(A:A)" = "=SUM(A:A)") "whole column must not change on a row insert"
ok "delete-col-whole-col" (adjustFormulaTextForEdit (DeleteCols (0, 1)) Option.isNone "=SUM(B:B)" = "=SUM(A:A)") (adjustFormulaTextForEdit (DeleteCols (0, 1)) Option.isNone "=SUM(B:B)")

// A whole-column dependency is one rectangle, never a million cells.
let colDeps, colRects = referencesRanged None "sheet1" (match parseFormula "=SUM(A:A)" with Ok x -> x | _ -> failwith "parse")
ok "whole-col-dep-rect" (Set.isEmpty colDeps && List.length colRects = 1) (sprintf "%d cells, %d rects" (Set.count colDeps) (List.length colRects))

// ----- item 7: OFFSET / INDIRECT return references ---------------------------
check "offset-range-sum"    "=SUM(OFFSET(A1,0,0,2,1))"      (n 30.0)
check "offset-range-shift"  "=SUM(OFFSET(A1,1,1,3,1))"      (n 9.0)
check "offset-default-size" "=SUM(OFFSET(A1:A2,1,0))"       (n 50.0)
check "offset-negative"     "=SUM(OFFSET(A5,0,0,-2,1))"     (n 90.0)
check "offset-scalar"       "=OFFSET(A1,2,1)"               (n 3.0)
check "offset-off-sheet"    "=OFFSET(A1,-1,0)"              (e "REF!")
check "offset-count"        "=COUNTA(OFFSET(D1,0,0,5,1))"   (n 5.0)
check "offset-average"      "=AVERAGE(OFFSET(B1,0,0,COUNT(B:B),1))" (n 3.0)
check "offset-in-array"     "=SUM(OFFSET(A1,0,0,2,1)*2)"    (n 60.0)
check "indirect-range-sum"  "=SUM(INDIRECT(\"A1:A2\"))"     (n 30.0)
check "indirect-cell"       "=INDIRECT(\"B3\")"             (n 3.0)
check "indirect-built"      "=SUM(INDIRECT(\"A\"&2&\":A\"&4))" (n 90.0)
check "indirect-whole-col"  "=SUM(INDIRECT(\"B:B\"))"       (n 15.0)
check "indirect-bad"        "=INDIRECT(\"not a ref\")"      (e "REF!")
check "indirect-vlookup"    "=VLOOKUP(2,INDIRECT(\"B1:C5\"),2,FALSE)" (n 5.0)

// INDIRECT must not lag a recalculation behind the cell it points at.
let fml f = cell (CellValue.Formula (f, CellValue.Empty))
let valueOf r c (m: SpreadsheetModel) = getCellValue (at r c) m |> unwrapFormulaValue
let lagModel =
    { createEmpty 100 26 with
        Cells = Map.ofList [ at 0 0, fml "=INDIRECT(\"A3\")+1"; at 1 0, fml "=A1*2"; at 2 0, fml "=A4*10"; at 3 0, cell (n 5.0) ] }
let lag1 = recalcAndSpill None None lagModel
ok "indirect-no-lag"       (valueOf 0 0 lag1 = n 51.0) (sprintf "%A" (valueOf 0 0 lag1))
ok "indirect-dependents"   (valueOf 1 0 lag1 = n 102.0) (sprintf "%A" (valueOf 1 0 lag1))
let lag2 = lag1 |> setCell (at 3 0) (cell (n 7.0)) |> recalcAndSpill None None
ok "indirect-follows-edit" (valueOf 0 0 lag2 = n 71.0 && valueOf 1 0 lag2 = n 142.0) (sprintf "%A %A" (valueOf 0 0 lag2) (valueOf 1 0 lag2))

// Same through the incremental engine: INDIRECT and OFFSET cells are volatile.
let wbOf (m: SpreadsheetModel) : Workbook =
    { Sheets = Map.ofList [ "sheet1", { Id = "sheet1"; Name = "Sheet1"; Model = m; Visible = true } ]
      SheetOrder = [ "sheet1" ]; ActiveSheet = "sheet1"; Tables = Map.empty; Names = Map.empty; NameInfo = Map.empty }
let incWb0 = wbOf lag1
match trySingleCellEdit { Sheet = "sheet1"; Cell = at 3 0 } (cell (n 9.0)) incWb0 None with
| IncrementalApplied (wb, _) ->
    let m = wb.Sheets.["sheet1"].Model
    ok "indirect-incremental" (valueOf 0 0 m = n 91.0 && valueOf 1 0 m = n 182.0) (sprintf "%A %A" (valueOf 0 0 m) (valueOf 1 0 m))
| IncrementalFallback reason -> ok "indirect-incremental" false ("fell back: " + reason)

// Structured references have dependency edges: a calculated column feeding
// =SUM(T[Dbl]) is recalculated first, and an edit reaches the total.
let tableModel =
    { createEmpty 100 26 with
        Cells =
            Map.ofList
                [ at 0 0, fml "=SUM(T[Dbl])"
                  at 2 0, cell (s "Qty"); at 2 1, cell (s "Dbl")
                  at 3 0, cell (n 1.0);   at 3 1, fml "=[@Qty]*2"
                  at 4 0, cell (n 2.0);   at 4 1, fml "=[@Qty]*2" ] }
let tableWb =
    { wbOf tableModel with
        Tables = Map.ofList [ "T", { Name = "T"; SheetId = "sheet1"; StartRow = 2; StartCol = 0; EndRow = 4; EndCol = 1
                                     Columns = [ "Qty"; "Dbl" ]; HasHeaders = true; HasTotals = false } ] }
let table1 = recalcAndSpill (Some tableWb) (Some "sheet1") tableModel
ok "table-sum-ordered" (valueOf 0 0 table1 = n 6.0) (sprintf "%A" (valueOf 0 0 table1))
let tableWb1 = updateSheetModel "sheet1" table1 tableWb
let table2 = table1 |> setCell (at 3 0) (cell (n 10.0)) |> recalcAndSpill (Some tableWb1) (Some "sheet1")
ok "table-sum-after-edit" (valueOf 0 0 table2 = n 24.0) (sprintf "%A" (valueOf 0 0 table2))
match trySingleCellEdit { Sheet = "sheet1"; Cell = at 4 0 } (cell (n 5.0)) tableWb1 None with
| IncrementalApplied (wb, _) ->
    let m = wb.Sheets.["sheet1"].Model
    ok "table-sum-incremental" (valueOf 0 0 m = n 12.0) (sprintf "%A" (valueOf 0 0 m))
| IncrementalFallback reason -> ok "table-sum-incremental" false ("fell back: " + reason)

// ----- item 9: array expressions ---------------------------------------------
check "sumproduct-mask"     "=SUMPRODUCT((C1:C5>3)*C1:C5)"      (n 17.0)
check "sumproduct-dashdash" "=SUMPRODUCT(--(D1:D5=\"fig\"))"    (n 2.0)
check "sumproduct-pair"     "=SUMPRODUCT(A1:A5,B1:B5)"          (n 550.0)
check "sumproduct-aligned"  "=SUMPRODUCT(E1:E5,A1:A5)"          (n 10.0)
check "sumproduct-two-masks" "=SUMPRODUCT((D1:D5=\"fig\")*(C1:C5>6),A1:A5)" (n 40.0)
check "sumproduct-shape"    "=SUMPRODUCT(A1:A5,B1:B4)"          (e "VALUE!")
check "sumproduct-bool-range" "=SUMPRODUCT(C1:C5>3)"            (n 0.0)
check "if-lifted-sum"       "=SUM(IF(C1:C5>3,C1:C5,0))"         (n 17.0)
check "if-lifted-count"     "=COUNT(IF(D1:D5=\"fig\",A1:A5))"   (n 2.0)
check "if-lifted-max"       "=MAX(IF(D1:D5=\"fig\",A1:A5,0))"   (n 40.0)
check "len-lifted"          "=SUMPRODUCT(LEN(D1:D5))"           (n 19.0)
check "abs-lifted"          "=SUM(ABS(A1:A2-25))"               (n 20.0)
check "index-of-sort"       "=INDEX(SORT(D1:D5,1,-1),1)"        (s "pear")
check "index-of-sort-3"     "=INDEX(SORT(C1:C5),3)"             (n 4.0)
check "index-of-expr"       "=INDEX(A1:A5*2,2)"                 (n 40.0)
check "index-range-2d"      "=INDEX(A1:B5,2,2)"                 (n 2.0)
check "index-out-of-range"  "=INDEX(A1:A5,9)"                   (e "REF!")

// ----- item 8: *IF criteria do not coerce blanks or booleans ------------------
check "countif-lt-numbers-only" "=COUNTIF(E1:E5,\"<5\")"   (n 2.0)
check "countif-zero"            "=COUNTIF(E1:E5,0)"         (n 1.0)
check "countif-one"             "=COUNTIF(E1:E5,1)"         (n 1.0)
check "countif-numeric-text"    "=COUNTIF(E1:E5,2)"         (n 1.0)
check "countif-eq-text-number"  "=COUNTIF(E1:E5,\"=2\")"   (n 1.0)
check "countif-true"            "=COUNTIF(E1:E5,TRUE)"      (n 1.0)
check "countif-ne-number"       "=COUNTIF(E1:E5,\"<>1\")"  (n 4.0)
check "countif-ge"              "=COUNTIF(E1:E5,\">=0\")"  (n 2.0)
check "countif-wildcard-text"   "=COUNTIF(D1:E5,\"*\")"    (n 6.0)
check "countif-blank"           "=COUNTIF(E1:E5,\"\")"     (n 1.0)
check "averageif-skips-blank"   "=AVERAGEIF(G1:G3,\"<5\")" (n 3.0)
check "sumif-lt"                "=SUMIF(E1:E5,\"<5\",A1:A5)" (n 60.0)
check "sumifs-zero"             "=SUMIFS(A1:A5,E1:E5,0)"    (n 50.0)
check "countifs-pair"           "=COUNTIFS(E1:E5,\"<5\",A1:A5,\">10\")" (n 1.0)
check "countif-text-order"      "=COUNTIF(D1:D5,\">=g\")"  (n 2.0)

// ----- item 10: aggregates coerce direct arguments, not range values ----------
check "sum-direct-coerced"   "=SUM(\"3\",TRUE)"         (n 4.0)
check "sum-direct-text"      "=SUM(1,\"a\")"            (e "VALUE!")
check "sum-range-uncoerced"  "=SUM(E1:E3)"              (n 1.0)
check "sum-cell-ref-text"    "=SUM(E2)"                 (n 0.0)
check "sum-array-booleans"   "=SUM({-1,2.5,TRUE})"      (n 1.5)
check "sum-computed-bool"    "=SUM(1=1,2)"              (n 3.0)
check "count-direct"         "=COUNT(1,\"2\",\"a\",TRUE)" (n 3.0)
check "count-range"          "=COUNT(E1:E5)"            (n 2.0)
check "average-direct"       "=AVERAGE(\"4\",TRUE,1)"   (n 2.0)
check "average-range"        "=AVERAGE(E1:E5)"          (n 0.5)
check "max-range-bool"       "=MAX(E1:E5)"              (n 1.0)
check "min-direct-text"      "=MIN(5,\"x\")"            (e "VALUE!")
check "product-direct"       "=PRODUCT(\"2\",3,TRUE)"   (n 6.0)
check "median-range"         "=MEDIAN(E1:E5)"           (n 0.5)
check "sum-mixed"            "=SUM(A1:A2,\"5\",E3)"     (n 35.0)

// ----- item 11 -----------------------------------------------------------------
check "vlookup-wildcard"   "=VLOOKUP(\"ki*\",D1:E5,1,FALSE)"  (s "kiwi")
check "vlookup-question"   "=VLOOKUP(\"f?g\",D1:D5,1,FALSE)"  (s "fig")
check "vlookup-tilde"      "=VLOOKUP(\"a~*\",D1:D5,1,FALSE)"  (e "N/A")
check "hlookup-wildcard"   "=HLOOKUP(\"app*\",D1:E2,2,FALSE)" (s "fig")
check "match-wildcard"     "=MATCH(\"p*\",D1:D5,0)"           (n 5.0)
check "match-wildcard-mid" "=MATCH(\"*iw*\",D1:D5,0)"         (n 3.0)
check "match-exact-still"  "=MATCH(\"fig\",D1:D5,0)"          (n 2.0)
check "search-wildcard"    "=SEARCH(\"b*d\",\"abcde\")"      (n 2.0)
check "search-question"    "=SEARCH(\"c?e\",\"abcde\")"      (n 3.0)
check "search-tilde"       "=SEARCH(\"~?\",\"is it?\")"      (n 6.0)
check "search-plain"       "=SEARCH(\"CD\",\"abcde\")"       (n 3.0)
check "search-from"        "=SEARCH(\"a*\",\"banana\",3)"    (n 4.0)
check "weekday-default"    "=WEEKDAY(45292)"                  (n 2.0)   // 2024-01-01, a Monday
check "weekday-type2"      "=WEEKDAY(45292,2)"                (n 1.0)
check "weekday-type3"      "=WEEKDAY(45292,3)"                (n 0.0)
check "weekday-type11"     "=WEEKDAY(45292,11)"               (n 1.0)
check "weekday-type16"     "=WEEKDAY(45292,16)"               (n 3.0)   // week starts Saturday
check "weekday-type17"     "=WEEKDAY(45298,17)"               (n 1.0)   // 2024-01-07, a Sunday
check "weekday-sunday-2"   "=WEEKDAY(45298,2)"                (n 7.0)
check "weekday-bad-type"   "=WEEKDAY(45292,5)"                (e "NUM!")
check "overflow-mul"       "=1E308*10"                        (e "NUM!")
check "overflow-add"       "=1E308+1E308"                     (e "NUM!")
check "overflow-exp"       "=EXP(1000)"                       (e "NUM!")
check "exp-ok"             "=ROUND(EXP(1),4)"                 (n 2.7183)
check "textjoin-ignore"    "=TEXTJOIN(\",\",TRUE,\"a\",\"\",\"b\",G1:G3)" (s "a,b,2,4")
check "textjoin-keep"      "=TEXTJOIN(\",\",FALSE,\"a\",\"\",\"b\")"      (s "a,,b")
check "mid-negative"       "=MID(\"abc\",1,-1)"               (e "VALUE!")
check "left-negative"      "=LEFT(\"abc\",-1)"                (e "VALUE!")
check "right-negative"     "=RIGHT(\"abc\",-1)"               (e "VALUE!")
check "rept-negative"      "=REPT(\"a\",-1)"                  (e "VALUE!")
check "left-zero"          "=LEFT(\"abc\",0)"                 (s "")
check "blank-eq-false"     "=E4=FALSE"                        (b true)
check "blank-eq-zero"      "=E4=0"                            (b true)
check "blank-eq-empty"     "=E4=\"\""                         (b true)
check "blank-ne-true"      "=E4=TRUE"                         (b false)
check "false-eq-blank"     "=FALSE=E4"                        (b true)

printfn ""
printfn "Stage 22 semantics tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
