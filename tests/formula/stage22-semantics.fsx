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
// A rewritten formula keeps how it was written: only the references that moved are written anew
// (lyric-multifamily, poi-47804: brackets the tree does not need, a space before a bracket).
let insertRow = adjustFormulaTextForEdit (InsertRows (1, 1)) Option.isNone
let kept name formula want = ok name (insertRow formula = want) (insertRow formula)
kept "spelling-redundant-brackets" "=+E361/((1+Discount_Rate)^Q383)" "=+E362/((1+Discount_Rate)^Q384)"
kept "spelling-space-before-bracket" "=CELL(\"width\",C3 )" "=CELL(\"width\",C4 )"
kept "spelling-lower-case" "=sum( a2 : a5 )*2" "=sum( A3 : A6 )*2" // (a moved reference is written as Excel stores it)
kept "spelling-unmoved-reference" "=((A1))+B5" "=((A1))+B6"
let onMySheet = adjustFormulaTextForEdit (InsertRows (1, 1)) (fun sheet -> sheet = Some "My Sheet")
ok "spelling-quoted-sheet" (onMySheet "=('My Sheet'!B3 )+B3" = "=('My Sheet'!B4 )+B3") (onMySheet "=('My Sheet'!B3 )+B3")
kept "spelling-whole-column-untouched" "=SUM(A:A )" "=SUM(A:A )"
ok "spelling-deleted-ref" (adjustFormulaTextForEdit (DeleteRows (1, 1)) Option.isNone "=(A2)+A3" = "=(#REF!)+A2") (adjustFormulaTextForEdit (DeleteRows (1, 1)) Option.isNone "=(A2)+A3")
ok "spelling-deleted-range-falls-back" (adjustFormulaTextForEdit (DeleteRows (1, 2)) Option.isNone "=SUM((A2:A3))+A4" = "=SUM(#REF!)+A2") (adjustFormulaTextForEdit (DeleteRows (1, 2)) Option.isNone "=SUM((A2:A3))+A4")
ok "spelling-shift-copy" (shiftFormulaText 1 0 "=((A1 ))*2" = "=((A2 ))*2") (shiftFormulaText 1 0 "=((A1 ))*2")

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
      SheetOrder = [ "sheet1" ]; ActiveSheet = "sheet1"; Tables = Map.empty; Names = Map.empty; NameInfo = Map.empty; Iteration = None; External = Map.empty }
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

// ----- Excel 2019's results (tools/excel-parity/eval-cells.vbs) -------------------------------
// A formula whose value is an empty cell is 0, and what reads it sees 0, not an empty cell.
let blankModel =
    { createEmpty 100 26 with
        Cells = Map.ofList [ at 0 0, fml "=B9"; at 1 0, fml "=A1=\"\""; at 2 0, fml "=IF(TRUE,B9)"; at 3 0, fml "=A3&\"x\""
                             at 4 0, fml "=ISBLANK(A1)"; at 5 0, fml "=CONCATENATE(B9,\" \",A1)"; at 6 0, fml "=-MIN(0,5)" ] }
let blanks = recalcAndSpill None None blankModel
ok "formula-of-blank-is-0" ([ for r in 0 .. 6 -> valueOf r 0 blanks ] = [ n 0.0; b false; n 0.0; s "0x"; b false; s " 0"; n 0.0 ]) (sprintf "%A" [ for r in 0 .. 6 -> valueOf r 0 blanks ])
ok "no-negative-zero" (match valueOf 6 0 blanks with CellValue.Number z -> 1.0 / z > 0.0 | _ -> false) ""
check "blank-direct"        "=B9"                              (n 0.0)
// The end of a formula's addition: 7 units in the last place of the larger operand are 0, 8 are not;
// SUM's last addition the same, wherever the SUM is; not inside a function, not a deeper operation.
let units k = sprintf "(1.5+%d*2^-52)" k
let off k = n (1.5 - (1.5 + float k * 2.0 ** -52.0))
check "cancel-7-ulps"       ("=1.5-" + units 7)                (n 0.0)
check "cancel-8-ulps"       ("=1.5-" + units 8)                (off 8)
check "cancel-1.9-7-ulps"   "=1.9-(1.9+7*2^-52)"               (n 0.0)
check "cancel-0+a-b"        ("=0+1.5-" + units 3)              (n 0.0)
check "cancel-not-in-if"    ("=IF(TRUE,1.5-" + units 3 + ")") (off 3)
check "cancel-not-times"    ("=1*(1.5-" + units 3 + ")")      (off 3)
check "cancel-last-op-only" ("=1.5-" + units 3 + "+0")        (off 3)
// Parentheses around the whole formula are its last operation: no addition at its end.
check "cancel-not-wrapped"  ("=(1.5-" + units 3 + ")")        (off 3)
check "cancel-not-wrapped2" ("=((1.5-" + units 3 + "))")      (off 3)
check "cancel-inner-parens" ("=(1.5)-" + units 3)             (n 0.0)
check "cancel-wrapped-left" ("=(1.5-0)-" + units 3)           (n 0.0)
check "sum-wrapped-cancels" ("=(SUM(1.5,-" + units 3 + "))")  (n 0.0)
check "cancel-0.1+0.2-0.3"  "=0.1+0.2-0.3"                     (n 0.0)
check "wrapped-0.1+0.2-0.3" "=(0.1+0.2-0.3)"                   (n (0.1 + 0.2 - 0.3))
check "sum-cancels"         ("=SUM(1.5,-" + units 7 + ")")    (n 0.0)
check "sum-cancels-inside"  ("=1*SUM(1.5,-" + units 5 + ")")  (n 0.0)
check "sum-8-ulps"          ("=SUM(1.5,-" + units 8 + ")")    (off 8)
check "sum-last-addition"   ("=SUM(1.5,-" + units 3 + ",0)")  (off 3)
check "sum-cancels-last"    ("=SUM(0,1.5,-" + units 3 + ")")  (n 0.0)
check "cancel-real-sum"     "=SUM(1000000.1,234.567,-(1000000.1+234.567))" (n 0.0)
// SUMIF / SUMIFS: an error among the cells added is the result; one among those not added is not.
let sumifModel =
    { createEmpty 100 26 with
        Cells = Map.ofList [ at 0 0, cell (n 1.0); at 0 1, cell (n 2.0); at 0 2, cell (n 3.0)
                             at 1 0, cell (n 10.0); at 1 1, cell (e "DIV/0!"); at 1 2, cell (n 30.0) ] }
let sumif f = evaluateFormula f sumifModel None
ok "sumif-error-added"      (sumif "=SUMIF(A1:C1,2,A2:C2)" = e "DIV/0!") (sprintf "%A" (sumif "=SUMIF(A1:C1,2,A2:C2)"))
ok "sumif-error-not-added"  (sumif "=SUMIF(A1:C1,3,A2:C2)" = n 30.0 && sumif "=SUMIF(A1:C1,1,A2:C2)" = n 10.0) ""
ok "sumif-all"              (sumif "=SUMIF(A1:C1,\">0\",A2:C2)" = e "DIV/0!") ""
ok "sumifs-error-added"     (sumif "=SUMIFS(A2:C2,A1:C1,2)" = e "DIV/0!" && sumif "=SUMIFS(A2:C2,A1:C1,\">2\")" = n 30.0) ""
// What a deleted range leaves, on another sheet: #REF!.
check "sheet-ref-error"     "='Other Sheet'!#REF!"            (e "REF!")
check "sheet-ref-error-sum" "=SUM(Other!#REF!)"               (e "REF!")

// Comparisons (Excel 2019, eval-cells.vbs): the operators, SWITCH and the criteria of COUNTIF and
// its family compare numbers rounded to 15 significant digits (a half away from zero, from the exact
// binary value; no absolute tolerance); MATCH, the LOOKUPs, RANK, MODE, FREQUENCY, DELTA and GESTEP
// compare the binary values. A number written in a formula or typed, and text read as a number, keeps
// its first 15 significant digits, the rest cut.
check "cmp-0.1+0.2"         "=0.1+0.2=0.3"                     (b true)
check "cmp-0.1+0.2-gt"      "=0.1+0.2>0.3"                     (b false)
check "cmp-0.1+0.2-le"      "=0.1+0.2<=0.3"                    (b true)
check "cmp-0.1+0.2-ne"      "=0.3<>0.1+0.2"                    (b false)
check "cmp-22-ulps"         "=(1+22*2^-52)=1"                  (b true)
check "cmp-23-ulps"         "=(1+23*2^-52)=1"                  (b false)
check "cmp-23-ulps-gt"      "=(1+23*2^-52)>1"                  (b true)
check "cmp-one-ulp-apart"   "=(1+22*2^-52)=(1+23*2^-52)"       (b false)
check "cmp-23-45-ulps"      "=(1+23*2^-52)=(1+45*2^-52)"       (b true)
check "cmp-tie-up"          "=(1E15+5)=(1E15+10)"              (b true)
check "cmp-tie-not-even"    "=(1E15+25)=(1E15+30)"             (b true)
check "cmp-tie-not-down"    "=(1E15+25)=(1E15+20)"             (b false)
check "cmp-tie-half"        "=(1E14+2.5)=(1E14+3)"             (b true)
check "cmp-tie-negative"    "=-(1E15+5)=-(1E15+10)"            (b true)
check "cmp-straddle"        "=(1E15+4)=(1E15+6)"               (b false)
check "cmp-to-10"           "=(10-2*2^-49)=10"                 (b true)
check "cmp-not-10"          "=(10-3*2^-49)=10"                 (b false)
check "cmp-1e-16"           "=1E-16=0"                         (b false)
check "cmp-1e-16-gt"        "=1E-16>0"                         (b true)
check "cmp-residue"         "=1*(0.1+0.2-0.3)=0"               (b false)
check "cmp-tiny"            "=1E-300*(1+22*2^-52)=1E-300"      (b true)
check "cmp-tiny-23"         "=1E-300*(1+23*2^-52)=1E-300"      (b false)
check "cmp-if"              "=IF(0.1+0.2=0.3,1,0)"             (n 1.0)
check "cmp-switch"          "=SWITCH(0.1+0.2,0.3,1,0)"         (n 1.0)
check "literal-cut"         "=1.0000000000000099"              (n 1.0)
check "literal-cut-int"     "=123456789012345.6"               (n 123456789012345.0)
check "literal-cut-small"   "=0.1234567890123456789"           (n 0.123456789012345)
check "literal-cut-16"      "=1000000000000005"                (n 1e15)
check "literal-cut-20"      "=12345678901234567890"            (n 1.23456789012345e19)
check "literal-cut-9s"      "=9.9999999999999999"              (n 9.99999999999999)
check "literal-exp"         "=1.234567890123459E+300"          (n 1.23456789012345e300)
check "value-cut"           "=VALUE(\"1.0000000000000099\")"   (n 1.0)
check "coerce-cut"          "=--\"1000000000000005\""          (n 1e15)
check "numbervalue-cut"     "=NUMBERVALUE(\"1.0000000000000099\")" (n 1.0)
check "array-constant-cut"  "=MATCH(0.3,{0.1,0.30000000000000004,0.5},0)" (n 2.0)
//   D1:D3 = 0.1, 0.1+0.2, 0.5   E1:E3 = 1, 2, 3   H1:H3 = 0.1, 0.3, 0.5   J1:J2 = 1+45 ulp, 1
//   G1:G2 = "k", 0.3 (a database criteria range over K1:L4 = k/v, D, E)
let cmpModel =
    let col c values = values |> List.mapi (fun r v -> at r c, cell v)
    { createEmpty 100 26 with
        Cells =
            col 3 [ n 0.1; n (0.1 + 0.2); n 0.5 ] @ col 4 [ n 1.0; n 2.0; n 3.0 ] @ col 7 [ n 0.1; n 0.3; n 0.5 ]
            @ col 9 [ n (1.0 + 45.0 * 2.0 ** -52.0); n 1.0 ] @ col 6 [ s "k"; n 0.3 ]
            @ col 10 [ s "k"; n 0.1; n (0.1 + 0.2); n 0.5 ] @ col 11 [ s "v"; n 1.0; n 2.0; n 3.0 ]
            |> Map.ofList }
let checkCmp label formula (expected: CellValue) =
    let got = evaluateFormula formula cmpModel None
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-24s  %-40s  expected %A  got %A" label formula expected got
checkCmp "match-exact-binary"  "=MATCH(0.3,D1:D3,0)"           (e "N/A")
checkCmp "match-approx-binary" "=MATCH(0.3,D1:D3,1)"           (n 1.0)
checkCmp "match-desc-binary"   "=MATCH(D2,{0.5;0.3;0.1},-1)"   (n 1.0)
checkCmp "vlookup-exact"       "=VLOOKUP(0.3,D1:E3,2,FALSE)"   (e "N/A")
checkCmp "vlookup-approx"      "=VLOOKUP(0.3,D1:E3,2,TRUE)"    (n 1.0)
checkCmp "lookup-binary"       "=LOOKUP(0.3,D1:D3,E1:E3)"      (n 1.0)
checkCmp "countif-number"      "=COUNTIF(D1:D3,0.3)"           (n 1.0)
checkCmp "countif-eq-text"     "=COUNTIF(D1:D3,\"=0.3\")"      (n 1.0)
checkCmp "countif-text"        "=COUNTIF(D1:D3,\"0.3\")"       (n 1.0)
checkCmp "countif-le"          "=COUNTIF(D1:D3,\"<=0.3\")"     (n 2.0)
checkCmp "countif-gt"          "=COUNTIF(D1:D3,\">0.3\")"      (n 1.0)
checkCmp "countif-ne"          "=COUNTIF(D1:D3,\"<>0.3\")"     (n 2.0)
checkCmp "countif-lt"          "=COUNTIF(D1:D3,\"<0.3\")"      (n 1.0)
checkCmp "countif-ge"          "=COUNTIF(D1:D3,\">=0.3\")"     (n 2.0)
checkCmp "countif-cell"        "=COUNTIF(H1:H3,D2)"            (n 1.0)
checkCmp "countif-cut"         "=COUNTIF(J1:J2,\"<1.0000000000000099\")" (n 0.0)
checkCmp "sumif"               "=SUMIF(D1:D3,0.3,E1:E3)"       (n 2.0)
checkCmp "sumifs-gt"           "=SUMIFS(E1:E3,D1:D3,\">0.3\")" (n 3.0)
checkCmp "countifs-le"         "=COUNTIFS(D1:D3,\"<=0.3\")"    (n 2.0)
checkCmp "averageif"           "=AVERAGEIF(D1:D3,0.3,E1:E3)"   (n 2.0)
checkCmp "maxifs"              "=MAXIFS(E1:E3,D1:D3,0.3)"      (n 2.0)
checkCmp "minifs-gt"           "=MINIFS(E1:E3,D1:D3,\">0.3\")" (n 3.0)
checkCmp "dsum"                "=DSUM(K1:L4,\"v\",G1:G2)"      (n 2.0)
checkCmp "sumproduct-eq"       "=SUMPRODUCT(--(D1:D3=0.3))"    (n 1.0)
checkCmp "sumproduct-ranges"   "=SUMPRODUCT(--(D1:D3=H1:H3))"  (n 3.0)
checkCmp "sumproduct-gt"       "=SUMPRODUCT((D1:D3>0.3)*1)"    (n 1.0)
checkCmp "sum-if-array"        "=SUM(IF(D1:D3=H1:H3,1,0))"     (n 3.0)
checkCmp "filter-eq"           "=ROWS(FILTER(D1:D3,D1:D3=0.3))" (n 1.0)
checkCmp "filter-sum"          "=SUM(FILTER(E1:E3,D1:D3<=0.3))" (n 3.0)
checkCmp "sequence-eq"         "=SUM(--(SEQUENCE(3)*0.1=0.3))"  (n 1.0)
checkCmp "sequence-sumproduct" "=SUMPRODUCT(--(SEQUENCE(3)*0.1=0.3))" (n 1.0)
checkCmp "makearray-eq"        "=SUM(MAKEARRAY(3,1,LAMBDA(r,c,--(r*0.1=0.3))))" (n 1.0)
checkCmp "makearray-le"        "=SUM(--(MAKEARRAY(3,1,LAMBDA(r,c,r*0.1))<=0.3))" (n 3.0)
checkCmp "makearray-gt"        "=SUM(--(MAKEARRAY(3,1,LAMBDA(r,c,r*0.1))>0.3))" (n 0.0)
checkCmp "exact-text"          "=EXACT(D2,0.3)"                (b true)
checkCmp "delta-binary"        "=DELTA(D2,0.3)"                (n 0.0)
checkCmp "gestep-binary"       "=GESTEP(0.3,D2)"               (n 0.0)
checkCmp "rank-binary"         "=RANK(0.3,D1:D3)"              (e "N/A")
checkCmp "mode-binary"         "=MODE(D2,0.3)"                 (e "N/A")
checkCmp "frequency-binary"    "=INDEX(FREQUENCY(D1:D3,0.3),1)" (n 1.0)

printfn ""
printfn "Stage 22 semantics tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
