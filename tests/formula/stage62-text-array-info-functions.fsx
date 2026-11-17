#load "function-table-kit.fsx"
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open ``Function-table-kit``
module Wb = Rbt.Spreadsheet.Workbook

// Text, dynamic-array and information functions of the function table.
let m = empty

// ----- text
let cells =
    modelOf [ (0, 0), B true; (0, 1), E Err.value; (1, 0), N 1234.01234; (1, 1), S "Seattle"
              (2, 0), S "Hello"; (3, 0), S "a\"b" ]
text cells "=ARRAYTOTEXT(A1:B3)" "TRUE, #VALUE!, 1234.01234, Seattle, Hello, "
text cells "=ARRAYTOTEXT(A1:B3,1)" "{TRUE,#VALUE!;1234.01234,\"Seattle\";\"Hello\",}"
text cells "=ARRAYTOTEXT({1,2;3,4})" "1, 2, 3, 4"
text cells "=ARRAYTOTEXT({1,2;3,4},1)" "{1,2;3,4}"
error cells "=ARRAYTOTEXT(A1:B3,2)" Err.value
text cells "=VALUETOTEXT(B2)" "Seattle"
text cells "=VALUETOTEXT(B2,1)" "\"Seattle\""
text cells "=VALUETOTEXT(A4,1)" "\"a\"\"b\""
text cells "=VALUETOTEXT(B1)" "#VALUE!"
text cells "=VALUETOTEXT(A1,1)" "TRUE"
text cells "=VALUETOTEXT(A2)" "1234.01234"
eq "VALUETOTEXT over an array" (arrayOf cells "=VALUETOTEXT({\"a\",1},1)") [| [| S "\"a\""; S "1" |] |]
text m "=DOLLAR(1234.567,2)" "$1,234.57"
text m "=DOLLAR(1234.567,-2)" "$1,200"
text m "=DOLLAR(-1234.567,-2)" "($1,200)"
text m "=DOLLAR(-0.123,4)" "($0.1230)"
text m "=DOLLAR(99.888)" "$99.89"
text m "=DOLLAR(0)" "$0.00"
error m "=DOLLAR(\"x\")" Err.value
text m "=ENCODEURL(\"http://contoso.sharepoint.com/Finance/Profit and Loss Statement.xlsx\")" "http%3A%2F%2Fcontoso.sharepoint.com%2FFinance%2FProfit%20and%20Loss%20Statement.xlsx"
text m "=ENCODEURL(\"a-b_c.d~e\")" "a-b_c.d~e"
text m "=ENCODEURL(\"\u00E9\")" "%C3%A9"
number m "=LENB(\"abc\")" 3.0
text m "=LEFTB(\"abc\",2)" "ab"
text m "=LEFTB(\"abc\")" "a"
text m "=RIGHTB(\"abc\",2)" "bc"
text m "=MIDB(\"abcdef\",2,3)" "bcd"
number m "=FINDB(\"c\",\"abcabc\",4)" 6.0
number m "=SEARCHB(\"C\",\"abcabc\")" 3.0
text m "=REPLACEB(\"abcdef\",2,3,\"X\")" "aXef"
eq "LENB over an array" (arrayOf m "=LENB({\"a\",\"abc\"})") [| [| N 1.0; N 3.0 |] |]

// ----- dynamic arrays
let na = E Err.na
eq "EXPAND" (arrayOf m "=EXPAND({1,2;3,4},3,3)") [| [| N 1.0; N 2.0; na |]; [| N 3.0; N 4.0; na |]; [| na; na; na |] |]
eq "EXPAND padded" (arrayOf m "=EXPAND({1,2;3,4},3,3,\"-\")") [| [| N 1.0; N 2.0; S "-" |]; [| N 3.0; N 4.0; S "-" |]; [| S "-"; S "-"; S "-" |] |]
eq "EXPAND rows only" (arrayOf m "=EXPAND({1,2},2)") [| [| N 1.0; N 2.0 |]; [| na; na |] |]
eq "EXPAND columns only" (arrayOf m "=EXPAND({1;2},,2,0)") [| [| N 1.0; N 0.0 |]; [| N 2.0; N 0.0 |] |]
error m "=EXPAND({1,2;3,4},1,3)" Err.value
eq "WRAPROWS" (arrayOf m "=WRAPROWS({1,2,3,4,5},2)") [| [| N 1.0; N 2.0 |]; [| N 3.0; N 4.0 |]; [| N 5.0; na |] |]
eq "WRAPROWS padded" (arrayOf m "=WRAPROWS({1,2,3},2,0)") [| [| N 1.0; N 2.0 |]; [| N 3.0; N 0.0 |] |]
eq "WRAPCOLS" (arrayOf m "=WRAPCOLS({1;2;3;4;5},2)") [| [| N 1.0; N 3.0; N 5.0 |]; [| N 2.0; N 4.0; na |] |]
error m "=WRAPROWS({1,2;3,4},2)" Err.value
error m "=WRAPCOLS({1,2,3},0)" Err.num
// In a scalar context a spilling function is its top-left value.
number m "=EXPAND({7,8},2,2)" 7.0
number m "=SUM(WRAPROWS({1,2,3,4,5},2,0))" 15.0

// ----- information: a workbook of three sheets
let at r c : CellAddress = { Row = r; Col = c }
let sheetOf (entries: (CellAddress * CellValue) list) =
    { createEmpty 20 8 with Cells = entries |> List.map (fun (a, v) -> a, cell v) |> Map.ofList }
let firstBare =
    { sheetOf [ at 0 0, N 1.0; at 1 0, CellValue.Formula ("=SUM(A1:A1)*2", N 2.0); at 2 0, S "text"; at 3 0, N 0.25 ]
        with ColumnWidths = Map.ofList [ 1, 75.0 ] }
let first =
    { firstBare with
        Cells =
            firstBare.Cells
            |> Map.add (at 3 0) { (cell (N 0.25)) with Format = { defaultFormat with NumberFormat = Some "0.00%" } }
            |> Map.add (at 4 0) { (cell (N 1234.5)) with Format = { defaultFormat with NumberFormat = Some "#,##0.00" } }
            |> Map.add (at 5 0) { (cell (N 45000.0)) with Format = { defaultFormat with NumberFormat = Some "d-mmm-yy" } }
            |> Map.add (at 6 0) { (cell (S "left")) with Format = { defaultFormat with HorizontalAlign = Some "left" } } }
let sheetEntry id name model : Wb.SheetId * Wb.Sheet = id, { Id = id; Name = name; Model = model; Visible = true }
open Rbt.Spreadsheet.Workbook
let wb : Wb.Workbook =
    { Sheets = Map.ofList [ sheetEntry "s1" "Data" first; sheetEntry "s2" "Summary" (sheetOf [ at 0 0, N 5.0 ]); sheetEntry "s3" "My Sheet" (sheetOf []) ]
      SheetOrder = [ "s1"; "s2"; "s3" ]; ActiveSheet = "s1"; Tables = Map.empty; Names = Map.empty; NameInfo = Map.empty; Iteration = None }
    |> Wb.setDefinedName None "Total" "Summary!$A$1" None
let inWb (sheet: string) (formula: string) =
    match parseFormula formula with
    | Ok expr -> unwrapValue (eval expr { Model = wb.Sheets.[sheet].Model; Workbook = Some wb; Sheet = Some sheet; CurrentCell = Some (at 9 2) })
    | Result.Error e -> failwith e
let wbIs (formula: string) (expected: CellValue) = eq ("workbook " + formula) (inWb "s1" formula) expected
wbIs "=SHEET()" (N 1.0)
eq "SHEET() on the second sheet" (inWb "s2" "=SHEET()") (N 2.0)
wbIs "=SHEET(Summary!A1)" (N 2.0)
wbIs "=SHEET(\"My Sheet\")" (N 3.0)
wbIs "=SHEET(Total)" (N 2.0)
wbIs "=SHEET(\"Nope\")" (E Err.na)
wbIs "=SHEETS()" (N 3.0)
wbIs "=SHEETS(A1:B2)" (N 1.0)
wbIs "=SUM('Data:My Sheet'!A1)" (N 6.0)
wbIs "=SHEETS('Data:My Sheet'!A1)" (N 3.0)
wbIs "=SHEETS(Data:Summary!A1)" (N 2.0)
wbIs "=SHEETS('Summary:My Sheet'!A1:B3)" (N 2.0)
wbIs "=FORMULATEXT(A2)" (S "=SUM(A1:A1)*2")
wbIs "=FORMULATEXT(A1)" (E Err.na)
wbIs "=FORMULATEXT(OFFSET(A1,1,0))" (S "=SUM(A1:A1)*2")
wbIs "=FORMULATEXT(1)" (E Err.na)
wbIs "=ISOMITTED(1)" (B false)
wbIs "=LET(f,LAMBDA(a,b,IF(ISOMITTED(b),\"no b\",a+b)),f(1,))" (S "no b")
wbIs "=LET(f,LAMBDA(a,b,IF(ISOMITTED(b),\"no b\",a+b)),f(1))" (S "no b")
wbIs "=LET(f,LAMBDA(a,b,IF(ISOMITTED(b),\"no b\",a+b)),f(1,2))" (N 3.0)
wbIs "=CELL(\"address\",B3)" (S "$B$3")
wbIs "=CELL(\"address\",Summary!A1)" (S "Summary!$A$1")
wbIs "=CELL(\"address\",'My Sheet'!C7)" (S "'My Sheet'!$C$7")
wbIs "=CELL(\"address\")" (S "$C$10")
wbIs "=CELL(\"row\",B3:D9)" (N 3.0)
wbIs "=CELL(\"col\",B3:D9)" (N 2.0)
wbIs "=CELL(\"contents\",A2)" (N 2.0)
wbIs "=CELL(\"type\",A1)" (S "v")
wbIs "=CELL(\"type\",A3)" (S "l")
wbIs "=CELL(\"type\",H20)" (S "b")
wbIs "=CELL(\"format\",A1)" (S "G")
wbIs "=CELL(\"format\",A4)" (S "P2")
wbIs "=CELL(\"format\",A5)" (S ",2")
wbIs "=CELL(\"format\",A6)" (S "D1")
wbIs "=CELL(\"prefix\",A7)" (S "'")
wbIs "=CELL(\"prefix\",A3)" (S "")
wbIs "=CELL(\"protect\",A1)" (N 1.0)
wbIs "=CELL(\"width\",B1)" (N 10.0)
wbIs "=CELL(\"filename\")" (S "")
wbIs "=CELL(\"nonsense\",A1)" (E Err.value)
eq "format codes" ([ Some "0"; Some "0.000"; Some "$#,##0_);($#,##0)"; Some "0.00E+00"; Some "h:mm AM/PM"; Some "mm/dd"; Some "0;[Red]-0"; Some "#,##0_);(#,##0)" ] |> List.map Rbt.Spreadsheet.Functions.FnInformation.formatCodeOf)
    [ "F0"; "F3"; "C0"; "S2"; "D7"; "D5"; "F0-"; ",0" ]
wbIs "=INFO(\"system\")" (S "pcdos")
wbIs "=INFO(\"numfile\")" (N 3.0)
wbIs "=INFO(\"recalc\")" (S "Automatic")
wbIs "=INFO(\"origin\")" (S "$A:$A$1")
wbIs "=INFO(\"bogus\")" (E Err.value)
setHostInfo "directory" (Some "C:\\Books\\")
setHostInfo "filename" (Some "Plan.xlsx")
setHostInfo "osversion" (Some "Test OS 1.0")
wbIs "=INFO(\"directory\")" (S "C:\\Books\\")
wbIs "=INFO(\"osversion\")" (S "Test OS 1.0")
wbIs "=CELL(\"filename\",Summary!A1)" (S "C:\\Books\\[Plan.xlsx]Summary")
setHostInfo "filename" None

printfn "stage62 text, array and information functions: %d passed, %d failed" pass fail
if fail > 0 then exit 1
