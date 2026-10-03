#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Workbook
open Rbt.Spreadsheet.Formula

// Structure edits: paste/fill reference shifting, the Expr → text printer,
// row/column insert + delete, sort, and workbook-wide recalculation.

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
let modelOf cells = { createEmpty 100 26 with Cells = Map.ofList cells }
let formulaAt r c (m: SpreadsheetModel) =
    match getCellValue (at r c) m with
    | CellValue.Formula (f, _) -> f
    | other -> sprintf "<%A>" other
let valueAt r c (m: SpreadsheetModel) = getCellValue (at r c) m |> unwrapFormulaValue

// --- Paste shifts references through the parser, not by scanning text -------
// This block runs before anything else touches the Formula module, so it also
// proves Model.pasteClipboard finds the rewriter on a cold start.
let pasteOf formula =
    let source = modelOf [ at 0 2, fml formula ]
    let clip = serializeSelection (SingleCell (0, 2)) source false
    pasteClipboard (at 1 3) clip source |> formulaAt 1 3

eq "paste-cold-start"     (pasteOf "=LOG10(A1)")             "=LOG10(B2)"
eq "paste-function-name"  (pasteOf "=ATAN2(A1,B1)")          "=ATAN2(B2,C2)"
eq "paste-sheet-name"     (pasteOf "=SUM(Sheet1!A1:A2)")     "=SUM(Sheet1!B2:B3)"
eq "paste-quoted-sheet"   (pasteOf "=SUM('My Sheet'!A1:A2)") "=SUM('My Sheet'!B2:B3)"
eq "paste-string-literal" (pasteOf "=IF(A1=\"B2\",1,2)")     "=IF(B2=\"B2\",1,2)"
eq "paste-col-locked"     (pasteOf "=$A1+1")                 "=$A2+1"
eq "paste-row-locked"     (pasteOf "=A$1+1")                 "=B$1+1"
eq "paste-fully-locked"   (pasteOf "=$A$1+1")                "=$A$1+1"
eq "paste-no-equals"      (pasteOf "A1*2")                   "B2*2"
eq "paste-name-kept"      (pasteOf "=Rate*A1")               "=Rate*B2"
eq "paste-unparsable"     (pasteOf "=SUM(A1")                "=SUM(A1"

// Pasting up/left past the sheet edge gives #REF!, as Excel does.
let pasteUp =
    let source = modelOf [ at 1 1, fml "=A1+B2" ]
    let clip = serializeSelection (SingleCell (1, 1)) source false
    pasteClipboard (at 0 0) clip source |> formulaAt 0 0
eq "paste-off-sheet" pasteUp "=#REF!+A1"

// Cut-then-paste moves the formula as written.
let cutPaste =
    let source = modelOf [ at 0 2, fml "=A1+$B$2" ]
    let clip = serializeSelection (SingleCell (0, 2)) source true
    pasteClipboard (at 5 5) clip source |> formulaAt 5 5
eq "cut-paste-unshifted" cutPaste "=A1+$B$2"

// A copied block keeps each formula relative to its own cell.
let blockPaste =
    let source = modelOf [ at 0 0, num 1.0; at 1 0, num 2.0; at 0 1, fml "=A1*10"; at 1 1, fml "=A2*10" ]
    let clip = serializeSelection (Range (0, 1, 1, 1)) source false
    pasteClipboard (at 0 2) clip source |> recalcAndSpill None None
eq "block-paste-formula" (formulaAt 1 2 blockPaste) "=B2*10"
eq "block-paste-value"   (valueAt 1 2 blockPaste) (CellValue.Number 200.0)

eq "fill-shift" (shiftFormulaText 2 0 "=SUM(A1:A3)/COUNT($A$1:A3)") "=SUM(A3:A5)/COUNT($A$1:A5)"

// --- Printer round trip -----------------------------------------------------
let roundTrip (formula: string) =
    match parseFormula formula with
    | Result.Error e -> ok ("print " + formula) false ("does not parse: " + e)
    | Ok expr ->
        let text = exprToFormula expr
        match parseFormula text with
        | Ok again -> ok ("print " + formula) (again = expr) (sprintf "printed %s reparsed differently" text)
        | Result.Error e -> ok ("print " + formula) false (sprintf "printed %s does not parse: %s" text e)
let prints (formula: string) (expected: string) =
    match parseFormula formula with
    | Ok expr -> eq ("text " + formula) (exprToFormula expr) expected
    | Result.Error e -> ok ("text " + formula) false e

for f in
    [ "=1+2*3"; "=(1+2)*3"; "=1-(2-3)"; "=1-2-3"; "=2^3^2"; "=2^(3^2)"; "=-2^2"; "=-(2^2)"; "=2^-3"
      "=(1+2)%"; "=-A1%"; "=(-A1)%"; "=--A1"; "=A1&B1&\"x\""; "=(A1=B1)=(C1<>D1)"; "=A1<=B1"
      "=1/(2/3)"; "=1/2/3"; "=(A1&B1)+1"; "=\"say \"\"hi\"\"\""; "=1.5E-7+2.25"; "=1E+30"
      "=$A$1+A$1+$A1+A1"; "=SUM(A1:B2,C3)"; "=Sheet2!A1+'My Sheet'!$B$2"; "='It''s'!A1:B2"
      "='2024'!A1"; "=#REF!+#N/A"; "=#DIV/0!"; "=A1#"; "=SUM(A1#)"; "=TRUE"; "=IF(FALSE,1,2)"
      "=Sales[Amount]"; "=Sales[[#Totals],[Amount]]"; "=Sales[@Amount]"; "=Sales[@[Unit Price]]"
      "=Sales[[#All],[Qty]:[Amount]]"; "=Sales[#Headers]"; "=[@Qty]*[@Price]"; "=Sales[[Qty]:[Amount]]"
      "=Sales[@]"; "=SUM(Sales[Unit Price])"
      "=LET(x,A1,y,x*2,x+y)"; "=MyName*2"; "=SUM(MyRange)"; "=IF(A1>0,\"a,b\",\"(\")"
      "=1+2+3+4+5+6+7+8+9+10"; "=A1*(B1+C1)*(D1-E1)"; "=-(-(1))"; "=+A1"; "=A1%%" ] do
    roundTrip f

prints "= sum( A1 , B1 )" "sum(A1,B1)"
prints "=(1+2)*3" "(1+2)*3"
prints "=1+(2*3)" "1+2*3"
prints "=Sheet1!A1:B2" "Sheet1!A1:B2"
prints "='My Sheet'!A1" "'My Sheet'!A1"
prints "=_xlfn.XLOOKUP(A1,B1:B9,C1:C9)" "XLOOKUP(A1,B1:B9,C1:C9)"
prints "=A1!B2+XFD1!A1+XFE1!A1" "'A1'!B2+'XFD1'!A1+XFE1!A1"

// A chain thousands of operators long must print without recursing per operator.
let longChain = "=" + String.concat "+" [ for i in 1 .. 4000 -> sprintf "A%d" i ]
ok "print-long-chain" (shiftFormulaText 1 0 longChain = "=" + String.concat "+" [ for i in 2 .. 4001 -> sprintf "A%d" i ]) "long chain shifted wrongly"

printfn ""
printfn "Stage 21 structure tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
