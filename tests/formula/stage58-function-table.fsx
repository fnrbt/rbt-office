#load "function-table-kit.fsx"
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open ``Function-table-kit``
module Wb = Rbt.Spreadsheet.Workbook

// The function table itself: Excel 365's newer error values, the generated catalog
// against the modules, names nobody provides, provider functions (with and without a
// provider, answering at once or later) and the cached values of their formulas.

// ----- the error values
let m = empty
for code in [ Err.busy; Err.connect; Err.blocked; Err.field; Err.unknown; Err.gettingData ] do
    is "literal" m ("=#" + code) (E code)
    is "ISERROR" m ("=ISERROR(#" + code + ")") (B true)
    is "ISERR" m ("=ISERR(#" + code + ")") (B true)
    is "ISNA" m ("=ISNA(#" + code + ")") (B false)
    is "IFERROR" m ("=IFERROR(#" + code + ",1)") (N 1.0)
    is "propagates" m ("=1+#" + code) (E code)
    eq ("typed " + code) (Rbt.Spreadsheet.Input.parseInputFor false 2026 ("#" + code)).Value (E code)
    eq ("typed lower-case " + code) (Rbt.Spreadsheet.Input.parseInputFor false 2026 ("#" + code.ToLowerInvariant())).Value (E code)
    eq ("shown " + code) (Rbt.Spreadsheet.CellText.cellDisplayText None 100.0 defaultFormat (E code)) ("#" + code)
number m "=ERROR.TYPE(#NULL!)" 1.0
number m "=ERROR.TYPE(#N/A)" 7.0
number m "=ERROR.TYPE(#GETTING_DATA)" 8.0
number m "=ERROR.TYPE(#SPILL!)" 9.0
number m "=ERROR.TYPE(#CONNECT!)" 10.0
number m "=ERROR.TYPE(#BLOCKED!)" 11.0
number m "=ERROR.TYPE(#UNKNOWN!)" 12.0
number m "=ERROR.TYPE(#FIELD!)" 13.0
number m "=ERROR.TYPE(#CALC!)" 14.0
// (#BUSY! has no number Microsoft documents; until the Excel oracle says one, #N/A.)
error m "=ERROR.TYPE(#BUSY!)" Err.na
error m "=ERROR.TYPE(1)" Err.na
number m "=COUNTIF({1,2},\"<>#BUSY!\")" 2.0

// ----- the catalog is what the modules say, and every function of it is installed
let modules =
    [ Rbt.Spreadsheet.Functions.FnMath.moduleName, Rbt.Spreadsheet.Functions.FnMath.functions
      Rbt.Spreadsheet.Functions.FnDatabase.moduleName, Rbt.Spreadsheet.Functions.FnDatabase.functions
      Rbt.Spreadsheet.Functions.FnEngineering.moduleName, Rbt.Spreadsheet.Functions.FnEngineering.functions
      Rbt.Spreadsheet.Functions.FnStatistical.moduleName, Rbt.Spreadsheet.Functions.FnStatistical.functions
      Rbt.Spreadsheet.Functions.FnStatDist.moduleName, Rbt.Spreadsheet.Functions.FnStatDist.functions
      Rbt.Spreadsheet.Functions.FnFinancial.moduleName, Rbt.Spreadsheet.Functions.FnFinancial.functions
      Rbt.Spreadsheet.Functions.FnText.moduleName, Rbt.Spreadsheet.Functions.FnText.functions
      Rbt.Spreadsheet.Functions.FnArrays.moduleName, Rbt.Spreadsheet.Functions.FnArrays.functions
      Rbt.Spreadsheet.Functions.FnInformation.moduleName, Rbt.Spreadsheet.Functions.FnInformation.functions ]
let fromModules =
    [ for (moduleName, defs) in modules do
        for def in defs ->
            def.Name, moduleName, def.Signature, (if def.Volatile then 1 else 0) ||| (match def.Impl with ArrayFunction _ -> 2 | _ -> 0) ]
    |> List.sort
let catalog =
    [ for entry in Rbt.Spreadsheet.FunctionCatalog.entries -> entry.Name, entry.Module, entry.Signature, entry.Flags ] |> List.sort
eq "catalog = modules (run scripts/gen-function-catalog.fsx)" catalog fromModules
eq "modules of the catalog" (List.ofArray Rbt.Spreadsheet.FunctionCatalog.modules |> List.sort) (modules |> List.map fst |> List.sort)
for (name, moduleName, _, _) in catalog do
    eq ("installed " + name) ((tryFunctionDef name).IsSome && isFunctionModuleInstalled moduleName) true
eq "no provider function is in a module" (catalog |> List.filter (fun (name, _, _, _) -> isProviderFunction name)) []
eq "no catalog name twice" (catalog |> List.countBy (fun (name, _, _, _) -> name) |> List.filter (fun (_, n) -> n > 1)) []
printfn "function table: %d functions in %d modules" catalog.Length modules.Length
eq "MUNIT spills" (isArrayFunction "munit") true
eq "WRAPROWS spills" (isArrayFunction "WRAPROWS") true
eq "STOCKHISTORY spills" (isArrayFunction "STOCKHISTORY") true
eq "ACOT does not" (isArrayFunction "ACOT") false
eq "CELL is volatile" (match parseFormula "=CELL(\"row\")" with Ok e -> containsVolatile e | _ -> false) true
eq "ACOT is not" (match parseFormula "=ACOT(1)" with Ok e -> containsVolatile e | _ -> true) false
eq "names a formula calls" (calledFunctionNames "=_xlfn.ACOT(1)+sum(\"ROMAN(\",'a(b'!A1)+_xlfn._xlws.SORT(A1:A2)") [ "ACOT"; "SUM"; "SORT" ]
let mutable readyCalls = 0
requestFunctionsFor [ "=ROMAN(1)"; "=SUM(1)"; "=NOSUCH(2)" ] (fun () -> readyCalls <- readyCalls + 1)
eq "on .NET every module is there: ready at once" readyCalls 1
error m "=NOSUCHFUNCTION(1)" Err.name
error m "=ROMANS(1)" Err.name
// An omitted argument of a table function: MissingArg, its default (DELTA's second is 0).
number m "=DELTA(0,)" 1.0
error m "=ACOT()" Err.value
error m "=ACOT(1,2)" Err.value

// ----- provider functions
error m "=WEBSERVICE(\"http://example.com\")" Err.connect
error m "=CALL(\"x\")" Err.blocked
error m "=REGISTER.ID(\"x\",\"y\")" Err.blocked
error m "=CUBEVALUE(\"c\",\"m\")" Err.na
error m "=IMPORTCSV(\"a.csv\")" Err.blocked
eq "SUM has no provider" (registerFunctionProvider "SUM" { Invoke = fun _ answer -> answer [| [| N 1.0 |] |] }) false
let mutable calls = 0
let mutable lastCall : ProviderCall option = None
let echo = { Invoke = fun (call: ProviderCall) answer ->
                calls <- calls + 1
                lastCall <- Some call
                answer [| [| S ("got " + (match call.Arguments.[0].[0].[0] with CellValue.String s -> s | v -> sprintf "%A" v)) |] |] }
eq "register WEBSERVICE" (registerFunctionProvider "webservice" echo) true
text m "=WEBSERVICE(\"u\")" "got u"
text m "=WEBSERVICE(\"u\")" "got u"
eq "an answer is kept for the same arguments" calls 1
text m "=WEBSERVICE(\"v\")&\"!\"" "got v!"
eq "the arguments are handed over" (lastCall |> Option.map (fun c -> c.Name, c.Arguments.Length)) (Some ("WEBSERVICE", 1))
// A provider that answers later: #BUSY! until then, and listeners hear of the answer.
let mutable pending : (CellValue[][] -> unit) list = []
eq "register STOCKHISTORY" (registerFunctionProvider "STOCKHISTORY" { Invoke = fun _ answer -> pending <- answer :: pending }) true
let mutable arrivals = 0
let stop = onFunctionsArrived (fun () -> arrivals <- arrivals + 1)
error m "=STOCKHISTORY(\"MSFT\",45000)" Err.busy
error m "=STOCKHISTORY(\"MSFT\",45000)" Err.busy
eq "asked once while waiting" pending.Length 1
pending.Head [| [| S "Date"; S "Close" |]; [| N 45000.0; N 250.0 |] |]
eq "the answer was told" arrivals 1
eq "STOCKHISTORY spills its answer" (arrayOf m "=STOCKHISTORY(\"MSFT\",45000)") [| [| S "Date"; S "Close" |]; [| N 45000.0; N 250.0 |] |]
text m "=STOCKHISTORY(\"MSFT\",45000)" "Date"
// Answering again replaces the value (RTD-style streams).
pending.Head [| [| S "Again" |] |]
eq "a second answer is told" arrivals 2
text m "=STOCKHISTORY(\"MSFT\",45000)" "Again"
stop ()
// A provider that throws is #CONNECT!.
eq "register TRANSLATE" (registerFunctionProvider "TRANSLATE" { Invoke = fun _ _ -> failwith "offline" }) true
error m "=TRANSLATE(\"Hallo\",\"de\",\"en\")" Err.connect
unregisterFunctionProvider "TRANSLATE"
unregisterFunctionProvider "STOCKHISTORY"
error m "=STOCKHISTORY(\"MSFT\",45000)" Err.connect

// ----- cached values: a formula of a provider function nobody provides keeps the file's value
unregisterFunctionProvider "WEBSERVICE"
let at r c : CellAddress = { Row = r; Col = c }
let sheet =
    { createEmpty 10 4 with
        Cells =
            [ at 0 0, CellValue.Formula ("=WEBSERVICE(\"u\")", S "cached text")
              at 1 0, CellValue.Formula ("=LEN(WEBSERVICE(\"u\"))", N 11.0)
              at 2 0, CellValue.Formula ("=WEBSERVICE(\"u\")", CellValue.Empty)
              at 3 0, CellValue.Formula ("=A2*2", CellValue.Empty)
              at 4 0, CellValue.Formula ("=ACOT(1)*4", CellValue.Empty)
              at 5 0, CellValue.Formula ("=WEBSERVICE(\"u\")", E Err.busy) ]
            |> List.map (fun (a, v) -> a, cell v) |> Map.ofList }
let book : Wb.Workbook =
    let wb0 = Wb.createWorkbook ()
    Wb.updateSheetModel wb0.ActiveSheet sheet wb0
let valueIn (wb: Wb.Workbook) r c = unwrapFormulaValue (getCellValue (at r c) wb.Sheets.[wb.ActiveSheet].Model)
let calculated = recalculateWorkbook book
eq "kept: the file's value" (valueIn calculated 0 0) (S "cached text")
eq "kept inside another function" (valueIn calculated 1 0) (N 11.0)
eq "no value to keep: the error" (valueIn calculated 2 0) (E Err.connect)
eq "what reads a kept value" (valueIn calculated 3 0) (N 22.0)
eq "a function of a module" (valueIn calculated 4 0) (N System.Math.PI)
eq "#BUSY! is no value to keep" (valueIn calculated 5 0) (E Err.connect)
eq "register WEBSERVICE again" (registerFunctionProvider "WEBSERVICE" echo) true
let provided = recalculateWorkbook book
eq "with a provider: its answer" (valueIn provided 0 0) (S "got u")
eq "with a provider: inside another function" (valueIn provided 1 0) (N 5.0)
unregisterFunctionProvider "WEBSERVICE"
// The incremental engine keeps it too.
let full, engine = Rbt.Spreadsheet.Recalc.recalculateAll book
eq "full pass keeps it" (valueIn full 0 0) (S "cached text")
let edited = Wb.updateSheetModel full.ActiveSheet (setCell (at 7 0) (cell (N 1.0)) full.Sheets.[full.ActiveSheet].Model) full
match Rbt.Spreadsheet.Recalc.recalculate engine edited None with
| Some again -> eq "an edit keeps it" (valueIn again 0 0) (S "cached text")
| None -> eq "incremental pass ran" false true

printfn "stage58 function table: %d passed, %d failed" pass fail
if fail > 0 then exit 1
