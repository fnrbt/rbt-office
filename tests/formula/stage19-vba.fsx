#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/VBScriptParser.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

// VBA runs natively (cross-platform) AND writes cells back via runOnModel.

open RedBat.Spreadsheet.Model
open System.Reflection
module V = RedBat.Spreadsheet.Vba

let cell v = { Value = v; Format = defaultFormat; Comment = None }
let model =
    { createEmpty 20 8 with
        Cells = Map.ofList [ { Row=0; Col=0 }, cell (CellValue.Number 10.) ] }

let mutable pass = 0
let mutable fail = 0
let check label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-26s expected %A got %A" label expected got

let disp m r c = displayValue (getCellValue { Row=r; Col=c } m)

// A macro that reads A1, writes a computed value, a string, and a formula.
let src =
    String.concat "\n"
        [ "Range(\"B1\").Value = Range(\"A1\").Value * 2"
          "Cells(1, 3).Value = \"hello\""
          "Range(\"D1\").Formula = \"=B1+5\"" ]

match V.runOnModel model src with
| Result.Error e -> fail <- fail + 1; printfn "FAIL macro errored: %s" e
| Result.Ok m ->
    check "B1 = A1*2"        (disp m 0 1) "20"
    check "C1 = string"      (disp m 0 2) "hello"
    check "D1 formula compute" (disp m 0 3) "25"   // =B1+5 = 20+5
    check "A1 untouched"     (disp m 0 0) "10"

// Loop macro: fill A3:A5 with 1,2,3.
let src2 =
    String.concat "\n"
        [ "Dim i"
          "For i = 1 To 3"
          "  Cells(i + 2, 1).Value = i"
          "Next i" ]
match V.runOnModel model src2 with
| Result.Error e -> fail <- fail + 1; printfn "FAIL loop macro errored: %s" e
| Result.Ok m ->
    check "loop A3" (disp m 2 0) "1"
    check "loop A4" (disp m 3 0) "2"
    check "loop A5" (disp m 4 0) "3"

// --- events: hasSub + fireEvent ---
let openModel = { model with Macros = "Sub Workbook_Open()\n  Range(\"A1\").Value = 99\nEnd Sub" }
check "hasSub true"  (V.hasSub "Workbook_Open" openModel.Macros) true
check "hasSub false" (V.hasSub "Worksheet_Change" openModel.Macros) false
check "Workbook_Open fires" (disp (V.fireEvent "Workbook_Open" "" openModel) 0 0) "99"
check "no-sub no-op"        (disp (V.fireEvent "Workbook_Open" "" model) 0 0) "10"

// Worksheet_Change receiving the changed cell as Target.
let wcModel = { model with Macros = "Sub Worksheet_Change(ByVal Target As Range)\n  Range(\"Z1\").Value = Target.Address\nEnd Sub" }
check "Worksheet_Change Target" (disp (V.fireEvent "Worksheet_Change" "Range(\"B2\")" wcModel) 0 25) "$B$2"

// --- compile cache: same source compiles once, reused thereafter ---
match V.compileCached "Range(\"A1\").Value = 1", V.compileCached "Range(\"A1\").Value = 1" with
| Result.Ok a, Result.Ok b -> check "compile cache hit (same instance)" (System.Object.ReferenceEquals(a, b)) true
| _ -> fail <- fail + 1; printfn "FAIL compile cache compile error"
let denseMacroSource index =
    System.String.Join(
        "\n",
        [|
            yield "Dim x"
            yield sprintf "' dense source %02d" index
            for _ in 1 .. 1400 do
                yield "x = 1"
        |])
let compileCacheSource label source =
    match V.compileCached source with
    | Result.Ok bytecode -> Some bytecode
    | Result.Error e ->
        fail <- fail + 1
        printfn "FAIL %s errored: %s" label e
        None
match compileCacheSource "compile cache dense first" (denseMacroSource 0),
      compileCacheSource "compile cache dense repeat" (denseMacroSource 0) with
| Some first, Some repeat ->
    check "compile cache dense hit" (System.Object.ReferenceEquals(first, repeat)) true
    for index in 1 .. 23 do
        compileCacheSource (sprintf "compile cache dense %02d" index) (denseMacroSource index) |> ignore
    match compileCacheSource "compile cache dense first after fill" (denseMacroSource 0) with
    | Some afterFill ->
        check "compile cache evicts dense bytecode" (System.Object.ReferenceEquals(first, afterFill)) false
    | None -> ()
| _ -> ()
let largeMacroSource index =
    "Range(\"A1\").Value = 1\n"
    + sprintf "' large source %02d " index
    + System.String('x', 100000)
let largeMacroRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for index in 0 .. 39 do
        match V.compileCached (largeMacroSource index) with
        | Result.Ok _ -> ()
        | Result.Error e ->
            fail <- fail + 1
            printfn "FAIL compile cache large source errored: %s" e
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
check "compile cache bounded source memory" (largeMacroRetainedBytes < 6000000L) true
let touchOversizedMacro () =
    let source =
        "Range(\"A1\").Value = 1\n' oversized "
        + System.String('x', 2000000)
    match V.compileCached source with
    | Result.Ok _ -> ()
    | Result.Error e ->
        fail <- fail + 1
        printfn "FAIL compile cache oversized source errored: %s" e
let oversizedMacroRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    touchOversizedMacro ()
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
check "compile cache skips oversized source" (oversizedMacroRetainedBytes < 2000000L) true

let retainedUnitsMethod =
    typeof<SpreadsheetModel>.Assembly.GetType("RedBat.Spreadsheet.Vba").GetMethod("bytecodeProgramRetainedUnits", BindingFlags.NonPublic ||| BindingFlags.Static)
if isNull retainedUnitsMethod then
    fail <- fail + 1
    printfn "FAIL compile cache retained-unit reflection"
else
    let oversizedInstructionCount = 1000000
    let instruction = Bytecode.makeInstruction Bytecode.Nop 0
    let bytecode : Bytecode.BytecodeProgram =
        { Code = Array.create oversizedInstructionCount instruction
          Constants = [||]
          Globals = 0
          GlobalNames = [||]
          Functions = [||]
          Classes = [||] }
    let sw = System.Diagnostics.Stopwatch.StartNew()
    let retainedUnits = retainedUnitsMethod.Invoke(null, [| box 1; box bytecode |]) :?> int
    sw.Stop()
    check "compile cache retained units cap"
        (retainedUnits > 1000000 && sw.Elapsed.TotalMilliseconds < 10.0)
        true

let largeCompileHitSources =
    [| for index in 0 .. 15 ->
        "Range(\"A1\").Value = 1\n"
        + sprintf "' large repeated source %02d " index
        + System.String('x', 20000) |]
for source in largeCompileHitSources do
    V.compileCached source |> ignore
let largeCompileHitSw = System.Diagnostics.Stopwatch.StartNew()
let mutable largeCompileHitCount = 0
for _pass in 1 .. 2000 do
    for source in largeCompileHitSources do
        match V.compileCached source with
        | Result.Ok _ -> largeCompileHitCount <- largeCompileHitCount + 1
        | Result.Error _ -> ()
largeCompileHitSw.Stop()
check "compile cache large source hits"
    (largeCompileHitCount = 32000 && largeCompileHitSw.Elapsed.TotalMilliseconds < 40.0)
    true

let largeLexerSource =
    System.String.Join("\n", [| for _ in 1 .. 6000 -> "Range(\"A1\").Value = 1" |])
let largeLexerSw = System.Diagnostics.Stopwatch.StartNew()
let mutable largeLexerTokenCount = 0
let largeLexerOk =
    try
        largeLexerTokenCount <- (Lexer.tokenize largeLexerSource).Length
        true
    with ex ->
        printfn "FAIL lexer large source errored: %s" ex.Message
        false
largeLexerSw.Stop()
check "lexer large macro stack"
    (largeLexerOk && largeLexerTokenCount > 40000 && largeLexerSw.Elapsed.TotalMilliseconds < 1000.0)
    true
let largeParseSource =
    System.String.Join("\n", [| for _ in 1 .. 2000 -> "Range(\"A1\").Value = 1" |])
let largeParseSw = System.Diagnostics.Stopwatch.StartNew()
let largeParseTopLevelCount =
    match ParserSimple.parse_dialect Ast.VBA largeParseSource with
    | Result.Ok program -> program.TopLevels.Length
    | Result.Error e ->
        printfn "FAIL parser large source errored: %s" e
        -1
largeParseSw.Stop()
check "parser large macro perf"
    (largeParseTopLevelCount = 2000 && largeParseSw.Elapsed.TotalMilliseconds < 2500.0)
    true
let astString text : Ast.Expression =
    Ast.Expression.Literal (Ast.String text)
let astInteger value : Ast.Expression =
    Ast.Expression.Literal (Ast.Integer value)
let distinctConstantProgram count : Ast.Program =
    { Dialect = Ast.VBA
      TopLevels =
        [ for value in 1 .. count ->
            Ast.TopLevel.TopLevelStatement (
                Ast.Statement.MemberAssignment (
                    Ast.Expression.Call (Ast.Expression.Identifier "Range", [ astString "A1" ]),
                    "Value",
                    astInteger value)) ] }
let distinctConstantCompileCount = 8000
let distinctConstantCompileSw = System.Diagnostics.Stopwatch.StartNew()
let distinctConstantBytecode =
    Compiler.compileWithGlobalNames
        [ "Application"; "ActiveWorkbook"; "ActiveSheet"; "Range"; "Cells" ]
        (distinctConstantProgram distinctConstantCompileCount)
distinctConstantCompileSw.Stop()
check "compile distinct constants perf"
    (distinctConstantBytecode.Constants.Length = distinctConstantCompileCount + 1
     && distinctConstantCompileSw.Elapsed.TotalMilliseconds < 1000.0)
    true

check "like wildcard star" (StackVm.vbLikeMatch "a*c" "abbbc") true
check "like wildcard digit" (StackVm.vbLikeMatch "file###" "file123") true
check "like wildcard digit miss" (StackVm.vbLikeMatch "file###" "file12x") false
check "like wildcard class" (StackVm.vbLikeMatch "[!ab]at" "cat") true
check "like wildcard class miss" (StackVm.vbLikeMatch "[!ab]at" "aat") false
check "like wildcard case" (StackVm.vbLikeMatch "A?C" "abc") true
let largeLikePattern = System.String('?', 20000)
let largeLikeText = System.String('x', 20000)
let largeLikeSw = System.Diagnostics.Stopwatch.StartNew()
let largeLikeResult = StackVm.vbLikeMatch largeLikePattern largeLikeText
largeLikeSw.Stop()
check "like wildcard large stack"
    (largeLikeResult && largeLikeSw.Elapsed.TotalMilliseconds < 1000.0)
    true

// Read-only model snapshots should handle sparse used bounds without requiring
// a dense snapshot just to answer a single blank-cell lookup.
let sparseReadModel =
    { createEmpty 2000 100 with
        Cells = Map.ofList [ { Row=1999; Col=99 }, cell (CellValue.String "far") ] }
match V.runModel sparseReadModel "Dim x\nx = Range(\"A1\").Value\nx = Cells(1, 1).Text" with
| Result.Ok _ -> check "sparse runModel read" true true
| Result.Error e -> fail <- fail + 1; printfn "FAIL sparse runModel read: %A" e

let sparseSnapshot =
    V.snapshotFromModel
        { createEmpty 3 3 with
            Cells = Map.ofList [ { Row=2; Col=2 }, cell (CellValue.String "far") ] }
check "snapshot remains dense" sparseSnapshot.Length 9

// --- bytecode persistence: compile→serialize→deserialize→run == direct run ---
let bcSrc =
    String.concat "\n"
        [ "Dim i"
          "For i = 1 To 3"
          "  Cells(i + 2, 2).Value = i * i"
          "Next i"
          "Range(\"B1\").Value = \"done\"" ]
match V.compileToBytes bcSrc with
| Result.Error e -> fail <- fail + 1; printfn "FAIL compileToBytes: %s" e
| Result.Ok blob ->
    check "bytecode blob non-empty" (blob.Length > 8) true
    match V.runBytesOnModel model blob with
    | Result.Error e -> fail <- fail + 1; printfn "FAIL runBytesOnModel: %s" e
    | Result.Ok m ->
        // squares in B3:B5, "done" in B1 — same as running the source directly
        check "bc B3 = 1"  (disp m 2 1) "1"
        check "bc B4 = 4"  (disp m 3 1) "4"
        check "bc B5 = 9"  (disp m 4 1) "9"
        check "bc B1 done" (disp m 0 1) "done"
        // identical to the source path
        match V.runOnModel model bcSrc with
        | Result.Ok m2 ->
            check "bytecode == source path"
                [ for r in 2..4 -> disp m r 1 ] [ for r in 2..4 -> disp m2 r 1 ]
        | Result.Error e -> fail <- fail + 1; printfn "FAIL source path: %s" e

printfn ""
printfn "Stage 19 VBA tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
