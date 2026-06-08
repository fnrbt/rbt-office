#r "/home/dave/Dev/browser/office/src/Office.Formats/bin/Debug/netstandard2.0/OfficeFormat.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/VBScriptParser.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

// vbaProject.bin: write a source-only macro project, read it back through the
// CFB + MS-OVBA path, and run the recovered macro on a model.

open RedBat.Spreadsheet.Model
module V = RedBat.Spreadsheet.Vba

let mutable pass = 0
let mutable fail = 0
let check label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-28s expected %A got %A" label expected got
let disp m r c = displayValue (getCellValue { Row=r; Col=c } m)

let source =
    String.concat "\n"
        [ "Sub Main()"
          "  Range(\"A1\").Value = 42"
          "  Range(\"A2\").Value = Range(\"A1\").Value * 2"
          "End Sub"
          "Main()" ]

// 1. write → valid CFB
let bin = VbaProject.writeProject "Module1" source
check "CFB signature"  (bin.[0], bin.[1]) (0xD0uy, 0xCFuy)
check "stream count"   (VbaProject.cfbRead bin |> List.length) 5
check "module recovered" (VbaProject.readModules bin |> List.map fst) [ "Module1" ]

// 2. read back the source (Attribute lines stripped) — round-trips
let recovered = VbaProject.readMacroSource bin
check "source round-trips" recovered source

// 3. run the recovered macro on a model
match V.runOnModel (createEmpty 10 8) recovered with
| Result.Error e -> fail <- fail + 1; printfn "FAIL recovered macro errored: %s" e
| Result.Ok m ->
    check "recovered A1" (disp m 0 0) "42"
    check "recovered A2" (disp m 1 0) "84"

// 4. larger source crossing an OVBA chunk boundary still round-trips
let big = String.concat "\n" [ for i in 1..400 -> sprintf "' comment line %d padding padding padding" i ] + "\nRange(\"A1\").Value = 7"
let bigBin = VbaProject.writeProject "M" big
check "big source round-trips" (VbaProject.readMacroSource bigBin) big

printfn ""
printfn "Stage 20 vbaProject.bin tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
