#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Vb.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// Excel's object model for macros (MacroHost): macros as people write them, run on a workbook, and
// what the workbook holds afterwards.

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Workbook
module M = Rbt.Spreadsheet.MacroHost

let mutable pass = 0
let mutable fail = 0
let check label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-50s expected %A got %A" label expected got

let output = ResizeArray<string>()
let hooks = { M.defaultHooks with Output = output.Add; FileName = "Book.xlsm" }

/// A workbook of two sheets: Data (A1:B4 a small table) and Report (empty).
let start () =
    let wb = createWorkbook ()
    let wb = renameSheet wb.ActiveSheet "Data" wb
    let wb = addSheet "Report" wb
    let data = wb.SheetOrder.[0]
    let set (r, c, value) (wb: Workbook) =
        let model = wb.Sheets.[data].Model
        updateSheetModel data (setCell { Row = r; Col = c } { Value = value; Format = defaultFormat; Comment = None } model) wb
    [ 0, 0, String "Name"; 0, 1, String "Amount"
      1, 0, String "Ann"; 1, 1, Number 10.0
      2, 0, String "Bob"; 2, 1, Number 20.0
      3, 0, String "Cy"; 3, 1, Number 30.0 ]
    |> List.fold (fun acc cell -> set cell acc) (setActiveSheet data wb)

let program (source: string) =
    match M.MacroProgram.Compile [ Rbt.Vb.Api.ModuleSource.standard "Module1" source ] with
    | Result.Ok compiled -> compiled
    | Result.Error failures -> failwithf "compile: %s" (failures |> List.map Rbt.Vb.Api.VbFailure.format |> String.concat "; ")
let tryRun source (procedure: string option) = (program source).Run(start (), hooks, defaultArg procedure "Main")
let run source procedure =
    match tryRun source procedure with
    | M.MacroDone (wb, _) -> wb
    | M.MacroFailed failure -> failwithf "macro failed: %s" (Rbt.Vb.Api.VbFailure.format failure)

let sheetNamed (wb: Workbook) name = wb.SheetOrder |> List.map (fun id -> wb.Sheets.[id]) |> List.find (fun s -> s.Name = name)
let valueAt (wb: Workbook) sheet (r, c) = getCellValue { Row = r; Col = c } (sheetNamed wb sheet).Model |> unwrapFormulaValue
let formatAt (wb: Workbook) sheet (r, c) = getCell { Row = r; Col = c } (sheetNamed wb sheet).Model |> Option.map (fun c -> c.Format) |> Option.defaultValue defaultFormat

// ---- values are typed ----
let typedSource = """
Sub Main()
    Range("C2").Value = Range("B2").Value + 1
    Range("C3").Value = Range("A2").Value & "!"
    Range("C4").Value = TypeName(Range("B3").Value)
End Sub"""
let typed = run typedSource (Some "Main")
check "a number reads as a number (10 + 1)" (valueAt typed "Data" (1, 2)) (Number 11.0)
check "text reads as text" (valueAt typed "Data" (2, 2)) (String "Ann!")
check "TypeName of a number cell" (valueAt typed "Data" (3, 2)) (String "Double")

// ---- writes are entered as typing would enter them; formulas calculate before the next read ----
let enteredSource = """
Sub Main()
    Range("D1").Value = "123"
    Range("D2").Value = "=SUM(B2:B4)"
    Range("D3").Value = Range("D2").Value * 2
    Range("D4").Formula = "=D1+1"
    Range("E1").Value = Range("D4").Formula
End Sub"""
let entered = run enteredSource (Some "Main")
check "\"123\" becomes a number" (valueAt entered "Data" (0, 3)) (Number 123.0)
check "a formula is calculated" (valueAt entered "Data" (1, 3)) (Number 60.0)
check "and read back calculated within the macro" (valueAt entered "Data" (2, 3)) (Number 120.0)
check "Formula read back and entered again: the formula" (getCellValue { Row = 0; Col = 4 } (sheetNamed entered "Data").Model) (Formula ("=D1+1", Number 124.0))

// ---- loops over cells, Rows.Count, End(xlUp) ----
let loopSource = """
Sub Main()
    Dim r, last, total
    last = Cells(Rows.Count, 2).End(xlUp).Row
    For r = 2 To last
        total = total + Cells(r, 2).Value
        Cells(r, 3).Value = Cells(r, 2).Value * 2
    Next
    Range("F1").Value = last
    Range("F2").Value = total
    Range("F3").Value = Range("A1").End(xlDown).Address
    Range("F4").Value = Range("A1").End(xlToRight).Address(False, False)
End Sub"""
let loop = run loopSource (Some "Main")
check "End(xlUp) from the bottom finds the last row" (valueAt loop "Data" (0, 5)) (Number 4.0)
check "the total of the loop" (valueAt loop "Data" (1, 5)) (Number 60.0)
check "Cells(r, 3) written in the loop" (valueAt loop "Data" (3, 2)) (Number 60.0)
check "End(xlDown)" (valueAt loop "Data" (2, 5)) (String "$A$4")
check "End(xlToRight), relative address" (valueAt loop "Data" (3, 5)) (String "B1")

// ---- several sheets ----
let sheetsSource = """
Sub Main()
    Worksheets("Report").Range("A1").Value = "Total"
    Worksheets("Report").Range("B1").Value = WorksheetFunction.Sum(Worksheets("Data").Range("B2:B4"))
    Sheets(2).Range("A2").Value = Worksheets.Count
    Worksheets("Report").Name = "Summary"
    Dim ws
    Set ws = Worksheets.Add(, Worksheets("Summary"))
    ws.Name = "Extra"
    ws.Range("A1").Value = ActiveSheet.Name
End Sub"""
let sheets = run sheetsSource (Some "Main")
check "a write to another sheet" (valueAt sheets "Summary" (0, 0)) (String "Total")
check "WorksheetFunction.Sum over a range of another sheet" (valueAt sheets "Summary" (0, 1)) (Number 60.0)
check "Worksheets.Count" (valueAt sheets "Summary" (1, 0)) (Number 2.0)
check "renamed, and one added after it" (sheets.SheetOrder |> List.map (fun id -> sheets.Sheets.[id].Name)) [ "Data"; "Summary"; "Extra" ]
check "the added sheet is the active one" (valueAt sheets "Extra" (0, 0)) (String "Extra")

// ---- formats ----
let formatsSource = """
Sub Main()
    With Range("A1:B1")
        .Font.Bold = True
        .Font.Color = RGB(255, 0, 0)
        .Interior.Color = RGB(0, 0, 255)
        .HorizontalAlignment = xlCenter
    End With
    Range("B2:B4").NumberFormat = "0.00"
    Debug.Print Range("B2").Text
    Range("C2").Value = Range("A1").Font.Bold
End Sub"""
let formats = run formatsSource (Some "Main")
check "Font.Bold" (formatAt formats "Data" (0, 1)).Bold true
check "Font.Color (BGR) as #RRGGBB" (formatAt formats "Data" (0, 0)).TextColor (Some "#FF0000")
check "Interior.Color" (formatAt formats "Data" (0, 0)).BgColor (Some "#0000FF")
check "HorizontalAlignment" (formatAt formats "Data" (0, 0)).HorizontalAlign (Some "center")
check "NumberFormat" (formatAt formats "Data" (2, 1)).NumberFormat (Some "0.00")
check "Text shows the format" (output |> Seq.tryLast) (Some "10.00")
check "Font.Bold reads back" (valueAt formats "Data" (1, 2)) (Boolean true)

// ---- arrays to and from ranges ----
let arraysSource = """
Sub Main()
    Dim v
    v = Range("A2:B4").Value
    Range("E1").Value = UBound(v, 1) & "x" & UBound(v, 2)
    Range("E2").Value = v(3, 1)
    Range("G1:H2").Value = Range("A1:B2").Value
    Range("G4:I4").Value = Array(1, 2, 3)
End Sub"""
let arrays = run arraysSource (Some "Main")
check "Value of a block is a 2-D array" (valueAt arrays "Data" (0, 4)) (String "3x2")
check "1-based items" (valueAt arrays "Data" (1, 4)) (String "Cy")
check "a block assigned a block" (valueAt arrays "Data" (1, 7)) (Number 10.0)
check "a row assigned a 1-D array" (valueAt arrays "Data" (3, 8)) (Number 3.0)

// ---- Offset, Resize, CurrentRegion, UsedRange, Find ----
let navigationSource = """
Sub Main()
    Range("A1").Offset(1, 1).Value = 99
    Range("E1").Value = Range("A1").Resize(2, 3).Address
    Range("E2").Value = Range("A2").CurrentRegion.Address
    Range("E3").Value = ActiveSheet.UsedRange.Rows.Count
    Range("E4").Value = Range("A1:A4").Find("bob").Row
End Sub"""
let navigation = run navigationSource (Some "Main")
check "Offset" (valueAt navigation "Data" (1, 1)) (Number 99.0)
check "Resize" (valueAt navigation "Data" (0, 4)) (String "$A$1:$C$2")
check "CurrentRegion" (valueAt navigation "Data" (1, 4)) (String "$A$1:$B$4")
check "UsedRange.Rows.Count (before E3 itself)" (valueAt navigation "Data" (2, 4)) (Number 4.0)
check "Find, case-insensitive" (valueAt navigation "Data" (3, 4)) (Number 3.0)

// ---- clearing, rows, copying ----
let editsSource = """
Sub Main()
    Rows(3).Delete
    Range("B2").ClearContents
    Range("A1:B1").Copy Range("D1")
    Range("A2").EntireRow.Insert
End Sub"""
let edits = run editsSource (Some "Main")
check "row 3 deleted, row 4 moved up and then down by the insert" (valueAt edits "Data" (3, 0)) (String "Cy")
check "the inserted row is empty" (valueAt edits "Data" (1, 0)) CellValue.Empty
check "ClearContents" (valueAt edits "Data" (2, 1)) CellValue.Empty
check "Copy to a destination" (valueAt edits "Data" (0, 4)) (String "Amount")

// ---- Evaluate, names, selection ----
let miscSource = """
Sub Main()
    Range("E1").Value = Evaluate("SUM(B2:B4)*2")
    ActiveWorkbook.Names.Add "Amounts", Range("B2:B4")
    Range("E2").Value = WorksheetFunction.Max(Range("Amounts"))
    Range("B3").Select
    ActiveCell.Offset(0, 1).Value = "next to Bob"
    Range("E3").Value = Selection.Address
End Sub"""
let misc = run miscSource (Some "Main")
check "Evaluate" (valueAt misc "Data" (0, 4)) (Number 120.0)
check "a name added and used" (valueAt misc "Data" (1, 4)) (Number 30.0)
check "ActiveCell after Select" (valueAt misc "Data" (2, 2)) (String "next to Bob")
check "Selection" (valueAt misc "Data" (2, 4)) (String "$B$3")

// ---- what Excel raises ----
let failed = tryRun "Sub Main()\n    Worksheets(\"Nope\").Range(\"A1\").Value = 1\nEnd Sub" (Some "Main")
check "a missing sheet stops the macro" (match failed with M.MacroFailed _ -> true | _ -> false) true
let unknown = tryRun "Sub Main()\n    Range(\"A1\").Frobnicate = 1\nEnd Sub" (Some "Main")
check "an unknown member stops the macro" (match unknown with M.MacroFailed _ -> true | _ -> false) true
check "nothing of a failed macro is kept" (match tryRun "Sub Main()\n    Range(\"A1\").Value = 5\n    Worksheets(\"Nope\").Activate\nEnd Sub" (Some "Main") with M.MacroFailed _ -> true | _ -> false) true

// ---- what the engine's 0.3.0 host bridge allows: default members, For Each, named and omitted arguments ----
let bridgeSource = """
Sub Main()
    Dim c, total, ws, names
    For Each c In Range("B2:B4")
        total = total + c
        c.Offset(0, 1) = c * 10
    Next
    Range("E1") = total
    Range("E2") = Range("E1") + 1
    For Each ws In Worksheets
        names = names & ws.Name & ";"
    Next
    Range("E3") = names
    Range("E4") = Range("A1:A4").Find(What:="cy").Row
    Worksheets.Add After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "Last"
    Worksheets("Data").Range("E5") = Worksheets("Data").Range("A2").Address(, False)
End Sub"""
let bridge = run bridgeSource (Some "Main")
check "For Each over a range, reading its default member" (valueAt bridge "Data" (0, 4)) (Number 60.0)
check "assigning through the default member" (valueAt bridge "Data" (3, 2)) (Number 300.0)
check "Range(..) = .. and Range(..) + 1" (valueAt bridge "Data" (1, 4)) (Number 61.0)
check "For Each over Worksheets" (valueAt bridge "Data" (2, 4)) (String "Data;Report;")
check "a named argument (Find What:=)" (valueAt bridge "Data" (3, 4)) (Number 4.0)
check "Worksheets.Add After:=" (bridge.SheetOrder |> List.map (fun id -> bridge.Sheets.[id].Name)) [ "Data"; "Report"; "Last" ]
check "an omitted argument (Address(, False))" (valueAt bridge "Data" (4, 4)) (String "A$2")

// ---- errors by number, trapped by On Error ----
let errorsSource = """
Sub Main()
    On Error Resume Next
    Worksheets("Nope").Activate
    Worksheets("Data").Range("E1") = Err.Number
    Err.Clear
    Range("A1").Frobnicate = 1
    Worksheets("Data").Range("E2") = Err.Number
    Err.Clear
    Range("A1").Offset(-5, 0).Value = 1
    Worksheets("Data").Range("E3") = Err.Number
End Sub"""
let errors = run errorsSource (Some "Main")
check "a missing sheet: 9" (valueAt errors "Data" (0, 4)) (Number 9.0)
check "an unknown member: 438" (valueAt errors "Data" (1, 4)) (Number 438.0)
check "off the sheet: 1004" (valueAt errors "Data" (2, 4)) (Number 1004.0)
match tryRun "Sub Main()\n    Range(\"A1\").Value = 5\n    Worksheets(\"Nope\").Activate\nEnd Sub" (Some "Main") with
| M.MacroFailed failure -> check "an untrapped error: its number, procedure and line" (failure.Number, failure.Procedure, failure.Line) (9, "Main", 3)
| M.MacroDone _ -> check "an untrapped error stops the macro" false true

// ---- a program keeps its module state between runs; the Immediate window; events; a runaway ----
let counter = program "Dim count\nSub Bump()\n    count = count + 1\n    Range(\"F1\") = count\nEnd Sub\nSub Spin()\n    Do\n    Loop\nEnd Sub\nSub Worksheet_Change(ByVal Target As Range)\n    Target.Offset(0, 1) = Target.Value * 2\nEnd Sub"
let afterBump wb = match counter.Run(wb, hooks, "Bump") with M.MacroDone (next, _) -> next | M.MacroFailed f -> failwith f.Description
let twice = afterBump (afterBump (start ()))
check "module variables survive between runs" (valueAt twice "Data" (0, 5)) (Number 2.0)
output.Clear()
match counter.Immediate(start (), hooks, "?Range(\"B2\").Value * 2") with
| M.MacroDone _ -> check "an Immediate expression prints its value" (output |> Seq.map (fun s -> s.Trim()) |> Seq.toList) [ "20" ]
| M.MacroFailed f -> check "Immediate" f.Description ""
let changed = match counter.RunEvent(start (), hooks, "Worksheet_Change", Some ((start ()).SheetOrder.[0], 2, 1, 2, 1)) with Some (M.MacroDone (wb, _)) -> wb | _ -> failwith "event"
check "an event handler gets its Target" (valueAt changed "Data" (2, 2)) (Number 40.0)
check "no handler, no run" (counter.RunEvent(start (), hooks, "Workbook_Open", None)).IsNone true
match counter.Run(start (), hooks, "Spin", maxInstructions = 1_000_000) with
| M.MacroFailed failure -> check "a runaway macro is stopped" failure.Number 18
| M.MacroDone _ -> check "a runaway macro is stopped" false true

// ---- 0.4.0: dates, errors, Format, collections ----
let datesSource = """
Sub Main()
    Range("D1").Value = DateSerial(2024, 1, 31)
    Range("D2").Value = Range("D1").Value + 1
    Range("E1").Value = TypeName(Range("D1").Value) & "/" & TypeName(Range("D1").Value2)
    Range("E2").Value = Format(Range("D2").Value, "yyyy-mm-dd")
    Range("E3").Value = Year(Range("D1").Value)
    Range("D3").Formula = "=1/0"
    Range("E4").Value = IsError(Range("D3").Value) & "/" & Range("D3").Text
    Range("D4").Value = CVErr(xlErrNA)
    Dim d
    Set d = CreateObject("Scripting.Dictionary")
    d("Ann") = 1: d("Bob") = 2
    Range("E5").Value = d.Count & ":" & Join(d.Keys, ",")
    Range("E6").Value = WorksheetFunction.Sum(Range("D1:D2")) = 2 * CDbl(DateSerial(2024, 1, 31)) + 1
End Sub"""
let dates = run datesSource (Some "Main")
check "a Date written gets a date format" ((formatAt dates "Data" (0, 3)).NumberFormat) (Some "m/d/yyyy")
check "and its serial" (valueAt dates "Data" (0, 3)) (Number 45322.0)
check "Value is a Date, Value2 a Double" (valueAt dates "Data" (0, 4)) (String "Date/Double")
check "date arithmetic and Format (the text entered as a date, as typing would)" (valueAt dates "Data" (1, 4)) (Number 45323.0)
check "Year of a cell's date" (valueAt dates "Data" (2, 4)) (Number 2024.0)
check "an error cell reads as an Error value" (valueAt dates "Data" (3, 4)) (String "True/#DIV/0!")
check "CVErr written is the error" (valueAt dates "Data" (3, 3)) (Error "N/A")
check "Scripting.Dictionary" (valueAt dates "Data" (4, 4)) (String "2:Ann,Bob")
check "dates into WorksheetFunction" (valueAt dates "Data" (5, 4)) (Boolean true)

// ---- the macro list of a project of several modules ----
let modules =
    [ M.macroModule "Module1" "Attribute VB_Name = \"Module1\"\nSub Main()\nEnd Sub\nPrivate Sub Hidden()\nEnd Sub\nSub WithArgs(x)\nEnd Sub\nFunction F()\nEnd Function"
      M.macroModule "Sheet1" "Attribute VB_Name = \"Sheet1\"\nAttribute VB_Base = \"0{00020820-0000-0000-C000-000000000046}\"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = True\nSub Hello()\nEnd Sub"
      M.macroModule "Class1" "Attribute VB_Name = \"Class1\"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = False\nAttribute VB_Exposed = False\nPublic Value\n" ]
check "module kinds from their Attribute lines" (modules |> List.map (fun m -> m.Kind)) [ Rbt.Vb.Api.Standard; Rbt.Vb.Api.Document; Rbt.Vb.Api.Class ]
match M.MacroProgram.Compile modules with
| Result.Ok project -> check "the macros: standard ones by name, a sheet's as Sheet.Name" project.Macros [ "Main"; "Sheet1.Hello" ]
| Result.Error f -> check "the project compiles" (f |> List.map (fun x -> x.Description)) []

printfn "Stage 57 macro host tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
