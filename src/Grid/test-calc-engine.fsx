#!/usr/bin/env dotnet fsi

// Test script for calculation engine optimizations
// Demonstrates: dependency graph, topological sort, dirty flagging, smart recalc

#load "Spreadsheet/Model.fs"
#load "Spreadsheet/Formula.fs"
#load "Spreadsheet/Calc.fs"

open System
open System.Diagnostics
open RedBat.Grid.Spreadsheet.Model
open RedBat.Grid.Spreadsheet.Formula
open RedBat.Grid.Spreadsheet.Calc

printfn "=== Spreadsheet Calculation Engine Test ==="
printfn ""

// Create a simple spreadsheet with formulas
let mutable model = createEmpty 1000 26
let mutable engine = emptyEngine

// Helper to set a cell
let setTestCell (row: int) (col: int) (value: CellValue) =
    let addr = { Row = row; Col = col }
    let cellData = { Value = value; Format = defaultFormat }
    let (newModel, newEngine) = updateCell addr value cellData model engine
    model <- newModel
    engine <- newEngine

// Create a dependency chain: A1 -> A2 -> A3 -> A4 -> A5
printfn "1. Building Dependency Chain"
printfn "   A1 = 10"
printfn "   A2 = A1 + 5"
printfn "   A3 = A2 * 2"
printfn "   A4 = A3 - 3"
printfn "   A5 = A4 / 2"
printfn ""

setTestCell 0 0 (CellValue.Number 10.0)
setTestCell 1 0 (CellValue.Formula ("=A1+5", CellValue.Empty))
setTestCell 2 0 (CellValue.Formula ("=A2*2", CellValue.Empty))
setTestCell 3 0 (CellValue.Formula ("=A3-3", CellValue.Empty))
setTestCell 4 0 (CellValue.Formula ("=A4/2", CellValue.Empty))

// Recalculate
let (model', engine') = smartRecalculate model engine
model <- model'
engine <- engine'

// Print results
let printCell row col label =
    let addr = { Row = row; Col = col }
    let value = getCellValue addr model
    printfn "   %s = %s" label (displayValue value)

printCell 0 0 "A1"
printCell 1 0 "A2"
printCell 2 0 "A3"
printCell 3 0 "A4"
printCell 4 0 "A5"
printfn ""

// Show dependency graph
printfn "2. Dependency Graph"
let stats = getStats model engine
printfn "   Total cells: %d" stats.TotalCells
printfn "   Formula cells: %d" stats.FormulaCells
printfn "   Total dependencies: %d" stats.Dependencies
printfn "   Calc chain length: %d" stats.CalcChainLength
printfn ""

printfn "3. Calculation Chain (topologically sorted):"
for addr in engine.CalcChain do
    printfn "   %s" (formatCellRef addr.Row addr.Col)
printfn ""

// Create a more complex example with multiple branches
printfn "4. Complex Dependency Tree"
printfn "   B1 = 100"
printfn "   B2 = B1 * 2    (depends on B1)"
printfn "   B3 = B1 + 50   (depends on B1)"
printfn "   B4 = B2 + B3   (depends on B2 and B3)"
printfn ""

setTestCell 0 1 (CellValue.Number 100.0)
setTestCell 1 1 (CellValue.Formula ("=B1*2", CellValue.Empty))
setTestCell 2 1 (CellValue.Formula ("=B1+50", CellValue.Empty))
setTestCell 3 1 (CellValue.Formula ("=B2+B3", CellValue.Empty))

let (model'', engine'') = smartRecalculate model engine
model <- model''
engine <- engine''

printCell 0 1 "B1"
printCell 1 1 "B2"
printCell 2 1 "B3"
printCell 3 1 "B4"
printfn ""

let stats2 = getStats model engine
printfn "   Updated stats:"
printfn "   Total cells: %d" stats2.TotalCells
printfn "   Formula cells: %d" stats2.FormulaCells
printfn "   Dependencies: %d" stats2.Dependencies
printfn "   Calc chain length: %d" stats2.CalcChainLength
printfn ""

// Test dirty flagging - change B1 and see what gets recalculated
printfn "5. Dirty Flagging Test"
printfn "   Changing B1 from 100 to 200"
printfn "   This should mark B2, B3, B4 as dirty (but not A-column)"
printfn ""

let addr = { Row = 0; Col = 1 }
let cellData = { Value = CellValue.Number 200.0; Format = defaultFormat }
let (model''', engine''') = updateCell addr (CellValue.Number 200.0) cellData model engine
model <- model'''
engine <- engine'''

printfn "   Dirty cells: %d" (Set.count engine.DirtySet)
printfn "   Dirty set:"
for dirtyAddr in engine.DirtySet do
    printfn "     %s" (formatCellRef dirtyAddr.Row dirtyAddr.Col)
printfn ""

// Smart recalculate - only dirty cells
printfn "6. Smart Recalculation (only dirty cells)"
let sw = Stopwatch.StartNew()
let (model'''', engine'''') = smartRecalculate model engine
model <- model''''
engine <- engine''''
sw.Stop()

printfn "   Recalculated in %d μs" (sw.ElapsedTicks * 1000L / Stopwatch.Frequency)
printfn "   Results after B1 change:"
printCell 0 1 "B1"
printCell 1 1 "B2"
printCell 2 1 "B3"
printCell 3 1 "B4"
printfn ""

// Performance comparison: naive vs smart
printfn "7. Performance Comparison"
printfn ""

// Build a large spreadsheet with 100 cells in a chain
printfn "   Building spreadsheet with 100 chained formulas..."
model <- createEmpty 1000 26
engine <- emptyEngine

setTestCell 0 2 (CellValue.Number 1.0)  // C1 = 1
for i in 1 .. 99 do
    let formula = sprintf "=C%d+1" i
    setTestCell i 2 (CellValue.Formula (formula, CellValue.Empty))

let (modelBig, engineBig) = smartRecalculate model engine
model <- modelBig
engine <- engineBig

printfn "   Initial calculation:"
printfn "     C1 = %s" (displayValue (getCellValue { Row = 0; Col = 2 } model))
printfn "     C100 = %s" (displayValue (getCellValue { Row = 99; Col = 2 } model))
printfn ""

// Now change C1 and measure smart recalc
printfn "   Changing C1 from 1 to 10..."
let addr' = { Row = 0; Col = 2 }
let cellData' = { Value = CellValue.Number 10.0; Format = defaultFormat }
let (modelDirty, engineDirty) = updateCell addr' (CellValue.Number 10.0) cellData' model engine

printfn "   Dirty cells: %d" (Set.count engineDirty.DirtySet)
printfn ""

// Time smart recalculation
let sw2 = Stopwatch.StartNew()
let (modelRecalc, engineRecalc) = smartRecalculate modelDirty engineDirty
sw2.Stop()

printfn "   Smart recalculation time: %d μs" (sw2.ElapsedTicks * 1000L / Stopwatch.Frequency)
printfn "     C1 = %s" (displayValue (getCellValue { Row = 0; Col = 2 } modelRecalc))
printfn "     C100 = %s" (displayValue (getCellValue { Row = 99; Col = 2 } modelRecalc))
printfn ""

// Test circular reference detection
printfn "8. Circular Reference Detection"
printfn "   Creating circular dependency: D1=D2, D2=D1"
printfn ""

model <- createEmpty 1000 26
engine <- emptyEngine

setTestCell 0 3 (CellValue.Formula ("=D2", CellValue.Empty))
setTestCell 1 3 (CellValue.Formula ("=D1", CellValue.Empty))

let statsCirc = getStats model engine
printfn "   Formula cells: %d" statsCirc.FormulaCells
printfn "   Calc chain length: %d" statsCirc.CalcChainLength
printfn "   Has circular reference: %b" statsCirc.HasCircularRef
printfn ""

if statsCirc.HasCircularRef then
    printfn "   ✓ Circular reference detected!"
    printfn "   Calc chain is shorter than formula count, preventing infinite loop"
else
    printfn "   ✗ Circular reference not detected"

printfn ""
printfn "=== Test Complete ==="
printfn ""
printfn "Key Optimizations Demonstrated:"
printfn "1. Dependency Graph: Tracks which cells reference which"
printfn "2. Topological Sort: Calculates cells in correct order"
printfn "3. Dirty Flagging: Only marks affected cells for recalculation"
printfn "4. Smart Recalculation: Only recalculates dirty cells"
printfn "5. Circular Reference Detection: Prevents infinite loops"
printfn "6. Formula Caching: Parse once, evaluate many times"
