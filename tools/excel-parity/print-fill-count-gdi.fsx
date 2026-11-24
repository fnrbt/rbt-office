// Hinted GDI advances of the fill-count faces (check-print-fill-count.py). dotnet fsi tools/excel-parity/print-fill-count-gdi.fsx > print-fill-count-gdi.tsv
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"
open Rbt.Spreadsheet
printfn "face\tsize\tscale\tppem\tchar\tadv"
for (f, s, sc) in ["Aptos Narrow",14.,100.; "Arial",11.,60.; "Calibri",11.,100.] do
  let px = int (floor (s * sc / 100.0 * 600.0 / 72.0 + 0.5))
  for c in "0123456789.- ,$#" do
    match ExcelText.advance f false false px false (int c) with
    | ValueSome w -> printfn "%s\t%g\t%g\t%d\t%c\t%d" f s sc px c w
    | ValueNone -> ()
