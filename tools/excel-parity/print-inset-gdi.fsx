// Hinted GDI widths of the inset fixture texts (lock-print-inset.py). dotnet fsi tools/excel-parity/print-inset-gdi.fsx > GDI.tsv
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"
open Rbt.Spreadsheet
let faces = ["Aptos Narrow";"Arial";"Calibri";"Georgia";"Tahoma";"Verdana"]
printfn "face\tsize\tscale\tppem\tdigit\tw_pos\tw_neg\tw_zero"
for f in faces do
  for s in [9.;10.;11.;12.;14.] do
    for sc in [60.;75.;100.] do
      let px = int (floor (s * sc / 100.0 * 600.0 / 72.0 + 0.5))
      let adv c = match ExcelText.advance f false false px false (int c) with ValueSome w -> w | ValueNone -> -1000
      let d = [ for c in '0'..'9' -> adv c ] |> List.max
      let w (t:string) = t |> Seq.sumBy adv
      printfn "%s\t%g\t%g\t%d\t%d\t%d\t%d\t%d" f s sc px d (w "1234.50") (w "-1234.50") (w "0.00")
