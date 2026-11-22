#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open System
module Cf = Rbt.Spreadsheet.ConditionalFormatting

// The icons of Excel's 20 icon sets as Excel 2019 draws them: 16 x 16 px in Excel's flat colours.
// tests/test-documents-excel/synthetic-icon-sets.xlsx (tools/excel-parity/build-icon-sets.vbs) holds every set on a
// row of 1 .. 5; the colours below are the most common pixel of each icon in Excel's picture of it.

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-56s %s" label detail)

let red, amber, green, grey, dark, blue = "#D65532", "#EAC282", "#68A490", "#808080", "#505050", "#4A7DB1"
let excel =
    [ "3Arrows", [ red; amber; green ]; "3ArrowsGray", [ grey; grey; grey ]; "3Flags", [ red; amber; green ]
      "3TrafficLights1", [ red; amber; green ]; "3TrafficLights2", [ red; amber; green ]; "3Signs", [ red; amber; green ]
      "3Symbols", [ red; amber; green ]; "3Symbols2", [ red; amber; green ]; "4Arrows", [ red; amber; amber; green ]
      "4ArrowsGray", [ grey; grey; grey; grey ]; "4RedToBlack", [ dark; "#B1B1B1"; "#EDB2A3"; "#D86344" ]
      "4Rating", [ blue; blue; blue; blue ]; "4TrafficLights", [ dark; red; amber; green ]
      "5Arrows", [ red; amber; amber; amber; green ]; "5ArrowsGray", [ grey; grey; grey; grey; grey ]
      "5Rating", [ "#B3B3B3"; blue; blue; blue; blue ]; "5Quarters", [ dark; dark; dark; dark; dark ]
      "3Stars", [ "#727272"; amber; amber ]; "3Triangles", [ red; amber; green ]; "5Boxes", [ "#B3B3B3"; blue; blue; blue; blue ] ]
for family, colors in excel do
    colors |> List.iteri (fun index color ->
        let svg = Cf.iconSvg family index
        ok (sprintf "%s icon %d: 16 px in Excel's %s" family index color)
            (svg.Contains "width='16' height='16'" && svg.Contains color) svg)

printfn "stage66-icon-sets: %d passed, %d failed" pass fail
if fail > 0 then exit 1
