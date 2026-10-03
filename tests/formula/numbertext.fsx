#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// NumberText.formatInvariant is the portable formatter the Fable build uses in
// place of Double.ToString. It must agree with the .NET runtime formatter, so
// every pattern/value pair is checked against it.
open System
open Rbt.Spreadsheet

let inv = Globalization.CultureInfo.InvariantCulture

let formats =
    [ "G"; "G15"; "G11"; "G10"; "G7"; "G5"; "G3"; "G2"; "G1"; "g4"
      "F0"; "F2"; "F5"; "N0"; "N2"; "N4"; "E"; "E0"; "E2"; "e3"; "E10"
      "0"; "0.0"; "0.00"; "0.000"; "#"; "#.##"; "#.00"; "0.##"; "0.0#"; "#,##0"; "#,##0.00"; "#,###"; "#,#"; "#,##0.0#"
      "000"; "00000"; "000.00"; "0%"; "0.0%"; "0.00%"; "#,##0%"; "0.00‰"
      "0.00E+00"; "0.0E+0"; "#.##E+00"; "0E+00"; "0.00e-00"; "0.000E-0"; "##0.0E+0"
      "0,"; "0,,"; "0.0,"; "#,##0,"; "0.00,,\"M\""; "#,##0.00,\"K\""
      "$#,##0.00"; "\"$\"#,##0.00"; "(#,##0.00)"; "0.00 \"units\""; "'USD' 0.00"; "\\#0.0"; "0.00\\%"; "[0.00]"; "0.0x"
      "# ##0.00"; "0-00-00"; "00:00"; "0.0.0"; "0.00E"; "0.00E+"; "E0"; "abc"; ".00"; "#." ]

let values =
    [ 0.0; 1.0; -1.0; 0.5; -0.5; 0.1; 0.2; 0.3; 0.1 + 0.2; 1.5; 2.5; -2.5; 0.05; 0.005; 0.0049; 0.995; 0.9995; 9.995; 99.5; 999.5
      1.005; 2.675; 1.115; 12.345; 123.456; 1234.5; 1234.5678; -1234.5678; 12345.6789; 123456.789; 1234567.891; 999999.999; 9999999.5
      1e-5; 1.5e-5; 1e-4; 1.234e-4; 0.000999; 1e-7; -1e-7; 1e-10; 1.23456789e-12
      1e10; 1e14; 1e15; 123456789012345.0; 1234567890123456.0; 1e16; 1.5e20; 1e21; 1e22; 5e-324
      45474.0; 45474.75; 43465.0; 0.080209174; 61000.0; 20500000.0; 3.14159265358979; 2.718281828459045; 1.0 / 3.0; 2.0 / 3.0; 100.0; 1000.0; 0.9; 0.01; -0.001; -0.004; -0.0049
      0.125; 0.375; 3.5; 4.5; 0.25; 0.75; 1234.5; 12345.5; 0.0625; 1.0e23; 8.5e-3 ]

let mutable passed = 0
let mutable failed = 0
for format in formats do
    for value in values do
        let expected = value.ToString(format, inv)
        let actual = NumberText.formatInvariant format value
        if actual = expected then passed <- passed + 1
        else
            failed <- failed + 1
            if failed <= 40 then printfn "FAIL %-14s %-24s expected %-28s got %s" format (value.ToString("R", inv)) expected actual

for value in [ nan; infinity; -infinity ] do
    for format in [ "G"; "0.00"; "N2" ] do
        if NumberText.formatInvariant format value = value.ToString(format, inv) then passed <- passed + 1
        else
            failed <- failed + 1
            printfn "FAIL %s %f" format value

printfn "NumberText tests: %d passed, %d failed" passed failed
if failed > 0 then exit 1
