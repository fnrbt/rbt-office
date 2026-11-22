#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open System
open System.IO
open Rbt.Spreadsheet.Model
module Bridge = Rbt.Spreadsheet.XlsxBridge
module Paint = Rbt.Spreadsheet.Paint
module Layout = Rbt.Spreadsheet.Layout

// Pivot styles as Excel 2019 draws them. A pivot table's cells hold its values only; its look is its style's
// (`<pivotTableStyleInfo>`: Excel's built-in PivotStyleLight1-28 / Medium1-28 / Dark1-28 of presetTableStyles.xml,
// whose dxfs a pivot style numbers from 0, or the workbook's own), laid over the parts of the table: the whole
// table and its page fields, the header rows, the row-label columns, banded rows and columns, the subheadings of
// outer row / column items, subtotal rows and columns, the grand total row and column.
// tests/test-documents-excel/synthetic-pivot-styles.xlsx was made in Excel (tools/excel-parity/build-pivot-styles.vbs:
// 12 pivot tables - compact, outline and tabular forms, page fields, subtotals at the top and the bottom, two levels of
// rows and columns, stripes, headers off - in 12 styles), and synthetic-pivot-styles.excel.tsv is what Excel shows of
// every cell around each (Range.DisplayFormat: fill, bold, font colour, the four edges).

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-56s %s" label detail)

let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let book =
    let wb = XlsxReader.loadBytesNative (File.ReadAllBytes(Path.Combine(root, "test-documents-excel/synthetic-pivot-styles.xlsx"))) |> Option.get
    Bridge.Import.importWorkbook wb
let sheetOf (name: string) = resolveSheet (book.Sheets |> Array.find (fun sheet -> sheet.Name = name)).Model

let cellAt (reference: string) =
    let letters = reference |> Seq.takeWhile Char.IsLetter |> Seq.toArray |> System.String
    let col = letters |> Seq.fold (fun acc ch -> acc * 26 + int ch - int 'A' + 1) 0
    int (reference.Substring letters.Length) - 1, col - 1

/// An edge as Excel's DisplayFormat gives it ("linestyle/weight/RRGGBB") in our terms: kind and colour.
let excelLine (text: string) =
    if text = "-" then None
    else
        let parts = text.Split '/'
        let kind =
            match parts.[0], parts.[1] with
            | "-4119", _ -> "double"
            | _, "-4138" -> "medium"
            | _, "4" -> "thick"
            | _ -> "thin"
        Some (kind, "#" + parts.[2])
let ourLine (edge: Layout.SharedEdge) =
    match edge with
    | Layout.EdgeBorder border when border.Style <> "none" -> Some (border.Style, defaultArg border.Color "#000000")
    | _ -> None

let rows =
    File.ReadAllLines(Path.Combine(root, "test-documents-excel/synthetic-pivot-styles.excel.tsv"))
    |> Array.map (fun line -> line.Split '\t')
    |> Array.filter (fun parts -> parts.Length >= 11 && parts.[1] <> "#")
let models = rows |> Array.map (fun parts -> parts.[0]) |> Array.distinct |> Array.map (fun name -> name, sheetOf name) |> Map.ofArray

ok "the fixture's pivot tables are read with their styles" (models |> Map.forall (fun _ model -> model.Pivots |> List.forall (fun pivot -> match pivot.Style with Some style -> not style.Elements.IsEmpty | None -> false)))
    (sprintf "%A" (models |> Map.map (fun _ model -> model.Pivots |> List.map (fun pivot -> pivot.Style.IsSome))))

for sheetName, model in Map.toList models do
    let wrong = ResizeArray<string>()
    let mine = rows |> Array.filter (fun parts -> parts.[0] = sheetName)
    for parts in mine do
        let row, col = cellAt parts.[1]
        let look = Paint.cellLook model row col (getCell { Row = row; Col = col } model) None 64.0
        let fill = look.Format.BgColor |> Option.map (fun c -> c.TrimStart('#').ToUpperInvariant()) |> Option.defaultValue "-"
        let bold = if look.Format.Bold then "True" else "False"
        let text = look.Format.TextColor |> Option.map (fun c -> c.TrimStart('#').ToUpperInvariant()) |> Option.defaultValue "000000"
        let edges = [ "top", ourLine (Layout.edgeAbove model row col), excelLine parts.[7]
                      "bottom", ourLine (Layout.edgeBelow model row col), excelLine parts.[8]
                      "left", ourLine (Layout.edgeLeft model row col), excelLine parts.[9]
                      "right", ourLine (Layout.edgeRight model row col), excelLine parts.[10] ]
        let differences =
            [ if fill <> parts.[3] then sprintf "fill %s, Excel %s" fill parts.[3]
              if bold <> parts.[4] then sprintf "bold %s, Excel %s" bold parts.[4]
              if text <> parts.[6] then sprintf "text %s, Excel %s" text parts.[6]
              for side, got, expected in edges do
                  if got <> expected then sprintf "%s %A, Excel %A" side got expected ]
        if not differences.IsEmpty then wrong.Add(sprintf "%s %s" parts.[1] (String.Join("; ", differences)))
    ok (sprintf "%s: every cell as Excel shows it (%d cells)" sheetName mine.Length) (wrong.Count = 0)
        (sprintf "%d differ: %s" wrong.Count (String.Join(" | ", wrong |> Seq.truncate 6)))

printfn "stage64-pivot-styles: %d passed, %d failed" pass fail
if fail > 0 then exit 1
