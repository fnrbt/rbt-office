#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open System
open System.IO
open Rbt.Spreadsheet.Model
module Bridge = Rbt.Spreadsheet.XlsxBridge
module Ops = Rbt.Spreadsheet.SheetOps
module Paint = Rbt.Spreadsheet.Paint
module Layout = Rbt.Spreadsheet.Layout

// Table styles as Excel draws them: the built-in styles are Excel's own definitions (the
// `presetTableStyles.xml` of ECMA-376: every element of each style, its fill, font and lines, in
// the workbook's theme colours), a custom style is the `<tableStyle>` of styles.xml with its dxfs;
// the elements are laid over each other in Excel's order (whole table, column and row stripes, last
// and first column, header and total rows, their first and last cells), each element's lines on the
// edges of the block it covers (a stripe two rows wide is one block), a cell's own format above all.
// tests/test-documents-excel/synthetic-table-styles.xlsx (tools/excel-parity/make-table-styles-fixture.py)
// holds a table in each family and one in a custom style that uses every element.

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-56s %s" label detail)
let eq label (got: 'a) (expected: 'a) = ok label (got = expected) (sprintf "expected %A  got %A" expected got)

let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let sheetOf (file: string) (name: string) =
    let wb = XlsxReader.loadBytesNative (File.ReadAllBytes(Path.Combine(root, file))) |> Option.get
    let book = Bridge.Import.importWorkbook wb
    let sheet = book.Sheets |> Array.find (fun sheet -> sheet.Name = name)
    resolveSheet sheet.Model

/// "B4" to a 0-based cell.
let cellAt (reference: string) =
    let letters = reference |> Seq.takeWhile Char.IsLetter |> Seq.toArray |> System.String
    let col = letters |> Seq.fold (fun acc ch -> acc * 26 + int ch - int 'A' + 1) 0
    { Row = int (reference.Substring letters.Length) - 1; Col = col - 1 }

/// The format the painter shows a cell of a table in: the cell's own with its table's look under it.
let shown (model: SpreadsheetModel) (reference: string) : CellFormat =
    let addr = cellAt reference
    let own = getCell addr model |> Option.map (fun cell -> cell.Format) |> Option.defaultValue defaultFormat
    match model.Tables |> List.tryPick (fun table -> tableCellLook table addr.Row addr.Col) with
    | Some look -> formatWithTableLook look own
    | None -> own
let fill model reference = (shown model reference).BgColor
let text model reference = (shown model reference).TextColor
let bold model reference = (shown model reference).Bold
let line (edge: BorderStyle option) = edge |> Option.map (fun b -> b.Style, defaultArg b.Color "")
let top model reference = line (shown model reference).BorderTop
let bottom model reference = line (shown model reference).BorderBottom
let left model reference = line (shown model reference).BorderLeft
let right model reference = line (shown model reference).BorderRight

// A theme colour with a tint is what Excel shows: excel-2019-tints.tsv is Excel 2019's own colour for
// every theme colour of the Office theme with 25 tints (tools/excel-parity/probe-tints.vbs: theme index,
// tint, RRGGBB), among them the tints files write (0.79998168889431442 ...).
do
    let office = [| "FFFFFF"; "000000"; "E7E6E6"; "44546A"; "4472C4"; "ED7D31"; "A5A5A5"; "FFC000"; "5B9BD5"; "70AD47" |]
    let rows =
        File.ReadAllLines(Path.Combine(__SOURCE_DIRECTORY__, "excel-2019-tints.tsv"))
        |> Array.filter (fun line -> line.Trim() <> "")
        |> Array.map (fun line -> let parts = line.Split '\t' in int parts.[0], float (parts.[1].Replace(',', '.')), parts.[2].Trim())
    let wrong = rows |> Array.filter (fun (index, amount, excel) -> XlsxReader.Xlsx.applyColorTint office.[index] amount <> excel)
    ok (sprintf "a tint is Excel's colour: %d of %d theme colours with a tint" (rows.Length - wrong.Length) rows.Length) (rows.Length = 250 && wrong.Length = 0)
        (sprintf "%A" (wrong |> Array.truncate 8 |> Array.map (fun (index, amount, excel) -> office.[index], amount, "Excel " + excel, "ours " + XlsxReader.Xlsx.applyColorTint office.[index] amount)))

let synthetic = sheetOf "test-documents-excel/synthetic-table-styles.xlsx" "Styles"
eq "the fixture's eleven tables are read with their styles"
    (synthetic.Tables |> List.map (fun table -> table.Style |> Option.map (fun style -> style.Name) |> Option.defaultValue "-"))
    [ "TableStyleLight1"; "TableStyleLight9"; "TableStyleLight15"; "TableStyleMedium2"; "TableStyleMedium7"; "TableStyleMedium9"
      "TableStyleMedium16"; "TableStyleMedium28"; "TableStyleDark1"; "TableStyleDark9"; "Boxed" ]

// ---- Light1 (A3:D8): no header fill, a line above and below the header and below the table, grey bands --------
eq "Light1: the header is bold, unfilled, with lines above and below" (fill synthetic "B3", bold synthetic "B3", top synthetic "B3", bottom synthetic "B3")
    (None, true, Some ("thin", "#000000"), Some ("thin", "#000000"))
eq "Light1: banded rows in grey (white darker 15%)" (fill synthetic "B4", fill synthetic "B5", fill synthetic "B6") (Some "#D9D9D9", None, Some "#D9D9D9")
eq "Light1: a line below the last row, none inside" (bottom synthetic "B8", bottom synthetic "B5") (Some ("thin", "#000000"), None)
eq "Light1: the cells' own black text stays black" (text synthetic "B4") (Some "#000000")

// ---- Light9 (F3:I9, total row, first column): a filled header, a line above each band -------------------------
eq "Light9: the header in the accent, white bold text" (fill synthetic "G3", text synthetic "G3", bold synthetic "G3") (Some "#4472C4", Some "#FFFFFF", true)
eq "Light9: each band of rows has a line above it" (top synthetic "G5", top synthetic "G6") (Some ("thin", "#4472C4"), Some ("thin", "#4472C4"))
eq "Light9: the total row is bold under a double line" (bold synthetic "G9", top synthetic "G9") (true, Some ("double", "#4472C4"))
eq "Light9: the first column is bold, the others not" (bold synthetic "F5", bold synthetic "G5") (true, false)
eq "Light9: the outline of the table" (left synthetic "F5", right synthetic "I5", right synthetic "G5") (Some ("thin", "#4472C4"), Some ("thin", "#4472C4"), None)

// ---- Medium2 (F13:I19, total row): the default style - light lines between the rows ---------------------------
eq "Medium2: header, bands, plain rows" ([ "G13"; "G14"; "G15"; "G16" ] |> List.map (fill synthetic)) [ Some "#4472C4"; Some "#D9E1F2"; None; Some "#D9E1F2" ]
eq "Medium2: a line of the accent lighter 40% between the rows" (top synthetic "G15", bottom synthetic "G15") (Some ("thin", "#8EA9DB"), Some ("thin", "#8EA9DB"))
eq "Medium2: no line between the columns" (left synthetic "H15") None
eq "Medium2: the total row is bold under a double line, unfilled" (fill synthetic "G19", bold synthetic "G19", top synthetic "G19") (None, true, Some ("double", "#4472C4"))

// ---- Medium7 (A23:D28, column stripes only): the accent of the theme's sixth accent ----------------------------
eq "Medium7: banded columns, not rows" ([ "A25"; "B25"; "C25"; "D25" ] |> List.map (fill synthetic)) [ Some "#E2EFDA"; None; Some "#E2EFDA"; None ]
eq "Medium7: the header in accent 6" (fill synthetic "B23") (Some "#70AD47")

// ---- Medium9 (F23:I29, total row, first and last column): every cell filled, white lines ------------------------
eq "Medium9: the body lighter 80%, the bands lighter 60%" (fill synthetic "G24", fill synthetic "G25") (Some "#B4C6E7", Some "#D9E1F2")
eq "Medium9: white lines between the cells" (top synthetic "G25", left synthetic "H25") (Some ("thin", "#FFFFFF"), Some ("thin", "#FFFFFF"))
eq "Medium9: the first and last columns in the accent, white bold" ([ "F25"; "I25" ] |> List.map (fun cell -> fill synthetic cell, text synthetic cell, bold synthetic cell))
    [ Some "#4472C4", Some "#FFFFFF", true; Some "#4472C4", Some "#FFFFFF", true ]
eq "Medium9: a thick white line under the header, over the total row" (bottom synthetic "G23", top synthetic "G29", fill synthetic "G29") (Some ("thick", "#FFFFFF"), Some ("thick", "#FFFFFF"), Some "#4472C4")

// ---- Medium16 (A33:D38, first column): grey bands, a medium line above and below ------------------------------
eq "Medium16: grey bands, a medium line round the table's top and bottom" (fill synthetic "B34", fill synthetic "B35", top synthetic "B33", bottom synthetic "B38")
    (Some "#D9D9D9", None, Some ("medium", "#000000"), Some ("medium", "#000000"))
eq "Medium16: the first column in the accent, white bold" (fill synthetic "A35", text synthetic "A35", bold synthetic "A35") (Some "#4472C4", Some "#FFFFFF", true)

// ---- Dark1 (A43:D49, total row, first column): dark fills, white text ------------------------------------------
eq "Dark1: the body in black lighter 45%, white text" (fill synthetic "B45", text synthetic "B45") (Some "#737373", Some "#FFFFFF")
eq "Dark1: bands in black lighter 25%" (fill synthetic "B44") (Some "#404040")
eq "Dark1: the header in black under a medium white line" (fill synthetic "B43", bottom synthetic "B43") (Some "#000000", Some ("medium", "#FFFFFF"))
eq "Dark1: the first column has a white line on its right" (fill synthetic "A45", right synthetic "A45", bold synthetic "A45") (Some "#404040", Some ("medium", "#FFFFFF"), true)
eq "Dark1: the total row in black lighter 15%" (fill synthetic "B49", top synthetic "B49") (Some "#262626", Some ("medium", "#FFFFFF"))

// ---- Dark9 (F43:I49, total row): the header in the second accent of the pair ----------------------------------
eq "Dark9: header in accent 2, body accent 1 lighter 80%, bands lighter 60%" (fill synthetic "G43", fill synthetic "G44", fill synthetic "G45") (Some "#ED7D31", Some "#B4C6E7", Some "#D9E1F2")
eq "Dark9: header text white" (text synthetic "G43") (Some "#FFFFFF")

// ---- Boxed (A53:D59), a custom style with every element -------------------------------------------------------
eq "custom: the header row's fill and white bold text" (fill synthetic "B53", text synthetic "B53", bold synthetic "B53") (Some "#2F75B5", Some "#FFFFFF", true)
eq "custom: the first header cell over the header row" (fill synthetic "A53", text synthetic "A53") (Some "#FFC000", Some "#FFFFFF")
eq "custom: row stripes two rows wide, then one row without" ([ "B54"; "B55"; "B56"; "B57"; "B58" ] |> List.map (fill synthetic))
    [ Some "#DDEBF7"; Some "#DDEBF7"; None; Some "#DDEBF7"; Some "#DDEBF7" ]
eq "custom: a column stripe where no row stripe is" (fill synthetic "C56", fill synthetic "D56") (Some "#F2F2F2", None)
eq "custom: the first column bold, the last red" (bold synthetic "A56", text synthetic "D56", bold synthetic "B56") (true, Some "#C00000", false)
eq "custom: the outline is medium, the inside thin" (left synthetic "A55", right synthetic "D55", top synthetic "B53", top synthetic "B55", left synthetic "C55")
    (Some ("medium", "#5B9BD5"), Some ("medium", "#5B9BD5"), Some ("medium", "#5B9BD5"), Some ("thin", "#9BC2E6"), Some ("thin", "#9BC2E6"))
eq "custom: the total row italic under a double line" ((shown synthetic "B59").Italic, top synthetic "B59") (true, Some ("double", "#5B9BD5"))
eq "custom: the last total cell" (fill synthetic "D59", bold synthetic "D59") (Some "#C6E0B4", true)

// ---- the painter draws what the look says ------------------------------------------------------------------------
let items = Paint.buildPaintItems synthetic (Layout.defaultViewport 1400.0 1400.0)
let fills = items |> List.choose (function Paint.FillRect (_, color) -> Some color | _ -> None) |> List.countBy id |> Map.ofList
eq "painter: the custom first header cell, the Dark1 body" (Map.tryFind "#FFC000" fills, Map.tryFind "#737373" fills |> Option.map (fun n -> n > 0)) (Some 1, Some true)

// ---- a style made in the app: the built-in definition in the Office theme ---------------------------------------
let made : Table =
    { Name = "Made"; StartRow = 1; StartCol = 1; EndRow = 5; EndCol = 3; HeaderRowCount = 1
      Style = Some (Ops.tableStyleNamed "TableStyleMedium2"); AutoFilter = None; HasTotals = false }
let lookOf (table: Table) row col = tableCellLook table row col |> Option.map (fun look -> formatWithTableLook look defaultFormat)
eq "a built-in style named in the app is Excel's" ((lookOf made 0 0).Value.BgColor, (lookOf made 2 1).Value.BorderTop |> line) (Some "#4472C4", Some ("thin", "#8EA9DB"))

// ---- a look is worked out once a kind of row, and is the same as worked out for the cell itself ----------
do
    let boxed = synthetic.Tables |> List.find (fun table -> table.Style |> Option.exists (fun style -> style.Name = "Boxed"))
    for rows, totals in [ 1, false; 2, true; 5, true; 7, false; 40, true; 41, false ] do
        let table = { boxed with StartRow = 3; EndRow = 3 + rows + (if totals then 1 else 0); HasTotals = totals }
        let style = table.Style.Value
        let direct row col =
            Rbt.Spreadsheet.TableStyles.cellLook style.Elements (table.StartRow - 1, table.StartCol - 1, table.EndRow - 1, table.EndCol - 1)
                table.HeaderRowCount (if totals then 1 else 0) (style.ShowFirstColumn, style.ShowLastColumn, style.ShowRowStripes, style.ShowColumnStripes) row col
        let differing =
            [ for row in table.StartRow - 2 .. table.EndRow do
                for col in table.StartCol - 2 .. table.EndCol do
                    if tableCellLook table row col <> direct row col then yield row, col ]
        eq (sprintf "the looked-up look is the cell's own: %d rows%s" rows (if totals then " and a total row" else "")) differing []

// ---- the e-commerce dashboard's table, in its custom "Data-style" ---------------------------------------------
let ecommerce = sheetOf "real-world/dashboard-ecommerce-sales.xlsx" "Data"
eq "Data-style: the header in accent 1 of the workbook's theme, the text its own" (fill ecommerce "B1", bold ecommerce "B1") (Some "#1CADE4", (getCell (cellAt "B1") ecommerce).Value.Format.Bold)
eq "Data-style: both stripes the same light blue" (fill ecommerce "B2", fill ecommerce "B3") (Some "#DEEAF6", Some "#DEEAF6")
eq "Data-style: first and last column are on, but the style has no such element" (bold ecommerce "A5", bold ecommerce "V5") ((getCell (cellAt "A5") ecommerce).Value.Format.Bold, (getCell (cellAt "V5") ecommerce).Value.Format.Bold)

// ---- the lines of a table's style are the edges of its cells, on a sheet with merged cells too --------------
// (The web grid drew only the lines above and below a table's cells there; the painter and Excel all four.)
do
    let model =
        { createEmpty 20 10 with
            Tables = [ made ]
            MergedCells = [ { StartRow = 15; StartCol = 5; EndRow = 16; EndCol = 6 } ] }
    let look = lookOf made 2 1
    eq "a table's left line on a sheet with merged cells" (Layout.edgeLeft model 2 0) (Layout.EdgeBorder (lookOf made 2 0).Value.BorderLeft.Value)
    eq "... and on one without" (Layout.edgeLeft { model with MergedCells = [] } 2 0) (Layout.EdgeBorder (lookOf made 2 0).Value.BorderLeft.Value)
    ignore look
// ---- a table without a style: plain cells (no look of its own, as Excel draws one) ---------------------------
do
    let plain = { made with Style = None }
    let model = { createEmpty 20 10 with Tables = [ plain ] } |> setCell { Row = 0; Col = 0 } { Value = CellValue.String "Name"; Format = defaultFormat; Comment = None }
    let look = Paint.cellLook model 0 0 (getCell { Row = 0; Col = 0 } model) None 64.0
    eq "a table without a style: its header is a plain cell" (look.Fill, look.Bold, look.Align) (None, false, Layout.AlignLeft)

printfn ""
printfn "Stage 54 table style tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
