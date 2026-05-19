#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office-spreadsheet/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"
#r "/home/dave/Dev/browser/office-spreadsheet/src/Spreadsheet.Pdf/bin/Debug/net10.0/RedBat.Spreadsheet.Pdf.dll"

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Pdf

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label detail)

// ----- HeaderFooter -----

let ctx : HeaderFooter.PageContext =
    { PageNumber = 2
      TotalPages = 5
      SheetName = "Sheet One"
      FileName = "report.xlsx"
      FilePath = "/var/reports"
      Now = System.DateTime(2026, 3, 15, 14, 30, 0) }

let sec1 = HeaderFooter.resolve ctx "&LLeft&CCenter&RRight"
ok "hf-3-sections" (sec1.Left.Text = "Left" && sec1.Center.Text = "Center" && sec1.Right.Text = "Right") (sprintf "%A" sec1)

let sec2 = HeaderFooter.resolve ctx "Page &P of &N"
ok "hf-page-fields" (sec2.Center.Text = "Page 2 of 5") sec2.Center.Text

let sec3 = HeaderFooter.resolve ctx "&A — &F"
ok "hf-sheet-file" (sec3.Center.Text = "Sheet One — report.xlsx") sec3.Center.Text

let sec4 = HeaderFooter.resolve ctx "&D at &T"
ok "hf-date-time" (sec4.Center.Text = "3/15/2026 at 2:30 PM") sec4.Center.Text

let sec5 = HeaderFooter.resolve ctx "AT&&T"
ok "hf-ampersand-escape" (sec5.Center.Text = "AT&T") sec5.Center.Text

let sec6 = HeaderFooter.resolve ctx "&L&BBold left&R&\"Arial,Bold\"&12Right title"
ok "hf-strip-style-codes" (sec6.Left.Text = "Bold left" && sec6.Left.Bold && sec6.Right.Text = "Right title" && sec6.Right.FontName = Some "Arial" && sec6.Right.Bold && sec6.Right.FontSize = Some 12.0) (sprintf "%A" sec6)

let sec7 = HeaderFooter.resolve ctx "&L&A&CPage &P of &N"
ok "hf-poi49156-footer" (sec7.Left.Text = "Sheet One" && sec7.Center.Text = "Page 2 of 5") (sprintf "%A" sec7)

let sec8 = HeaderFooter.resolve ctx "&C&\"Times New Roman,Regular\"&12&A"
ok "hf-times-12-style" (sec8.Center.Text = "Sheet One" && sec8.Center.FontName = Some "Times New Roman" && sec8.Center.FontSize = Some 12.0 && not sec8.Center.Bold) (sprintf "%A" sec8)

let sec9 = HeaderFooter.resolve ctx "&CPlain &Bbold&B plain &KFF0000red &K000000&\"Arial,Italic\"italic &16big&\"Calibri,Regular\"&10 normal"
let sec9Runs = sec9.Center.Runs
ok "hf-mid-run-styles"
   (sec9.Center.Text = "Plain bold plain red italic big normal"
    && sec9Runs.Length = 7
    && sec9Runs.[0].Text = "Plain "
    && not sec9Runs.[0].Bold
    && sec9Runs.[1].Text = "bold"
    && sec9Runs.[1].Bold
    && sec9Runs.[2].Text = " plain "
    && not sec9Runs.[2].Bold
    && sec9Runs.[3].Text = "red "
    && sec9Runs.[3].Color = Some "#FF0000"
    && sec9Runs.[4].Text = "italic "
    && sec9Runs.[4].FontName = Some "Arial"
    && sec9Runs.[4].Italic
    && sec9Runs.[4].Color = Some "#000000"
    && sec9Runs.[5].Text = "big"
    && sec9Runs.[5].FontSize = Some 16.0
    && sec9Runs.[6].Text = " normal"
    && sec9Runs.[6].FontName = Some "Calibri"
    && not sec9Runs.[6].Italic
    && sec9Runs.[6].FontSize = Some 10.0)
   (sprintf "%A" sec9)

// ----- Pagination -----

let cell v = { Value = v; Format = defaultFormat; Comment = None }

// Build a model with 50 rows × 8 cols of numeric data, uniform 100 px
// wide / 25 px tall.
let buildModel (rows: int) (cols: int) (ps: PageSetup) : SpreadsheetModel =
    let cells =
        [ for r in 0 .. rows - 1 do
            for c in 0 .. cols - 1 ->
                { Row = r; Col = c }, cell (CellValue.Number (float (r * cols + c))) ]
        |> Map.ofList
    { createEmpty (max rows 1000) (max cols 26) with
        Cells = cells
        PageSetup = ps }

// Small page: 200x200 pt printable (after margins, ~80x80 px each side gone).
let tinyPageSetup =
    { defaultPageSetup with
        PageWidth = 200.0     // very small
        PageHeight = 200.0
        Margins = { Top = 0.1; Bottom = 0.1; Left = 0.1; Right = 0.1; Header = 0.0; Footer = 0.0 } }

let m1 = buildModel 50 8 tinyPageSetup
let p1 = Pagination.paginate m1
ok "pagi-multi-row-page"
   (List.length p1.Pages >= 2)
   (sprintf "pages=%d" (List.length p1.Pages))
ok "pagi-row-pages-fit"
   (p1.Pages |> List.forall (fun p -> p.RowEnd >= p.RowStart && p.ColEnd >= p.ColStart))
   "every page has a non-empty rectangle"

// Manual row break at row 5 (1-based) → first page must end at row 4 (0-based).
let m2 = buildModel 30 4
                   { defaultPageSetup with
                       ManualRowBreaks = [5]
                       Margins = { defaultPageMargins with Top = 0.1; Bottom = 0.1 } }
let p2 = Pagination.paginate m2
ok "pagi-manual-row-break"
   (p2.Pages |> List.exists (fun p -> p.RowStart = 0 && p.RowEnd = 4))
   (sprintf "first page rows = %A" (p2.Pages |> List.head))

// Print area takes precedence over used range.
let m3 = buildModel 20 6
                   { defaultPageSetup with
                       PrintArea = Some { From = { Row = 2; Col = 1 }; To = { Row = 8; Col = 3 } } }
let p3 = Pagination.paginate m3
ok "pagi-print-area"
   (p3.PrintBounds.From = { Row = 2; Col = 1 } && p3.PrintBounds.To = { Row = 8; Col = 3 })
   (sprintf "%A" p3.PrintBounds)

// "Down then over" ordering: every row page within col-page A precedes
// col-page B.
let m4 = buildModel 40 20 tinyPageSetup
let p4 = Pagination.paginate m4
let firstColPage = p4.Pages |> List.head |> fun p -> p.ColStart
let lastSameCol = p4.Pages |> List.takeWhile (fun p -> p.ColStart = firstColPage) |> List.length
ok "pagi-down-then-over"
   (lastSameCol > 1 && lastSameCol = (p4.Pages |> List.filter (fun p -> p.ColStart = firstColPage) |> List.length))
   (sprintf "first col-page covers %d row pages out of %d total" lastSameCol (List.length p4.Pages))

// Fit-to-page shrinks scale so all content fits on one page.
let m5 = buildModel 200 30
                   { defaultPageSetup with
                       FitToPage = true; FitToWidth = 1; FitToHeight = 1 }
let p5 = Pagination.paginate m5
ok "pagi-fit-to-one-page" (List.length p5.Pages = 1) (sprintf "pages=%d scale=%f" (List.length p5.Pages) p5.Scale)
ok "pagi-fit-shrinks-scale" (p5.Scale < 1.0) (sprintf "scale=%f" p5.Scale)

printfn ""
printfn "Print-helper tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
