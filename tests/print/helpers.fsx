#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"
#r "../../src/Spreadsheet.Pdf/bin/Debug/net10.0/Rbt.Spreadsheet.Pdf.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Pdf

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
// ----- pictures in the PDF -----
// A picture (PNG / JPEG) is embedded as an image object and drawn over its
// rectangle; a labelled frame stands only for what cannot be embedded.
module Paint = Rbt.Spreadsheet.Paint
let pngBytes = System.Convert.FromBase64String "iVBORw0KGgoAAAANSUhEUgAAABgAAAAQCAIAAACDRijCAAAAYUlEQVR42mP4TyXAAMQfekSASCHhBNkIahCQBTGLbOMQBkEQ2WahG0S207AYRJ7TcBpEqtPwGUSS0wgbRKTTiDKIGKeRYBB+p5FmEB6nkWMQVqeRaRCm0ygyCNlpUIOoAgBIlHezofwiaQAAAABJRU5ErkJggg=="
let pdfText (bytes: byte[]) = System.Text.Encoding.Latin1.GetString bytes
let countOf (needle: string) (text: string) =
    let mutable n = 0
    let mutable at = text.IndexOf needle
    while at >= 0 do
        n <- n + 1
        at <- text.IndexOf (needle, at + needle.Length)
    n
let pictureItem (data: byte[]) : Paint.PaintItem =
    Paint.DrawImage ({ X = 40.0; Y = 30.0; Width = 96.0; Height = 64.0 }, { ContentType = "image/png"; Data = data; Label = "Picture: Logo" })
let withPicture = pdfText (Render.sceneToBytes 400.0 300.0 [ pictureItem pngBytes ])
ok "pdf picture: an image object" (countOf "/Subtype /Image" withPicture = 1) (sprintf "%d image objects" (countOf "/Subtype /Image" withPicture))
ok "pdf picture: its size" (withPicture.Contains "/Width 24" && withPicture.Contains "/Height 16") ""
ok "pdf picture: drawn" (withPicture.Contains " Do") ""
let samePictureTwice = pdfText (Render.sceneToBytes 400.0 300.0 [ pictureItem pngBytes; Paint.DrawImage ({ X = 200.0; Y = 30.0; Width = 48.0; Height = 32.0 }, { ContentType = "image/png"; Data = pngBytes; Label = "Picture: Logo" }) ])
ok "pdf picture: one object for a picture drawn twice" (countOf "/Subtype /Image" samePictureTwice = 1 && countOf " Do" samePictureTwice = 2) ""
let junk = pdfText (Render.sceneToBytes 400.0 300.0 [ pictureItem [| 137uy; 80uy; 78uy; 71uy; 1uy; 2uy; 3uy; 250uy |] ])
ok "pdf picture: what cannot be decoded keeps its frame" (countOf "/Subtype /Image" junk = 0 && countOf " re" junk >= 1) ""

// ----- a print area of several areas: each is its own run of pages -----
let area r1 c1 r2 c2 : CellRange = { From = { Row = r1; Col = c1 }; To = { Row = r2; Col = c2 } }
let twoAreas = { tinyPageSetup with PrintArea = Some (area 0 0 9 2); ExtraPrintAreas = [ area 0 4 24 6 ] }
let pm = Pagination.paginate (buildModel 30 8 twoAreas)
ok "areas: both are printed, the first is the bounds" (pm.PrintAreas = [ area 0 0 9 2; area 0 4 24 6 ] && pm.PrintBounds = area 0 0 9 2) (sprintf "%A" pm.PrintAreas)
let pagesOfArea index = pm.Pages |> List.filter (fun page -> page.Area = index)
ok "areas: every page lies in its area"
   (pm.Pages |> List.forall (fun page -> let a = pm.PrintAreas.[page.Area] in page.RowStart >= a.From.Row && page.RowEnd <= a.To.Row && page.ColStart >= a.From.Col && page.ColEnd <= a.To.Col))
   (sprintf "%A" pm.Pages)
ok "areas: the pages of the first area come first" ((pm.Pages |> List.map (fun page -> page.Area)) = (List.replicate (pagesOfArea 0).Length 0 @ List.replicate (pagesOfArea 1).Length 1) && not (pagesOfArea 0).IsEmpty && (pagesOfArea 1).Length > 1) (sprintf "%A" (pm.Pages |> List.map (fun page -> page.Area)))
let covered index = pagesOfArea index |> List.collect (fun page -> [ for r in page.RowStart .. page.RowEnd do for c in page.ColStart .. page.ColEnd -> r, c ]) |> List.sort
ok "areas: each area's cells are on its pages, once" (covered 0 = [ for r in 0 .. 9 do for c in 0 .. 2 -> r, c ] && covered 1 = [ for r in 0 .. 24 do for c in 4 .. 6 -> r, c ]) ""
ok "areas: the cells between the areas are on no page" (pm.Pages |> List.forall (fun page -> page.ColEnd < 3 || page.ColStart > 3)) ""
ok "areas: one area is area 0" ((Pagination.paginate (buildModel 30 8 { tinyPageSetup with PrintArea = Some (area 0 0 9 2) })).Pages |> List.forall (fun page -> page.Area = 0)) ""
let fitted = Pagination.paginate (buildModel 30 8 { twoAreas with FitToPage = true; FitToWidth = 1; FitToHeight = 1 })
let alone index = (Pagination.paginate (buildModel 30 8 { twoAreas with PrintArea = Some (fitted.PrintAreas.[index]); ExtraPrintAreas = []; FitToPage = true; FitToWidth = 1; FitToHeight = 1 })).Scale
ok "areas: fit to one page is one page for each area, at the scale the larger one needs"
   (fitted.Pages.Length = 2 && abs (fitted.Scale - min (alone 0) (alone 1)) < 1e-9 && alone 1 < alone 0)
   (sprintf "%d pages, scale %f (%f, %f)" fitted.Pages.Length fitted.Scale (alone 0) (alone 1))

printfn "Print-helper tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
