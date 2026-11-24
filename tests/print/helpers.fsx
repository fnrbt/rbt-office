#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"
#r "../../src/Spreadsheet.Pdf/bin/Debug/net10.0/Rbt.Spreadsheet.Pdf.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Pdf

module Icons = Rbt.Spreadsheet.ConditionalFormatting
module Painter = Rbt.Spreadsheet.Paint

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

let secTheme = HeaderFooter.resolve ctx "&L&\"Arial,Regular\"&8&K01+034https://example.com/x.html"
ok "hf-theme-colour-code" (secTheme.Left.Text = "https://example.com/x.html") secTheme.Left.Text

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

// Excel 2019 blockhold.xlsx: an inherited printable column style can hold a trailing column-page only
// when a visible, used row on it shares an absolute 16-row block with a used extent reaching that column.
do
    let setup =
        { defaultPageSetup with
            PageWidth = 600.0; PageHeight = 1100.0; ManualRowBreaks = [ 20; 41 ]
            Margins = { defaultPageMargins with Top = 0.1; Bottom = 0.1; Left = 0.1; Right = 0.1 } }
    let x = cell (CellValue.String "x")
    let addr r c = { Row = r; Col = c }
    let at r c data = addr r c, data
    let fill = { defaultFormat with BgColor = Some "#CCCCCC" }
    let font = { defaultFormat with Bold = true }
    let style = { First = 10; Last = 10; Format = fill; Source = -1 }
    let baseCells = Map.ofList [ at 0 0 x; at 1 10 x; at 59 0 x ]
    let model markRow targetRow =
        { createEmpty 100 26 with
            PageSetup = setup; DefaultRowHeight = 20.0; ColumnStyles = [ style ]
            Cells = baseCells |> Map.add (addr markRow 10) x |> Map.add (addr targetRow 1) x }
    let held m = Pagination.paginate m |> fun p ->
        p.Pages |> List.exists (fun page -> page.RowStart = 41 && page.ColStart <= 10 && page.ColEnd >= 10)
    let check name expected m =
        let actual = held m
        ok ("print block: " + name) (actual = expected) (sprintf "held=%b expected=%b" actual expected)
    check "K32 / B48 misses block" false (model 31 47)
    check "K33 / B48 shares block" true (model 32 47)
    check "K39 / B41 before break" false (model 38 40)
    check "K39 / B42 across break" true (model 38 41)
    check "K39 / B49 outside block" false (model 38 48)
    let baseMarked = { model 38 44 with Cells = baseCells |> Map.add (addr 38 10) x }
    check "no target row" false baseMarked
    check "custom 15pt target row" true { baseMarked with RowPoints = Map.ofList [ 44, 15.0 ]; RowHeights = Map.ofList [ 44, 20.0 ] }
    check "hidden target row" false { model 38 44 with RowHeights = Map.ofList [ 44, 0.0 ] }
    check "hidden source still marks" true { model 38 44 with RowHeights = Map.ofList [ 38, 0.0 ] }
    check "font-only column is not printable" false { model 38 44 with ColumnStyles = [ { style with Format = font } ] }
    check "font-formatted source marks" true { baseMarked with Cells = baseMarked.Cells |> Map.add (addr 38 10) { cell CellValue.Empty with Format = font } |> Map.add (addr 44 1) x }
    check "font-formatted target participates" true { baseMarked with Cells = baseMarked.Cells |> Map.add (addr 44 1) { cell CellValue.Empty with Format = font } }
    check "formula-empty source marks" true { baseMarked with Cells = baseMarked.Cells |> Map.add (addr 38 10) (cell (CellValue.Formula ("=\"\"", CellValue.String ""))) |> Map.add (addr 44 1) x }
    check "formula-empty target participates" true { baseMarked with Cells = baseMarked.Cells |> Map.add (addr 44 1) (cell (CellValue.Formula ("=\"\"", CellValue.String ""))) }
    check "blank merged source marks" true { baseMarked with Cells = Map.add (addr 44 1) x baseCells; MergedCells = [ { StartRow = 38; StartCol = 10; EndRow = 38; EndCol = 11 } ] }
    check "blank merged target participates" true { baseMarked with MergedCells = [ { StartRow = 44; StartCol = 1; EndRow = 44; EndCol = 2 } ] }
    check "row font style does not participate" false { baseMarked with RowStyles = Map.ofList [ 44, font ] }
    check "row font style overrides value" false { model 38 44 with RowStyles = Map.ofList [ 44, font ] }
    check "row fill style participates" true { baseMarked with RowStyles = Map.ofList [ 44, fill ] }

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

// ----- The pages of real workbooks: what the host API reports (Pagination.paginate, behind
// GetPrintPages) is what the PDF has. Every sheet of several workbooks is rendered and its
// pages counted in the PDF. -----
let pdfPageCount (bytes: byte[]) =
    let text = System.Text.Encoding.Latin1.GetString bytes
    System.Text.RegularExpressions.Regex.Matches(text, "/Type\\s*/Page(?![a-zA-Z])").Count
let mutable sheetsPaged = 0
let mutable pagesCounted = 0
for file in [ "real-world/milk-street-office.xlsx"; "real-world/lyric-multifamily.xlsx"; "real-world/dcf-stock-analysis.xlsx"; "real-world/dashboard-ecommerce-sales.xlsx"; "real-world/jumeirah-hotel.xlsx"; "test-documents-excel/poi-49156.xlsx" ] do
    let path = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", file)
    let imported = Rbt.Spreadsheet.XlsxBridge.Import.importWorkbook (XlsxReader.loadBytesNative (System.IO.File.ReadAllBytes path) |> Option.get)
    for sheet in imported.Sheets do
        // (A sheet of a few hundred thousand cells is a few thousand pages: left to the PDF runner.)
        if sheet.Model.Cells.Count < 20000 then
            let pagination = Pagination.paginatePrinted sheet.Model
            let options : Render.RenderOptions = { DefaultFontFamily = "Helvetica"; Fonts = []; CjkFallback = []; SheetName = sheet.Name; FileName = "book" }
            let pdf = Render.toBytes options sheet.Model
            sheetsPaged <- sheetsPaged + 1
            pagesCounted <- pagesCounted + pagination.Pages.Length
            ok (sprintf "pages: %s / %s" (System.IO.Path.GetFileName file) sheet.Name) (pdfPageCount pdf = max 1 pagination.Pages.Length)
                (sprintf "the PDF has %d pages, the pagination %d" (pdfPageCount pdf) pagination.Pages.Length)
            // Every cell of every print area is on one page of its area at most, and a cell on none holds nothing (the
            // pages at the end with nothing on them are left out, as Excel 2019 leaves them).
            let covered =
                pagination.PrintAreas |> List.mapi (fun index area ->
                    let pages = pagination.Pages |> List.filter (fun page -> page.Area = index)
                    let cells = pages |> List.sumBy (fun page -> (page.RowEnd - page.RowStart + 1) * (page.ColEnd - page.ColStart + 1))
                    // (The repeat rows / columns print where they are on the pages that begin at or before them.)
                    let rows = area.To.Row - area.From.Row + 1
                    let cols = area.To.Col - area.From.Col + 1
                    let onPage (addr: CellAddress) = pages |> List.exists (fun page -> addr.Row >= page.RowStart && addr.Row <= page.RowEnd && addr.Col >= page.ColStart && addr.Col <= page.ColEnd)
                    let leftOut =
                        sheet.Model.Cells |> Map.exists (fun addr data ->
                            addr.Row >= area.From.Row && addr.Row <= area.To.Row && addr.Col >= area.From.Col && addr.Col <= area.To.Col
                            && (match data.Value with CellValue.Empty | CellValue.String "" -> false | _ -> true) && not (onPage addr))
                    cells <= max 0 rows * max 0 cols && (cells = max 0 rows * max 0 cols || not leftOut))
            ok (sprintf "pages: %s / %s cover their areas" (System.IO.Path.GetFileName file) sheet.Name) (List.forall id covered) (sprintf "%A" covered)
// A landscape sheet is paginated the wider way round, whether its setup holds the paper's size
// in portrait (a file read by the editor) or already turned (the native runner).
do
    let cells = [ for r in 0 .. 199 do for c in 0 .. 19 -> { Row = r; Col = c }, { Value = CellValue.Number 1.0; Format = defaultFormat; Comment = None } ] |> Map.ofList
    let sheet setup = { createEmpty 200 20 with Cells = cells; PageSetup = setup }
    let portrait = Pagination.paginatePrinted (sheet defaultPageSetup)
    let landscape = Pagination.paginatePrinted (sheet { defaultPageSetup with Orientation = LandscapeOrientation })
    let turned = Pagination.paginatePrinted (sheet { defaultPageSetup with Orientation = LandscapeOrientation; PageWidth = defaultPageSetup.PageHeight; PageHeight = defaultPageSetup.PageWidth })
    let first (p: Pagination.Pagination) = let page = p.Pages.Head in page.RowEnd, page.ColEnd
    ok "landscape: fewer rows and more columns to a page" (fst (first landscape) < fst (first portrait) && snd (first landscape) > snd (first portrait)) (sprintf "%A vs %A" (first landscape) (first portrait))
    ok "landscape: the same pages with the turned size given" (landscape.Pages = turned.Pages) ""
    let size (bytes: byte[]) = let m = System.Text.RegularExpressions.Regex.Match(System.Text.Encoding.Latin1.GetString bytes, "/MediaBox\\s*\\[\\s*0\\s+0\\s+([0-9.]+)\\s+([0-9.]+)") in float m.Groups.[1].Value, float m.Groups.[2].Value
    let options : Render.RenderOptions = { DefaultFontFamily = "Helvetica"; Fonts = []; CjkFallback = []; SheetName = "S"; FileName = "f" }
    let pdf = Render.toBytes options (sheet { defaultPageSetup with Orientation = LandscapeOrientation })
    ok "landscape: the PDF's pages are the wider way round" (size pdf = (792.0, 612.0) && pdfPageCount pdf = landscape.Pages.Length) (sprintf "%A, %d pages (%d)" (size pdf) (pdfPageCount pdf) landscape.Pages.Length)
ok "pages: sheets were paginated" (sheetsPaged >= 12 && pagesCounted > sheetsPaged) (sprintf "%d sheets, %d pages" sheetsPaged pagesCounted)

// Charts and shapes in print: a scene as paint items, flattened, with its text.
let printScene : Rbt.Spreadsheet.Charts.Scene =
    { Width = 200.0; Height = 100.0
      Items =
        [ Rbt.Spreadsheet.Charts.PRect (0.0, 0.0, 200.0, 100.0, Some "#FFFFFF", 1.0, None, 0.0, 0.0)
          Rbt.Spreadsheet.Charts.PRect (5.0, 5.0, 10.0, 10.0, None, 1.0, None, 0.0, 0.0)
          Rbt.Spreadsheet.Charts.PGroup (10.0, 20.0, 50.0, 50.0,
            [ Rbt.Spreadsheet.Charts.PText (1.0, 2.0, "inside", 12.0, "#000000", "start", false, false, 0.0)
              Rbt.Spreadsheet.Charts.PRect (0.0, 0.0, 4.0, 4.0, Some "gradient(90;0 #112233 1;1 #445566 1)", 1.0, None, 0.0, 0.0) ]) ] }
let printItems = Paint.scenePaintItems printScene
let printTexts = printItems |> List.choose (function Paint.DrawText text -> Some text | _ -> None)
ok "print scene: a text in a group is a top-level text at its place" (printTexts |> List.map (fun text -> text.Text, text.Rect.X) = [ "inside", 11.0 ]) (sprintf "%A" printTexts)
ok "print scene: no path with neither fill nor outline (it would be painted with the next one)" (printItems |> List.forall (function Paint.VecPath (_, None, None) -> false | _ -> true)) ""
ok "print scene: a gradient is its first colour" (printItems |> List.exists (function Paint.VecPath (_, Some "#112233", _) -> true | _ -> false)) (sprintf "%A" printItems)

// Print sizes: rows and columns without a size of their own print at the sheet's defaults, and drawings count.
let sizedSheet =
    let cell = { Value = CellValue.Number 1.0; Format = defaultFormat; Comment = None }
    { (createEmpty 200 20 |> setCell { Row = 0; Col = 0 } cell |> setCell { Row = 59; Col = 0 } cell) with RowHeights = Map.empty; DefaultRowHeight = 40.0 }
ok "print sizes: a tall default row makes more pages than the short one"
    ((Pagination.paginatePrinted sizedSheet).Pages.Length > (Pagination.paginatePrinted { sizedSheet with DefaultRowHeight = 15.0 }).Pages.Length) ""

// Several sheets in one PDF (the browser component's "workbook"): each by its own setup, one after the other.
do
    let cells n = [ for r in 0 .. n - 1 -> { Row = r; Col = 0 }, { Value = CellValue.Number 1.0; Format = defaultFormat; Comment = None } ] |> Map.ofList
    let short = { createEmpty 200 20 with Cells = cells 10 }
    let long = { createEmpty 200 20 with Cells = cells 150; PageSetup = { defaultPageSetup with Orientation = LandscapeOrientation } }
    let options : Render.RenderOptions = { DefaultFontFamily = "Helvetica"; Fonts = []; CjkFallback = []; SheetName = ""; FileName = "f" }
    let pdf = Render.toBytesOfSheets options [ Render.pdfSheet short "Short" []; Render.pdfSheet long "Long" [] ]
    let expected = (Pagination.paginatePrinted short).Pages.Length + (Pagination.paginatePrinted long).Pages.Length
    ok "sheets: one PDF of every sheet's pages" (pdfPageCount pdf = expected && expected > 2) (sprintf "%d pages, %d expected" (pdfPageCount pdf) expected)
    let widths =
        System.Text.RegularExpressions.Regex.Matches(System.Text.Encoding.Latin1.GetString pdf, @"/MediaBox\s*\[\s*0\s+0\s+([0-9.]+)")
        |> Seq.map (fun m -> float m.Groups.[1].Value) |> Seq.distinct |> Seq.sort |> List.ofSeq
    ok "sheets: each in its own orientation" (widths = [ 612.0; 792.0 ]) (sprintf "%A" widths)
    ok "sheets: none is a blank page" (pdfPageCount (Render.toBytesOfSheets options []) = 1) ""
    // A picture the browser hands as JPEG instead of its PNG (pngPictures / withPictures).
    let png = Array.append [| 0x89uy; 0x50uy; 0x4Euy; 0x47uy; 0x0Duy; 0x0Auy; 0x1Auy; 0x0Auy |] (Array.zeroCreate 20)
    let anchor : DrawingAnchor =
        { Kind = "twoCell"; FromRow = 0; FromCol = 1; FromOffsetX = 0.0; FromOffsetY = 0.0; ToRow = Some 3; ToCol = Some 3; ToOffsetX = 0.0; ToOffsetY = 0.0
          X = 0.0; Y = 0.0; Width = None; Height = None; EditAs = None }
    let pictured = { short with Drawings = [ { Id = 1; Anchor = anchor; Name = None; Description = None; Content = Picture ("image/png", "png", png) } ] }
    let listed = Render.pngPictures [ Render.pdfSheet pictured "P" [] ]
    let jpeg = [| 0xFFuy; 0xD8uy; 0xFFuy |]
    let swapped = Render.withPictures (fun data -> if obj.ReferenceEquals(data, png) then Some ("image/jpeg", jpeg) else None) [ Render.pdfSheet pictured "P" [] ]
    let content = match swapped.Head.Model.Drawings.Head.Content with Picture (kind, _, data) -> kind, data.Length | _ -> "", 0
    ok "pictures: the PNG listed, then drawn from the JPEG handed for it" (listed.Length = 1 && obj.ReferenceEquals(listed.[0], png) && content = ("image/jpeg", 3)) (sprintf "%d, %A" listed.Length content)

// ----- cell text in print, as Excel 2019's PDF sets it (make-print-text-fixture.py) -----
do
    let paint = Rbt.Spreadsheet.Paint.buildPaintItems
    let textItems (width: float) (text: string) (wrap: bool) =
        let model =
            createEmpty 4 4
            |> setCell { Row = 0; Col = 1 } { Value = String text; Format = { defaultFormat with WrapText = wrap }; Comment = None }
            |> setColumnWidth 1 width
        let viewport : Rbt.Spreadsheet.Layout.Viewport =
            { ScrollX = 0.0; ScrollY = 0.0; Width = 400.0; Height = 100.0; HeaderHeight = 0.0; RowHeaderWidth = 0.0
              Measurer = Some (fun s _ _ _ _ -> 6.0 * float s.Length) }
        [ for item in paint model viewport do
            match item with
            | Rbt.Spreadsheet.Paint.DrawText t when t.Text.Contains "a" || t.Text.Contains "b" -> yield t
            | _ -> () ]
    // (Column A is 100 px: B starts there; the headers' letters are left out.) A line 2.56 px in from the cell's left side, 16/600 in.
    let one = textItems 100.0 "abc" false
    ok "print: text 2.56 px from the cell's left side" (one.Length = 1 && abs (one.Head.Rect.X - (100.0 + 2.56)) < 1e-6) (sprintf "%A" (one |> List.map (fun t -> t.Rect.X)))
    // A wrapped line fits while no wider than the column less 4.48 px: 15 characters of 6 px (90) in 95 px, not in 94.
    let lines width = textItems width "aaaaaaa bbbbbbb" true |> List.map (fun t -> t.Text)
    ok "print: wrapped at the column less 4.48 px" (lines 95.0 = [ "aaaaaaa bbbbbbb" ] && lines 94.0 = [ "aaaaaaa"; "bbbbbbb" ]) (sprintf "%A / %A" (lines 95.0) (lines 94.0))
    // The text's inset by the Normal font's widest digit in print: ceil(d / 4) + 4 dots (Excel 2019's PDF of 33
    // Normal fonts: Calibri 11 16 dots, Calibri 9 14, Calibri 14 19, Cambria 14 21).
    let inset digit = System.Math.Round (Rbt.Spreadsheet.Paint.printPadXOf { createEmpty 1 1 with PrintDigitWidth = digit } * 6.25, 6)
    ok "print: the text's inset by the Normal font" ([ 47.0; 38.0; 59.0; 65.0 ] |> List.map inset = [ 16.0; 14.0; 19.0; 21.0 ]) (sprintf "%A" ([ 47.0; 38.0; 59.0; 65.0 ] |> List.map inset))
    // A format's `_x` at the end / start of the section shown leaves the width of x ("Pads": 0_) 28 dots of Calibri 11).
    let pads code value = Rbt.Spreadsheet.CellText.widthPads { defaultFormat with NumberFormat = Some code } value
    ok "print: a format's widths of characters" (pads "0_)" (Number 1.0) = ("", ")") && pads "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"??_);_(@_)" (Number -5.0) = ("(", "")
                                                && pads "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\)" (Number 5.0) = ("(", ")") && pads "0_W_)" (Number 1.0) = ("", "W)")
                                                && pads "_(@_)" (String "a") = ("(", ")") && pads "[>=100]0_)" (Number 1.0) = ("", "") && pads "0\\)" (Number 1.0) = ("", "")) ""
    // Rows an AutoFilter hides print nothing (houseflipping's REPAIR ESTIMATOR): synthetic-print-pages-filter, "print pages:" below.
    // Where a line stands in its cell (synthetic-print-pages-place "Place", Excel's PDF): the text at the bottom of row 3 (30 pt,
    // Calibri 11 to 20, Arial, Segoe UI, Times, Trebuchet, Tahoma) on one baseline, 51.5 dots above the row's bottom in Excel
    // (Calibri 20's descent, 7 px, and 1, at 6.05 dots a pixel: 48.4); at the top / in the middle by the font's own metrics.
    let place =
        let book = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "test-documents-excel", "synthetic-print-pages-place.xlsx")
        let imported = Rbt.Spreadsheet.XlsxBridge.Import.importWorkbook (XlsxReader.loadBytesNative (System.IO.File.ReadAllBytes book) |> Option.get)
        Pagination.withPrintSizes imported.Sheets.[0].Model
    let viewport : Rbt.Spreadsheet.Layout.Viewport =
        { ScrollX = 0.0; ScrollY = 0.0; Width = 2000.0; Height = 600.0; HeaderHeight = 0.0; RowHeaderWidth = 0.0
          Measurer = Some (fun s _ _ _ _ -> 6.0 * float s.Length) }
    let placed = [ for item in paint place viewport do match item with Rbt.Spreadsheet.Paint.DrawText t when t.Text = "xH" -> yield t | _ -> () ]
    let rowTop (row: int) = Seq.sum [ for r in 0 .. row - 1 -> Rbt.Spreadsheet.SheetOps.axisSize true r place ]
    let inRow (row: int) = placed |> List.filter (fun t -> abs (t.Rect.Y - rowTop row) < 0.5)
    let above (t: Rbt.Spreadsheet.Paint.PaintText) = (t.Rect.Height - t.Baseline) * 6.25
    let bottoms = inRow 2 |> List.map above
    ok "print: text at the bottom of a row on one baseline" (bottoms.Length = 8 && bottoms |> List.forall (fun b -> abs (b - bottoms.Head) < 1e-6 && abs (b - 51.5) <= 4.0)) (sprintf "%A" bottoms)
    ok "print: text at the top and in the middle by the font" (inRow 4 |> List.forall (fun t -> t.Baseline < 0.0) && inRow 3 |> List.forall (fun t -> t.Baseline < 0.0 && t.Baseline > -0.01)) ""
    // Printer GDI's internal leading can exceed descent. At Calibri 11 Normal, n=5 stands 37
    // dots above the grid, n=7 stands 50; the fonts on a mixed row share that baseline.
    let level face pt = Rbt.Spreadsheet.FontMetrics.printLevelPx (Some face) pt false false
    ok "print: GDI internal leading, not screen descent"
        (level "Segoe UI" 10.0 = 4.0 && level "Trebuchet MS" 10.0 = 5.0
         && level "Trebuchet MS" 20.0 = 8.0
         && Rbt.Spreadsheet.FontMetrics.descentPx (Some "Segoe UI") 10.0 false false = 3.0) ""
    let metricModel =
        { createEmpty 3 3 with
            DefaultRowHeight = 60.0; NormalFontName = "Calibri"; NormalFontSize = 11.0
            Cells =
                Map.ofList [
                    { Row = 0; Col = 0 }, { cell (String "metric-a") with Format = { defaultFormat with FontName = Some "Segoe UI"; FontSize = Some 10.0 } }
                    { Row = 0; Col = 1 }, { cell (String "metric-b") with Format = { defaultFormat with FontName = Some "Trebuchet MS"; FontSize = Some 10.0 } }
                    { Row = 1; Col = 0 }, { cell (String "metric-c") with Format = { defaultFormat with FontName = Some "Calibri"; FontSize = Some 20.0 } } ] }
    let metricText =
        [ for item in Paint.buildPrintPaintItems (0.0, 0.0, 1.0) metricModel viewport do
            match item with
            | Rbt.Spreadsheet.Paint.DrawText t when t.Text.StartsWith "metric-" ->
                yield t.Text, (t.Rect.Height - t.Baseline) * 6.25
            | _ -> () ] |> Map.ofList
    ok "print: mixed-font row printer baseline"
        (metricText.Count = 3
         && abs (metricText.["metric-a"] - 37.0) < 0.01
         && abs (metricText.["metric-b"] - 37.0) < 0.01
         && abs (metricText.["metric-c"] - 50.0) < 0.01)
        (sprintf "%A" metricText)


// ----- Excel 2019 Page Break Preview's breaks (legacy *.excel.tsv fields;
// fresh probes namespace them preview.*) and direct PDF page counts (direct.pages). -----
do
    let dir = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "test-documents-excel")
    let mutable sheets = 0
    for tsv in System.IO.Directory.GetFiles(dir, "synthetic-print-pages-*.excel.tsv") |> Array.sort do
        let book = tsv.Replace(".excel.tsv", ".xlsx")
        let facts = System.IO.File.ReadAllLines tsv |> Array.map (fun line -> line.Split '\t') |> Array.filter (fun parts -> parts.Length = 4)
        let imported = Rbt.Spreadsheet.XlsxBridge.Import.importWorkbook (XlsxReader.loadBytesNative (System.IO.File.ReadAllBytes book) |> Option.get)
        imported.Sheets |> Array.iteri (fun index sheet ->
            let mine = facts |> Array.filter (fun parts -> parts.[1] = string (index + 1))
            let get key = mine |> Array.tryPick (fun parts -> if parts.[2] = key then Some parts.[3] else None)
            let preview = Pagination.paginatePreview sheet.Model
            let area = preview.PrintBounds
            let starts (pick: Pagination.PageSlice -> int) = preview.Pages |> List.map (fun page -> pick page + 1) |> List.distinct |> List.sort |> (fun xs -> if xs.IsEmpty then xs else xs.Tail)
            let breaks key last =
                let key = if Option.isSome (get ("preview." + key)) then "preview." + key else key
                mine |> Array.choose (fun parts -> if parts.[2] = key then Some (int (parts.[3].Split(' ').[0])) else None)
                |> Array.distinct |> Array.sort |> Array.filter (fun b -> b <= last) |> List.ofArray
            let previewPages = get "pages" |> Option.map int
            let directPages = get "direct.pages" |> Option.map int
            let excelRows, excelCols = breaks "hbreak" (area.To.Row + 1), breaks "vbreak" (area.To.Col + 1)
            sheets <- sheets + 1
            ok (sprintf "print pages: %s / %s" (System.IO.Path.GetFileNameWithoutExtension book) sheet.Name)
                ((previewPages |> Option.forall ((=) preview.Pages.Length))
                 && (directPages |> Option.forall ((=) (Pagination.paginatePrinted sheet.Model).Pages.Length))
                 && starts (fun page -> page.RowStart) = excelRows && starts (fun page -> page.ColStart) = excelCols)
                (sprintf "preview pages %d (%A), direct pages %A, rows %A (%A), columns %A (%A), scale %g"
                    preview.Pages.Length previewPages directPages
                    (starts (fun page -> page.RowStart)) excelRows (starts (fun page -> page.ColStart)) excelCols preview.Scale))
    let refitBook = System.IO.Path.Combine(dir, "synthetic-print-pages-bestfit-refit.xlsx")
    let refit = Rbt.Spreadsheet.XlsxBridge.Import.importWorkbook (XlsxReader.loadBytesNative (System.IO.File.ReadAllBytes refitBook) |> Option.get)
    let wide, narrow, textOnly = refit.Sheets.[0].Model, refit.Sheets.[1].Model, refit.Sheets.[2].Model
    let screen = Pagination.withBestFitWidths
    ok "print bestFit: imported flags" (Set.contains 1 wide.BestFitColumns && Set.contains 4 narrow.BestFitColumns && Set.contains 1 textOnly.BestFitColumns) ""
    ok "print bestFit: no narrow number" (getColumnWidth 1 (screen wide) = 84.0 && getColumnWidth 3 (screen wide) = 84.0) ""
    ok "print bestFit: whole numeric sheet re-fits"
        ([ 1; 2; 3 ] |> List.forall (fun col -> getColumnWidth col (screen narrow) = 85.0)
         && getColumnWidth 4 (screen narrow) = 19.0
         && getColumnWidth 1 (screen textOnly) = getColumnWidth 1 textOnly
         && getColumnWidth 2 (screen textOnly) = 85.0)
        (sprintf "wide %A narrow %A text %A" (screen wide).ColumnWidths (screen narrow).ColumnWidths (screen textOnly).ColumnWidths)
    // Excel 2019's trigger sweep (print scale 41 and 100 alike): a Calibri 12 General 2010
    // at 37 px starts re-fitting the entire sheet (51 px; a wide companion 96 -> 97);
    // at 38 px nothing is re-fitted, even though the 60%-zoom target is still 51.
    let number format value = { Value = CellValue.Number value; Format = format; Comment = None }
    let calibri12 = { defaultFormat with FontName = Some "Calibri"; FontSize = Some 12.0 }
    let boundary =
        { createEmpty 4 4 with
            NormalFontName = "Calibri"; NormalFontSize = 12.0; MaxDigitWidth = 8.0
            BestFitColumns = Set.ofList [ 0; 1 ]
            ColumnWidths = Map.ofList [ 0, 37.0; 1, 96.0 ]
            Cells = Map.ofList [ { Row = 0; Col = 0 }, number calibri12 2010.0
                                 { Row = 0; Col = 1 }, number calibri12 1.0 ] }
    let at37 = screen boundary
    let at38 = screen { boundary with ColumnWidths = Map.add 0 38.0 boundary.ColumnWidths }
    ok "print bestFit: screen trigger at 37 not 38"
        (getColumnWidth 0 at37 = 51.0 && getColumnWidth 1 at37 = 97.0
         && getColumnWidth 0 at38 = 38.0 && getColumnWidth 1 at38 = 96.0)
        (sprintf "at37 %A at38 %A" at37.ColumnWidths at38.ColumnWidths)
    let boldNumber = { calibri12 with Bold = true; NumberFormat = Some "#,##0" }
    let boldCells = Map.add { Row = 0; Col = 0 } (number boldNumber 785977.0) boundary.Cells
    let boldBoundary = { boundary with Cells = boldCells; ColumnWidths = Map.add 0 57.0 boundary.ColumnWidths }
    let at57 = screen boldBoundary
    let at58 = screen { boldBoundary with ColumnWidths = Map.add 0 58.0 boldBoundary.ColumnWidths }
    ok "print bestFit: formatted trigger at 57 not 58"
        (getColumnWidth 0 at57 = 91.0 && getColumnWidth 1 at57 = 97.0
         && getColumnWidth 0 at58 = 58.0 && getColumnWidth 1 at58 = 96.0)
        (sprintf "at57 %A at58 %A" at57.ColumnWidths at58.ColumnWidths)
    ok "print bestFit: refit changes the break"
        ((Pagination.paginatePreview wide).Pages.Length = 1
         && (Pagination.paginatePreview narrow).Pages.Length = 2
         && (Pagination.paginatePreview narrow).Pages.[1].ColStart = 3)
        (sprintf "wide %A narrow %A" (Pagination.paginatePreview wide).Pages (Pagination.paginatePreview narrow).Pages)
    // Direct Worksheet.ExportAsFixedFormat of the same public fixture leaves every
    // stored width unchanged. The three wide numeric columns fit page one; only
    // the narrow trigger moves to page two. Page Break Preview, by contrast,
    // re-fits them and moves the third wide column to page two.
    let directWidths = Pagination.withPrintSizes narrow
    let directPages = Pagination.paginatePrinted narrow
    let directPitchPt = getColumnWidth 1 directWidths * 0.75
    ok "print bestFit: direct PDF retains stored widths"
        (getColumnWidth 1 narrow = 84.0
         && abs (directPitchPt - 67.70) < 0.05
         && getColumnWidth 1 directWidths < getColumnWidth 1 (Pagination.withPrintSizes (screen narrow)))
        (sprintf "stored %g pt, refitted %g pt" directPitchPt (getColumnWidth 1 (Pagination.withPrintSizes (screen narrow)) * 0.75))
    ok "print bestFit: direct PDF and preview break at different columns"
        (directPages.Pages.Length = 2
         && directPages.Pages.[0].ColEnd = 3 && directPages.Pages.[1].ColStart = 4
         && (Pagination.paginatePreview narrow).Pages.[0].ColEnd = 2)
        (sprintf "direct %A preview %A" directPages.Pages (Pagination.paginatePreview narrow).Pages)
    // A direct export does not widen a best-fit number merely to avoid overflow.
    // In the public BestFit fixture, B2 and D2 both contain 123456789 in
    // 5-character-wide columns: Excel's direct PDF shows #### in both.
    let narrowNumber =
        { createEmpty 3 3 with
            ColumnWidths = Map.ofList [ 1, 35.0 ]
            BestFitColumns = Set.singleton 1
            Cells = Map.ofList [ { Row = 0; Col = 1 }, number defaultFormat 123456789.0 ] }
    let numberViewport : Rbt.Spreadsheet.Layout.Viewport =
        { ScrollX = 0.0; ScrollY = 0.0; Width = 600.0; Height = 120.0
          HeaderHeight = 0.0; RowHeaderWidth = 0.0
          Measurer = Some (fun text _ _ _ _ -> 7.0 * float text.Length) }
    let directText =
        Rbt.Spreadsheet.Paint.buildPrintPaintItems (0.0, 0.0, 1.0)
            (Pagination.withPrintSizes narrowNumber) numberViewport
        |> List.choose (function Rbt.Spreadsheet.Paint.DrawText text -> Some text.Text | _ -> None)
    ok "print bestFit: direct numeric overflow stays as hashes"
        (List.contains "####" directText)
        (sprintf "%A" directText)
    // Fresh direct Excel PDF: formatted numbers fit in printer-dot geometry,
    // not in the screen's GDI metrics and not across an empty neighbour.
    // The public synthetic-print-overflow.xlsx sweeps column widths; its
    // print-overflow-direct.excel.tsv includes a separate percentage holdout.
    let overflow =
        Rbt.Spreadsheet.XlsxBridge.Import.importWorkbook
            (XlsxReader.loadBytesNative
                (System.IO.File.ReadAllBytes (System.IO.Path.Combine(dir, "synthetic-print-overflow.xlsx"))) |> Option.get)
    let printed = overflow.Sheets |> Array.map (fun sheet -> Pagination.withPrintSizes sheet.Model)
    let printMeasure scale : Rbt.Spreadsheet.Layout.TextMeasurer =
        fun text pt family bold italic ->
            let dots = int (System.Math.Floor(pt * scale * 600.0 / 72.0 + 0.5))
            let face = defaultArg family "Calibri"
            Rbt.Spreadsheet.ExcelText.textWidth (fun _ -> 0.0) face bold italic dots false text * 96.0 / 600.0
    let printShown sheet col scale =
        let model = printed.[sheet]
        let address = { Row = 2; Col = col }
        let cell = Map.tryFind address model.Cells
        let width = getColumnWidth col model * scale
        (Painter.cellLookFitted (Some { Zoom = scale; Overhang = 0.0; Print = true })
            model 2 col cell (Some (printMeasure scale)) width).Text
    let currencyShown = [ printShown 0 7 1.0; printShown 0 8 1.0; printShown 0 1 1.0 ]
    ok "print overflow: currency threshold and hash budget"
        (currencyShown = [ "########"; "$1,234,567"; "#######" ]) (sprintf "%A" currencyShown)
    ok "print overflow: accounting reserves a leading placeholder"
        (printShown 1 1 1.0 = "######" && printShown 1 12 1.0 = "#########"
         && (printShown 1 13 1.0).Contains("1,234,567"))
        (sprintf "%A" [ printShown 1 1 1.0; printShown 1 12 1.0; printShown 1 13 1.0 ])
    ok "print overflow: General is not currency"
        (printShown 2 1 1.0 = "1234567") ""
    let percentShown = [ printShown 3 15 1.0; printShown 3 16 1.0 ]
    ok "print overflow: percentage holdout"
        (percentShown = [ "###########"; "1234567.89%" ]) (sprintf "%A" percentShown)
    ok "print overflow: trailing placeholder changes fit, not hashes"
        (printShown 8 9 1.0 = "#########"
         && (printShown 8 10 1.0).Trim() = "$1,234,567")
        (sprintf "%A" [ printShown 8 9 1.0; printShown 8 10 1.0 ])
    let scaledShown =
        [ printShown 15 1 0.75; printShown 15 2 0.75; printShown 15 3 0.75
          printShown 18 3 0.75; printShown 18 4 0.75 ]
    ok "print overflow: 75% font advances and column scale"
        (scaledShown = [ "########"; "#########"; "$1,234,567"; "#########"; "$1,234,567" ])
        (sprintf "%A" scaledShown)
    ok "print overflow: 75% leading reserve, trailing reserve"
        (printShown 16 2 0.75 = "#########"
         && printShown 17 1 0.75 = "#######"
         && printShown 17 2 0.75 = "########"
         && (printShown 16 8 0.75).Trim() = "$1,234,567")
        (sprintf "%A" [ printShown 16 2 0.75; printShown 17 1 0.75;
                         printShown 17 2 0.75; printShown 16 8 0.75 ])
    let resized = Rbt.Spreadsheet.SheetOps.setAxisSize false 1 1 40.0 narrow
    let moved = insertColumn 1 narrow
    ok "print bestFit: resize clears flag, insert moves it"
        (not (Set.contains 1 resized.BestFitColumns) && Set.contains 2 resized.BestFitColumns
         && not (Set.contains 1 moved.BestFitColumns) && Set.contains 2 moved.BestFitColumns && Set.contains 5 moved.BestFitColumns) ""
    let output = Rbt.Spreadsheet.XlsxBridge.Export.toWorkbookOut
                    { Sheets = [| wide; resized; textOnly |]
                      Names = [| "WideOnly"; "NarrowTrigger"; "TextOnly" |]
                      Visible = [| true; true; true |]
                      ActiveSheetIndex = 0; NameInfo = Map.empty; DefinedNames = Map.empty
                      Source = Some refit.Source; Macros = None; MacroProject = None }
    let saved = XlsxReader.loadBytesNative (XlsxPackageWriter.writeWorkbookPackageBytes output) |> Option.get
    let keptFlag col = saved.Sheets.[1].Columns |> List.exists (fun c -> c.Index <= col && c.MaxIndex >= col && c.BestFit)
    ok "print bestFit: save clears resized flag" (not (keptFlag 2) && keptFlag 3 && keptFlag 5) ""
    ok "print pages: the fixtures were read" (sheets >= 40) (sprintf "%d sheets" sheets)
    // The sizes: a column of 64 px of Calibri 11 (7 px a digit on the screen, 47 dots in print) prints at 430 dots; 12 px of
    // Arial 10 (46) at 79; 12 px of Calibri 12 (8, 51) at 77 (76.5 up); a row of 20 px of Calibri 11 at 121 dots, of 546 px at 3303.
    let dots (px: float) = System.Math.Round (px * 6.25, 6)
    ok "print sizes: columns" (dots (Pagination.printColumnWidthPx 7.0 47.0 64.0) = 430.0 && dots (Pagination.printColumnWidthPx 7.0 46.0 12.0) = 79.0 && dots (Pagination.printColumnWidthPx 8.0 51.0 12.0) = 77.0) ""
    // The page scale is applied before the 600-dpi snap: a 765-dot column at 90 % prints 689 dots
    // (82.68 pt), not 688.5; an 806-dot column at 75 % prints 605. Page breaks already budget these
    // snapped dots, so this keeps the drawn edges and the page breaks on the same grid.
    let pageDots (px: float) (scale: float) = System.Math.Round (px * 6.25 * scale, 6)
    ok "print sizes: columns snap at the page scale"
        (pageDots (Pagination.printColumnWidthPxAtScale 1.0 122.4) 1.0 = 765.0
         && pageDots (Pagination.printColumnWidthPxAtScale 0.9 122.4) 0.9 = 689.0
         && pageDots (Pagination.printColumnWidthPxAtScale 0.75 (806.0 / 6.25)) 0.75 = 605.0
         && Pagination.printColumnWidthPxAtScale 0.0 122.4 = 122.4
         && Pagination.printColumnWidthPxAtScale 0.9 0.0 = 0.0) ""
    ok "print sizes: the widest digit in print" ([ "Calibri", 11.0; "Arial", 10.0; "Calibri", 12.0; "Arial", 12.0; "Arial MT", 12.0 ] |> List.map (fun (f, p) -> Rbt.Spreadsheet.ExcelText.printDigitWidth f p) = [ 47.0; 46.0; 51.0; 56.0; 56.0 ]) ""
    // A row of 20 pt shows at 26 px and prints as 26.67 px would: 161 dots of Calibri 11.
    ok "print sizes: a row's points" (dots (Pagination.printRowHeightPx 6.05 (Pagination.exactRowPx 20.0 26.0)) = 161.0 && Pagination.exactRowPx 20.0 27.0 = 27.0) ""
    // The grid 0.48 pt in from the margins: the heights fixture's first lines (margins 0.7 / 0.75 in) at x 50.88, y 54.48 from the top.
    let heights = Rbt.Spreadsheet.XlsxBridge.Import.importWorkbook (XlsxReader.loadBytesNative (System.IO.File.ReadAllBytes (System.IO.Path.Combine(dir, "synthetic-print-pages-heights.xlsx"))) |> Option.get)
    let options : Render.RenderOptions = { DefaultFontFamily = "Helvetica"; Fonts = []; CjkFallback = []; SheetName = "Heights"; FileName = "f" }
    let text = System.Text.Encoding.Latin1.GetString (Render.toBytes options heights.Sheets.[0].Model)
    // A page prints its own rows' text, not the row after its last nor before its first (rows of 121 dots summed in
    // floating point): synthetic-print-pages-titles' page 1 is rows 1-47, page 3 rows 94-137 under row 20.
    let calibri = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "..", "artifacts", "fonts", "windows", "calibri.ttf")
    if System.IO.File.Exists calibri then
        let titles = Rbt.Spreadsheet.XlsxBridge.Import.importWorkbook (XlsxReader.loadBytesNative (System.IO.File.ReadAllBytes (System.IO.Path.Combine(dir, "synthetic-print-pages-titles.xlsx"))) |> Option.get)
        let face = Render.fontFace "Calibri" "Calibri" false false (System.IO.File.ReadAllBytes calibri)
        let withFont : Render.RenderOptions = { DefaultFontFamily = "Calibri"; Fonts = [ face ]; CjkFallback = []; SheetName = "TitlesInside"; FileName = "f" }
        let pdf = Render.toBytes withFont titles.Sheets.[0].Model
        let path = System.IO.Path.GetTempFileName()
        System.IO.File.WriteAllBytes(path, pdf)
        let pageText (n: int) =
            let p = System.Diagnostics.Process.Start(System.Diagnostics.ProcessStartInfo("pdftotext", sprintf "-f %d -l %d \"%s\" -" n n path, RedirectStandardOutput = true))
            let t = p.StandardOutput.ReadToEnd() in p.WaitForExit(); t
        let p1, p3 = pageText 1, pageText 3
        ok "print: a page's own rows" (p1.Contains "r47c1" && not (p1.Contains "r48c1") && p3.Contains "r94c1" && p3.Contains "r20c1" && not (p3.Contains "r93c1")) ""
    ok "print: the grid 0.48 pt inside the margins" (text.Contains "50.88 " && text.Contains (sprintf "%g " (792.0 - 54.48))) ""
    ok "print sizes: rows" (dots (Pagination.printRowHeightPx 6.05 20.0) = 121.0 && dots (Pagination.printRowHeightPx 6.05 546.0) = 3303.0 && dots (Pagination.printRowHeightPx (130.0 / 21.0) 10.0) = 62.0) ""


// ----- Printed hash count: the old-failing public case, through the PDF consumer -----
// Aptos Narrow 14 in a 8.25-character column (58 px) printed 5 hashes before the printer-unit recount; Excel
// prints 6. The same sheet keeps ordinary, full-number and literal-hash controls.
do
    let aptos = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "..", "artifacts", "fonts", "windows-extra", "aptosnarrow-27970306811.ttf")
    if System.IO.File.Exists aptos then
        let fmt : CellFormat =
            { defaultFormat with FontName = Some "Aptos Narrow"; FontSize = Some 14.0; HorizontalAlign = Some "right"; NumberFormat = Some "0" }
        let cell value format = { Value = value; Format = format; Comment = None }
        let widthPx = Rbt.Spreadsheet.FontMetrics.columnWidthPx 7.0 8.25
        let model =
            { createEmpty 4 1 with
                ColumnWidths = Map.ofList [ 0, widthPx ]
                Cells =
                    Map.ofList
                        [ { Row = 0; Col = 0 }, cell (CellValue.Number 1123456789012.0) fmt
                          { Row = 1; Col = 0 }, cell (CellValue.Number -98.41) { fmt with NumberFormat = Some "0.00" }
                          { Row = 2; Col = 0 }, cell (CellValue.Number 42.0) fmt
                          { Row = 3; Col = 0 }, cell (CellValue.String "#####") { fmt with NumberFormat = None } ] }
        let face = Render.fontFace "Aptos Narrow" "Aptos Narrow" false false (System.IO.File.ReadAllBytes aptos)
        let options : Render.RenderOptions =
            { DefaultFontFamily = "Aptos Narrow"; Fonts = [ face ]; CjkFallback = []; SheetName = "HashCount"; FileName = "f" }
        let path = System.IO.Path.GetTempFileName()
        System.IO.File.WriteAllBytes(path, Render.toBytes options model)
        let p = System.Diagnostics.Process.Start(System.Diagnostics.ProcessStartInfo("pdftotext", sprintf "-layout \"%s\" -" path, RedirectStandardOutput = true))
        let text = p.StandardOutput.ReadToEnd() in p.WaitForExit()
        let lines = text.Split('\n') |> Array.map (fun line -> line.Trim()) |> Array.filter ((<>) "")
        ok "print hash count: Aptos 14 at 8.25 characters prints 6" (lines |> Array.contains "######") text
        ok "print hash count: ordinary number is unchanged" (lines |> Array.exists (fun line -> line.Contains "-98.41")) text
        ok "print hash count: full number is unchanged" (lines |> Array.contains "42") text
        ok "print hash count: literal hashes keep their text" (lines |> Array.contains "#####") text

// ----- Icon and sparkline artwork shared by DOM, native raster and PDF -----
do
    let rect : Rbt.Spreadsheet.Layout.Rect = { X = 100.0; Y = 20.0; Width = 75.0; Height = 30.0 }
    let arrow = Painter.iconPaintItems "3Arrows" 2 rect Painter.VAlignBottom
    let hasArrow =
        match arrow with
        | [ Painter.VecGroup ((_,_,_,_,x,y), [Painter.VecPath (segments,Some _,Some _)]) ] ->
            x = 102.0 && y = 33.0 && segments.Length >= 7
        | _ -> false
    ok "icons: 16px paths at Excel's gutter" hasArrow (sprintf "%A" arrow)
    ok "icons: half star clips same artwork"
        (match Painter.iconPaintItems "3Stars" 1 rect Painter.VAlignMiddle with
         | [ Painter.VecGroup (_, [_;Painter.Clipped (_,[_])]) ] -> true
         | _ -> false)
        (Icons.iconSvg "3Stars" 1)
    let addr = { Row = 2; Col = 3 }
    let group : SparklineGroup =
        { Kind = "line"; SeriesColor = "#376092"; NegativeColor = "#FF0000"; AxisColor = "#000000"
          MarkersColor = "#000000"; FirstColor = "#000000"; LastColor = "#000000"
          HighColor = "#000000"; LowColor = "#000000"; Markers = false; High = false
          Low = false; First = false; Last = false; Negative = false; DisplayXAxis = false
          EmptyAs = "gap"; LineWeight = 1.0; MinAxis = "individual"; MaxAxis = "individual"
          ManualMin = None; ManualMax = None; RightToLeft = false
          Items = [addr, "'Sheet ''B'''!$A$1:$C$1"] }
    let source : Rbt.Spreadsheet.Sparklines.Source =
        fun name ->
            if name = Some "Sheet 'B'" then
                Some (fun row col ->
                    if row = 0 && col <> 1 then Some (cell (Number (float (col+1)))) else None)
            else None
    let resolved = Rbt.Spreadsheet.Sparklines.inCell source [group] addr
    ok "sparklines: quoted cross-sheet, gaps"
        (match resolved with Some (_,v,(lo,hi)) -> v = [|Some 1.0;None;Some 3.0|] && lo = 1.0 && hi = 3.0 | _ -> false)
        (sprintf "%A" resolved)
    let absolute : Rbt.Spreadsheet.Sparklines.Source =
        fun name ->
            if name = Some "Sparklines" then
                Some (fun row col -> if row = 18 && col = 2 then Some (cell (Number 42.0)) else None)
            else None
    ok "sparklines: absolute $C$19 source"
        (Rbt.Spreadsheet.Sparklines.values absolute "Sparklines!$C$19:$D$19" = [|Some 42.0;None|])
        "the native PNG fixture's absolute range no longer crashes"
    let grouped =
        { group with MinAxis = "group"; MaxAxis = "group"
                     Items = group.Items @ [{ Row = 3; Col = 3 }, "'Sheet ''B'''!$A$2:$C$2"] }
    let groupSource : Rbt.Spreadsheet.Sparklines.Source =
        fun name ->
            if name = Some "Sheet 'B'" then
                let own = source name |> Option.get
                Some (fun row col ->
                    if row = 0 then own row col
                    elif row = 1 then Some (cell (Number (if col = 0 then -4.0 else 9.0)))
                    else None)
            else None
    ok "sparklines: group bounds"
        (match Rbt.Spreadsheet.Sparklines.inCell groupSource [grouped] addr with
         | Some (_,_,(lo,hi)) -> lo = -4.0 && hi = 9.0
         | _ -> false)
        "group scaling includes other destinations' source data"
    let model = { createEmpty 5 5 with Sparklines = [{ group with Markers = true }] }
    let viewport : Rbt.Spreadsheet.Layout.Viewport =
        { ScrollX = 0.0; ScrollY = 0.0; Width = 500.0; Height = 100.0
          HeaderHeight = 0.0; RowHeaderWidth = 0.0; Measurer = None }
    ok "sparklines: unpopulated destination"
        (Painter.sparklineCellItems source model viewport |> List.length = 1
         && Painter.sparklineOverlays source model model |> List.length = 1)
        "a blank cell still carries its sparkline in PNG and PDF"
    let printed = setColumnWidth addr.Col (2.0 * getColumnWidth addr.Col model) model
    ok "sparklines: screen artwork scaled for PDF"
        (match Painter.sparklineOverlays source model printed with
         | [(_,_,_,_,w,_,[Painter.VecGroup ((sx,_,_,_,_,_),_)])] ->
             w = getColumnWidth addr.Col printed && sx = 2.0
         | _ -> false)
        "PDF scales the screen-cell picture across the printed column"


// ----- printed italic reserve (docs/excel-parity-tracker.md, "Print italic reserve") -----
do
    let shift win upm pt scale = Rbt.Spreadsheet.FontMetrics.printItalicShiftPt win upm pt scale
    let near a b = abs (a - b) <= 0.001
    ok "print italic: Calibri 11 is 2.04 pt" (near (shift 1950.0 2048.0 11.0 1.0) 2.04) (sprintf "%g" (shift 1950.0 2048.0 11.0 1.0))
    ok "print italic: Calibri 8 is 1.44 pt" (near (shift 1950.0 2048.0 8.0 1.0) 1.44) (sprintf "%g" (shift 1950.0 2048.0 8.0 1.0))
    ok "print italic: Arial 14 is 2.52 pt" (near (shift 1854.0 2048.0 14.0 1.0) 2.52) (sprintf "%g" (shift 1854.0 2048.0 14.0 1.0))
    ok "print italic: Arial 14 at 75 % is 1.92 pt (GDI 1.80: one dot of ascent)" (near (shift 1854.0 2048.0 14.0 0.75) 1.92) (sprintf "%g" (shift 1854.0 2048.0 14.0 0.75))
    ok "print italic: Aptos Narrow 14 is 2.76 pt" (near (shift 2068.0 2048.0 14.0 1.0) 2.76) (sprintf "%g" (shift 2068.0 2048.0 14.0 1.0))
    ok "print italic: Georgia 15 is 2.76 pt" (near (shift 1878.0 2048.0 15.0 1.0) 2.76) (sprintf "%g" (shift 1878.0 2048.0 15.0 1.0))
    ok "print italic: Tahoma 16 at 80 % is 2.52 pt" (near (shift 2049.0 2048.0 16.0 0.8) 2.52) (sprintf "%g" (shift 2049.0 2048.0 16.0 0.8))
    // The format controls (print-italic-format-controls): an unquoted `*` fill in the active section suppresses the
    // reserve; `_)` alone does not; quotes and inactive sections do not.
    let fill (code: string) (n: float) = Rbt.Spreadsheet.CellText.numberFormatUsesFillForValue code n
    ok "print italic: 0.00_) is not a fill layout" (not (fill "0.00_)" 0.5)) ""
    ok "print italic: * 0.00 is a fill layout" (fill "* 0.00" 0.5) ""
    ok "print italic: a quoted *) is not a fill" (not (fill "0.00\"*\"" 0.5)) ""
    ok "print italic: inactive-section * does not count" (not (fill "0.00;[Red]* -0.00" 0.5) && fill "0.00;[Red]* -0.00" -0.5) ""
    ok "print italic: standard accounting is a fill layout" (fill "_($* #,##0.00_);_($* \\(#,##0.00\\);_($* \"-\"??_);_(@_)" -0.5) ""
    let eligible (value: CellValue) (format: CellFormat) (width: float) =
        let model =
            createEmpty 3 3
            |> setCell { Row = 0; Col = 1 } { Value = value; Format = format; Comment = None }
            |> setColumnWidth 1 width
            |> Pagination.withPrintSizes
        let viewport : Rbt.Spreadsheet.Layout.Viewport =
            { ScrollX = 0.0; ScrollY = 0.0; Width = 500.0; Height = 100.0
              HeaderHeight = 0.0; RowHeaderWidth = 0.0
              Measurer = Some (fun text _ _ _ _ -> 7.0 * float text.Length) }
        [ for item in Rbt.Spreadsheet.Paint.buildPrintPaintItems (0.0, 0.0, 1.0) model viewport do
            match item with
            | Rbt.Spreadsheet.Paint.DrawText t when t.Text <> "" && t.Align <> Rbt.Spreadsheet.Layout.AlignCenter ->
                yield t.Text, t.PrintItalicEligible, t.PrintOverflowHashes
            | _ -> () ]
    let italicEligibleOf (_, flag, _) = flag
    let textOf (text, _, _) = text
    let italic = { defaultFormat with Italic = true; FontName = Some "Calibri"; FontSize = Some 11.0; HorizontalAlign = Some "right" }
    let ordinary = eligible (Number 0.5) { italic with NumberFormat = Some "0.00" } 100.0
    ok "print italic: right-aligned italic number is eligible" (ordinary.Length = 1 && italicEligibleOf ordinary.Head) (sprintf "%A" ordinary)
    let padOnly = eligible (Number 0.5) { italic with NumberFormat = Some "0.00_)" } 100.0
    ok "print italic: trailing _) alone stays eligible" (padOnly.Length = 1 && italicEligibleOf padOnly.Head) (sprintf "%A" padOnly)
    let accounting = eligible (Number -0.5) { italic with NumberFormat = Some "_($* #,##0.00_);_($* \\(#,##0.00\\);_($* \"-\"??_);_(@_)" } 100.0
    ok "print italic: numeric accounting is not eligible" (accounting.Length = 1 && not (italicEligibleOf accounting.Head)) (sprintf "%A" accounting)
    let literalHashes = eligible (String "#######") { italic with NumberFormat = Some "\"#######\"" } 100.0
    ok "print italic: literal/formatted hashes stay eligible" (literalHashes.Length = 1 && italicEligibleOf literalHashes.Head) (sprintf "%A" literalHashes)
    let overflowHashes = eligible (Number -0.9841) { italic with NumberFormat = Some "0.000%" } 20.0
    ok "print italic: ### overflow is not eligible"
        (overflowHashes.Length = 1 && textOf overflowHashes.Head |> Seq.forall ((=) '#') && not (italicEligibleOf overflowHashes.Head))
        (sprintf "%A" overflowHashes)
    let left = eligible (Number 0.5) { italic with HorizontalAlign = Some "left"; NumberFormat = Some "0.00" } 100.0
    ok "print italic: left-aligned is not eligible" (left.Length = 1 && not (italicEligibleOf left.Head)) (sprintf "%A" left)
    // Excel prints a true numeric overflow placeholder centred whatever the cell's alignment; literal text
    // hashes and a custom quoted-hash number format keep the ordinary alignment path.
    let flags (value: CellValue) (format: CellFormat) (width: float) =
        let model =
            createEmpty 3 3
            |> setCell { Row = 0; Col = 1 } { Value = value; Format = format; Comment = None }
            |> setColumnWidth 1 width
            |> Pagination.withPrintSizes
        let viewport : Rbt.Spreadsheet.Layout.Viewport =
            { ScrollX = 0.0; ScrollY = 0.0; Width = 500.0; Height = 100.0
              HeaderHeight = 0.0; RowHeaderWidth = 0.0
              Measurer = Some (fun text _ _ _ _ -> 7.0 * float text.Length) }
        [ for item in Rbt.Spreadsheet.Paint.buildPrintPaintItems (0.0, 0.0, 1.0) model viewport do
            match item with
            | Rbt.Spreadsheet.Paint.DrawText t when t.Text <> "" -> yield t.Text, t.Align, t.PrintOverflowHashes
            | _ -> () ]
    let hashFormat (align: string) : CellFormat =
        { defaultFormat with FontName = Some "Calibri"; FontSize = Some 11.0; HorizontalAlign = Some align; NumberFormat = Some "0.000%" }
    let overflowFlag align =
        match flags (Number -0.9841) (hashFormat align) 20.0
              |> List.filter (fun (text, _, _) -> text <> "" && text |> Seq.forall ((=) '#')) with
        | [ _, _, isOverflow ] -> isOverflow
        | _ -> false
    ok "print hash geometry: true overflow is a placeholder for right alignment" (overflowFlag "right") ""
    ok "print hash geometry: true overflow is a placeholder for left alignment" (overflowFlag "left") ""
    ok "print hash geometry: true overflow is a placeholder for centred alignment" (overflowFlag "center") ""
    let customFlag =
        match flags (Number 123.0) { defaultFormat with FontName = Some "Calibri"; FontSize = Some 11.0; HorizontalAlign = Some "right"; NumberFormat = Some "\"#####\"" } 100.0
              |> List.tryFind (fun (text, _, _) -> text = "#####") with
        | Some (_, _, isOverflow) -> not isOverflow
        | None -> false
    ok "print hash geometry: custom quoted hashes are not overflow" customFlag ""
    let literalFlag =
        match flags (String "#####") { defaultFormat with FontName = Some "Calibri"; FontSize = Some 11.0; HorizontalAlign = Some "right" } 100.0
              |> List.tryFind (fun (text, _, _) -> text = "#####") with
        | Some (_, _, isOverflow) -> not isOverflow
        | None -> false
    ok "print hash geometry: literal hashes are not overflow" literalFlag ""

// ----- Printed fit cutoff: run text, pads and the two blind anchor decisions -----
do
    let codeFormat (code: string) =
        { defaultFormat with
            FontName = Some "Verdana"
            FontSize = Some 13.0
            HorizontalAlign = Some "right"
            NumberFormat = Some code }
    let runOf value code formatted =
        Rbt.Spreadsheet.CellText.printFitRunFor (codeFormat code) value formatted
    let pct = runOf (Number 12345.6789) "0.00%" "12345.68%"
    ok "print fit run: percentage text, no pads" (pct = Some ("12345.68%", "")) (sprintf "%A" pct)
    let grouped = runOf (Number 1234567.0) "#,##0.00" "1,234,567.00"
    ok "print fit run: grouped decimal text, no pads" (grouped = Some ("1,234,567.00", "")) (sprintf "%A" grouped)
    let accounting =
        runOf (Number 1234567.0) "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)" " $ 1,234,567 "
    ok "print fit run: active accounting fill keeps the prior decision"
        (accounting = None) (sprintf "%A" accounting)
    let padOnly = runOf (Number 0.5) "0.00_)" "0.50 "
    ok "print fit run: a trailing pad keeps its own advance"
        (padOnly = Some ("0.50", ")")) (sprintf "%A" padOnly)
    ok "print fit run: General keeps the prior path"
        (runOf (Number 1234567.0) "General" "1234567" = None) (sprintf "%A" (runOf (Number 1234567.0) "General" "1234567"))
    ok "print fit run: dates keep the prior path"
        (runOf (Number 45000.0) "m/d/yyyy" "3/15/2023" = None) (sprintf "%A" (runOf (Number 45000.0) "m/d/yyyy" "3/15/2023"))
    ok "print fit run: literal hash text is not a number fit"
        (runOf (Number 123.0) "\"#####\"" "#####" = None) (sprintf "%A" (runOf (Number 123.0) "\"#####\"" "#####"))
    ok "print fit run: text values keep the prior path"
        (runOf (CellValue.String "#####") "0.00" "#####" = None) ""
    // The printer-unit decision itself (advance, pads, both directions and the hash clamp).
    let flat _ = 5.0
    let restored = Render.printedNumberRun flat 20.0 5.0 0 "42" "" "#####" "42"
    ok "print fit decision: a fitting number Paint had hashed is restored"
        (restored = ("42", false)) (sprintf "%A" restored)
    let hashed = Render.printedNumberRun flat 20.0 5.0 0 "1123456789012" "" "1123456789012" "1123456789012"
    ok "print fit decision: an overflowing number becomes four hashes"
        (hashed = (System.String('#', 4), true)) (sprintf "%A" hashed)
    let padded = Render.printedNumberRun flat 30.0 5.0 0 "$1,2" "()" "$1,2" " $ 1,2 "
    ok "print fit decision: pads keep their advance" (padded = ("$1,2", false)) (sprintf "%A" padded)
    let slotted = Render.printedNumberRun flat 20.0 5.0 2 "1123456789012" "" "1123456789012" "1123456789012"
    ok "print fit decision: leading pads cost hashes"
        (slotted = (System.String('#', 2), true)) (sprintf "%A" slotted)
    let clamped = Render.printedNumberRun flat 2.0 5.0 0 "1" "" "1" "1"
    ok "print fit decision: a flagged placeholder keeps one hash" (clamped = ("#", true)) (sprintf "%A" clamped)
    // Paint carries the run and pads only for the print path's plain, single-line, unrotated numbers.
    let fitFields (value: CellValue) (format: CellFormat) (width: float) =
        let model =
            createEmpty 3 3
            |> setCell { Row = 0; Col = 1 } { Value = value; Format = format; Comment = None }
            |> setColumnWidth 1 width
            |> Pagination.withPrintSizes
        let viewport : Rbt.Spreadsheet.Layout.Viewport =
            { ScrollX = 0.0; ScrollY = 0.0; Width = 500.0; Height = 100.0
              HeaderHeight = 0.0; RowHeaderWidth = 0.0
              Measurer = Some (fun text _ _ _ _ -> 7.0 * float text.Length) }
        [ for item in Rbt.Spreadsheet.Paint.buildPrintPaintItems (0.0, 0.0, 1.0) model viewport do
            match item with
            | Rbt.Spreadsheet.Paint.DrawText t when t.PrintFitText <> "" ->
                yield t.PrintFitText, t.PrintFitPads, t.PrintFitFullText, t.PrintFitItalicEligible
            | _ -> () ]
    let fitFieldsAt family points italic widthChars value formatCode =
        fitFields (Number value)
            { defaultFormat with
                FontName = Some family
                FontSize = Some points
                Italic = italic
                HorizontalAlign = Some "right"
                NumberFormat = Some formatCode }
            (Rbt.Spreadsheet.FontMetrics.columnWidthPx 7.0 widthChars)
    let plainFormat =
        { defaultFormat with
            FontName = Some "Calibri"
            FontSize = Some 11.0
            HorizontalAlign = Some "right"
            NumberFormat = Some "0.00" }
    let plain = fitFields (Number 0.5) plainFormat 100.0
    ok "print fit fields: plain number carries its run and full text"
        (plain = [ "0.50", "", "0.50", false ]) (sprintf "%A" plain)
    ok "print fit fields: General carries none"
        (fitFields (Number 1234567.0) { plainFormat with NumberFormat = None } 100.0 = []) ""
    ok "print fit fields: date carries none"
        (fitFields (Number 45000.0) { plainFormat with NumberFormat = Some "m/d/yyyy" } 100.0 = []) ""
    ok "print fit fields: literal hash text carries none"
        (fitFields (CellValue.String "#####") plainFormat 100.0 = []) ""
    ok "print fit fields: shrink-to-fit carries none"
        (fitFields (Number 0.5) { plainFormat with ShrinkToFit = true } 100.0 = []) ""
    let accountingCells =
        fitFields (Number 1234567.0)
            { plainFormat with NumberFormat = Some "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)" } 100.0
    ok "print fit fields: active accounting fill keeps the prior decision"
        (accountingCells = []) (sprintf "%A" accountingCells)
    let paddedCells =
        fitFields (Number 0.5) { plainFormat with NumberFormat = Some "0.00_)" } 100.0
    ok "print fit fields: a non-fill pad carries its run and pad"
        (match paddedCells with [ text, pads, _, _ ] -> text = "0.50" && pads = ")" | _ -> false)
        (sprintf "%A" paddedCells)
    let italicFormat =
        { plainFormat with FontName = Some "Calibri"; Italic = true; NumberFormat = Some "0.00" }
    let italicFields = fitFields (Number 0.5) italicFormat 100.0
    ok "print fit fields: right-aligned italic number is reserve-eligible"
        (match italicFields with [ _, _, _, eligible ] -> eligible | _ -> false) (sprintf "%A" italicFields)
    let italicAccountingFields =
        fitFields (Number 1234567.0)
            { italicFormat with NumberFormat = Some "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)" } 100.0
    ok "print fit fields: italic active-fill accounting keeps the prior decision"
        (italicAccountingFields = []) (sprintf "%A" italicAccountingFields)
    // The two blind-anchor transitions through the PDF consumer: Verdana 13 upright at 17.375 characters
    // prints the full percentage where the screen-shaped rule hashed it; Georgia 15 italic at 90 % in
    // 16.6875 characters prints 8 hashes where that rule kept the full number (docs/excel-parity-tracker.md).
    let fontDir = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "..", "artifacts", "fonts")
    let renderedLines fontFile family italic points scale widthChars value formatCode =
        let model =
            createEmpty 2 1
            |> setCell { Row = 0; Col = 0 }
                { Value = Number value
                  Format = { defaultFormat with
                                FontName = Some family
                                FontSize = Some points
                                Italic = italic
                                HorizontalAlign = Some "right"
                                NumberFormat = Some formatCode }
                  Comment = None }
            |> setColumnWidth 0 (Rbt.Spreadsheet.FontMetrics.columnWidthPx 7.0 widthChars)
            |> fun m -> { m with PageSetup = { m.PageSetup with Scale = scale } }
        let face = Render.fontFace family family false italic (System.IO.File.ReadAllBytes fontFile)
        let options : Render.RenderOptions =
            { DefaultFontFamily = family; Fonts = [ face ]; CjkFallback = []; SheetName = "FitCutoff"; FileName = "f" }
        let path = System.IO.Path.GetTempFileName()
        System.IO.File.WriteAllBytes(path, Render.toBytes options model)
        let p = System.Diagnostics.Process.Start(
                    System.Diagnostics.ProcessStartInfo("pdftotext", sprintf "-layout \"%s\" -" path,
                                                        RedirectStandardOutput = true))
        let text = p.StandardOutput.ReadToEnd() in p.WaitForExit()
        text.Split('\n') |> Array.map (fun line -> line.Trim()) |> Array.filter ((<>) "")
    // The bundled Verdana face does not round-trip through this in-process PDF subset (all glyphs map to
    // zero), so its anchor is asserted at the Paint level; the Georgia PDF case below exercises the
    // rendered decision end to end.
    let verdana = System.IO.Path.Combine(fontDir, "windows", "verdana.ttf")
    if System.IO.File.Exists verdana then
        let verdanaFields =
            fitFieldsAt "Verdana" 13.0 false 17.375 12345.6789 "0.00%"
        ok "print fit cutoff: Verdana 13 at 17.375 carries the full percentage run"
            (match verdanaFields with [ text, _, full, _ ] -> text = "1234567.89%" && full = "1234567.89%" | _ -> false)
            (sprintf "%A" verdanaFields)
    let georgiaItalic = System.IO.Path.Combine(fontDir, "windows-extra", "georgiai.ttf")
    if System.IO.File.Exists georgiaItalic then
        let lines = renderedLines georgiaItalic "Georgia" true 15.0 90 16.6875 1234567.0 "#,##0.00"
        ok "print fit cutoff: Georgia 15 italic at 16.6875 prints 8 hashes"
            (lines |> Array.exists (fun line -> line = "########")) (String.concat "|" lines)
    // The printed column snaps at the scaled 600-dpi pitch: Georgia 15 upright at 90 % in 16.2188
    // characters is 82.68 pt (not the 82.62 pt of 765 dots x 0.9), so it prints the full number;
    // one step lower (16.1875, 81.96 pt) still prints hashes. The independent border-edge fixture
    // measures 82.68-82.71 pt for these columns (docs/excel-parity-tracker.md).
    let georgia = System.IO.Path.Combine(fontDir, "windows-extra", "georgia.ttf")
    if System.IO.File.Exists georgia then
        let wide = renderedLines georgia "Georgia" false 15.0 90 16.2188 1234567.0 "#,##0.00"
        ok "print fit cutoff: Georgia 15 upright at 16.2188 prints the full number"
            (wide |> Array.exists (fun line -> line.Contains "1,234,567.00" && not (line.Contains "#")))
            (String.concat "|" wide)
        let narrow = renderedLines georgia "Georgia" false 15.0 90 16.1875 1234567.0 "#,##0.00"
        ok "print fit cutoff: Georgia 15 upright at 16.1875 prints 8 hashes"
            (narrow |> Array.exists (fun line -> line = "########")) (String.concat "|" narrow)
    // Active `*` fill layouts (accounting) keep the prior print decision through the PDF: a narrow cell
    // prints the counted hashes and a wide one the fill-expanded number - never the raw unexpanded text.
    let calibri = System.IO.Path.Combine(fontDir, "windows", "calibri.ttf")
    if System.IO.File.Exists calibri then
        let accounting = "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)"
        let narrow = renderedLines calibri "Calibri" false 11.0 100 10.0 1234567.0 accounting
        ok "print fit cutoff: narrow accounting keeps its hash run"
            (narrow |> Array.exists (fun line -> line.Length > 0 && line |> Seq.forall ((=) '#')))
            (String.concat "|" narrow)
        let wide = renderedLines calibri "Calibri" false 11.0 100 22.0 1234567.0 accounting
        ok "print fit cutoff: wide accounting prints its fill-expanded number"
            (wide |> Array.exists (fun line -> line.Contains "1,234,567.00" && not (line.Contains "#")))
            (String.concat "|" wide)

// ----- printed text under a custom number format's fourth section -----
do
    let text = Rbt.Spreadsheet.CellText.textUnderFormat
    // Excel formats a text value only by a four-section format's fourth section, or by a one-section
    // code with an active `@`. Two or three sections leave it as stored; a quoted/escaped `@` is literal.
    ok "print text section: the fourth section formats text" (text "0;[Red]-0;0;@_)" "Ab1" = "Ab1 ") (text "0;[Red]-0;0;@_)" "Ab1")
    ok "print text section: an empty fourth section hides text" (text "0;0;0;" "Ab1" = "" && text ";;;" "Ab1" = "") ""
    ok "print text section: two or three sections leave text as stored" (text "0;@_)" "Ab1" = "Ab1" && text "0;0;@_)" "Ab1" = "Ab1") ""
    ok "print text section: a one-section @ replaces the text" (text "@" "Ab1" = "Ab1" && text "\"x\"@" "Ab1" = "xAb1") ""
    ok "print text section: a quoted or escaped @ is literal" (text "\"@\"" "Ab1" = "Ab1" && text "\\@" "Ab1" = "Ab1") ""
    ok "print text section: quoted and escaped underscores are literal" (text "0;0;0;\"_\"@" "Ab1" = "_Ab1" && text "0;0;0;\\_@" "Ab1" = "_Ab1") ""
    ok "print text section: literals before @ print" (text "0;0;0;\"Pre \"@" "Ab1" = "Pre Ab1") (text "0;0;0;\"Pre \"@" "Ab1")
    ok "print text section: the accounting fourth pads the text" (text "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)" "Ab1" = " Ab1 ") ""
    // Numeric cells under escaped literal formats keep the literal and do not read the escape as
    // a pad/placeholder/fill token (synthetic-print-textpad-escapes, E01-E10).
    let num = Rbt.Spreadsheet.CellText.formatNumberWithCode
    ok "print text section: escaped underscore is a literal in a number"
        (num "\\_0.00" 1234.5 = "_1234.50" && num "0.00\\_" 1234.5 = "1234.50_" && num "\\_#,##0.00_)" 1234.5 = "_1,234.50 ") ""
    ok "print text section: escaped @ is not a text placeholder"
        (num "\\@0.00" 1234.5 = "@1234.50" && num "0.00\\@" 1234.5 = "1234.50@") (sprintf "%A / %A" (num "\\@0.00" 1234.5) (num "0.00\\@" 1234.5))
    ok "print text section: escaped minus, backslash and space stay literal"
        (num "0.00\\-" 1234.5 = "1234.50-" && num "0.00\\\\" 1234.5 = "1234.50\\" && num "0.00\\ " 1234.5 = "1234.50 ") ""
    let fmt code align = { defaultFormat with NumberFormat = Some code; FontName = Some "Aptos Narrow"; FontSize = Some 14.0; HorizontalAlign = Some align }
    let pads code = Rbt.Spreadsheet.CellText.textWidthPads (fmt code "right") (CellValue.String "Ab1")
    ok "print text pads: trailing and leading pads" (pads "0;0;0;@_)" = ("", ")") && pads "0;0;0;_(@" = ("(", "") && pads "0;0;0;_(@_)" = ("(", ")")) ""
    ok "print text pads: repeated pads and no pad section" (pads "0;0;0;__(@__" = ("_", "_") && pads "0;@_)" = ("", "") && pads "0;0;@_)" = ("", "")) ""
    ok "print text pads: quotes and escapes are not pads" (pads "0;0;0;\"_\"@" = ("", "") && pads "0;0;0;\\_@" = ("", "")) ""
    let viewport : Rbt.Spreadsheet.Layout.Viewport =
        { ScrollX = 0.0; ScrollY = 0.0; Width = 600.0; Height = 200.0; HeaderHeight = 0.0; RowHeaderWidth = 0.0
          Measurer = Some (fun s _ _ _ _ -> if s = " " then 4.0 else 9.0 * float s.Length) }
    let drawn (printing: bool) code align (value: CellValue) =
        let model =
            createEmpty 4 4
            |> setCell { Row = 0; Col = 1 } { Value = value; Format = fmt code align; Comment = None }
            |> setColumnWidth 1 100.0
            |> Pagination.withPrintSizes
        let items =
            if printing then Rbt.Spreadsheet.Paint.buildPrintPaintItems (0.0, 0.0, 1.0) model viewport
            else Rbt.Spreadsheet.Paint.buildPaintItems model viewport
        [ for item in items do
            match item with
            | Rbt.Spreadsheet.Paint.DrawText t when t.FontSize = 14.0 && t.Text <> "" -> yield t
            | _ -> () ]
    let plain = drawn true "0;0;0;@_" "right" (CellValue.String "Ab1") |> List.exactlyOne
    let screenPadded = drawn false "0;0;0;@_)" "right" (CellValue.String "Ab1") |> List.exactlyOne
    let printedPadded = drawn true "0;0;0;@_)" "right" (CellValue.String "Ab1") |> List.exactlyOne
    // The screen keeps the stored text; the print path adds the hidden pad (a space plus the pad's
    // advance less a space's) and right-aligns it.
    ok "print text paint: screen keeps the stored text" (screenPadded.Text = "Ab1") screenPadded.Text
    ok "print text paint: printed text carries the fourth-section pad" (printedPadded.Text = "Ab1 ") printedPadded.Text
    ok "print text paint: a right pad narrows the printed run"
        (abs (printedPadded.Rect.X - plain.Rect.X) < 1e-9 && abs (plain.Rect.Width - printedPadded.Rect.Width - 5.0) < 1e-9)
        (sprintf "%A / %A" plain.Rect printedPadded.Rect)
    let printedLead = drawn true "0;0;0;_(@" "left" (CellValue.String "Ab1") |> List.exactlyOne
    let plainLead = drawn true "0;0;0;@_" "left" (CellValue.String "Ab1") |> List.exactlyOne
    ok "print text paint: a leading pad indents printed text"
        (printedLead.Text = " Ab1" && abs (printedLead.Rect.X - (plainLead.Rect.X + 5.0)) < 1e-9)
        (sprintf "%A / %A" plainLead.Rect printedLead.Rect)
    let printedCenter = drawn true "0;0;0;_(@_)" "center" (CellValue.String "Ab1") |> List.exactlyOne
    let plainCenter = drawn true "0;0;0;@_" "center" (CellValue.String "Ab1") |> List.exactlyOne
    ok "print text paint: both pads narrow and indent a centred run"
        (abs (printedCenter.Rect.X - (plainCenter.Rect.X + 5.0)) < 1e-9
         && abs (plainCenter.Rect.Width - printedCenter.Rect.Width - 10.0) < 1e-9)
        (sprintf "%A / %A" plainCenter.Rect printedCenter.Rect)
    let printedPrefix = drawn true "0;0;0;\"Pre \"@" "left" (CellValue.String "Ab1") |> List.exactlyOne
    ok "print text paint: printed literals appear" (printedPrefix.Text = "Pre Ab1") printedPrefix.Text
    let hidden = drawn true "0;0;0;" "left" (CellValue.String "Ab1")
    ok "print text paint: an empty fourth section prints nothing" (List.isEmpty hidden) ""
    // Numeric controls are not text: the fourth section never applies to a number.
    let numeric = drawn true "0;0;0;@_)" "right" (CellValue.Number 1234.5) |> List.exactlyOne
    let numericPlain = drawn true "0;0;0;0" "right" (CellValue.Number 1234.5) |> List.exactlyOne
    ok "print text paint: a numeric control keeps its number section" (numeric.Text = numericPlain.Text && abs (numeric.Rect.Width - numericPlain.Rect.Width) < 1e-9) (sprintf "%A / %A" numericPlain numeric)

printfn "Print-helper tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
