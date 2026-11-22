#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
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
            // Every cell of every print area is on exactly one page of its area.
            let covered =
                pagination.PrintAreas |> List.mapi (fun index area ->
                    let pages = pagination.Pages |> List.filter (fun page -> page.Area = index)
                    let cells = pages |> List.sumBy (fun page -> (page.RowEnd - page.RowStart + 1) * (page.ColEnd - page.ColStart + 1))
                    // (The repeat rows / columns print where they are on the pages that begin at or before them.)
                    let rows = area.To.Row - area.From.Row + 1
                    let cols = area.To.Col - area.From.Col + 1
                    cells = max 0 rows * max 0 cols)
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


// ----- pages as Excel 2019's PDF lays them out (make-print-pages-fixture.py; Excel's facts in *.excel.tsv: Page Break
// Preview's breaks and each sheet's pages in its PDF, probe-print-pages.vbs) -----
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
            let pagination = Pagination.paginatePrinted sheet.Model
            let area = pagination.PrintBounds
            let starts (pick: Pagination.PageSlice -> int) = pagination.Pages |> List.map (fun page -> pick page + 1) |> List.distinct |> List.sort |> (fun xs -> if xs.IsEmpty then xs else xs.Tail)
            let breaks key last = mine |> Array.choose (fun parts -> if parts.[2] = key then Some (int (parts.[3].Split(' ').[0])) else None) |> Array.distinct |> Array.sort |> Array.filter (fun b -> b <= last) |> List.ofArray
            let excelPages = match get "pages" with Some n -> int n | None -> 0
            let excelRows, excelCols = breaks "hbreak" (area.To.Row + 1), breaks "vbreak" (area.To.Col + 1)
            sheets <- sheets + 1
            ok (sprintf "print pages: %s / %s" (System.IO.Path.GetFileNameWithoutExtension book) sheet.Name)
                (pagination.Pages.Length = excelPages && starts (fun page -> page.RowStart) = excelRows && starts (fun page -> page.ColStart) = excelCols)
                (sprintf "pages %d (Excel %d), rows %A (%A), columns %A (%A), scale %g" pagination.Pages.Length excelPages
                    (starts (fun page -> page.RowStart)) excelRows (starts (fun page -> page.ColStart)) excelCols pagination.Scale))
    ok "print pages: the fixtures were read" (sheets >= 40) (sprintf "%d sheets" sheets)
    // The sizes: a column of 64 px of Calibri 11 (7 px a digit on the screen, 47 dots in print) prints at 430 dots; 12 px of
    // Arial 10 (46) at 79; 12 px of Calibri 12 (8, 51) at 77 (76.5 up); a row of 20 px of Calibri 11 at 121 dots, of 546 px at 3303.
    let dots (px: float) = System.Math.Round (px * 6.25, 6)
    ok "print sizes: columns" (dots (Pagination.printColumnWidthPx 7.0 47.0 64.0) = 430.0 && dots (Pagination.printColumnWidthPx 7.0 46.0 12.0) = 79.0 && dots (Pagination.printColumnWidthPx 8.0 51.0 12.0) = 77.0) ""
    ok "print sizes: the widest digit in print" ([ "Calibri", 11.0; "Arial", 10.0; "Calibri", 12.0; "Arial", 12.0; "Arial MT", 12.0 ] |> List.map (fun (f, p) -> Rbt.Spreadsheet.ExcelText.printDigitWidth f p) = [ 47.0; 46.0; 51.0; 56.0; 56.0 ]) ""
    // A row of 20 pt shows at 26 px and prints as 26.67 px would: 161 dots of Calibri 11.
    ok "print sizes: a row's points" (dots (Pagination.printRowHeightPx 6.05 (Pagination.exactRowPx 20.0 26.0)) = 161.0 && Pagination.exactRowPx 20.0 27.0 = 27.0) ""
    // The grid 0.48 pt in from the margins: the heights fixture's first lines (margins 0.7 / 0.75 in) at x 50.88, y 54.48 from the top.
    let heights = Rbt.Spreadsheet.XlsxBridge.Import.importWorkbook (XlsxReader.loadBytesNative (System.IO.File.ReadAllBytes (System.IO.Path.Combine(dir, "synthetic-print-pages-heights.xlsx"))) |> Option.get)
    let options : Render.RenderOptions = { DefaultFontFamily = "Helvetica"; Fonts = []; CjkFallback = []; SheetName = "Heights"; FileName = "f" }
    let text = System.Text.Encoding.Latin1.GetString (Render.toBytes options heights.Sheets.[0].Model)
    ok "print: the grid 0.48 pt inside the margins" (text.Contains "50.88 " && text.Contains (sprintf "%g " (792.0 - 54.48))) ""
    ok "print sizes: rows" (dots (Pagination.printRowHeightPx 6.05 20.0) = 121.0 && dots (Pagination.printRowHeightPx 6.05 546.0) = 3303.0 && dots (Pagination.printRowHeightPx (130.0 / 21.0) 10.0) = 62.0) ""


printfn "Print-helper tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
