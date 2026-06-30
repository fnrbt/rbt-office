#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/pdf/bin/Debug/net10.0/FsPdf.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet.Pdf/bin/Debug/net10.0/RedBat.Spreadsheet.Pdf.dll"

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

let retainedBytesFor (f: unit -> 'a) =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    let result = f()
    System.GC.KeepAlive result
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before

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
ok "pagi-cache-keeps-first-source"
   (obj.ReferenceEquals(Pagination.paginate m1, p1))
   "pagination should retain the first source after caching a second model"

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

let largePaginationModel docIndex =
    let payload cellIndex =
        sprintf "pagination-retained-%02d-%03d-%s" docIndex cellIndex (System.String('x', 512 * 1024))
    let cells =
        [ for cellIndex in 0 .. 23 ->
            { Row = cellIndex; Col = 0 }, cell (CellValue.String (payload cellIndex)) ]
        |> Map.ofList
    { createEmpty 1000 26 with Cells = cells }

let paginationRetainedBytes =
    retainedBytesFor (fun () ->
        let mutable total = 0
        for docIndex in 0 .. 7 do
            let pagination = Pagination.paginate (largePaginationModel docIndex)
            total <- total + pagination.Pages.Length
        total)
ok "pagi-cache-retained-text-memory"
   (paginationRetainedBytes < 48L * 1024L * 1024L)
   (sprintf "retained %.1f MB" (float paginationRetainedBytes / 1024.0 / 1024.0))
let paginationOversizedSourceCellsWeakReference () =
    let payload = System.String('P', 6 * 1024 * 1024)
    let cells =
        [ { Row = 0; Col = 0 }, cell (CellValue.String payload) ]
        |> Map.ofList
    let model = { createEmpty 10 10 with Cells = cells }
    Pagination.paginate model |> ignore
    System.WeakReference(cells)
let paginationOversizedSourceCellsWeak = paginationOversizedSourceCellsWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "pagi-cache-skips-oversized-source"
   (not paginationOversizedSourceCellsWeak.IsAlive)
   "pagination cache should not retain an oversized source cell map"
let paginationOversizedPageSetupWeakReference () =
    let payload = System.String('H', 6 * 1024 * 1024)
    let pageSetup =
        { defaultPageSetup with
            HeaderFooter = { defaultHeaderFooter with FirstHeader = Some payload } }
    let model = { createEmpty 10 10 with PageSetup = pageSetup }
    Pagination.paginate model |> ignore
    System.WeakReference(payload)
let paginationOversizedPageSetupWeak = paginationOversizedPageSetupWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "pagi-cache-skips-oversized-page-setup"
   (not paginationOversizedPageSetupWeak.IsAlive)
   "pagination cache should not retain oversized page setup metadata"
let paginationOversizedItemSourceCellCount = 300001
let paginationOversizedItemSourceModel =
    let cells =
        [ for row in 0 .. paginationOversizedItemSourceCellCount - 1 ->
            { Row = row; Col = 0 }, cell (CellValue.Number (float row)) ]
        |> Map.ofList
    { createEmpty paginationOversizedItemSourceCellCount 1 with
        Cells = cells
        PageSetup =
            { defaultPageSetup with
                PrintArea = Some { From = { Row = 0; Col = 0 }; To = { Row = 0; Col = 0 } } } }
let paginationOversizedItemSourceMs, paginationOversizedItemSourcePages =
    let sw = System.Diagnostics.Stopwatch.StartNew()
    let mutable total = 0
    for _ in 1 .. 40 do
        total <- total + (Pagination.paginate paginationOversizedItemSourceModel).Pages.Length
    sw.Stop()
    sw.Elapsed.TotalMilliseconds, total
ok "pagi-oversized-item-source-fast"
   (paginationOversizedItemSourcePages = 40
    && paginationOversizedItemSourceMs < 600.0)
   (sprintf "pagination oversized item source took %.1f ms" paginationOversizedItemSourceMs)

// ----- PDF writer -----

let pdfBuilder = FsPdf.Doc.create ()
for i in 1 .. 64 do
    FsPdf.Doc.addPage pdfBuilder (100.0, 100.0) [] [] [ FsPdf.Draw.Raw (sprintf "%% marker page-%02d" i) ] |> ignore

let pdfBytes = FsPdf.Doc.toBytes pdfBuilder
let pdfDoc = FsPdf.Pdf.openData pdfBytes
let pdfPageText pageNumber =
    match FsPdf.Pdf.getPage pdfDoc pageNumber with
    | Some pageRef -> FsPdf.PageOps.getPageContents pdfDoc pageRef |> System.Text.Encoding.UTF8.GetString
    | None -> ""

ok "pdf-writer-page-count"
   (FsPdf.Pdf.numPages pdfDoc = 64)
   (sprintf "pages=%d" (FsPdf.Pdf.numPages pdfDoc))
ok "pdf-writer-page-order"
   ((pdfPageText 1).Contains("page-01") && (pdfPageText 64).Contains("page-64"))
   (sprintf "first=%s last=%s" (pdfPageText 1) (pdfPageText 64))

let allocatedBytesFor (f: unit -> 'a) =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetAllocatedBytesForCurrentThread()
    let result = f()
    System.GC.KeepAlive result
    System.GC.GetAllocatedBytesForCurrentThread() - before

let noCommentPdfCells =
    seq {
        for r in 0 .. 49999 do
            yield { Row = r; Col = 8 }, cell (CellValue.Number (float r))
        yield { Row = 0; Col = 0 }, cell (CellValue.Number 1.0)
    }
    |> Map.ofSeq
let noCommentPdfSetup =
    { defaultPageSetup with
        PrintArea = Some { From = { Row = 0; Col = 0 }; To = { Row = 0; Col = 0 } }
        CellComments = CellCommentsNone }
let noCommentPdfModel =
    { createEmpty 50000 10 with
        Cells = noCommentPdfCells
        PageSetup = noCommentPdfSetup
        Selection = SingleCell(100000, 100000) }
let noCommentPdfOptions: Render.RenderOptions =
    { DefaultFontFamily = "Calibri"
      Fonts = []
      CjkFallback = []
      SheetName = "Sheet1"
      FileName = "bench.xlsx" }
Render.toBytes noCommentPdfOptions noCommentPdfModel |> ignore
let noCommentPdfAllocated =
    allocatedBytesFor (fun () -> Render.toBytes noCommentPdfOptions noCommentPdfModel)
ok "pdf-no-comment-allocation"
   (noCommentPdfAllocated < 500000L)
   (sprintf "allocated %.2f MB" (float noCommentPdfAllocated / 1024.0 / 1024.0))

let offscreenCommentFormat = { defaultFormat with HasComment = true }
let offscreenCommentPdfCells =
    noCommentPdfCells
    |> Map.add { Row = 49999; Col = 8 } { Value = CellValue.Number 49999.0; Format = offscreenCommentFormat; Comment = None }
let offscreenCommentPdfModel =
    { noCommentPdfModel with Cells = offscreenCommentPdfCells }
Render.toBytes noCommentPdfOptions offscreenCommentPdfModel |> ignore
let offscreenCommentPdfAllocated =
    allocatedBytesFor (fun () -> Render.toBytes noCommentPdfOptions offscreenCommentPdfModel)
ok "pdf-offscreen-comment-allocation"
   (offscreenCommentPdfAllocated < 500000L)
   (sprintf "allocated %.2f MB" (float offscreenCommentPdfAllocated / 1024.0 / 1024.0))

let printableCommentPdfCellCount = 300001
let printableCommentPdfCells =
    seq {
        yield { Row = 0; Col = 0 }, { cell (CellValue.Number 1.0) with Format = offscreenCommentFormat }
        for r in 0 .. printableCommentPdfCellCount - 1 do
            yield { Row = r; Col = 8 }, cell (CellValue.Number (float r))
    }
    |> Map.ofSeq
let printableCommentPdfModel =
    { createEmpty printableCommentPdfCellCount 10 with
        Cells = printableCommentPdfCells
        PageSetup = noCommentPdfSetup
        Selection = SingleCell(100000, 100000) }
Render.toBytes noCommentPdfOptions printableCommentPdfModel |> ignore
let printableCommentPdfAllocated =
    allocatedBytesFor (fun () -> Render.toBytes noCommentPdfOptions printableCommentPdfModel)
ok "pdf-printable-comment-allocation"
   (printableCommentPdfAllocated < 1000000L)
   (sprintf "allocated %.2f MB" (float printableCommentPdfAllocated / 1024.0 / 1024.0))

let printDimensionOverrideCount = 300001
let printDimensionOverrideModel =
    { noCommentPdfModel with
        RowHeights = [ for row in 0 .. printDimensionOverrideCount - 1 -> row, 20.0 ] |> Map.ofList
        ColumnWidths = [ for col in 0 .. printDimensionOverrideCount - 1 -> col, 100.0 ] |> Map.ofList }
Render.toBytes noCommentPdfOptions printDimensionOverrideModel |> ignore
let printDimensionOverrideAllocated =
    allocatedBytesFor (fun () -> Render.toBytes noCommentPdfOptions printDimensionOverrideModel)
ok "pdf-print-dimensions-small-area-allocation"
   (printDimensionOverrideAllocated < 1000000L)
   (sprintf "allocated %.2f MB" (float printDimensionOverrideAllocated / 1024.0 / 1024.0))

let implicitPrintDimensionOverrideModel =
    { createEmpty 1 1 with
        Cells = [ { Row = 0; Col = 0 }, cell (CellValue.Number 1.0) ] |> Map.ofList
        PageSetup = { defaultPageSetup with CellComments = CellCommentsNone }
        RowHeights = [ for row in 0 .. printDimensionOverrideCount - 1 -> row, 20.0 ] |> Map.ofList
        ColumnWidths = [ for col in 0 .. printDimensionOverrideCount - 1 -> col, 100.0 ] |> Map.ofList }
Render.toBytes noCommentPdfOptions implicitPrintDimensionOverrideModel |> ignore
let implicitPrintDimensionOverrideAllocated =
    allocatedBytesFor (fun () -> Render.toBytes noCommentPdfOptions implicitPrintDimensionOverrideModel)
ok "pdf-print-dimensions-implicit-used-area-allocation"
   (implicitPrintDimensionOverrideAllocated < 1000000L)
   (sprintf "allocated %.2f MB" (float implicitPrintDimensionOverrideAllocated / 1024.0 / 1024.0))

let multiPageNoFontModel =
    buildModel 300 12
        { defaultPageSetup with
            Options = { defaultPrintOptions with GridLines = true; Headings = true } }
let multiPageNoFontBytes = Render.toBytes noCommentPdfOptions multiPageNoFontModel
let multiPageNoFontDoc = FsPdf.Pdf.openData multiPageNoFontBytes
ok "pdf-multipage-no-font-page-count"
   (FsPdf.Pdf.numPages multiPageNoFontDoc = List.length (Pagination.paginate multiPageNoFontModel).Pages)
   (sprintf "pages=%d" (FsPdf.Pdf.numPages multiPageNoFontDoc))
let multiPageNoFontAllocated =
    allocatedBytesFor (fun () -> Render.toBytes noCommentPdfOptions multiPageNoFontModel)
ok "pdf-multipage-no-font-allocation"
   (multiPageNoFontAllocated < 5L * 1024L * 1024L)
   (sprintf "allocated %.2f MB" (float multiPageNoFontAllocated / 1024.0 / 1024.0))

let spreadsheetPdfFontBytes = System.IO.File.ReadAllBytes("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
let fontPdfOptions: Render.RenderOptions =
    { noCommentPdfOptions with
        DefaultFontFamily = "DejaVu Sans"
        Fonts = [
            { FaceName = "DejaVu Sans"
              Family = "DejaVu Sans"
              Bold = false
              Italic = false
              Data = spreadsheetPdfFontBytes }
        ] }
let multiPageFontModel =
    buildModel 80 8
        { tinyPageSetup with
            Options = { defaultPrintOptions with GridLines = true; Headings = true } }
let multiPageFontBytes = Render.toBytes fontPdfOptions multiPageFontModel
let multiPageFontDoc = FsPdf.Pdf.openData multiPageFontBytes
ok "pdf-multipage-font-page-count"
   (FsPdf.Pdf.numPages multiPageFontDoc = List.length (Pagination.paginate multiPageFontModel).Pages)
   (sprintf "pages=%d" (FsPdf.Pdf.numPages multiPageFontDoc))
let multiPageFontAllocated =
    allocatedBytesFor (fun () -> Render.toBytes fontPdfOptions multiPageFontModel)
ok "pdf-multipage-font-allocation"
   (multiPageFontAllocated < 40L * 1024L * 1024L)
   (sprintf "allocated %.2f MB" (float multiPageFontAllocated / 1024.0 / 1024.0))

let unusedFallbackPdfOptions: Render.RenderOptions =
    { fontPdfOptions with
        CjkFallback = [
            for index in 1 .. 12 ->
                { FaceName = sprintf "UnusedFallback%02d" index
                  Family = sprintf "UnusedFallback%02d" index
                  Bold = false
                  Italic = false
                  Data = spreadsheetPdfFontBytes }
        ] }
Render.toBytes unusedFallbackPdfOptions multiPageFontModel |> ignore
let unusedFallbackAllocated =
    allocatedBytesFor (fun () -> Render.toBytes unusedFallbackPdfOptions multiPageFontModel)
ok "pdf-unused-fallback-font-allocation"
   (unusedFallbackAllocated < 45L * 1024L * 1024L)
   (sprintf "allocated %.2f MB" (float unusedFallbackAllocated / 1024.0 / 1024.0))

let cjkFallbackStressCells =
    [ for r in 0 .. 39 do
        for c in 0 .. 5 ->
            { Row = r; Col = c }, cell (CellValue.String "漢") ]
    |> Map.ofList
let cjkFallbackStressModel =
    { createEmpty 1000 26 with
        Cells = cjkFallbackStressCells
        PageSetup =
            { tinyPageSetup with
                Options = { defaultPrintOptions with GridLines = true; Headings = true } } }
let invalidFallbackPdfOptions: Render.RenderOptions =
    { fontPdfOptions with
        CjkFallback = [
            { FaceName = "InvalidFallback"
              Family = "InvalidFallback"
              Bold = false
              Italic = false
              Data = Array.create (512 * 1024) 127uy }
        ] }
Render.toBytes invalidFallbackPdfOptions cjkFallbackStressModel |> ignore
let invalidFallbackAllocated =
    allocatedBytesFor (fun () -> Render.toBytes invalidFallbackPdfOptions cjkFallbackStressModel)
ok "pdf-invalid-fallback-font-allocation"
   (invalidFallbackAllocated < 25L * 1024L * 1024L)
   (sprintf "allocated %.2f MB" (float invalidFallbackAllocated / 1024.0 / 1024.0))

printfn ""
printfn "Print-helper tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
