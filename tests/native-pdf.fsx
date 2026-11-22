open System
open System.Diagnostics
open System.IO
open System.IO.Compression
open System.Text.RegularExpressions

// Exercise the actual CLI font adapters without a configured Windows font set.
// Native rendering still requires installed fallback fonts; PDF inspection uses Poppler.
let root = Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, ".."))
let scratch = Path.Combine(Path.GetTempPath(), "rbt-native-pdf-" + Guid.NewGuid().ToString("N"))
Directory.CreateDirectory scratch |> ignore

let run executable arguments =
    let start = ProcessStartInfo(executable)
    start.WorkingDirectory <- scratch
    start.UseShellExecute <- false
    start.RedirectStandardOutput <- true
    start.RedirectStandardError <- true
    start.Environment.["RB_WORD_FONT_DIR"] <- scratch
    for argument in arguments do start.ArgumentList.Add argument
    use child = Process.Start start
    let output = child.StandardOutput.ReadToEndAsync()
    let errors = child.StandardError.ReadToEndAsync()
    child.WaitForExit()
    let text, error = output.GetAwaiter().GetResult(), errors.GetAwaiter().GetResult()
    if child.ExitCode <> 0 then failwithf "%s failed (%d): %s%s" executable child.ExitCode text error
    text

let check label condition =
    if not condition then failwith label

try
    let workbook = Path.Combine(root, "tests/test-documents-excel/synthetic-drawing-extras.xlsx")
    let sheetPdf = Path.Combine(scratch, "sheet.pdf")
    run "dotnet" [ Path.Combine(root, "tools/spreadsheet-pdf-runner/bin/Debug/net10.0/SpreadsheetPdfRunner.dll"); workbook; sheetPdf ] |> ignore
    let sheetText = run "pdftotext" [ sheetPdf; "-" ]
    check "Spreadsheet PDF must retain A1's value without Windows fonts" (Regex.IsMatch(sheetText, @"\b42\b"))
    let images = run "pdfimages" [ "-list"; sheetPdf ]
    check "Spreadsheet PDF must paint the imported 24x16 RGB picture"
        (Regex.IsMatch(images, @"(?m)^\s*1\s+\d+\s+image\s+24\s+16\s+rgb\s+3\s+8\s"))

    // Excel's linked picture is a live camera of Source!A1:B2, not the
    // drawing's saved EMF preview. Its anchor straddles a manual page break.
    let cameraWorkbook = Path.Combine(root, "tests/test-documents-excel/synthetic-print-linked-picture.xlsx")
    let cameraPdf = Path.Combine(scratch, "camera.pdf")
    run "dotnet" [ Path.Combine(root, "tools/spreadsheet-pdf-runner/bin/Debug/net10.0/SpreadsheetPdfRunner.dll"); cameraWorkbook; cameraPdf; "1" ] |> ignore
    let firstCameraPage = run "pdftotext" [ "-f"; "1"; "-l"; "1"; cameraPdf; "-" ]
    let secondCameraPage = run "pdftotext" [ "-f"; "2"; "-l"; "2"; cameraPdf; "-" ]
    check "Camera must print both source columns on its first page"
        (firstCameraPage.Contains "LEFTTOP" && firstCameraPage.Contains "RIGHTTOP")
    check "Camera must print the source cells it reaches on the next page"
        (secondCameraPage.Contains "RIGHTTOP" && secondCameraPage.Contains "RIGHTBOTTOM")
    check "Camera must not fall back to a labelled drawing frame"
        (not ((firstCameraPage + secondCameraPage).Contains "Object: Picture"))
    let cameraWords = Regex.Split(firstCameraPage + secondCameraPage, @"\s+") |> Set.ofArray
    check "Camera must paint cells without sheet row/column headers"
        (Set.intersect cameraWords (set [ "A"; "B"; "1"; "2" ]) |> Set.isEmpty)

    let document = Path.Combine(scratch, "word.docx")
    do
        use package = ZipFile.Open(document, ZipArchiveMode.Create)
        let part name xml =
            use writer = new StreamWriter(package.CreateEntry(name).Open())
            writer.Write(xml: string)
        part "[Content_Types].xml" """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"""
        part "_rels/.rels" """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"""
        part "word/document.xml" """<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t xml:space="preserve">Native regular </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">bold </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r></w:p></w:body></w:document>"""
    let wordPdf = Path.Combine(scratch, "word.pdf")
    run "dotnet" [ Path.Combine(root, "tools/pdf-runner/bin/Debug/net10.0/PdfRunner.dll"); document; wordPdf ] |> ignore
    let wordText = run "pdftotext" [ wordPdf; "-" ]
    check "Word PDF must retain regular and styled text without Windows fonts"
        (Regex.Replace(wordText, @"\s+", " ").Trim() = "Native regular bold italic")
    printfn "Native PDF font regressions: 3 passed"
finally
    Directory.Delete(scratch, true)
