open System
open System.Diagnostics
open System.IO
open System.IO.Compression
open System.Text

// The PDF writer must keep the character a text run held even when the font
// draws two characters with one glyph.  Calibri draws the ASCII hyphen U+002D
// and the Unicode hyphen U+2010 with the same glyph (and the minus sign U+2212
// with another), so a one-entry-per-glyph ToUnicode map cannot say which was
// printed.  The pre-fix output extracted every U+002D as U+2010
// ("31-Jan-16" as "31\u2010Jan\u201016").
//
// The worksheet control is deterministic generated input, checked in as
// tests/test-documents-excel/synthetic-pdf-hyphen-control.xlsx by
// tools/excel-parity/make-pdf-unicode-fixture.py.  The Word control is the same
// characters as separate runs, built in the scratch directory.  Expected text
// is written here as literal code points, not read back from the fixture:
// the assertions check the product, not the fixture's own text.
let root = Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, ".."))
let scratch = Path.Combine(Path.GetTempPath(), "rbt-pdf-unicode-" + Guid.NewGuid().ToString("N"))
Directory.CreateDirectory scratch |> ignore

let hyphen = "\u002D"       // ASCII hyphen-minus, what Excel writes in a format
let unicodeHyphen = "\u2010" // U+2010 HYPHEN
let minusSign = "\u2212"     // U+2212 MINUS SIGN
let mixed = hyphen + unicodeHyphen + minusSign + hyphen + unicodeHyphen + minusSign

let fontsDir = Path.Combine(root, "artifacts", "fonts", "windows")
let calibri = Path.Combine(fontsDir, "calibri.ttf")

let run executable arguments =
    let start = ProcessStartInfo(executable)
    start.WorkingDirectory <- scratch
    start.UseShellExecute <- false
    start.RedirectStandardOutput <- true
    start.RedirectStandardError <- true
    if File.Exists calibri then
        start.Environment.["RB_WORD_FONT_DIR"] <- fontsDir
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

let checkEqual label expected actual =
    if expected <> actual then failwithf "%s: expected %A, got %A" label expected actual

let partText (package: string) (name: string) =
    use archive = ZipFile.OpenRead package
    let entry = archive.GetEntry name
    if isNull entry then failwithf "%s has no %s" package name
    use stream = entry.Open()
    use reader = new StreamReader(stream, Encoding.UTF8)
    reader.ReadToEnd()

let pdfLines path =
    run "pdftotext" [ "-layout"; path; "-" ]
    |> fun text -> text.Replace("\r", "").Split('\n')
    |> Array.map (fun line -> line.Trim())
    |> Array.filter (fun line -> line.Trim() <> "")

try
    // --- The worksheet control's source is the intended public input -------------
    let workbook = Path.Combine(root, "tests/test-documents-excel/synthetic-pdf-hyphen-control.xlsx")
    check "missing worksheet control (run tools/excel-parity/make-pdf-unicode-fixture.py)" (File.Exists workbook)
    check "worksheet generator" (File.Exists(Path.Combine(root, "tools/excel-parity/make-pdf-unicode-fixture.py")))
    let sheet = partText workbook "xl/worksheets/sheet1.xml"
    check "xlsx source: ASCII hyphen cell" (sheet.Contains(sprintf "<t>%s</t>" hyphen))
    check "xlsx source: U+2010 cell" (sheet.Contains(sprintf "<t>%s</t>" unicodeHyphen))
    check "xlsx source: U+2212 cell" (sheet.Contains(sprintf "<t>%s</t>" minusSign))
    check "xlsx source: three ASCII hyphens" (sheet.Contains(sprintf "<t>%s</t>" (String.replicate 3 hyphen)))
    check "xlsx source: three U+2010 hyphens" (sheet.Contains(sprintf "<t>%s</t>" (String.replicate 3 unicodeHyphen)))
    check "xlsx source: three U+2212 signs" (sheet.Contains(sprintf "<t>%s</t>" (String.replicate 3 minusSign)))
    check "xlsx source: mixed cell" (sheet.Contains(sprintf "<t>%s</t>" mixed))
    // The mixed-run cell holds six separate rich-text runs, not one string.
    let sixRuns =
        mixed.ToCharArray()
        |> Array.map (fun c -> sprintf "<r><t>%c</t></r>" c)
        |> String.concat ""
    check "xlsx source: mixed cell is six rich-text runs"
        (sheet.Contains(sprintf "<c r=\"B11\" t=\"inlineStr\"><is>%s</is></c>" sixRuns))
    check "xlsx source: formatted negative number"
        (sheet.Contains "<c r=\"B9\" s=\"1\"><v>-1234.5</v></c>")
    let styles = partText workbook "xl/styles.xml"
    check "xlsx source: date and its yyyy-mm-dd format"
        (sheet.Contains "<c r=\"B10\" s=\"2\"><v>45322</v></c>"
         && styles.Contains "<numFmt numFmtId=\"164\" formatCode=\"yyyy-mm-dd\"/>")

    // --- The Word control, the same characters in separate runs ------------------
    let document = Path.Combine(scratch, "unicode-control.docx")
    do
        use package = ZipFile.Open(document, ZipArchiveMode.Create)
        let part name xml =
            use writer = new StreamWriter(package.CreateEntry(name).Open(), UTF8Encoding(false))
            writer.Write(xml: string)
        part "[Content_Types].xml" """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"""
        part "_rels/.rels" """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"""
        let para (runs: string list) =
            "<w:p>" + (runs |> List.map (fun text -> sprintf "<w:r><w:t xml:space=\"preserve\">%s</w:t></w:r>" text) |> String.concat "") + "</w:p>"
        let body =
            [ para [ "ascii: "; hyphen ]
              para [ "u2010: "; unicodeHyphen ]
              para [ "u2212: "; minusSign ]
              para [ "repeat: "; String.replicate 3 hyphen; " / "; String.replicate 3 unicodeHyphen; " / "; String.replicate 3 minusSign ]
              para ([ "repeat runs: " ] @ [ for c in String.replicate 3 hyphen -> string c ]
                    @ [ " / " ] @ [ for c in String.replicate 3 unicodeHyphen -> string c ])
              para ([ "mixed runs: " ] @ [ for c in mixed -> string c ])
              para [ "date text: 2024-01-31" ] ]
            |> String.concat ""
        part "word/document.xml" (sprintf """<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>%s</w:body></w:document>""" body)
    let documentSource = partText document "word/document.xml"
    let sixWordRuns =
        mixed.ToCharArray()
        |> Array.map (fun c -> sprintf "<w:r><w:t xml:space=\"preserve\">%c</w:t></w:r>" c)
        |> String.concat ""
    check "docx source: mixed line is six runs"
        (documentSource.Contains(sprintf "<w:p><w:r><w:t xml:space=\"preserve\">mixed runs: </w:t></w:r>%s</w:p>" sixWordRuns))

    // --- Spreadsheet export ------------------------------------------------------
    let sheetPdf = Path.Combine(scratch, "unicode-control.pdf")
    run "dotnet" [ Path.Combine(root, "tools/spreadsheet-pdf-runner/bin/Debug/net10.0/SpreadsheetPdfRunner.dll"); workbook; sheetPdf ] |> ignore
    if File.Exists calibri then
        let fonts = run "pdffonts" [ sheetPdf ] |> fun text -> text.Replace("\r", "")
        check "spreadsheet control must be printed in Calibri, whose hyphen glyph is shared" (fonts.Contains "Calibri")
    let sheetRows =
        pdfLines sheetPdf
        |> Array.map (fun line -> line.Split([| ' '; '\t' |], StringSplitOptions.RemoveEmptyEntries))
        |> Array.filter (fun fields -> fields.Length >= 2)
    let valueOf label =
        sheetRows
        |> Array.tryPick (fun fields -> if fields.[0] = label then Some(String.concat " " fields.[1..]) else None)
        |> Option.defaultWith (fun () -> failwithf "spreadsheet PDF has no row %s" label)
    let expectedValues =
        [ "ascii", hyphen
          "u2010", unicodeHyphen
          "u2212", minusSign
          "repeat-ascii", String.replicate 3 hyphen
          "repeat-u2010", String.replicate 3 unicodeHyphen
          "repeat-u2212", String.replicate 3 minusSign
          "mixed", mixed
          "negative-number", hyphen + "1,234.50"
          "date", "2024-01-31"
          "mixed-runs", mixed
          "repeat-runs", String.replicate 3 hyphen
          "repeat-runs-u2010", String.replicate 3 unicodeHyphen ]
    for label, expected in expectedValues do
        checkEqual (sprintf "xlsx PDF value of %s" label) expected (valueOf label)
    // The labels use ASCII hyphens too, so a glyph-only mapping corrupts them.
    for label in [ "repeat-ascii"; "negative-number"; "mixed-runs"; "repeat-runs"; "repeat-runs-u2010" ] do
        check (sprintf "xlsx PDF row label %s" label) (sheetRows |> Array.exists (fun fields -> fields.[0] = label))

    // --- Word export -------------------------------------------------------------
    let wordPdf = Path.Combine(scratch, "unicode-control-word.pdf")
    run "dotnet" [ Path.Combine(root, "tools/pdf-runner/bin/Debug/net10.0/PdfRunner.dll"); document; wordPdf ] |> ignore
    if File.Exists calibri then
        let fonts = run "pdffonts" [ wordPdf ] |> fun text -> text.Replace("\r", "")
        check "Word control must be printed in Calibri, whose hyphen glyph is shared" (fonts.Contains "Calibri")
    let expectedWordLines =
        [ "ascii: " + hyphen
          "u2010: " + unicodeHyphen
          "u2212: " + minusSign
          "repeat: " + String.replicate 3 hyphen + " / " + String.replicate 3 unicodeHyphen + " / " + String.replicate 3 minusSign
          "repeat runs: " + String.replicate 3 hyphen + " / " + String.replicate 3 unicodeHyphen
          "mixed runs: " + mixed
          "date text: 2024-01-31" ]
    checkEqual "Word PDF text" expectedWordLines (pdfLines wordPdf |> List.ofArray)

    printfn "PDF Unicode regression control: 1 xlsx + 1 docx passed"
finally
    Directory.Delete(scratch, true)
