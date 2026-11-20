// The fonts a document embeds, against what Word 2019 does with the fixtures
// tools/word-parity/fixtures/embedded-fonts.py generates (tests/word/fixtures/embedded-fonts;
// Word's widths of every word from its PDFs in word/widths.tsv, the fonts its PDFs hold in
// word/fonts.tsv). What Word showed (docs/layout-parity-flow.md, ledger):
// - text in a family the machine has no font of is set in the face the document embeds of it,
//   with or without w:embedTrueTypeFonts (reg, reg-noset); a family the machine has is set in
//   the machine's font (calibri: an embedded "Calibri" loses);
// - the family is the font table's w:name (names: "Rbt Outer" is the embedded face, "Rbt Inner" -
//   the face's own name - is not);
// - fsType 2 (restricted) faces are not used (restricted); 4 and 8 are (noprint, editable);
// - a family sets its other styles in the faces it has: bold and italic of the regular face
//   with the regular's advances (reg, sans-reg), regular and italic text of a family that
//   embeds only its bold face in the bold face (bold-only);
// - a character the embedded face lacks (a subset face: subset*, glyphs-*) is set in Calibri at
//   Calibri's advance, whatever the defaults, the theme, the table's PANOSE or altName say: that
//   character alone when the family embeds one face (subset: "Quiz" 6.72 + 12.0 + 3.96pt), the run
//   from it to its end when it embeds more (subset-all4: "Quiz" all Calibri, glyphs-RB "uuQuu":
//   "uu" embedded, "Quu" Calibri). (In one Word session for all the fixtures Word had lost subset's
//   face: .notdef boxes; page-lines.vbs now opens each in a Word of its own.)
// - a line in an embedded face is that face's single line: its hhea ascender, descender and line
//   gap - Liberation Mono's 1.1328 of the size, Carlito's 1.2207, Ubuntu's 1.149, Poppins' 1.5 (the
//   showcase's exhibition-leaflet: 9pt single, 13.56pt; its Windows metrics would give 1.762).
// Ours: Office.Formats' EmbeddedFonts reads and de-obfuscates them; ShaperText measures with them
// added to a font set of fonts/' faces for the families it has none of (addEmbeddedFaces); the PDF
// writer embeds them (Pdf.withEmbeddedFonts); the layout takes their single lines (Layout.fontFileLineFactor);
// a document saved after an edit keeps them as they were.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Pdf.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Unicode.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Bidi.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.OpenType.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Shaper.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Text.Engine.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/ShaperText/bin/Debug/net10.0/Rbt.Word.ShaperText.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Word.Pdf.dll"

open System
open System.IO
open System.IO.Compression
module Shaper = Rbt.Word.ShaperText
module TB = Rbt.Word.TextBuffer
module Docx = Rbt.Word.Docx
module Pdf = Rbt.Word.Pdf

let root = Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, "../.."))
let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/embedded-fonts")
let mutable pass = 0
let mutable fail = 0
let check label ok detail =
    if ok then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-48s %s" label detail

let font file = File.ReadAllBytes(Path.Combine(root, "fonts", file))
let face family bold italic file : Shaper.FontFace = { Family = family; Bold = bold; Italic = italic; Data = font file }
// The machine's fonts: Calibri (Carlito, its widths) and Times New Roman (Liberation Serif).
let machine =
    Shaper.createFontSet "Calibri" [
        face "Calibri" false false "Carlito-Regular.ttf"; face "Calibri" true false "Carlito-Bold.ttf"
        face "Calibri" false true "Carlito-Italic.ttf"; face "Calibri" true true "Carlito-BoldItalic.ttf"
        face "Times New Roman" false false "LiberationSerif-Regular.ttf" ]

let load name = File.ReadAllBytes(Path.Combine(folder, name + ".docx")) |> Docx.loadBytesNative |> Option.get
let embeddedOf (doc: Rbt.Word.Document.Document) =
    doc.Package.Preserved |> Option.map EmbeddedFonts.ofPreserved |> Option.defaultValue []

// ---- what is read ----------------------------------------------------------------------------
let faces name = [ for f in embeddedOf (load name) -> f.Family, f.Bold, f.Italic ]
let regular family = [ family, false, false ]
for name, expected in [
    "reg", regular "Rbt Embed Mono"; "reg-noset", regular "Rbt Embed Mono"
    "all4", [ for b, i in [ false, false; true, false; false, true; true, true ] -> "Rbt Embed Mono", b, i ]
    "bold-only", [ "Rbt Embed Mono", true, false ]; "names", regular "Rbt Outer"
    "restricted", []; "noprint", regular "Rbt Embed Mono"; "editable", regular "Rbt Embed Mono"
    "subset", regular "Rbt Embed Mono"; "sans-reg", regular "Rbt Embed Sans" ] do
    let got = faces name
    check (sprintf "%s: the faces read" name) (got = expected) (sprintf "%A, expected %A" got expected)
let subsetFace = embeddedOf (load "subset") |> List.head
check "subset: w:subsetted read" subsetFace.Subsetted ""
check "de-obfuscated: a TrueType font" (subsetFace.Data.[0] = 0uy && subsetFace.Data.[1] = 1uy && subsetFace.Data.[2] = 0uy && subsetFace.Data.[3] = 0uy) (sprintf "%A" subsetFace.Data.[0..3])
check "obfuscation is its own inverse"
    (EmbeddedFonts.deobfuscate "{3EEE3167-E5B8-4798-AE48-EA6B71E31D4D}" (EmbeddedFonts.deobfuscate "{3EEE3167-E5B8-4798-AE48-EA6B71E31D4D}" subsetFace.Data |> Option.get) |> Option.get = subsetFace.Data) ""
let calibre = File.ReadAllBytes(Path.Combine(root, "tests/test-documents/calibre-demo.docx")) |> Docx.loadBytesNative |> Option.get
let calibreFaces = [ for f in embeddedOf calibre -> f.Family, f.Bold, f.Italic ]
check "calibre-demo: Ubuntu x4, Tahoma, Ubuntu Mono"
    (calibreFaces = [ "Ubuntu", false, false; "Ubuntu", true, false; "Ubuntu", false, true; "Ubuntu", true, true; "Tahoma", false, false; "Ubuntu Mono", false, false ])
    (sprintf "%A" calibreFaces)

// ---- widths, as Word set them ----------------------------------------------------------------
let rows =
    File.ReadAllLines(Path.Combine(folder, "word/widths.tsv"))
    |> Array.filter (fun l -> l.Length > 0)
    |> Array.map (fun l -> let c = l.Split('\t') in c.[0], c.[1], c.[2], float c.[3])
// (Not checked: "Rbt Inner", a family neither the machine nor the document has, and restricted-roman,
// whose face Word may not use and whose table entry is Times New Roman's PANOSE - Word set both in
// Cambria, ours in its default family; that is Word's substitution for a font it has not got, not the
// embedded fonts' rule.)
let mutable listed = 0
let run = "MMMMMMMMMM iiiiiiiiii Quiz"
for fixture, what, word, width in rows do
    let family, style = let i = what.LastIndexOf '/' in what.Substring(0, i), what.Substring(i + 1)
    if family = "Rbt Inner" || fixture = "restricted-roman" then listed <- listed + 1
    else
        let embedded = [ for f in embeddedOf (load fixture) -> ({ Family = f.Family; Bold = f.Bold; Italic = f.Italic; Data = f.Data }: Shaper.FontFace) ]
        let set = Shaper.addEmbeddedFaces machine embedded
        let textStyle = { TB.defaultStyle with FontFamily = Some family; FontSize = 10.0; Bold = style.StartsWith "Bold"; Italic = style.EndsWith "Italic" }
        let measureRange = Shaper.createRangeMeasurer set
        // (A word of the paragraph's one run, measured in the run as the layout measures it; a glyphs-*
        // paragraph's runs each whole.)
        let ours =
            if fixture.StartsWith "glyphs-" then word.Split('|') |> Array.sumBy (fun r -> measureRange r textStyle 0 r.Length)
            else measureRange run textStyle (run.IndexOf word) word.Length
        // (Word's widths are its PDF's word boxes: 1/600in steps, synthetic bold's stroke.)
        check (sprintf "%s %s %s" fixture what word) (abs (ours - width) <= 0.25) (sprintf "Word %.3f ours %.3f" width ours)

// The machine's font is kept when the document embeds one of its family.
check "addEmbeddedFaces keeps the set's own family"
    (Shaper.measureText (Shaper.addEmbeddedFaces machine [ face "Calibri" false false "LiberationMono-Regular.ttf" ]) "iiii" { TB.defaultStyle with FontFamily = Some "Calibri"; FontSize = 10.0 }
        = Shaper.measureText machine "iiii" { TB.defaultStyle with FontFamily = Some "Calibri"; FontSize = 10.0 }) ""
// (A family of the machine's that lacks a style keeps falling back as before: only an embedded family
// sets a style it has no face of in another of its faces.)
let boldOnly = Shaper.createFontSet "Calibri" [ face "Calibri" false false "Carlito-Regular.ttf"; face "Mono" true false "LiberationMono-Bold.ttf" ]
check "a machine family without the style: the default family"
    (Shaper.measureText boldOnly "iiii" { TB.defaultStyle with FontFamily = Some "Mono"; FontSize = 10.0 }
        = Shaper.measureText boldOnly "iiii" { TB.defaultStyle with FontFamily = Some "Calibri"; FontSize = 10.0 }) ""

// ---- lines as tall as the embedded face's single line ------------------------------------------
// (Word's PDFs: the four one-line paragraphs' baselines 33.96pt apart over three lines in reg - Liberation
// Mono's 1.1328 of 10pt - and 36.60pt in sans-reg, Carlito's 1.2207; calibre-demo's Ubuntu 12pt at 1.15
// is 15.84pt a line, 1.149 of its size.)
module Layout = Rbt.Word.Layout
for name, family, wordPitch in [ "reg", "Rbt Embed Mono", 33.96 / 3.0; "sans-reg", "Rbt Embed Sans", 36.60 / 3.0 ] do
    let doc = load name
    let embedded = [ for f in embeddedOf doc -> ({ Family = f.Family; Bold = f.Bold; Italic = f.Italic; Data = f.Data }: Shaper.FontFace) ]
    let set = Shaper.addEmbeddedFaces machine embedded
    let settings = Layout.documentSettings 16.0 24.0 0.25 doc
    let layouts, _ = Layout.layoutDocumentPages settings (Shaper.createMeasurer set) (Shaper.createRangeMeasurer set) doc
    let tops = [ for layout in layouts do for line in layout.Lines -> line.Y ]
    let pitch = (List.last tops - List.head tops) / float (tops.Length - 1) * 0.75
    check (sprintf "%s: a line of %s" name family) (abs (pitch - wordPitch) <= 0.1) (sprintf "Word %.3f ours %.3f" wordPitch pitch)
let ubuntu = embeddedOf calibre |> List.find (fun f -> f.Family = "Ubuntu" && not f.Bold && not f.Italic)
let ubuntuFactor = Layout.fontFileLineFactor ubuntu.Data |> Option.defaultValue 0.0
check "Ubuntu's single line: 1.149 of its size" (abs (ubuntuFactor - 1.149) < 0.0005) (sprintf "%.4f" ubuntuFactor)
let leaflet = File.ReadAllBytes(Path.Combine(root, "tests/test-documents-showcase/exhibition-leaflet.docx")) |> Docx.loadBytesNative |> Option.get
let poppins = embeddedOf leaflet |> List.find (fun f -> f.Family = "Poppins" && not f.Bold && not f.Italic)
let poppinsFactor = Layout.fontFileLineFactor poppins.Data |> Option.defaultValue 0.0
check "Poppins' single line: 1.5 of its size (hhea; not its Windows 1.762)" (abs (poppinsFactor - 1.5) < 0.0005) (sprintf "%.4f" poppinsFactor)
check "Calibri's (Carlito's) single line: 1.2207" (abs ((Layout.fontFileLineFactor (font "Carlito-Regular.ttf") |> Option.defaultValue 0.0) - 1.2207) < 0.0005) ""
check "Arial's (Liberation Sans') single line: 1.1499" (abs ((Layout.fontFileLineFactor (font "LiberationSans-Regular.ttf") |> Option.defaultValue 0.0) - 1.1499) < 0.0005) ""

// ---- the PDF holds the embedded face ----------------------------------------------------------
let pdfOptions : Pdf.RenderOptions =
    { DefaultFontFamily = "Calibri"
      Fonts = [ { FaceName = "Calibri"; Family = "Calibri"; Bold = false; Italic = false; Data = font "Carlito-Regular.ttf" }
                { FaceName = "TimesNewRoman"; Family = "Times New Roman"; Bold = false; Italic = false; Data = font "LiberationSerif-Regular.ttf" } ]
      MeasureText = None; ShapeText = None; Images = [] }
let pdfText (bytes: byte[]) = Text.Encoding.Latin1.GetString bytes
let regPdf = Pdf.toBytes pdfOptions (load "reg") |> pdfText
check "reg: the PDF embeds the document's face" (regPdf.Contains "Embedded-RbtEmbedMono") "no Embedded-RbtEmbedMono in the PDF"
let calibriPdf = Pdf.toBytes pdfOptions (load "calibri") |> pdfText
check "calibri: the PDF keeps the given Calibri" (not (calibriPdf.Contains "Embedded-Calibri")) "the embedded Calibri was used"
check "withEmbeddedFonts given twice adds nothing more"
    ((Pdf.withEmbeddedFonts (load "all4") (Pdf.withEmbeddedFonts (load "all4") pdfOptions)).Fonts.Length = pdfOptions.Fonts.Length + 4) ""

// ---- saved as it was ---------------------------------------------------------------------------
let entries (bytes: byte[]) =
    use zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read)
    [ for e in zip.Entries do
        use s = e.Open()
        use m = new MemoryStream()
        s.CopyTo m
        yield e.FullName, m.ToArray() ] |> Map.ofList
let original = entries (File.ReadAllBytes(Path.Combine(root, "tests/test-documents/calibre-demo.docx")))
let edited =
    let first = Rbt.Word.Document.paragraphsOfBlocks calibre.Blocks |> List.head
    calibre |> Rbt.Word.Document.setCursor first.Id 0 |> Rbt.Word.Document.insertText "Edited "
let package = Docx.exportTypedDocumentPackage edited
let saved = entries (DocumentWriter.writeTypedPackageBytes { package with Unchanged = false })
for path in original |> Map.toList |> List.map fst |> List.filter (fun p -> p.StartsWith "word/fonts/" || p.Contains "fontTable") do
    check (sprintf "saved after an edit: %s as it was" path) (Map.tryFind path saved = Map.tryFind path original) ""
check "saved after an edit: the faces read again" ([ for f in embeddedOf (Docx.loadBytesNative (DocumentWriter.writeTypedPackageBytes { package with Unchanged = false }) |> Option.get) -> f.Family, f.Bold, f.Italic ] = calibreFaces) ""

printfn "Embedded fonts against Word: %d as Word, %d not (%d listed, not checked)" pass fail listed
if fail > 0 then exit 1
