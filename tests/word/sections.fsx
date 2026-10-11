// Sections as a file states them. A section's properties stand at its end
// (the w:sectPr in its last paragraph's properties, the body's w:sectPr for
// the last section) and say how the section itself starts (w:type: on a new
// page unless "continuous"), its columns, its page and whether its first page
// has a header of its own (w:titlePg). A page shows the header of the first
// section on it.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open Rbt.Word.TextBuffer
module Doc = Rbt.Word.Document
module Layout = Rbt.Word.Layout
module Paint = Rbt.Word.Paint

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-50s expected %A got %A" label expected got

let root = IO.Path.GetFullPath(IO.Path.Combine(__SOURCE_DIRECTORY__, "../.."))
let importFile (path: string) =
    match DocxReader.loadBytesNative (IO.File.ReadAllBytes(IO.Path.Combine(root, path))) with
    | Some parsed -> Rbt.Word.Import.importDoc parsed
    | None -> failwithf "Could not parse %s" path

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let settingsOf (doc: Doc.Document) : Layout.LayoutSettings =
    let s = doc.PageSetup
    { Width = s.PageWidth; PageHeight = s.PageHeight; FontSize = 16.0; LineHeight = 24.0; PaddingLeft = s.MarginLeft; PaddingTop = s.MarginTop
      MarginRight = s.MarginRight; MarginBottom = s.MarginBottom; InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 3.5 }
let textOf (para: Doc.ParagraphBlock) = toString para.Buffer
let paragraphs (doc: Doc.Document) = doc.Blocks |> List.choose Doc.tryGetParagraph

// phpword-sample-11: paragraphs 1-6 are the first section (the sixth, empty, has
// <w:sectPr> ... <w:cols w:space="720"/><w:titlePg/> - one column, a title page, no w:type), the
// seventh alone the second (<w:type w:val="continuous"/><w:cols w:num="2" w:space="720"/>), the
// rest the last (the body's sectPr: continuous, one column).
let doc = importFile "tests/test-documents-wild/phpword-sample-11-readword2007.docx"
let paras = paragraphs doc
let settings = settingsOf doc
let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
let layoutOf (para: Doc.ParagraphBlock) = layouts |> List.find (fun layout -> layout.BlockId = para.Id)
let sectionEndsOf (found: Doc.ParagraphBlock list) = found |> List.indexed |> List.filter (fun (_, para) -> Doc.endsSection para.Metadata) |> List.map fst
check "the sixth and the seventh paragraph each end a section" (sectionEndsOf paras) [ 5; 6 ]
check "the first section: one column, a title page" (paras.[5].Metadata.SectionColumns |> Option.map (fun c -> c.Count) |> Option.defaultValue 1, paras.[5].Metadata.SectionPageSetup |> Option.map (fun s -> s.DifferentFirstPage)) (1, Some true)
check "the second: two columns, continuous" (paras.[6].Metadata.SectionColumns |> Option.map (fun c -> c.Count), paras.[6].Metadata.SectionBreak) (Some 2, Some "continuous")
let fullWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
let multicolumn = paras |> List.findIndex (fun para -> (textOf para).StartsWith "PHPWord can also create multicolumn")
check "the multicolumn paragraph is the seventh" multicolumn 6
// (11906 - 2 x 1418 twips = 9070 = 604.67 px; two columns with 720 twips between them: (9070 - 720) / 2 = 4175 = 278.33 px.)
check "paragraphs 1-6 are as wide as the page's text" [ for i in 0 .. 4 -> Math.Round((layoutOf paras.[i]).AvailableWidth, 1) ] [ for _ in 0 .. 4 -> Math.Round(fullWidth, 1) ]
check "the multicolumn paragraph is one column wide" (Math.Round((layoutOf paras.[6]).AvailableWidth, 1)) (Math.Round((fullWidth - 48.0) / 2.0, 1))
check "the paragraphs after it are as wide as the page again" (Math.Round((layoutOf paras.[8]).AvailableWidth, 1)) (Math.Round(fullWidth, 1))
check "the continuous sections go on on the first page" (Layout.pageIndexForYInPages pages (layoutOf paras.[6]).Y, Layout.pageIndexForYInPages pages (layoutOf paras.[7]).Y) (0, 0)
check "below the first section" ((layoutOf paras.[6]).Y > (layoutOf paras.[4]).Y, (layoutOf paras.[7]).Y >= (layoutOf paras.[6]).Y + (layoutOf paras.[6]).Height - 0.5) (true, true)
let sectionEnds = layouts |> List.filter (fun layout -> Doc.endsSection layout.Metadata)
let headerOf pageIndex =
    let target = (Paint.headerFooterTargetForPage settings doc sectionEnds pages pageIndex true).Value
    target.TargetKind, target.TargetSection, target.TargetPart |> Option.map (fun part -> part.Blocks |> List.choose Doc.tryGetParagraph |> List.map textOf |> String.concat "|")
check "the first page shows the first section's first-page header" (headerOf 0) ("first", 0, Some "First page header")

// Sections that start on a new page: the page after the one the section before ends on - a
// section's last paragraph is not followed by a page of its own.
let para id text : Doc.ParagraphBlock = { Id = id; Buffer = ofString text; Metadata = Doc.defaultMetadata }
let geometry : Doc.SectionPageSetup =
    { PageWidth = 816.0; PageHeight = 1056.0; MarginLeft = 96.0; MarginRight = 96.0; MarginTop = 96.0; MarginBottom = 96.0
      HeaderDistance = 48.0; FooterDistance = 48.0; DifferentFirstPage = false; EvenAndOddHeaders = false }
let ends id kind : Doc.Block = Doc.Paragraph { para id "" with Metadata = { Doc.defaultMetadata with SectionBreak = Some kind; SectionPageSetup = Some geometry } }
let three : Doc.Document =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "a" "section one"); ends "e1" "nextPage"; Doc.Paragraph (para "b" "section two"); ends "e2" "nextPage"; Doc.Paragraph (para "c" "section three") ]
        PageSetup = { Doc.defaultPageSetup with PageWidth = 816.0; PageHeight = 1056.0; MarginLeft = 96.0; MarginRight = 96.0; MarginTop = 96.0; MarginBottom = 96.0 } }
let threeSettings = settingsOf three
let threeLayouts, threePages = Layout.layoutDocumentPages threeSettings measureText measureRange three
let pageOf id = Layout.pageIndexForYInPages threePages (threeLayouts |> List.find (fun layout -> layout.BlockId = id)).Y
check "three sections on new pages: three pages" (threePages.Length, pageOf "a", pageOf "e1", pageOf "b", pageOf "e2", pageOf "c") (3, 0, 0, 1, 1, 2)

printfn "word sections: %d passed, %d failed" pass fail
if fail > 0 then exit 1
