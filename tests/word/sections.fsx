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
      MarginRight = s.MarginRight; MarginBottom = s.MarginBottom; InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 3.5; DefaultTabStop = s.DefaultTabStop; CompatibilityMode = s.CompatibilityMode; HangingPunctuation = Rbt.Word.Layout.NoHanging; SpaceShrink = 0.0; FirstPageExtraRoom = 0.0 }
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

// A section that starts on an odd page (w:type oddPage in its own properties) after a page with an
// odd number: a blank page goes before it, the section's own - numbered, and with the header the
// section's rules give a page of its number; the section's first page (w:titlePg) is the page its
// text starts on. A section that starts on an even page after an odd one needs none.
do
    let layoutOfDoc (document: Doc.Document) =
        let s = settingsOf document
        let found, foundPages = Layout.layoutDocumentPages s measureText measureRange document
        let pageOfBlock id = Layout.pageIndexForYInPages foundPages (found |> List.find (fun layout -> layout.BlockId = id)).Y
        s, found, foundPages, pageOfBlock
    let titled = { geometry with DifferentFirstPage = true }
    let endsWith id kind (setup: Doc.SectionPageSetup) start : Doc.Block =
        Doc.Paragraph { para id "" with Metadata = { Doc.defaultMetadata with SectionBreak = Some kind; SectionPageSetup = Some setup; SectionPageNumberStart = start } }
    let document secondStarts lastStarts secondNumber : Doc.Document =
        { three with
            Blocks = [ Doc.Paragraph (para "a" "section one"); endsWith "e1" "nextPage" geometry None; Doc.Paragraph (para "b" "section two"); endsWith "e2" secondStarts titled secondNumber; Doc.Paragraph (para "c" "section three") ]
            PageSetup = { three.PageSetup with SectionStart = lastStarts } }
    let _, _, oddPages, oddPage = layoutOfDoc (document "oddPage" (Some "evenPage") None)
    check "an odd-page section after page 1: a blank page 2, the section on page 3, the next on page 4" (oddPages.Length, oddPage "a", oddPage "b", oddPage "c") (4, 0, 2, 3)
    let _, _, evenPages, evenPage = layoutOfDoc (document "evenPage" (Some "oddPage") None)
    check "an even-page section after page 1, an odd-page one after page 2: no blank page" (evenPages.Length, evenPage "a", evenPage "b", evenPage "c") (3, 0, 1, 2)
    let _, _, bothPages, bothPage = layoutOfDoc (document "oddPage" (Some "oddPage") None)
    check "two odd-page sections: pages 1, 3 and 5" (bothPages.Length, bothPage "a", bothPage "b", bothPage "c") (5, 0, 2, 4)
    let _, _, restartPages, restartPage = layoutOfDoc (document "evenPage" None (Some 1))
    check "an even-page section that starts at number 1: the blank page is number 1" (restartPages.Length, restartPage "b", restartPage "c") (4, 2, 3)
    // The headers of the pages and their numbers.
    let oddDoc = document "oddPage" (Some "evenPage") None
    let oddSettings, oddLayouts, oddFrames, _ = layoutOfDoc oddDoc
    let oddEnds = oddLayouts |> List.filter (fun layout -> Doc.endsSection layout.Metadata)
    let target pageIndex = (Paint.headerFooterTargetForPage oddSettings oddDoc oddEnds oddFrames pageIndex true).Value
    check "the blank page is the second section's, not its first page; its text starts on its first page" ([ for page in 0 .. 3 -> (target page).TargetKind, (target page).TargetSection, (target page).TargetFields.FieldPage ]) [ ("default", 0, 1); ("default", 1, 2); ("first", 1, 3); ("default", 2, 4) ]
    check "the section's pages are the blank one and its own" ((target 2).TargetFields.FieldSectionPages) 2
    // Laid out again after an edit, with the layouts of the last pass kept: the same pages.
    let state = Layout.createIncrementalLayoutState ()
    let first, _ = Layout.layoutDocumentPagesIncremental state oddSettings measureText measureRange oddDoc
    let edited = { oddDoc with Blocks = oddDoc.Blocks |> List.map (function Doc.Paragraph p when p.Id = "c" -> Doc.Paragraph { p with Buffer = ofString "section three, edited" } | block -> block) }
    let again, againPages = Layout.layoutDocumentPagesIncremental state oddSettings measureText measureRange edited
    let whole, wholePages = Layout.layoutDocumentPages oddSettings measureText measureRange edited
    check "an edit after the blank page: the incremental layout is the whole one" (againPages = wholePages, again |> List.map (fun layout -> layout.BlockId, layout.Y), first.Length) (true, whole |> List.map (fun layout -> layout.BlockId, layout.Y), again.Length)

// w:vAlign: the text of a page that is not full stands in the middle between the margins, at the
// bottom, or spread from the top to the bottom (not on the section's last page).
do
    let body = { Doc.defaultPageSetup with PageWidth = 816.0; PageHeight = 1056.0; MarginLeft = 96.0; MarginRight = 96.0; MarginTop = 96.0; MarginBottom = 96.0 }
    let titlePage vAlign : Doc.Document =
        { Doc.empty with
            Blocks = [ Doc.Paragraph (para "t1" "Title"); Doc.Paragraph (para "t2" "Subtitle")
                       Doc.Paragraph { para "te" "" with Metadata = { Doc.defaultMetadata with SectionBreak = Some "nextPage"; SectionPageSetup = Some geometry; SectionVerticalAlign = vAlign } }
                       Doc.Paragraph (para "b1" "Body") ]
            PageSetup = body }
    let placed vAlign =
        let document = titlePage vAlign
        let found, foundPages = Layout.layoutDocumentPages (settingsOf document) measureText measureRange document
        let of' id = found |> List.find (fun layout -> layout.BlockId = id)
        (of' "t1").Y, (of' "te").Y + (of' "te").Height, (of' "b1").Y, foundPages
    let topFirst, topEnd, topBody, topPages = placed None
    let contentBottom = 1056.0 - 96.0
    let slack = contentBottom - topEnd
    check "top: the text starts at the top margin" (topFirst, slack > 500.0) (96.0, true)
    let centerFirst, centerEnd, centerBody, centerPages = placed (Some "center")
    check "center: the text is moved down by half of what is free" (Math.Round(centerFirst - topFirst, 2), Math.Round(contentBottom - centerEnd, 2)) (Math.Round(slack / 2.0, 2), Math.Round(slack / 2.0, 2))
    let bottomFirst, bottomEnd, bottomBody, _ = placed (Some "bottom")
    check "bottom: the last line ends at the bottom margin" (Math.Round(bottomEnd, 2), Math.Round(bottomFirst - topFirst, 2)) (Math.Round(contentBottom, 2), Math.Round(slack, 2))
    check "the next section is where it was, the pages too" (centerBody, bottomBody, centerPages) (topBody, topBody, topPages)
    // The body's own (the last section's): a document of one section.
    let single vAlign : Doc.Document = { Doc.empty with Blocks = [ Doc.Paragraph (para "s1" "One"); Doc.Paragraph (para "s2" "Two") ]; PageSetup = { body with VerticalAlign = vAlign } }
    let singleY vAlign =
        let document = single vAlign
        let found, _ = Layout.layoutDocumentPages (settingsOf document) measureText measureRange document
        (found |> List.find (fun layout -> layout.BlockId = "s1")).Y, (let last = found |> List.find (fun layout -> layout.BlockId = "s2") in last.Y + last.Height)
    let (plainTop, plainEnd), (centredTop, centredEnd) = singleY None, singleY (Some "center")
    check "a document of one section, centred" (Math.Round(centredTop - plainTop, 2), Math.Round(contentBottom - centredEnd, 2)) (Math.Round((contentBottom - plainEnd) / 2.0, 2), Math.Round((contentBottom - plainEnd) / 2.0, 2))
    // Spread ("both"): a section of two pages - on the first the blocks are spread, the last stays at the top.
    let many : Doc.Document =
        { Doc.empty with
            Blocks = [ for index in 1 .. 3 -> Doc.Paragraph (para (sprintf "j%d" index) (sprintf "Paragraph %d" index)) ]
                     @ [ Doc.Paragraph { para "jb" "after a page break" with Metadata = { Doc.defaultMetadata with PageBreakBefore = true } } ]
            PageSetup = { body with VerticalAlign = Some "both" } }
    let spread, spreadPages = Layout.layoutDocumentPages (settingsOf many) measureText measureRange many
    let spreadOf id = spread |> List.find (fun layout -> layout.BlockId = id)
    check "both: the first block at the top, the last at the bottom, the one between half way; the last page from the top"
        (Math.Round((spreadOf "j1").Y, 2), Math.Round((spreadOf "j3").Y + (spreadOf "j3").Height, 2), Math.Round((spreadOf "j2").Y - (spreadOf "j1").Y, 2) = Math.Round((spreadOf "j3").Y - (spreadOf "j2").Y, 2), Math.Round((spreadOf "jb").Y - spreadPages.[1].Top, 2), spreadPages.Length)
        (96.0, Math.Round(contentBottom, 2), true, 96.0, 2)

// A page break in a section with columns starts a new page; a column break goes on in the next column.
do
    let body = { Doc.defaultPageSetup with PageWidth = 816.0; PageHeight = 1056.0; MarginLeft = 96.0; MarginRight = 96.0; MarginTop = 96.0; MarginBottom = 96.0 }
    let twoColumns : Doc.SectionColumns = { Count = 2; Gap = 48.0; Separator = false; Widths = [] }
    let withBreak (metadata: Doc.BlockMetadata) : Doc.Document =
        { Doc.empty with
            Blocks = [ Doc.Paragraph (para "h" "A heading over the columns")
                       Doc.Paragraph { para "he" "" with Metadata = { Doc.defaultMetadata with SectionBreak = Some "nextPage"; SectionPageSetup = Some geometry } }
                       Doc.Paragraph { para "c1" "first column" with Metadata = metadata }
                       Doc.Paragraph (para "c2" "after the break") ]
            PageSetup = { body with Columns = Some twoColumns; SectionStart = Some "continuous" } }
    let place metadata =
        let document = withBreak metadata
        let found, foundPages = Layout.layoutDocumentPages (settingsOf document) measureText measureRange document
        let of' id = found |> List.find (fun layout -> layout.BlockId = id)
        Layout.pageIndexForYInPages foundPages (of' "c1").Y, Layout.pageIndexForYInPages foundPages (of' "c2").Y, (of' "c2").Y - (of' "c1").Y, foundPages.Length
    let firstPage, secondPage, _, pageCount = place { Doc.defaultMetadata with PageBreakAfter = true }
    check "a page break in two columns: the text after it is on the next page" (firstPage, secondPage, pageCount) (0, 1, 2)
    let firstPage, secondPage, down, pageCount = place { Doc.defaultMetadata with PageBreakAfter = true; ColumnBreakAfter = true }
    // (Below the mark the break's paragraph leaves at the column's top: one empty line in the mark's font and its
    // space after, no more than "first column"'s own - Word 2010 and 2019, fixtures float-wrap cbr-text.)
    let firstLine = (let found, _ = Layout.layoutDocumentPages (settingsOf (withBreak { Doc.defaultMetadata with PageBreakAfter = true; ColumnBreakAfter = true })) measureText measureRange (withBreak { Doc.defaultMetadata with PageBreakAfter = true; ColumnBreakAfter = true }) in found |> List.find (fun layout -> layout.BlockId = "c1"))
    check "a column break: the text after it is in the next column of the same page, a line below its top" (firstPage, secondPage, down > firstLine.Lines.Head.Height / 2.0 && down <= firstLine.Height + 0.01, pageCount) (0, 0, true, 1)
    let document = withBreak Doc.defaultMetadata
    let before = { document with Blocks = document.Blocks |> List.map (function Doc.Paragraph p when p.Id = "c2" -> Doc.Paragraph { p with Metadata = { p.Metadata with PageBreakBefore = true } } | block -> block) }
    let found, foundPages = Layout.layoutDocumentPages (settingsOf before) measureText measureRange before
    check "a page break before a paragraph in two columns: on the next page" (Layout.pageIndexForYInPages foundPages (found |> List.find (fun layout -> layout.BlockId = "c2")).Y, foundPages.Length) (1, 2)

// A section break at the caret says how the section that starts there starts: its type goes to the
// properties of the section after the caret's paragraph, and the paragraph ends a section that has
// the properties (and the start) of the section it was cut from.
do
    let body = { Doc.defaultPageSetup with PageWidth = 816.0; PageHeight = 1056.0; MarginLeft = 96.0; MarginRight = 96.0; MarginTop = 96.0; MarginBottom = 96.0; PageNumberStart = Some 5; VerticalAlign = Some "center" }
    let plain : Doc.Document =
        { Doc.empty with
            Blocks = [ Doc.Paragraph (para "p1" "one"); Doc.Paragraph (para "p2" "two"); Doc.Paragraph (para "p3" "three") ]
            PageSetup = body
            Selection = Some { Anchor = { BlockId = "p1"; Offset = 0 }; Focus = { BlockId = "p1"; Offset = 0 } } }
    let metadataOf (document: Doc.Document) id = document.Blocks |> List.pick (function Doc.Paragraph p when p.Id = id -> Some p.Metadata | _ -> None)
    let continuous = Doc.setCurrentSectionBreak (Some "continuous") plain
    check "a continuous break after the first paragraph: the last section is continuous, the first starts as the document did"
        ((metadataOf continuous "p1").SectionBreak, continuous.PageSetup.SectionStart, (metadataOf continuous "p1").SectionPageSetup |> Option.map (fun setup -> setup.PageWidth), (metadataOf continuous "p1").SectionVerticalAlign)
        (Some "nextPage", Some "continuous", Some 816.0, Some "center")
    check "the page number the section started with is the first part's" ((metadataOf continuous "p1").SectionPageNumberStart, continuous.PageSetup.PageNumberStart) (Some 5, None)
    let layouts, pages = Layout.layoutDocumentPages (settingsOf { continuous with PageSetup = { continuous.PageSetup with VerticalAlign = None } }) measureText measureRange { continuous with PageSetup = { continuous.PageSetup with VerticalAlign = None }; Blocks = continuous.Blocks |> List.map (function Doc.Paragraph p -> Doc.Paragraph { p with Metadata = { p.Metadata with SectionVerticalAlign = None } } | block -> block) }
    check "the text after a continuous break is on the same page" pages.Length 1
    // A second break further down: between the two, the section the first break started.
    let second = Doc.setCurrentSectionBreak (Some "oddPage") { continuous with Selection = Some { Anchor = { BlockId = "p2"; Offset = 0 }; Focus = { BlockId = "p2"; Offset = 0 } } }
    check "a break in the continuous section: that part stays continuous, what follows starts on an odd page"
        ((metadataOf second "p1").SectionBreak, (metadataOf second "p2").SectionBreak, second.PageSetup.SectionStart) (Some "nextPage", Some "continuous", Some "oddPage")
    // The type of an existing break changed: at the paragraph that ends a section.
    let changed = Doc.setCurrentSectionBreak (Some "nextPage") { second with Selection = Some { Anchor = { BlockId = "p1"; Offset = 0 }; Focus = { BlockId = "p1"; Offset = 0 } } }
    check "on a paragraph that ends a section: the type of the section after it" ((metadataOf changed "p1").SectionBreak, (metadataOf changed "p2").SectionBreak, changed.PageSetup.SectionStart) (Some "nextPage", Some "nextPage", Some "oddPage")
    // Taken away: one section, which starts as the first part did.
    let removed = Doc.setCurrentSectionBreak None { second with Selection = Some { Anchor = { BlockId = "p2"; Offset = 0 }; Focus = { BlockId = "p2"; Offset = 0 } } }
    check "a break taken away: the section after it starts as the one before it did" ((metadataOf removed "p2").SectionBreak, (metadataOf removed "p2").SectionPageSetup, removed.PageSetup.SectionStart) (None, None, Some "continuous")

printfn "word sections: %d passed, %d failed" pass fail
if fail > 0 then exit 1
