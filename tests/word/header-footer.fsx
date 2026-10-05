// Which header and footer a page shows, section by section: a section's own
// (the references on the paragraph that ends it), the one of the section
// before it when it states none ("link to previous"), the first-page header
// on the first page of a section with a title page, the even one on even
// pages when the document tells them apart - and none where no section up to
// the page's refers to one. The body of a section starts below its header
// when that reaches below the top margin. (Paint.headerFooterTargetForPage is
// what the editor enters a header by; the paint commands are held to the same.)
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"

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

let settings : Layout.LayoutSettings =
    { Width = 816.0; PageHeight = 1056.0; FontSize = 16.0; LineHeight = 24.0; PaddingLeft = 96.0; PaddingTop = 96.0
      MarginRight = 96.0; MarginBottom = 96.0; InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 3.5 }
let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let words prefix count = String.concat " " [ for j in 1 .. count -> sprintf "%s%d" prefix j ]
let para id text : Doc.ParagraphBlock = { Id = id; Buffer = ofString text; Metadata = Doc.defaultMetadata }
let geometry titlePage : Doc.SectionPageSetup =
    { PageWidth = 816.0; PageHeight = 1056.0; MarginLeft = 96.0; MarginRight = 96.0; MarginTop = 96.0; MarginBottom = 96.0
      HeaderDistance = 48.0; FooterDistance = 48.0; DifferentFirstPage = titlePage; EvenAndOddHeaders = false }
let sectionEnd id headerRefs footerRefs titlePage : Doc.Block =
    Doc.Paragraph { para id "end of section" with
                        Metadata = { Doc.defaultMetadata with SectionBreak = Some "nextPage"; SectionPageSetup = Some (geometry titlePage); HeaderReferences = headerRefs; FooterReferences = footerRefs } }
let part kind id (texts: string list) : Doc.HeaderFooterPart =
    { Kind = kind; RelationshipId = Some id; Blocks = texts |> List.mapi (fun i text -> Doc.Paragraph (para (sprintf "%s-p%d" id i) text)) }
let body name count = [ for i in 1 .. count -> Doc.Paragraph (para (sprintf "%s-%d" name i) (words name 80)) ]

// Section one: one page, header A. Section two: three pages, a title page with its own header (C), otherwise the
// header of section one; the only footer of the document is its own. Section three: header B, the footer of two.
// Section four (the last): a tall header of its own.
let pageSetup =
    { Doc.defaultPageSetup with
        PageWidth = 816.0; PageHeight = 1056.0; MarginLeft = 96.0; MarginRight = 96.0; MarginTop = 96.0; MarginBottom = 96.0
        HeaderDistance = 48.0; FooterDistance = 48.0; HeaderReferences = [ "default", "hTall" ] }
let doc : Doc.Document =
    { Doc.empty with
        Blocks =
            body "one" 3 @ [ sectionEnd "end1" [ "default", "hA" ] [] false ]
            @ body "two" 22 @ [ sectionEnd "end2" [ "first", "hC" ] [ "default", "fTwo" ] true ]
            @ body "three" 3 @ [ sectionEnd "end3" [ "default", "hB" ] [] false ]
            @ body "four" 3
        PageSetup = pageSetup
        Styles = []
        Selection = None
        HeaderParts = [ part "default" "hA" [ "header A" ]; part "default" "hB" [ "header B" ]; part "first" "hC" [ "title page header C" ]
                        part "default" "hTall" [ "tall 1"; "tall 2"; "tall 3"; "tall 4"; "tall 5" ] ]
        FooterParts = [ part "default" "fTwo" [ "footer of section two" ] ] }

let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
let sectionEnds = layouts |> List.filter (fun layout -> Doc.endsSection layout.Metadata)
let pageOf id = Layout.pageIndexForYInPages pages (layouts |> List.find (fun layout -> layout.BlockId = id)).Y
let target pageIndex isHeader = (Paint.headerFooterTargetForPage settings doc sectionEnds pages pageIndex isHeader).Value
let shown pageIndex isHeader =
    let t = target pageIndex isHeader
    (t.TargetPart |> Option.bind (fun p -> p.RelationshipId) |> Option.defaultValue "none"), t.TargetKind, t.TargetSection, t.TargetOwn

let twoFirst, twoLast, threePage, fourPage = pageOf "two-1", pageOf "end2", pageOf "three-1", pageOf "four-1"
check "the sections start on pages of their own" (pageOf "one-1", twoFirst, twoLast > twoFirst + 1, threePage = twoLast + 1, fourPage = threePage + 1) (0, 1, true, true, true)
check "section one shows its header" (shown 0 true) ("hA", "default", 0, true)
check "the title page of section two shows the first-page header" (shown twoFirst true) ("hC", "first", 1, true)
check "its other pages show the header of section one (it states no default of its own)" (shown (twoFirst + 1) true, shown twoLast true) (("hA", "default", 1, false), ("hA", "default", 1, false))
check "section three shows its own" (shown threePage true) ("hB", "default", 2, true)
check "the last section shows the one the page setup refers to" (shown fourPage true) ("hTall", "default", 3, true)
check "of four sections" ((target 0 true).TargetSectionCount, (target twoLast true).TargetSectionStartPage) (4, twoFirst)
check "no such page" (Paint.headerFooterTargetForPage settings doc sectionEnds pages pages.Length true, Paint.headerFooterTargetForPage settings doc sectionEnds pages -1 false) (None, None)
// The footer: section one has none (the document's only footer is section two's), sections three and four show two's.
check "section one has no footer" (shown 0 false) ("none", "default", 0, false)
check "section two has the footer (its title page shows it too: there is no first-page footer)" (shown (twoFirst + 1) false, shown twoFirst false) (("fTwo", "default", 1, true), ("fTwo", "first", 1, false))
check "the sections after it show it too" (shown threePage false, shown fourPage false) (("fTwo", "default", 2, false), ("fTwo", "default", 3, false))

// The paint commands: each page's header and footer text.
let _, commands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange doc layouts pages
let textsOn pageIndex (top: float) (bottom: float) =
    let page = List.item pageIndex pages
    commands
    |> List.choose (function
        | Paint.DrawText (text, _, _, baseline, _) when baseline >= page.Top + top && baseline <= page.Top + bottom -> Some text
        | _ -> None)
    |> String.concat " "
let headerOn pageIndex = textsOn pageIndex 0.0 ((List.item pageIndex pages).MarginTop - 2.0)
let footerOn pageIndex = textsOn pageIndex (1056.0 - 94.0) 1056.0
check "painted: the headers page by page" (headerOn 0, headerOn twoFirst, headerOn (twoFirst + 1), headerOn threePage) ("header A", "title page header C", "header A", "header B")
check "painted: the footers page by page" (footerOn 0, footerOn twoFirst, footerOn threePage, footerOn fourPage) ("", "footer of section two", "footer of section two", "footer of section two")

// The body of the last section starts below its five header paragraphs; the others at the top margin.
let firstLineOffset id = let layout = layouts |> List.find (fun layout -> layout.BlockId = id) in layout.Y - (List.item (pageOf id) pages).Top
check "the body starts at the top margin under a one-line header" (firstLineOffset "one-1" < 130.0, firstLineOffset "three-1" < 130.0) (true, true)
check "and below a header that reaches further down" (firstLineOffset "four-1" >= 48.0 + 150.0, (List.item fourPage pages).MarginTop >= 198.0, (List.item 0 pages).MarginTop) (true, true, 96.0)
// ... in an incremental layout too, and when the header gets shorter again.
let state = Layout.createIncrementalLayoutState ()
check "the same incrementally" (Layout.layoutDocumentPagesIncremental state settings measureText measureRange doc = (layouts, pages)) true
let shorter = { doc with HeaderParts = doc.HeaderParts |> List.map (fun p -> if p.RelationshipId = Some "hTall" then { p with Blocks = List.truncate 1 p.Blocks } else p) }
let shortLayouts, shortPages = Layout.layoutDocumentPagesIncremental state settings measureText measureRange shorter
check "a shorter header: the body moves up, as in a layout from nothing" ((shortLayouts, shortPages) = Layout.layoutDocumentPages settings measureText measureRange shorter, (List.item fourPage shortPages).MarginTop) (true, 96.0)

// Even and odd pages (a document setting): even pages show the even header where the section has one.
let evenDoc =
    { doc with
        PageSetup = { pageSetup with EvenAndOddHeaders = true; HeaderReferences = [ "default", "hA"; "even", "hB" ] }
        Blocks = body "even" 30 }
let evenLayouts, evenPages = Layout.layoutDocumentPages settings measureText measureRange evenDoc
let evenShown pageIndex = (Paint.headerFooterTargetForPage settings evenDoc [] evenPages pageIndex true).Value |> fun t -> (t.TargetPart |> Option.bind (fun p -> p.RelationshipId)), t.TargetKind
check "odd pages the default header, even pages the even one" (evenShown 0, evenShown 1, evenShown 2) ((Some "hA", "default"), (Some "hB", "even"), (Some "hA", "default"))

// A story's layouts are where the page paints the header.
let story = Paint.storyBlockLayouts settings measureText measureRange (Paint.headerFooterStartY settings doc.PageSetup (List.item fourPage pages) true) (doc.HeaderParts.[3].Blocks)
check "the layouts of a header's blocks, from the header distance down" (story.Length, abs (story.Head.Y - ((List.item fourPage pages).Top + 48.0)) < 0.01, story |> List.pairwise |> List.forall (fun (a, b) -> b.Y >= a.Y + a.Height - 0.01)) (5, true, true)

printfn "word header-footer: %d passed, %d failed" pass fail
if fail > 0 then exit 1
