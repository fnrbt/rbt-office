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
      MarginRight = 96.0; MarginBottom = 96.0; InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 3.5; DefaultTabStop = 48.0 }
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
// (ECMA-376 17.10.2: a title page without a first-page footer - its section's or an earlier section's - has a blank one, not the default footer.)
check "section two has the footer; its title page, without a first-page footer, shows none" (shown (twoFirst + 1) false, shown twoFirst false) (("fTwo", "default", 1, true), ("none", "first", 1, false))
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
check "painted: the footers page by page" (footerOn 0, footerOn twoFirst, footerOn (twoFirst + 1), footerOn threePage, footerOn fourPage) ("", "", "footer of section two", "footer of section two", "footer of section two")

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
let story = Paint.storyBlockLayouts settings measureText measureRange doc (Paint.headerFooterStartY settings doc.PageSetup (List.item fourPage pages) true) (doc.HeaderParts.[3].Blocks)
check "the layouts of a header's blocks, from the header distance down" (story.Length, abs (story.Head.Y - ((List.item fourPage pages).Top + 48.0)) < 0.01, story |> List.pairwise |> List.forall (fun (a, b) -> b.Y >= a.Y + a.Height - 0.01)) (5, true, true)

// ---------------------------------------------------------------------------
// A title page or an even page without a header of that kind shows a blank
// one - and the first-page header of an earlier section when there is one.
// ---------------------------------------------------------------------------
do
    let titleDoc =
        { doc with
            Blocks =
                body "a" 3 @ [ sectionEnd "endA" [ "default", "hA"; "first", "hC" ] [] true ]
                @ body "b" 3 @ [ sectionEnd "endB" [] [] true ]
                @ body "c" 60
            PageSetup = { pageSetup with DifferentFirstPage = false; EvenAndOddHeaders = true; HeaderReferences = [ "default", "hB" ] } }
    let titleLayouts, titlePages = Layout.layoutDocumentPages settings measureText measureRange titleDoc
    let ends = titleLayouts |> List.filter (fun layout -> Doc.endsSection layout.Metadata)
    let show pageIndex =
        let t = (Paint.headerFooterTargetForPage settings titleDoc ends titlePages pageIndex true).Value
        (t.TargetPart |> Option.bind (fun p -> p.RelationshipId) |> Option.defaultValue "none"), t.TargetKind
    let pageOfBlock id = Layout.pageIndexForYInPages titlePages (titleLayouts |> List.find (fun layout -> layout.BlockId = id)).Y
    let cFirst = pageOfBlock "c-1"
    check "a title page shows the first-page header of the section before it when it has none of its own" (show (pageOfBlock "a-1"), show (pageOfBlock "b-1")) (("hC", "first"), ("hC", "first"))
    // (Page indexes are 0-based: an even page has an odd index.)
    let evenPage = if (cFirst + 1) % 2 = 0 then cFirst else cFirst + 1
    check "an even page without an even header shows none; the odd page after it the default one" (show evenPage, show (evenPage + 1)) (("none", "even"), ("hB", "default"))
    let _, titleCommands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange titleDoc titleLayouts titlePages
    let headerText pageIndex =
        let page = List.item pageIndex titlePages
        titleCommands |> List.choose (function Paint.DrawText (text, _, _, baseline, _) when baseline >= page.Top && baseline <= page.Top + page.MarginTop - 2.0 -> Some text | _ -> None) |> String.concat " "
    check "painted: nothing on the even page, the default header on the odd one" (headerText evenPage, headerText (evenPage + 1)) ("", "header B")

// ---------------------------------------------------------------------------
// A header's text has the document's styles, as the body's has: the default
// paragraph style where it names none or one the document does not define
// ("Header" on a header made in the editor), a defined style where it names one.
// ---------------------------------------------------------------------------
do
    let named id isDefault size : Doc.NamedStyle =
        { Id = id; Name = Some id; StyleType = "paragraph"; IsDefault = isDefault; BasedOn = None; Next = None; Linked = None; IsQuickStyle = false; Aliases = None
          UiPriority = None; SemiHidden = false; UnhideWhenUsed = false; Text = Some { defaultStyle with FontSize = size }
          TextBold = None; TextItalic = None; TextUnderline = None; TextAllCaps = None; TextSmallCaps = None; TextStrike = None
          SpacingBefore = None; SpacingAfter = None; ContextualSpacing = None; LineSpacing = None; LineSpacingRule = None; Alignment = None
          IndentLeft = None; IndentRight = None; FirstLineIndent = None; BackgroundColor = None; OutlineLevel = None
          KeepNext = None; KeepLines = None; PageBreakBefore = None; WidowControl = None }
    let styled styleId id text = Doc.Paragraph { para id text with Metadata = { Doc.defaultMetadata with Style = styleId } }
    let styledDoc =
        { doc with
            Blocks = body "s" 3
            Styles = [ named "Normal" true 21.0; named "Big" false 32.0 ]
            PageSetup = { pageSetup with HeaderReferences = [ "default", "hS" ] }
            HeaderParts = [ { Kind = "default"; RelationshipId = Some "hS"; Blocks = [ styled None "hs-plain" "plain"; styled (Some "Header") "hs-missing" "missing"; styled (Some "Big") "hs-big" "big" ] } ]
            FooterParts = [] }
    let styledLayouts, styledPages = Layout.layoutDocumentPages settings measureText measureRange styledDoc
    let _, styledCommands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange styledDoc styledLayouts styledPages
    let sizeOf text = styledCommands |> List.pick (function Paint.DrawText (drawn, style, _, _, _) when drawn = text -> Some style.FontSize | _ -> None)
    check "painted in the size of the default paragraph style, of a style that is not defined, of a defined one; the body's" (sizeOf "plain", sizeOf "missing", sizeOf "big", styledCommands |> List.pick (function Paint.DrawText (drawn, style, _, _, _) when drawn.StartsWith "s1" -> Some style.FontSize | _ -> None)) (21.0, 21.0, 32.0, 21.0)
    let storyLayouts = Paint.storyBlockLayouts settings measureText measureRange styledDoc 48.0 styledDoc.HeaderParts.Head.Blocks
    let storySize (layout: Layout.BlockLayout) =
        layout.Lines |> List.collect (fun line -> line.Segments) |> List.pick (function Layout.TextSegment (_, style, _, _, _) -> Some style.FontSize | _ -> None)
    check "the caret's layouts are laid out with the same sizes" (storyLayouts |> List.map storySize) [ 21.0; 21.0; 32.0 ]
    check "and the selection format says the size that is painted" ((Rbt.Word.WordStyles.selectionFormat { styledDoc with Blocks = styledDoc.HeaderParts.Head.Blocks; Selection = Some { Anchor = { BlockId = "hs-missing"; Offset = 1 }; Focus = { BlockId = "hs-missing"; Offset = 1 } } }).FontSize) (Rbt.Word.WordStyles.Uniform 21.0)

// ---------------------------------------------------------------------------
// The running fields: PAGE counts from a section's start (w:pgNumType
// w:start) in the section's format (w:fmt) or the field's own (\* roman),
// NUMPAGES is the number of pages, SECTIONPAGES that of the page's section.
// ---------------------------------------------------------------------------
do
    let field code text : Run = Text { Start = 0; Length = 0; Style = { defaultStyle with FieldCode = Some code; FieldResult = Some text } }
    let fieldPara id (pieces: (string option * string) list) : Doc.Block =
        let mutable buffer = ofString ""
        for (code, text) in pieces do
            let at = length buffer
            buffer <- insert at text buffer
            match code with
            | Some code -> buffer <- applyStyle at text.Length (fun style -> { style with FieldCode = Some code; FieldResult = Some text }) buffer
            | None -> ()
        Doc.Paragraph { Id = id; Buffer = buffer; Metadata = Doc.defaultMetadata }
    let footer = fieldPara "ff" [ None, "Page "; Some "PAGE", "1"; None, " of "; Some "NUMPAGES", "1"; None, " / "; Some "SECTIONPAGES", "1"; None, " / "; Some " PAGE \\* ROMAN ", "1"; None, " s"; Some "SECTION", "1" ]
    let numbered id start format : Doc.Block =
        Doc.Paragraph { para id "end of section" with
                            Metadata = { Doc.defaultMetadata with SectionBreak = Some "nextPage"; SectionPageSetup = Some (geometry false); SectionPageNumberStart = start; SectionPageNumberFormat = format; FooterReferences = [ "default", "fF" ] } }
    let fieldDoc =
        { doc with
            Blocks = body "front" 12 @ [ numbered "endFront" (Some 1) (Some "lowerRoman") ] @ body "main" 22 @ [ numbered "endMain" (Some 1) None ] @ body "back" 3
            PageSetup = { pageSetup with HeaderReferences = []; FooterReferences = []; PageNumberStart = Some 5; PageNumberFormat = Some "upperLetter" }
            HeaderParts = []
            FooterParts = [ { Kind = "default"; RelationshipId = Some "fF"; Blocks = [ footer ] } ] }
    let fieldLayouts, fieldPages = Layout.layoutDocumentPages settings measureText measureRange fieldDoc
    let ends = fieldLayouts |> List.filter (fun layout -> Doc.endsSection layout.Metadata)
    let pageOfBlock id = Layout.pageIndexForYInPages fieldPages (fieldLayouts |> List.find (fun layout -> layout.BlockId = id)).Y
    let frontLast, mainFirst, mainLast, backFirst = pageOfBlock "endFront", pageOfBlock "main-1", pageOfBlock "endMain", pageOfBlock "back-1"
    let total = fieldPages.Length
    let _, fieldCommands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange fieldDoc fieldLayouts fieldPages
    let footerText pageIndex =
        let page = List.item pageIndex fieldPages
        fieldCommands |> List.choose (function Paint.DrawText (text, _, _, baseline, _) when baseline >= page.Top + page.Height - 94.0 && baseline <= page.Top + page.Height -> Some text | _ -> None) |> String.concat ""
    let frontPages, mainPages = frontLast + 1, mainLast - mainFirst + 1
    check "the sections have several pages" (frontPages >= 2, mainPages >= 3, backFirst = total - 1) (true, true, true)
    check "the front matter: roman numbers, the pages of the document and of the section" (footerText 0, footerText 1) (sprintf "Page i of %d / %d / I s1" total frontPages, sprintf "Page ii of %d / %d / II s1" total frontPages)
    check "the main part starts at 1 again, in decimal" (footerText mainFirst, footerText (mainFirst + 2)) (sprintf "Page 1 of %d / %d / I s2" total mainPages, sprintf "Page 3 of %d / %d / III s2" total mainPages)
    check "the last section: from the number and in the format of the page setup" (footerText backFirst) (sprintf "Page E of %d / 1 / V s3" total)
    let fields pageIndex = (Paint.headerFooterTargetForPage settings fieldDoc ends fieldPages pageIndex false).Value.TargetFields
    check "the values an editor enters a footer with" ((fields mainFirst).FieldPage, (fields mainFirst).FieldPageFormat, (fields mainFirst).FieldPageCount, (fields mainFirst).FieldSectionPages, (fields 1).FieldPageFormat) (1, "decimal", total, mainPages, "lowerRoman")
    check "page numbers in the formats of w:pgNumType" ([ "decimal"; "upperRoman"; "lowerRoman"; "upperLetter"; "lowerLetter"; "numberInDash"; "unknown" ] |> List.map (fun format -> Paint.formatPageNumber format 28)) [ "28"; "XXVIII"; "xxviii"; "BB"; "bb"; "- 28 -"; "28" ]
    // With one more page (pages of notes after the body count: the editor hands the paint the pages it shows).
    let longer = Layout.pageInfosCoveringBottom ((List.last fieldPages).Top + (List.last fieldPages).Height + 200.0) fieldPages
    let _, longerCommands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange fieldDoc fieldLayouts longer
    let longerFooter pageIndex =
        let page = List.item pageIndex longer
        longerCommands |> List.choose (function Paint.DrawText (text, _, _, baseline, _) when baseline >= page.Top + page.Height - 94.0 && baseline <= page.Top + page.Height -> Some text | _ -> None) |> String.concat ""
    check "a page of notes after the body is a page: counted, numbered on, with the footer" (longer.Length, longerFooter 0, longerFooter (longer.Length - 1)) (total + 1, sprintf "Page i of %d / %d / I s1" (total + 1) frontPages, sprintf "Page F of %d / 2 / VI s3" (total + 1))

// ---------------------------------------------------------------------------
// A PAGE field in the text of the body shows the page's number as its section
// writes page numbers (it was always decimal: "see page 2" beside a footer
// that said "ii"), unless the field states a format of its own.
// ---------------------------------------------------------------------------
do
    let fieldPara id (code: string) : Doc.Block =
        let buffer = ofString "see page 0 here" |> applyStyle 9 1 (fun style -> { style with FieldCode = Some code; FieldResult = Some "0" })
        Doc.Paragraph { Id = id; Buffer = buffer; Metadata = Doc.defaultMetadata }
    let ending id format : Doc.Block =
        Doc.Paragraph { para id "end of section" with Metadata = { Doc.defaultMetadata with SectionBreak = Some "nextPage"; SectionPageSetup = Some (geometry false); SectionPageNumberFormat = format } }
    let fieldDoc =
        { doc with
            Blocks = [ fieldPara "f-front" " PAGE "; fieldPara "f-own" " PAGE \\* ROMAN "; ending "f-end1" (Some "lowerRoman") ] @ [ fieldPara "f-main" "PAGE"; ending "f-end2" None ] @ [ fieldPara "f-back" " PAGE " ]
            PageSetup = { pageSetup with HeaderReferences = []; FooterReferences = []; PageNumberFormat = Some "upperLetter" }
            HeaderParts = []
            FooterParts = [] }
    let anchor id page : Doc.PageFieldAnchor = { PageFieldBlockId = id; PageFieldStartOffset = 9; PageFieldCode = "PAGE"; PageNumber = page }
    let pageMap : Doc.PageMap = { PageCount = 5; BlockAnchors = []; Fields = [ anchor "f-front" 2; anchor "f-own" 2; anchor "f-main" 3; anchor "f-back" 5 ]; Bookmarks = [] }
    let refreshed = Doc.refreshFieldsWithPageMap pageMap fieldDoc
    let textOf id = refreshed.Blocks |> List.pick (function Doc.Paragraph p when p.Id = id -> Some (toString p.Buffer) | _ -> None)
    check "the section's format: roman in the front matter" (textOf "f-front") "see page ii here"
    check "the field's own format wins" (textOf "f-own") "see page II here"
    check "a section that states none: decimal" (textOf "f-main") "see page 3 here"
    check "the last section: the format of the page setup" (textOf "f-back") "see page E here"
    let formats = Doc.sectionNumberFormats fieldDoc
    check "the format of each paragraph's section" ([ "f-front"; "f-end1"; "f-main"; "f-end2"; "f-back" ] |> List.map (fun id -> formats.[id])) [ "lowerRoman"; "lowerRoman"; "decimal"; "decimal"; "upperLetter" ]

// ---------------------------------------------------------------------------
// The copy of a header that "link to previous" off gives a section: every
// block has an id of its own - the paragraphs, and the tables too (they kept
// theirs: two tables with one id in the document).
// ---------------------------------------------------------------------------
do
    let position : Doc.Position = { BlockId = "c-p"; Offset = 0 }
    let made = Doc.insertTable 1 2 { Doc.empty with Blocks = [ Doc.Paragraph (para "c-p" "text") ]; Selection = Some { Anchor = position; Focus = position } }
    let table = made.Blocks |> List.pick (function Doc.Table table -> Some table | _ -> None)
    let withText (texts: string list) (table: Doc.TableBlock) =
        { table with Rows = table.Rows |> List.map (fun row -> { row with Cells = row.Cells |> List.mapi (fun i c -> { c with Content = [ para (sprintf "%s-p%d" table.Id i) texts.[i] ] }) }) }
    let inner = { table with Id = "t-inner"; Rows = table.Rows |> List.map (fun row -> { row with Cells = List.truncate 1 row.Cells }) } |> withText [ "inner" ]
    let outer =
        { table with Id = "t-outer" } |> withText [ "one"; "two" ]
        |> fun t -> { t with Rows = t.Rows |> List.map (fun row -> { row with Cells = row.Cells |> List.mapi (fun i c -> if i = 0 then { c with NestedTables = [ inner ]; NestedTableAnchors = [ 1 ] } else c) }) }
    let source = [ Doc.Paragraph (para "c-p" "text"); Doc.Table outer ]
    let mutable serial = 0
    let copy = source |> Doc.copyBlocksWithNewIds (fun kind -> serial <- serial + 1; sprintf "new-%s-%d" kind serial)
    let rec tableIds (table: Doc.TableBlock) = table.Id :: (table.Rows |> List.collect (fun row -> row.Cells |> List.collect (fun c -> c.NestedTables |> List.collect tableIds)))
    let idsOf blocks = (Doc.paragraphsOfBlocks blocks |> List.map (fun p -> p.Id)) @ (blocks |> List.collect (function Doc.Table table -> tableIds table | _ -> []))
    let before, after = idsOf source, idsOf copy
    check "the copy has as many blocks, none with an id of the original, each id once" (after.Length, after |> List.filter (fun id -> List.contains id before), (List.distinct after).Length) (before.Length, [], before.Length)
    check "tables and paragraphs alike" (after |> List.filter (fun id -> id.StartsWith "new-table") |> List.length, after |> List.filter (fun id -> id.StartsWith "new-paragraph") |> List.length) (2, 4)
    check "with the same text" (Doc.paragraphsOfBlocks copy |> List.map (fun p -> toString p.Buffer)) [ "text"; "one"; "inner"; "two" ]

// ---------------------------------------------------------------------------
// A section's text stands at its own left margin: on screen a block is at the
// document's left margin (the final section's) plus its layout's x, so the
// layouts of a section with another left margin are moved by the difference -
// lines, list markers, the cells of a table and its borders, footnotes.
// ---------------------------------------------------------------------------
do
    let wide : Doc.SectionPageSetup = { geometry false with MarginLeft = 192.0; MarginRight = 48.0 }
    let marginDoc : Doc.Document =
        { Doc.empty with
            Blocks =
                [ Doc.Paragraph (para "wide-1" (words "wide" 60))
                  Doc.Paragraph { para "wide-end" "end" with Metadata = { Doc.defaultMetadata with SectionBreak = Some "nextPage"; SectionPageSetup = Some wide } }
                  Doc.Paragraph (para "last-1" (words "last" 60)) ]
            PageSetup = pageSetup
            Styles = []
            Selection = None }
    let marginLayouts, marginPages = Layout.layoutDocumentPages settings measureText measureRange marginDoc
    let lineExtent id =
        let layout = marginLayouts |> List.find (fun layout -> layout.BlockId = id)
        let xs = layout.Lines |> List.collect (fun line -> line.Segments |> List.map (fun segment -> Layout.segmentX segment, Layout.segmentX segment + Layout.segmentWidth segment))
        xs |> List.map fst |> List.min, xs |> List.map snd |> List.max
    let wideLeft, wideRight = lineExtent "wide-1"
    let lastLeft, lastRight = lineExtent "last-1"
    check "a section with a wider left margin: its lines start that much further right" (wideLeft, lastLeft) (96.0, 0.0)
    check "and end at its own right margin (the page is 816 wide, the margins 192 and 48)" (wideRight <= 96.0 + (816.0 - 192.0 - 48.0) + 8.0, wideRight > 96.0 + 400.0, lastRight <= 624.0 + 8.0) (true, true, true)
    check "the pages say whose margins they have" ((List.item 0 marginPages).MarginLeft, (List.item 1 marginPages).MarginLeft) (192.0, 96.0)
    let _, marginCommands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange marginDoc marginLayouts marginPages
    let textLeft (page: Layout.PageInfo) =
        marginCommands |> List.choose (function Paint.DrawText (_, _, x, baseline, _) when baseline > page.Top && baseline < page.Top + page.Height -> Some x | _ -> None) |> List.min
    check "painted at the section's own margin" (textLeft (List.item 0 marginPages), textLeft (List.item 1 marginPages)) (192.0, 96.0)
    let state = Layout.createIncrementalLayoutState ()
    check "the same incrementally, also after an edit in the section" 
        (Layout.layoutDocumentPagesIncremental state settings measureText measureRange marginDoc = (marginLayouts, marginPages),
         (let edited = marginDoc |> Doc.setCursor "wide-1" 3 |> Doc.insertText "typed "
          Layout.layoutDocumentPagesIncremental state settings measureText measureRange edited = Layout.layoutDocumentPages settings measureText measureRange edited))
        (true, true)

printfn "word header-footer: %d passed, %d failed" pass fail
if fail > 0 then exit 1
