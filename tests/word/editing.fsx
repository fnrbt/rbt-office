#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"

open Rbt.Word.TextBuffer
module Doc = Rbt.Word.Document
module TB = Rbt.Word.TextBuffer

// Editing commands of the model as the keys and the clipboard use them: copy
// and paste with formatting (Doc.copySelection / Doc.pasteClip).

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-30s expected %A got %A" label expected got

let paragraph id (text: string) : Doc.ParagraphBlock = { Id = id; Buffer = TB.empty |> TB.insert 0 text; Metadata = Doc.defaultMetadata }
let textOf (para: Doc.ParagraphBlock) = TB.toString para.Buffer
let docOf (blocks: Doc.Block list) = { Doc.empty with Blocks = blocks; Selection = None }
let caret id offset (doc: Doc.Document) = Doc.setCursor id offset doc
let range (fromId, fromOffset) (toId, toOffset) (doc: Doc.Document) = Doc.setSelection { BlockId = fromId; Offset = fromOffset } { BlockId = toId; Offset = toOffset } doc
let texts (doc: Doc.Document) = Doc.allParagraphs doc |> List.map textOf
let styleAt (doc: Doc.Document) id offset = (Doc.findParagraph doc id).Value.Buffer |> TB.getStyleAt offset
let boldOf (doc: Doc.Document) id = (Doc.findParagraph doc id).Value.Buffer |> fun b -> [ for i in 0 .. TB.length b - 1 -> (TB.getStyleAt i b).Bold ]
let selectionOf (doc: Doc.Document) = doc.Selection |> Option.map (fun s -> s.Anchor.BlockId, s.Anchor.Offset, s.Focus.BlockId, s.Focus.Offset)
let copy (doc: Doc.Document) = (Doc.copySelection doc).Value

// ---- a part of one paragraph: its text with its formatting, into the text at the caret ----
let source =
    let para = paragraph "s" "Hello bold world"
    docOf [ Doc.Paragraph { para with Buffer = TB.toggleBold 6 4 para.Buffer } ]
let clip = source |> range ("s", 4) ("s", 12) |> copy
check "copy: the plain text" clip.Text "o bold w"
check "copy: one part" clip.Paragraphs.Length 1
let target = docOf [ Doc.Paragraph (paragraph "t" "XY") ] |> caret "t" 1
let pasted = Doc.pasteClip clip target
check "paste: the text joins the paragraph" (texts pasted) [ "Xo bold wY" ]
check "paste: the bold text is bold, the rest is not" (boldOf pasted "t") [ false; false; false; true; true; true; true; false; false; false ]
check "paste: the caret after the pasted text" (selectionOf pasted) (Some ("t", 9, "t", 9))
check "paste: over a selection, it is replaced" (Doc.pasteClip clip (target |> range ("t", 0) ("t", 2)) |> texts) [ "o bold w" ]
check "nothing selected: nothing to copy" (Doc.copySelection (caret "s" 3 source)) None

// ---- several paragraphs: each with its properties, the last joined to the text after the caret ----
let centred = { Doc.defaultMetadata with Alignment = Doc.Center; ExplicitAlignment = true; Style = Some "Title" }
let righted = { Doc.defaultMetadata with Alignment = Doc.Right; ExplicitAlignment = true }
let listed = { Doc.defaultMetadata with ListStyle = Doc.Bullet; ListLevel = 1 }
let threeParagraphs =
    docOf [ Doc.Heading (1, { paragraph "a" "Title" with Metadata = centred })
            Doc.Paragraph { paragraph "b" "Body text" with Metadata = listed }
            Doc.Paragraph (paragraph "c" "Last one") ]
let multi = threeParagraphs |> range ("a", 2) ("c", 4) |> copy
check "copy: the plain text of three paragraphs" multi.Text "tle\nBody text\nLast"
check "copy: a part of each" (multi.Paragraphs |> List.map (fun p -> TB.toString p.Buffer)) [ "tle"; "Body text"; "Last" ]
let into =
    docOf [ Doc.Paragraph (paragraph "before" "Before")
            Doc.Paragraph { paragraph "d" "abcdef" with Metadata = righted }
            Doc.Paragraph (paragraph "after" "After") ]
    |> caret "d" 3
let multiPasted = Doc.pasteClip multi into
check "paste: three paragraphs" (texts multiPasted) [ "Before"; "abctle"; "Body text"; "Lastdef"; "After" ]
let paras = Doc.allParagraphs multiPasted
check "paste: the first has the properties of the copied paragraph (its mark)" (paras.[1].Id, paras.[1].Metadata.Alignment, paras.[1].Metadata.Style) ("d", Doc.Center, Some "Title")
check "paste: ... and is a heading as it was" (multiPasted.Blocks.[1] |> function Doc.Heading (level, _) -> Some level | _ -> None) (Some 1)
check "paste: the second is in its list" (paras.[2].Metadata.ListStyle, paras.[2].Metadata.ListLevel) (Doc.Bullet, 1)
check "paste: the last keeps the properties of the paragraph pasted into" (paras.[3].Metadata.Alignment, paras.[3].Metadata.Style, (multiPasted.Blocks.[3] |> function Doc.Paragraph _ -> true | _ -> false)) (Doc.Right, None, true)
check "paste: the caret after the pasted text" (selectionOf multiPasted) (Some (paras.[3].Id, 4, paras.[3].Id, 4))
check "paste: new paragraphs have ids of their own" (paras |> List.map (fun p -> p.Id) |> List.distinct |> List.length) 5

// A selection that ends at the start of a paragraph took the mark of the one before it.
let wholeParagraph = threeParagraphs |> range ("b", 0) ("c", 0) |> copy
let wholePasted = docOf [ Doc.Paragraph (paragraph "e" "xy") ] |> caret "e" 0 |> Doc.pasteClip wholeParagraph
check "paste: a whole paragraph is a paragraph of its own" (texts wholePasted) [ "Body text"; "xy" ]
check "paste: ... in its list; the paragraph pasted into is not" (Doc.allParagraphs wholePasted |> List.map (fun p -> p.Metadata.ListStyle)) [ Doc.Bullet; Doc.NoList ]

// ---- what the copy does not take: annotations of the source ----
let annotated =
    let para = paragraph "n" "link here [Comment 7] tail"
    let buffer =
        para.Buffer
        |> TB.setHyperlink 0 4 (Some "https://example.com")
        |> TB.setBookmark 5 4 "1" "mark"
        |> TB.setComment 10 11 (Some "7") (Some "note") None None
        |> TB.applyStyle 22 4 (fun style -> { style with SourceRun = Some 42 })
    docOf [ Doc.Paragraph { para with Buffer = buffer } ]
let annotatedClip = annotated |> range ("n", 0) ("n", 26) |> copy
let annotatedPasted = docOf [ Doc.Paragraph (paragraph "p" "") ] |> caret "p" 0 |> Doc.pasteClip annotatedClip
check "paste: the mark of a comment is not pasted" (texts annotatedPasted) [ "link here  tail" ]
check "paste: a link stays a link" ((styleAt annotatedPasted "p" 1).Hyperlink) (Some "https://example.com")
check "paste: a bookmark is not copied" ((styleAt annotatedPasted "p" 6).BookmarkStartId, (styleAt annotatedPasted "p" 6).BookmarkName) (None, None)
check "paste: pasted text is of no run of the file" ((styleAt annotatedPasted "p" 12).SourceRun) None

// ---- in a table cell: paragraphs split there as Enter splits them ----
let cellDoc =
    let withTable = docOf [ Doc.Paragraph (paragraph "p1" "Text") ] |> caret "p1" 4 |> Doc.insertTable 1 1
    let table = withTable.Blocks |> List.pick (function Doc.Table t -> Some t | _ -> None)
    let cell = { table.Rows.Head.Cells.Head with Content = [ paragraph "cell" "[]" ] }
    docOf [ Doc.Table { table with Rows = [ { table.Rows.Head with Cells = [ cell ] } ] } ]
let cellPasted = cellDoc |> caret "cell" 1 |> Doc.pasteClip multi
check "paste: in a table cell" (texts cellPasted) [ "[tle"; "Body text"; "Last]" ]

// ---- a section's end stays with the paragraph mark that ends it ----
let sectioned =
    docOf [ Doc.Paragraph { paragraph "s1" "end of section" with Metadata = { Doc.defaultMetadata with SectionBreak = Some "nextPage" } }
            Doc.Paragraph (paragraph "s2" "next section") ]
let sectionPasted = sectioned |> caret "s1" 3 |> Doc.pasteClip multi
check "paste: the section still ends once, after the text that was after the caret"
    (Doc.allParagraphs sectionPasted |> List.map (fun p -> textOf p, p.Metadata.SectionBreak))
    [ "endtle", None; "Body text", None; "Last of section", Some "nextPage"; "next section", None ]

// ---- Enter in the last paragraph of a section: one section end, after the second part ----
let entered = sectioned |> caret "s1" 3 |> Doc.splitBlock
check "Enter: the section still ends once, with the second part"
    (Doc.allParagraphs entered |> List.map (fun p -> textOf p, p.Metadata.SectionBreak))
    [ "end", None; " of section", Some "nextPage"; "next section", None ]
let broken =
    docOf [ Doc.Paragraph { paragraph "b1" "before break" with Metadata = { Doc.defaultMetadata with PageBreakAfter = true; ColumnBreakAfter = true } } ]
    |> caret "b1" 6 |> Doc.splitBlock
check "Enter: a column break after the text stays after it"
    (Doc.allParagraphs broken |> List.map (fun p -> textOf p, p.Metadata.PageBreakAfter, p.Metadata.ColumnBreakAfter))
    [ "before", false, false; " break", true, true ]
let rendered =
    docOf [ Doc.Paragraph { paragraph "r1" "started a column" with Metadata = { Doc.defaultMetadata with RenderedPageBreakBefore = Some 2 } } ]
    |> caret "r1" 7 |> Doc.splitBlock
check "Enter: only the first part started a page or column when Word laid it out"
    (Doc.allParagraphs rendered |> List.map (fun p -> textOf p, p.Metadata.RenderedPageBreakBefore))
    [ "started", Some 2; " a column", None ]

// ---- the clip as HTML for other programs ----
let html = Doc.clipToHtml multi
check "html: a heading, a list item, a paragraph"
    html "<h1 style=\"text-align:center\">tle</h1><ul><li>Body text</li></ul><p>Last</p>"
let boldHtml = Doc.clipToHtml clip
check "html: bold text in a span" boldHtml "<p>o <span style=\"font-weight:bold\">bold</span> w</p>"
let sized =
    let para = paragraph "z" "big red link"
    let buffer =
        para.Buffer
        |> TB.setFontSize 0 3 24.0
        |> TB.setColor 4 3 (Some "#FF0000")
        |> TB.setHyperlink 8 4 (Some "https://example.com/?a=1&b=2")
    docOf [ Doc.Paragraph { para with Buffer = buffer } ] |> range ("z", 0) ("z", 12) |> copy |> Doc.clipToHtml
check "html: a size in points, a colour, a link (escaped)"
    sized "<p><span style=\"font-size:18pt\">big</span> <span style=\"color:#FF0000\">red</span> <a href=\"https://example.com/?a=1&amp;b=2\">link</a></p>"

// ---- HTML pasted from elsewhere ----
let el tag attrs children = Doc.HtmlElement (tag, attrs, children)
let txt s = Doc.HtmlText s
let fromHtml nodes = (Doc.clipOfHtml nodes).Value
let webClip =
    fromHtml [
        el "h2" [] [ txt "A title" ]
        el "p" [ "style", "text-align: center" ] [ txt "Some "; el "b" [] [ txt "bold" ]; txt " and\n   "; el "span" [ "style", "font-style:italic; color: rgb(255, 0, 0); font-size: 12pt" ] [ txt "red italic" ] ]
        el "ul" [] [ el "li" [] [ txt "one" ]; el "li" [] [ txt "two"; el "ol" [] [ el "li" [] [ txt "inner" ] ] ] ]
        el "p" [] [ txt "line"; el "br" [] []; txt "break "; el "a" [ "href", "https://example.com" ] [ txt "link" ]; el "a" [ "href", "javascript:alert(1)" ] [ txt " bad" ] ]
        el "p" [] [ txt " " ] ]
check "html in: one paragraph a block" (webClip.Paragraphs |> List.map (fun p -> TB.toString p.Buffer)) [ "A title"; "Some bold and red italic"; "one"; "two"; "inner"; "line\nbreak link bad"; "" ]
check "html in: the plain text" webClip.Text "A title\nSome bold and red italic\none\ntwo\ninner\nline\nbreak link bad\n"
let piece i = webClip.Paragraphs.[i]
let styleIn i offset = TB.getStyleAt offset (piece i).Buffer
check "html in: a heading" ((piece 0).HeadingLevel) (Some 2)
check "html in: centred" ((piece 1).Metadata.Alignment, (piece 1).Metadata.ExplicitAlignment) (Doc.Center, true)
check "html in: bold" ((styleIn 1 5).Bold, (styleIn 1 4).Bold) (true, false)
check "html in: italic, red, 12 pt" ((styleIn 1 15).Italic, (styleIn 1 15).Color, (styleIn 1 15).FontSize, TB.ExplicitFormat.has TB.ExplicitFormat.FontSize (styleIn 1 15)) (true, Some "#FF0000", 16.0, true)
check "html in: list items, a nested list a level down"
    ([ 2; 3; 4 ] |> List.map (fun i -> (piece i).Metadata.ListStyle, (piece i).Metadata.ListLevel)) [ Doc.Bullet, 1; Doc.Bullet, 1; Doc.Numbered, 2 ]
check "html in: a link; one that may not be followed is text" ((styleIn 5 12).Hyperlink, (styleIn 5 17).Hyperlink) (Some "https://example.com", None)
let wordClip =
    fromHtml [
        el "p" [ "class", "MsoListParagraph"; "style", "mso-list:l0 level1 lfo1" ] [ el "span" [ "style", "mso-list:Ignore" ] [ txt "1." ]; txt "First" ]
        el "p" [ "class", "MsoListParagraph"; "style", "mso-list:l1 level2 lfo2" ] [ el "span" [ "style", "mso-list:Ignore" ] [ txt "·" ]; txt "Dot" ] ]
check "html in: Word's list paragraphs, without their numbers"
    (wordClip.Paragraphs |> List.map (fun p -> TB.toString p.Buffer, p.Metadata.ListStyle, p.Metadata.ListLevel)) [ "First", Doc.Numbered, 1; "Dot", Doc.Bullet, 2 ]
let htmlPasted = docOf [ Doc.Paragraph (paragraph "h" "[]") ] |> caret "h" 1 |> Doc.pasteClip (fromHtml [ el "p" [] [ el "b" [] [ txt "x" ] ]; el "p" [] [ txt "y" ] ])
check "html in: pasted" (texts htmlPasted) [ "[x"; "y]" ]

// ---- Word's keys for character and paragraph formatting ----
check "grow font: Word's size steps" ([ 8.0; 10.0; 10.5; 11.0; 28.0; 72.0; 75.0; 1638.0 ] |> List.map (Doc.nextFontSizePoints true true)) [ 9.0; 10.5; 11.0; 12.0; 36.0; 80.0; 80.0; 1638.0 ]
check "shrink font: Word's size steps" ([ 12.0; 11.0; 36.0; 80.0; 75.0; 8.0; 1.0 ] |> List.map (Doc.nextFontSizePoints false true)) [ 11.0; 10.5; 28.0; 72.0; 72.0; 7.0; 1.0 ]
check "a point larger / smaller" [ Doc.nextFontSizePoints true false 11.0; Doc.nextFontSizePoints false false 11.0; Doc.nextFontSizePoints true false 10.5 ] [ 12.0; 10.0; 11.0 ]
let raw (_: Doc.ParagraphBlock) (style: TextStyle) = style
let sizedDoc =
    let para = paragraph "f" "small big"
    docOf [ Doc.Paragraph { para with Buffer = para.Buffer |> TB.setFontSize 6 3 (24.0 * 96.0 / 72.0) } ] |> range ("f", 0) ("f", 9)
let grown = Doc.resizeFontWith raw true true sizedDoc
let pointsAt (doc: Doc.Document) offset = System.Math.Round((styleAt doc "f" offset).FontSize * 72.0 / 96.0, 2)
check "grow font: each run from its own size" (pointsAt grown 0, pointsAt grown 7) (Doc.nextFontSizePoints true true (TB.defaultStyle.FontSize * 72.0 / 96.0), 26.0)
check "grow font: stated on the runs" (TB.ExplicitFormat.has TB.ExplicitFormat.FontSize (styleAt grown "f" 0)) true
let sub = Doc.toggleVertAlignWith raw "subscript" sizedDoc
check "subscript on" ((styleAt sub "f" 0).VertAlign, (styleAt sub "f" 8).VertAlign) (Some "subscript", Some "subscript")
check "subscript off again" ((styleAt (Doc.toggleVertAlignWith raw "subscript" sub) "f" 0).VertAlign) None
check "superscript over subscript" ((styleAt (Doc.toggleVertAlignWith raw "superscript" sub) "f" 0).VertAlign) (Some "superscript")
let indentDoc =
    docOf [ Doc.Paragraph (paragraph "i1" "plain")
            Doc.Paragraph { paragraph "i2" "odd" with Metadata = { Doc.defaultMetadata with IndentLeft = Some 20.0 } }
            Doc.Paragraph { paragraph "i3" "item" with Metadata = listed } ]
    |> range ("i1", 0) ("i3", 2)
let indentOf (doc: Doc.Document) = Doc.allParagraphs doc |> List.map (fun p -> p.Metadata.IndentLeft, p.Metadata.ListLevel)
let indented = Doc.indentParagraphs true indentDoc
check "Ctrl+M: to the next half inch; a list item a level down" (indentOf indented) [ Some 48.0, 0; Some 48.0, 0; None, 2 ]
check "Ctrl+M: the indent is the paragraph's own" ((Doc.allParagraphs indented).Head.Metadata.Explicit &&& Doc.ParagraphExplicit.IndentLeft <> 0) true
check "Ctrl+Shift+M: back" (indentOf (Doc.indentParagraphs false indented)) [ Some 0.0, 0; Some 0.0, 0; None, 1 ]
check "Ctrl+Shift+M: not below the margin, nor out of the list" (indentOf (Doc.indentParagraphs false indentDoc)) [ None, 0; Some 0.0, 0; None, 1 ]

// ---- Backspace at the start of a list item: the bullet goes first ----
let listDoc =
    docOf [ Doc.Paragraph { paragraph "l1" "first" with Metadata = listed }
            Doc.Paragraph { paragraph "l2" "second" with Metadata = { listed with ListLevel = 2 } } ]
let unlisted = listDoc |> caret "l2" 0 |> Doc.deleteBackward
check "Backspace at the start of a list item: out of the list, the text as it was"
    (texts unlisted, Doc.allParagraphs unlisted |> List.map (fun p -> p.Metadata.ListStyle), selectionOf unlisted) ([ "first"; "second" ], [ Doc.Bullet; Doc.NoList ], Some ("l2", 0, "l2", 0))
check "Backspace again: joined to the item before" (texts (Doc.deleteBackward unlisted)) [ "firstsecond" ]

printfn ""
printfn "Word editing tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
