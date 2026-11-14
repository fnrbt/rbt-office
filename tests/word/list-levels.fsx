// The levels of a list. The file's numbering part says for each level of a
// list its number format, the number's text ("%1.%2."), where it starts, its
// indents, what follows the number, how the number stands at its position and
// how it looks. A paragraph is read at one level; all the levels of its list
// come with it (BlockMetadata.ListLevels), so a paragraph moved to another
// level (Tab / Shift+Tab, setListLevel) is at once what it is when the file
// is read with the paragraph at that level: its number, its text, its indents.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open Rbt.Word.TextBuffer
module Doc = Rbt.Word.Document
module Layout = Rbt.Word.Layout

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-50s expected %A got %A" label expected got

let w = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
let package (body: string) (numbering: string) =
    use stream = new IO.MemoryStream()
    do
        use archive = new IO.Compression.ZipArchive(stream, IO.Compression.ZipArchiveMode.Create, true)
        let add (name: string) (text: string) =
            use writer = new IO.StreamWriter(archive.CreateEntry(name).Open(), Text.UTF8Encoding(false))
            writer.Write text
        add "[Content_Types].xml" """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"""
        add "word/document.xml" (sprintf "<w:document %s><w:body>%s<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr></w:body></w:document>" w body)
        add "word/numbering.xml" (sprintf "<w:numbering %s>%s</w:numbering>" w numbering)
    stream.ToArray()
let import body numbering = package body numbering |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc

let level ilvl (inner: string) = sprintf "<w:lvl w:ilvl=\"%d\">%s</w:lvl>" ilvl inner
let ind left hanging = sprintf "<w:pPr><w:ind w:left=\"%d\" w:hanging=\"%d\"/></w:pPr>" left hanging
// List 1 (abstract 0): I. / I.a) / 1.1.1 (legal) / a level that never starts again.
// List 2 (abstract 1): bullets in Symbol and Wingdings.
// List 3: list 1's definition, starting at 5 (w:startOverride).
let numbering =
    "<w:abstractNum w:abstractNumId=\"0\">"
    + level 0 ("<w:start w:val=\"1\"/><w:numFmt w:val=\"upperRoman\"/><w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"right\"/>" + ind 720 360 + "<w:rPr><w:b/><w:color w:val=\"FF0000\"/></w:rPr>")
    + level 1 ("<w:start w:val=\"1\"/><w:numFmt w:val=\"lowerLetter\"/><w:lvlText w:val=\"%1.%2)\"/><w:suff w:val=\"nothing\"/>" + ind 1440 432)
    + level 2 ("<w:start w:val=\"3\"/><w:numFmt w:val=\"decimal\"/><w:isLgl/><w:lvlText w:val=\"%1.%2.%3\"/>" + ind 2160 504 + "<w:rPr><w:i/><w:sz w:val=\"32\"/></w:rPr>")
    + level 3 ("<w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlRestart w:val=\"0\"/><w:lvlText w:val=\"(%4)\"/>" + ind 2880 648)
    + "</w:abstractNum>"
    + "<w:abstractNum w:abstractNumId=\"1\">"
    + level 0 ("<w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"\"/>" + ind 720 360 + "<w:rPr><w:rFonts w:ascii=\"Symbol\" w:hAnsi=\"Symbol\"/></w:rPr>")
    + level 1 ("<w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"\"/>" + ind 1440 360 + "<w:rPr><w:rFonts w:ascii=\"Wingdings\" w:hAnsi=\"Wingdings\"/></w:rPr>")
    + "</w:abstractNum>"
    + "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>"
    + "<w:num w:numId=\"2\"><w:abstractNumId w:val=\"1\"/></w:num>"
    + "<w:num w:numId=\"3\"><w:abstractNumId w:val=\"0\"/><w:lvlOverride w:ilvl=\"0\"><w:startOverride w:val=\"5\"/></w:lvlOverride></w:num>"
let item numId ilvl (text: string) = sprintf "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"%d\"/><w:numId w:val=\"%d\"/></w:numPr></w:pPr><w:r><w:t>%s</w:t></w:r></w:p>" ilvl numId text
let itemWith numId ilvl (pPr: string) (text: string) = sprintf "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"%d\"/><w:numId w:val=\"%d\"/></w:numPr>%s</w:pPr><w:r><w:t>%s</w:t></w:r></w:p>" ilvl numId pPr text

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let settingsOf (doc: Doc.Document) : Layout.LayoutSettings =
    let s = doc.PageSetup
    { Width = s.PageWidth; PageHeight = s.PageHeight; FontSize = 16.0; LineHeight = 24.0; PaddingLeft = s.MarginLeft; PaddingTop = s.MarginTop
      MarginRight = s.MarginRight; MarginBottom = s.MarginBottom; InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 3.5; DefaultTabStop = 48.0; CompatibilityMode = 12 }
let paragraphs (doc: Doc.Document) = doc.Blocks |> List.choose Doc.tryGetParagraph
let layoutsOf (doc: Doc.Document) = Layout.layoutDocumentPages (settingsOf doc) measureText measureRange doc |> fst
let prefixes (doc: Doc.Document) = layoutsOf doc |> List.map (fun layout -> layout.Prefix)
let caretIn (doc: Doc.Document) (index: int) =
    let id = (paragraphs doc).[index].Id
    { doc with Selection = Some { Anchor = { BlockId = id; Offset = 0 }; Focus = { BlockId = id; Offset = 0 } } }
/// What of a paragraph says where it is in its list and how it stands.
let listed (para: Doc.ParagraphBlock) =
    let m = para.Metadata
    m.ListStyle, m.ListLevel, m.ListId, m.ListNumberFormat, m.ListLevelText, m.ListStart, m.IndentLeft, m.FirstLineIndent, m.ListLevelIndents
let placed (doc: Doc.Document) (index: int) =
    let layout = (layoutsOf doc).[index]
    layout.Prefix, Math.Round(layout.Indent, 2), Math.Round(layout.PrefixWidth, 2), layout.PrefixStyle |> Option.map (fun style -> style.Bold, style.Italic, style.Color, style.FontFamily, style.FontSize)

// ---- all the levels come with a paragraph ---------------------------------
let flat = import (item 1 0 "one" + item 1 0 "two" + item 1 0 "three") numbering
let flatParas = paragraphs flat
check "a paragraph of the list has the list's levels" (flatParas.[0].Metadata.ListLevels |> List.map (fun l -> l.Level, l.NumberFormat, l.LevelText)) [ (1, "upperRoman", Some "%1."); (2, "lowerLetter", Some "%1.%2)"); (3, "decimal", Some "%1.%2.%3"); (4, "decimal", Some "(%4)") ]
check "the levels' indents (px), what follows the number, how it stands, legal, restart"
    (flatParas.[0].Metadata.ListLevels |> List.map (fun l -> l.IndentLeft, l.FirstLineIndent, l.Suffix, l.Justification, l.Legal, l.Restart, l.Start))
    [ (Some 48.0, Some -24.0, None, Some "right", false, None, Some 1); (Some 96.0, Some -28.8, Some "nothing", None, false, None, Some 1); (Some 144.0, Some -33.6, None, None, true, None, Some 3); (Some 192.0, Some -43.2, None, None, false, Some 0, Some 1) ]
check "the look of the numbers" (flatParas.[0].Metadata.ListLevels |> List.map (fun l -> l.MarkBold, l.MarkItalic, l.MarkColor, l.MarkFont, l.MarkSize |> Option.map (fun size -> Math.Round(size, 2))) |> List.take 3) [ (Some true, None, Some "#FF0000", None, None); (None, None, None, None, None); (None, Some true, None, None, Some 21.33) ]
check "read: I. II. III." (prefixes flat) [ "I. "; "II. "; "III. " ]

// ---- a level change is what a reload at that level gives ----------------------
let moved = Doc.increaseIndent (caretIn flat 1)
let reread = import (item 1 0 "one" + item 1 1 "two" + item 1 0 "three") numbering
check "Tab on the second item: its level, format, text, start and indents are those of the file read with it at that level" (listed (paragraphs moved).[1]) (listed (paragraphs reread).[1])
check "and its number, its place and the look of its number" (placed moved 1) (placed reread 1)
check "the numbers: I. / I.a) / II." (prefixes moved) [ "I. "; "I.a)"; "II. " ]
check "what follows the number: nothing, where the level says so" ((prefixes moved).[1].EndsWith " ") false
let twice = Doc.increaseIndent (caretIn moved 1)
let rereadTwice = import (item 1 0 "one" + item 1 2 "two" + item 1 0 "three") numbering
check "Tab again: the third level" (listed (paragraphs twice).[1], placed twice 1) (listed (paragraphs rereadTwice).[1], placed rereadTwice 1)
check "a legal level: the numbers of the levels above it are decimal (1.1.3 - the level starts at 3)" (prefixes twice).[1] "1.1.3 "
check "the number of the third level is italic and 16 pt (21.33 px), as the level says" ((layoutsOf twice).[1].PrefixStyle |> Option.map (fun style -> style.Italic, Math.Round(style.FontSize, 2))) (Some (true, 21.33))
let back = Doc.decreaseIndent (caretIn (Doc.decreaseIndent (caretIn twice 1)) 1)
check "Shift+Tab twice: the paragraph is as it was read" (listed (paragraphs back).[1], placed back 1) (listed flatParas.[1], placed flat 1)
let set = Doc.setListLevel 2 (caretIn flat 2)
check "setListLevel 2 on the third item" (listed (paragraphs set).[2], (prefixes set).[2]) (listed (paragraphs (import (item 1 0 "one" + item 1 0 "two" + item 1 1 "three") numbering)).[2], "II.a)")
// The first level's number: bold and red, right-aligned.
check "the first level's number is bold and red" ((layoutsOf flat).[0].PrefixStyle |> Option.map (fun style -> style.Bold, style.Color)) (Some (true, Some "#FF0000"))
check "and right-aligned" (Rbt.Word.WordLayout.Numbering.markerJustification flatParas.[0].Metadata, Rbt.Word.WordLayout.Numbering.markerJustification (paragraphs moved).[1].Metadata) ("right", "left")

// An indent of the paragraph's own stays over the level's.
let ownIndent = import (item 1 0 "one" + itemWith 1 0 "<w:ind w:left=\"3000\"/>" "two") numbering
let ownMoved = Doc.increaseIndent (caretIn ownIndent 1)
check "a left indent the paragraph states itself stays; the first line is the new level's" ((paragraphs ownMoved).[1].Metadata.IndentLeft, (paragraphs ownMoved).[1].Metadata.FirstLineIndent) (Some 200.0, Some -28.8)

// ---- multi-level numbers, restarts ------------------------------------------
let deep = import (item 1 0 "a" + item 1 1 "b" + item 1 1 "c" + item 1 3 "d" + item 1 0 "e" + item 1 1 "f" + item 1 3 "g" + item 1 2 "h" + item 1 2 "i") numbering
check "each level's number is written as that level writes it; a level that never starts again counts on"
    (prefixes deep) [ "I. "; "I.a)"; "I.b)"; "(1) "; "II. "; "II.a)"; "(2) "; "2.1.3 "; "2.1.4 " ]
// A list that starts at 5 (w:startOverride).
let from5 = import (item 3 0 "five" + item 3 0 "six" + item 3 1 "six a") numbering
check "w:startOverride: V. VI. VI.a)" (prefixes from5) [ "V. "; "VI. "; "VI.a)" ]
check "Tab in it: the levels are the overridden list's" (prefixes (Doc.increaseIndent (caretIn from5 1))) [ "V. "; "V.a)"; "V.b)" ]

// ---- bullets ------------------------------------------------------------------
let bullets = import (item 2 0 "dot" + item 2 0 "another") numbering
check "a bullet in Symbol" ((layoutsOf bullets).[0].PrefixStyle |> Option.bind (fun style -> style.FontFamily), (paragraphs bullets).[0].Metadata.ListLevelText) (Some "Symbol", Some "")
let bulletMoved = Doc.increaseIndent (caretIn bullets 1)
check "Tab: the second level's bullet, in Wingdings, at its indent"
    ((paragraphs bulletMoved).[1].Metadata.ListLevelText, (layoutsOf bulletMoved).[1].PrefixStyle |> Option.bind (fun style -> style.FontFamily), (paragraphs bulletMoved).[1].Metadata.IndentLeft, (paragraphs bulletMoved).[1].Metadata.ListStyle)
    (Some "", Some "Wingdings", Some 96.0, Doc.Bullet)
check "as the file read at that level" (listed (paragraphs bulletMoved).[1], placed bulletMoved 1) (let again = import (item 2 0 "dot" + item 2 1 "another") numbering in listed (paragraphs again).[1], placed again 1)

// ---- a list made in the editor has no levels of a file: as before ---------------
let made =
    let para id text : Doc.ParagraphBlock = { Id = id; Buffer = ofString text; Metadata = Doc.defaultMetadata }
    { Doc.empty with Blocks = [ Doc.Paragraph (para "m1" "one"); Doc.Paragraph (para "m2" "two") ]; Selection = Some { Anchor = { BlockId = "m1"; Offset = 0 }; Focus = { BlockId = "m2"; Offset = 3 } } }
    |> Doc.toggleNumberedList
let madeMoved = Doc.increaseIndent { made with Selection = Some { Anchor = { BlockId = "m2"; Offset = 0 }; Focus = { BlockId = "m2"; Offset = 0 } } }
check "a list made in the editor: Tab changes the level" ((paragraphs madeMoved) |> List.map (fun para -> para.Metadata.ListLevel, para.Metadata.ListLevels.IsEmpty), prefixes made) ([ (1, true); (2, true) ], [ "1. "; "2. " ])

// ---- where the text after the number starts ------------------------------------
// The number stands at the first-line indent. A tab after it goes to the
// nearest of a custom tab stop (the level's own too) and the left indent when
// the paragraph hangs and the number ends before it; else to the next default
// tab stop (every 48 px from the margin). A number wider than its hanging
// indent so pushes the text on instead of overlapping it. The lines after
// the first start at the left indent.
let placement =
    "<w:abstractNum w:abstractNumId=\"5\">"
    + level 0 ("<w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/>" + ind 720 360)
    + level 1 ("<w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"Article %2.\"/>" + ind 720 360)
    + level 2 ("<w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%3.\"/><w:suff w:val=\"space\"/>" + ind 720 360)
    + level 3 ("<w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%4.\"/><w:pPr><w:tabs><w:tab w:val=\"num\" w:pos=\"1080\"/></w:tabs><w:ind w:left=\"1440\" w:hanging=\"720\"/></w:pPr>")
    + level 4 ("<w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%5.\"/><w:pPr><w:ind w:left=\"0\" w:firstLine=\"0\"/></w:pPr>")
    + "</w:abstractNum><w:num w:numId=\"5\"><w:abstractNumId w:val=\"5\"/></w:num>"
let longText = String.replicate 30 "word "
let placedDoc = import (item 5 0 "one" + item 5 1 "two" + item 5 2 "three" + item 5 3 "four" + item 5 4 longText) placement
let placedLayouts = layoutsOf placedDoc
/// The first line's text, from the left indent; the number's end, from the margin.
let textFromIndent (layout: Layout.BlockLayout) =
    match layout.Lines.Head.Segments with
    | Layout.TextSegment (_, _, x, _, _) :: _ -> Math.Round(x - layout.Indent, 2)
    | _ -> nan
let numberEnd (layout: Layout.BlockLayout) =
    layout.Indent + Rbt.Word.WordLayout.Paragraph.firstLineIndent layout.Metadata + measureText (layout.Prefix.TrimEnd()) layout.PrefixStyle.Value
let nextDefaultStop x = (floor (x / 48.0) + 1.0) * 48.0
check "a number that ends before the hanging indent: the text at the indent" (textFromIndent placedLayouts.[0]) 0.0
check "a number wider than the hanging indent: the text at the next default tab stop, past the number"
    (textFromIndent placedLayouts.[1]) (Math.Round(nextDefaultStop (numberEnd placedLayouts.[1]) - placedLayouts.[1].Indent, 2))
check "and that is past the number" (textFromIndent placedLayouts.[1] + placedLayouts.[1].Indent > numberEnd placedLayouts.[1]) true
check "w:suff space: a space after the number"
    (textFromIndent placedLayouts.[2]) (Math.Round(numberEnd placedLayouts.[2] + measureText " " placedLayouts.[2].PrefixStyle.Value - placedLayouts.[2].Indent, 2))
check "the level's tab stop before the indent: the text at it" (textFromIndent placedLayouts.[3]) -24.0
check "no hanging indent: the text at the next default tab stop" (textFromIndent placedLayouts.[4]) 48.0
check "no hanging indent: the lines after the first at the indent"
    (placedLayouts.[4].Lines.Length > 1
     && placedLayouts.[4].Lines.Tail |> List.forall (fun line -> match line.Segments with Layout.TextSegment (_, _, x, _, _) :: _ -> x = placedLayouts.[4].Indent | _ -> false))
    true
check "no hanging indent: the lines are as wide as the paragraph (the page's 624 px)" (placedLayouts.[4].AvailableWidth) 624.0

// Word's render of tests/repros/word/list-default-indent.docx (docs/images/word/list-default-indent-word.png):
// levels with no indents put every number at the margin and every text half an inch in.
let defaultIndentDoc =
    IO.File.ReadAllBytes(IO.Path.Combine(__SOURCE_DIRECTORY__, "../repros/word/list-default-indent.docx"))
    |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
let defaultIndentItems = layoutsOf defaultIndentDoc |> List.filter (fun layout -> layout.Prefix <> "")
check "list-default-indent: six numbered paragraphs" defaultIndentItems.Length 6
check "list-default-indent: the numbers at the margin, the texts half an inch in, as in Word"
    (defaultIndentItems |> List.map (fun layout -> Math.Round(layout.Indent + Rbt.Word.WordLayout.Paragraph.firstLineIndent layout.Metadata, 2), Math.Round(textFromIndent layout + layout.Indent, 2)) |> List.distinct)
    [ 0.0, 48.0 ]

// Dialog operations split only the affected list instance and remain one
// document value for the caller's transaction.
let keep: Doc.ListDialogChange = { Kind = None; Marker = None; Level = None; Sequence = "keep"; Start = 1 }
let edited result = match result with Ok doc -> doc | Error message -> failwith message
let noChange = Doc.applyListDialog (fun () -> failwith "no id should be allocated") keep (caretIn flat 1) |> edited
check "list dialog unchanged keeps document" (obj.ReferenceEquals(noChange, noChange |> Doc.applyListDialog (fun () -> failwith "unexpected allocation") keep |> edited)) true
let allocator first =
    let mutable next = first - 1
    fun () -> next <- next + 1; string next
let changedMarker =
    Doc.applyListDialog (allocator 100) { keep with Marker = Some "lowerLetter" } (caretIn flat 1) |> edited
check "partial marker edit leaves first list instance alone"
    ((paragraphs changedMarker).[0].Metadata.ListId, (paragraphs changedMarker).[1].Metadata.ListSourceId, (paragraphs changedMarker).[2].Metadata.ListSourceId)
    (Some "1", Some "1", Some "1")
check "partial marker edit retains ordinal continuity" (prefixes changedMarker) [ "I. "; "b. "; "III. " ]
let nestedSequence = import (item 1 0 "one" + item 1 0 "two" + item 1 1 "nested" + item 1 0 "four") numbering
let nestedEdited = Doc.applyListDialog (allocator 100) { keep with Marker = Some "lowerLetter" } (caretIn nestedSequence 1) |> edited
check "partial marker edit keeps unselected nested parent counters"
    (prefixes nestedEdited) [ "I. "; "b. "; "II.a)"; "III. " ]
let nestedRepeated = Doc.applyListDialog (allocator 200) { keep with Marker = Some "upperLetter" } (caretIn nestedEdited 1) |> edited
check "repeated partial marker edit preserves all ordinals"
    (prefixes nestedRepeated) [ "I. "; "B. "; "II.a)"; "III. " ]
let deepEdited = Doc.applyListDialog (allocator 300) { keep with Marker = Some "lowerRoman" } (caretIn deep 1) |> edited
check "partial marker edit preserves restarted and never-restarted child counters"
    (prefixes deepEdited) [ "I. "; "I.i)"; "I.b)"; "(1) "; "II. "; "II.a)"; "(2) "; "2.1.3 "; "2.1.4 " ]
let movedAndFormatted = Doc.applyListDialog (allocator 400) { keep with Marker = Some "lowerRoman"; Level = Some 2 } (caretIn flat 1) |> edited
check "format and level use the target imported template"
    ((paragraphs movedAndFormatted).[1].Metadata.ListLevelText, (paragraphs movedAndFormatted).[1].Metadata.ListLevels |> List.find (fun level -> level.Level = 2) |> fun level -> level.Suffix)
    (Some "%1.%2)", Some "nothing")
let rangeIn (doc: Doc.Document) first last =
    let a, b = (paragraphs doc).[first], (paragraphs doc).[last]
    { doc with Selection = Some { Anchor = { BlockId = a.Id; Offset = 0 }; Focus = { BlockId = b.Id; Offset = 1 } } }
let combined = Doc.applyListDialog (allocator 500) { keep with Marker = Some "lowerLetter"; Sequence = "restart"; Start = 5 } (rangeIn flat 1 2) |> edited
check "format and restart together number the selected run once"
    (prefixes combined) [ "I. "; "e. "; "f. " ]
let plain = Doc.setParagraphContents [ "first"; "second" ] Doc.empty |> fun doc -> rangeIn doc 0 1
let newRestarted = Doc.applyListDialog (allocator 600) { keep with Kind = Some Doc.Numbered; Marker = Some "lowerLetter"; Sequence = "restart"; Start = 5 } plain |> edited
check "ordinary paragraphs can become a restarted formatted list in one edit"
    (prefixes newRestarted) [ "e. "; "f. " ]
check "partial marker edit changes only selected formatting"
    (paragraphs changedMarker |> List.map (fun p -> p.Metadata.ListNumberFormat))
    [ Some "upperRoman"; Some "lowerLetter"; Some "upperRoman" ]
let restarted = Doc.applyListDialog (fun () -> "100") { keep with Sequence = "restart"; Start = 5 } (caretIn flat 1) |> edited
check "restart splits the contiguous tail" (paragraphs restarted |> List.map (fun p -> p.Metadata.ListId)) [ Some "1"; Some "100"; Some "100" ]
check "restart changes live counters" (prefixes restarted) [ "I. "; "V. "; "VI. " ]
let restartedAndFormatted = Doc.applyListDialog (allocator 200) { keep with Marker = Some "lowerLetter" } (caretIn restarted 1) |> edited
let restartedItem = (paragraphs restartedAndFormatted).[1].Metadata
check "formatting a restarted item retains its restart provenance"
    (restartedItem.ListSourceId, restartedItem.ListRestart, restartedItem.ListStart)
    (Some "1", true, Some 5)
let separated = import (item 1 0 "one" + "<w:p><w:r><w:t>between</w:t></w:r></w:p>" + item 3 0 "two" + item 3 0 "three") numbering
let continued = Doc.applyListDialog (fun () -> failwith "continue must not allocate") { keep with Sequence = "continue" } (caretIn separated 2) |> edited
check "continue reaches preceding list across ordinary paragraph"
    (paragraphs continued |> List.map (fun p -> p.Metadata.ListId)) [ Some "1"; None; Some "1"; Some "1" ]
check "continue uses preceding list counters" (prefixes continued) [ "I. "; ""; "II. "; "III. " ]
let continuedCaret = caretIn continued 2
let continuedAgain = Doc.applyListDialog (fun () -> failwith "continue no-op must not allocate") { keep with Sequence = "continue" } continuedCaret |> edited
check "continuing the same list keeps the document" (obj.ReferenceEquals(continuedAgain, continuedCaret)) true

// A bullet of Wingdings or Symbol (a private-use code, U+F0xx) is shown as the Unicode
// character it looks like: no host has those faces, and no other face has those codes
// (a checklist's Wingdings "o", U+F06F, showed as a missing-glyph box).
let marker = Rbt.Word.WordLayout.Numbering.displayMarkerText
check "a Wingdings checkbox bullet" (marker (Some "Wingdings") "\uF06F") "□"
check "a Wingdings arrow bullet" (marker (Some "Wingdings") "\uF0D8") "➢"
check "a Wingdings checked box" (marker (Some "Wingdings") "\uF0FE") "☑"
check "a Symbol bullet" (marker (Some "Symbol") "\uF0B7") "•"
check "Symbol's A8 is a diamond, not Wingdings' square" (marker (Some "Symbol") "\uF0A8") "♦"
check "the bullets known in any face" (marker None "\uF0A7") "▪"
check "a marker in another face is left as it is" (marker (Some "Calibri") "\uF06F") "\uF06F"
check "a number is left as it is" (marker (Some "Wingdings") "1.") "1."

printfn "word list levels: %d passed, %d failed" pass fail
if fail > 0 then exit 1
