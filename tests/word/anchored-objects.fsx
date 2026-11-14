// Anchored (floating) pictures: where their anchor puts them - relative to
// the page, the margins, the column, the anchor's paragraph, line or
// character, by an offset or an alignment (wp:positionH / wp:positionV) - and
// how text goes around them: beside the picture's box ("square", with its
// distance from text), above and below it only ("top and bottom"), as if it
// were not there (no wrapping: behind or in front of the text), and along the
// picture's outline ("tight" / "through" with a wp:wrapPolygon; the box
// without one).
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"

open Rbt.Word.Document
open Rbt.Word.Layout
module TB = Rbt.Word.TextBuffer

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-50s expected %A got %A" label expected got
let near (a: float) (b: float) = abs (a - b) < 0.75

let settings : LayoutSettings =
    { Width = 816.0; PageHeight = 1056.0; FontSize = 16.0; LineHeight = 24.0; PaddingLeft = 96.0; PaddingTop = 96.0
      MarginRight = 96.0; MarginBottom = 96.0; InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 3.5; DefaultTabStop = 48.0; CompatibilityMode = 12 }
let measureText (text: string) (style: TB.TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = rangeMeasurerFromTextMeasurer measureText
let para id (text: string) : ParagraphBlock = { Id = id; Buffer = TB.ofString text; Metadata = { defaultMetadata with SpacingBefore = Some 0.0; SpacingAfter = Some 0.0 } }
let words n = String.concat " " [ for i in 1 .. n -> sprintf "w%d" i ]
let textWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
let pitch = settings.PageHeight + pageGap

/// A paragraph of `text` with a floating picture of 120 x 90 anchored after `anchorAt` characters.
let withPicture id (text: string) (anchorAt: int) (shape: TB.ImageRun -> TB.ImageRun) : ParagraphBlock =
    let buffer =
        TB.ofString text
        |> TB.insertImageWithMetadata anchorAt "data:image/png;base64,AAAA" false (Some 120.0) (Some 90.0) None None None true (Some "square") (Some 0.0) (Some 0.0) None None None None None None None None None None None None None None None None None None
    let runs = buffer.Runs |> List.map (function TB.Image image -> TB.Image (shape image) | run -> run)
    { para id text with Buffer = { buffer with Runs = runs } }
let layoutOf (blocks: Block list) = layoutDocumentPages settings measureText measureRange { empty with Blocks = blocks; Selection = None } |> fst
/// The picture of a paragraph: its left edge from the left edge of the page, and its top down the canvas.
let pictureOf (layouts: BlockLayout list) (id: string) =
    layouts
    |> List.filter (fun layout -> layout.BlockId = id)
    |> List.collect (fun layout -> layout.Lines |> List.collect (fun line -> line.Segments |> List.choose (function ImageSegment (_, x, yOffset, _, _, _, _, _, _, _, _, _, _, _) -> Some (settings.PaddingLeft + layout.Indent * 0.0 + x, line.Y + yOffset) | _ -> None)))
    |> List.head
let linesOf (layouts: BlockLayout list) (id: string) = layouts |> List.filter (fun layout -> layout.BlockId = id) |> List.collect (fun layout -> layout.Lines)
/// Where a line's text starts and ends, from the left edge of the page.
let extentOf (line: MeasuredLine) =
    let texts = line.Segments |> List.choose (function TextSegment (_, _, x, width, _) -> Some (x, x + width) | _ -> None)
    if texts.IsEmpty then None else Some (settings.PaddingLeft + (texts |> List.map fst |> List.min), settings.PaddingLeft + (texts |> List.map snd |> List.max))

// ---------------------------------------------------------------------------
// Where the anchor puts the picture.
// ---------------------------------------------------------------------------
let filler = [ for i in 1 .. 4 -> Paragraph (para (sprintf "f%d" i) (words 30)) ]
let placed (shape: TB.ImageRun -> TB.ImageRun) =
    // (No wrapping: where the picture is does not move the text it is anchored in.)
    let layouts = layoutOf (filler @ [ Paragraph (withPicture "a" (words 60) 151 (fun image -> shape { image with Wrap = Some "none" })) ])
    let anchorLines = linesOf layouts "a"
    let anchorLine = anchorLines |> List.find (fun line -> line.Segments |> List.exists (function ImageSegment _ -> true | _ -> false))
    pictureOf layouts "a", (List.head anchorLines).Y, anchorLine.Y
let horizontal relativeFrom align offset = placed (fun image -> { image with PositionHRelativeFrom = relativeFrom; PositionHAlign = align; PositionX = Some offset }) |> fun ((x, _), _, _) -> x
check "from the page's left edge, by an offset" (near (horizontal (Some "page") None 40.0) 40.0) true
check "from the left margin (the text's edge), and from the column's" (near (horizontal (Some "margin") None 40.0) (96.0 + 40.0), near (horizontal (Some "column") None 25.0) (96.0 + 25.0)) (true, true)
check "aligned in the page: left, center, right" (near (horizontal (Some "page") (Some "left") 0.0) 0.0, near (horizontal (Some "page") (Some "center") 0.0) ((816.0 - 120.0) / 2.0), near (horizontal (Some "page") (Some "right") 0.0) (816.0 - 120.0)) (true, true, true)
check "aligned in the text area: left, center, right (an alignment goes before an offset)" (near (horizontal (Some "margin") (Some "left") 55.0) 96.0, near (horizontal (Some "margin") (Some "center") 55.0) (96.0 + (textWidth - 120.0) / 2.0), near (horizontal (Some "margin") (Some "right") 55.0) (96.0 + textWidth - 120.0)) (true, true, true)
check "in the left and in the right margin" (near (horizontal (Some "leftMargin") None 10.0) 10.0, near (horizontal (Some "rightMargin") None 10.0) (816.0 - 96.0 + 10.0), near (horizontal (Some "rightMargin") (Some "center") 0.0) (816.0 - 96.0 + (96.0 - 120.0) / 2.0)) (true, true, true)
let characterX = horizontal (Some "character") None 0.0
check "from the anchor's character: in the line, right of the left margin; an offset from there" (characterX > 96.0 + 1.0, near (horizontal (Some "character") None 30.0) (characterX + 30.0)) (true, true)
let vertical relativeFrom align offset = placed (fun image -> { image with PositionVRelativeFrom = relativeFrom; PositionVAlign = align; PositionY = Some offset })
let yOf ((_, y), _, _) = y
check "from the page's top, by an offset" (near (yOf (vertical (Some "page") None 300.0)) 300.0) true
check "from the top margin (it was from the anchor's line)" (near (yOf (vertical (Some "margin") None 50.0)) (96.0 + 50.0)) true
check "aligned in the page: top, center, bottom" (near (yOf (vertical (Some "page") (Some "top") 0.0)) 0.0, near (yOf (vertical (Some "page") (Some "center") 0.0)) ((1056.0 - 90.0) / 2.0), near (yOf (vertical (Some "page") (Some "bottom") 0.0)) (1056.0 - 90.0)) (true, true, true)
check "aligned in the text area: top, center, bottom" (near (yOf (vertical (Some "margin") (Some "top") 0.0)) 96.0, near (yOf (vertical (Some "margin") (Some "center") 0.0)) (96.0 + (864.0 - 90.0) / 2.0), near (yOf (vertical (Some "margin") (Some "bottom") 0.0)) (1056.0 - 96.0 - 90.0)) (true, true, true)
check "in the top margin and in the bottom margin" (near (yOf (vertical (Some "topMargin") None 5.0)) 5.0, near (yOf (vertical (Some "bottomMargin") None 3.0)) (1056.0 - 96.0 + 3.0)) (true, true)
let ((_, paragraphY), paragraphTop, anchorLineTop) = vertical (Some "paragraph") None 12.0
let ((_, lineY), _, _) = vertical (Some "line") None 12.0
check "from the top of the anchor's paragraph, and from the anchor's line (which is not the paragraph's first)" (near paragraphY (paragraphTop + 12.0), near lineY (anchorLineTop + 12.0), anchorLineTop > paragraphTop + 1.0) (true, true, true)
// The anchor takes no room in its line: the line does not end at it (a wide picture anchored late in a line broke the line there).
do
    let withAnchor = linesOf (layoutOf [ Paragraph (withPicture "b" (words 60) 99 (fun image -> { image with Wrap = Some "none" })) ]) "b"
    let without = linesOf (layoutOf [ Paragraph (para "b" (words 60)) ]) "b"
    check "the lines break where they break without the picture (one character further on, the anchor's)" (withAnchor |> List.map (fun line -> line.StartOffset), withAnchor.Length) (without |> List.map (fun line -> if line.StartOffset > 99 then line.StartOffset + 1 else line.StartOffset), without.Length)
// On the second page: relative to that page (it was relative to the first).
do
    let layouts = layoutOf ([ for i in 1 .. 90 -> Paragraph (para (sprintf "p%d" i) (words 30)) ] @ [ Paragraph (withPicture "late" (words 20) 10 (fun image -> { image with Wrap = Some "none"; PositionVRelativeFrom = Some "page"; PositionY = Some 200.0; PositionHRelativeFrom = Some "page"; PositionHAlign = Some "right" })) ])
    let anchorY = (linesOf layouts "late" |> List.head).Y
    let x, y = pictureOf layouts "late"
    check "an anchor on the second page: the picture is on that page, 200 below its top, at its right edge" (anchorY > pitch, near y (pitch + 200.0), near x (816.0 - 120.0)) (true, true, true)

// ---------------------------------------------------------------------------
// How text goes around it. The picture stands at the left margin, 30 below
// the top of its paragraph, which is long enough to pass it.
// ---------------------------------------------------------------------------
let wrapped (shape: TB.ImageRun -> TB.ImageRun) =
    let layouts =
        layoutOf [ Paragraph (withPicture "w" (words 220) 0 (fun image ->
                        shape { image with PositionHRelativeFrom = Some "margin"; PositionX = Some 0.0; PositionVRelativeFrom = Some "paragraph"; PositionY = Some 30.0 })) ]
    let _, top = pictureOf layouts "w"
    let lines = linesOf layouts "w"
    // The lines level with the picture, and where their text starts.
    lines, top, lines |> List.filter (fun line -> line.Y + line.Height > top + 1.0 && line.Y < top + 90.0 - 1.0) |> List.choose extentOf
do
    let lines, top, beside = wrapped (fun image -> { image with Wrap = Some "square" })
    check "square: the lines level with the picture start right of it" (beside.Length >= 3, beside |> List.forall (fun (left, _) -> left >= 96.0 + 120.0 - 0.5)) (true, true)
    check "the lines above and below it start at the margin" (lines |> List.filter (fun line -> line.Y + line.Height <= top + 1.0 || line.Y >= top + 90.0) |> List.choose extentOf |> List.forall (fun (left, _) -> near left 96.0)) true
    let _, _, withDistance = wrapped (fun image -> { image with Wrap = Some "square"; DistanceRight = Some 24.0 })
    check "with a distance from text: that much further right" (withDistance |> List.forall (fun (left, _) -> left >= 96.0 + 120.0 + 24.0 - 0.5), withDistance |> List.exists (fun (left, _) -> left < 96.0 + 120.0 + 24.0 + 20.0)) (true, true)
    let _, _, onTheRight = wrapped (fun image -> { image with Wrap = Some "square"; PositionHAlign = Some "right" })
    check "a picture at the right margin: the lines level with it end left of it" (onTheRight.Length >= 3, onTheRight |> List.forall (fun (left, right) -> near left 96.0 && right <= 96.0 + textWidth - 120.0 + 4.0)) (true, true)
do
    let lines, top, beside = wrapped (fun image -> { image with Wrap = Some "topAndBottom" })
    check "top and bottom: no line is level with the picture" beside.Length 0
    check "the text goes on below it" (lines |> List.exists (fun line -> line.Y >= top + 90.0 - 0.5), lines |> List.exists (fun line -> line.Y + line.Height <= top + 1.0)) (true, true)
    let _, topWithDistance, _ = wrapped (fun image -> { image with Wrap = Some "topAndBottom"; DistanceBottom = Some 30.0 })
    let linesWithDistance, _, _ = wrapped (fun image -> { image with Wrap = Some "topAndBottom"; DistanceBottom = Some 30.0 })
    check "with a distance below: the next line that much further down" (linesWithDistance |> List.filter (fun line -> line.Y > topWithDistance) |> List.forall (fun line -> line.Y >= topWithDistance + 90.0 + 30.0 - 0.5)) true
do
    let lines, top, beside = wrapped (fun image -> { image with Wrap = Some "none" })
    check "no wrapping (in front of the text): the lines are where they are without the picture" (beside.Length >= 3, beside |> List.forall (fun (left, _) -> near left 96.0)) (true, true)
    let _, _, behind = wrapped (fun image -> { image with Wrap = Some "square"; BehindDoc = Some true })
    check "behind the text: the same" (behind |> List.forall (fun (left, _) -> near left 96.0)) true
do
    // An outline that is the picture's full width at its top and a point at its bottom left.
    let triangle = [ (0.0, 0.0); (21600.0, 0.0); (0.0, 21600.0) ]
    let _, top, tight = wrapped (fun image -> { image with Wrap = Some "tight"; WrapPolygon = triangle })
    check "tight, with an outline: the lines level with the picture start right of the outline there - further left the further down" (tight.Length >= 3, tight |> List.pairwise |> List.forall (fun ((a, _), (b, _)) -> b <= a + 0.5), fst (List.head tight) > fst (List.last tight) + 20.0) (true, true, true)
    check "the first of them clear of the outline's wide top, the last close to the margin" (fst (List.head tight) >= 96.0 + 120.0 * 0.6, fst (List.last tight) < 96.0 + 120.0 * 0.5) (true, true)
    let _, _, through = wrapped (fun image -> { image with Wrap = Some "through"; WrapPolygon = triangle })
    check "through: the same outline" (through |> List.map fst) (tight |> List.map fst)
    let _, _, boxed = wrapped (fun image -> { image with Wrap = Some "tight" })
    check "tight without an outline: the picture's box" (boxed |> List.forall (fun (left, _) -> left >= 96.0 + 120.0 - 0.5)) true
    let _, _, tightDistance = wrapped (fun image -> { image with Wrap = Some "tight"; WrapPolygon = triangle; DistanceRight = Some 20.0 })
    check "an outline and a distance from text: that much further right, line by line" (List.zip tight tightDistance |> List.forall (fun ((a, _), (b, _)) -> b >= a + 20.0 - 6.5)) true
    // The same from an incremental layout (the picture's paragraph is laid out twice when its obstacle reaches up).
    let doc = { empty with Blocks = [ Paragraph (para "before" (words 40)); Paragraph (withPicture "w" (words 220) 0 (fun image -> { image with Wrap = Some "tight"; WrapPolygon = triangle; PositionHRelativeFrom = Some "margin"; PositionX = Some 0.0; PositionVRelativeFrom = Some "paragraph"; PositionY = Some 30.0 })); Paragraph (para "after" (words 40)) ]; Selection = None }
    let state = createIncrementalLayoutState ()
    check "the same incrementally" (layoutDocumentPagesIncremental state settings measureText measureRange doc = layoutDocumentPages settings measureText measureRange doc) true
    let edited = { doc with Blocks = doc.Blocks |> mapParagraphsOfBlocks (fun p -> if p.Id = "before" then para "before" (words 90) else p) }
    check "and after an edit above it" (layoutDocumentPagesIncremental state settings measureText measureRange edited = layoutDocumentPages settings measureText measureRange edited) true

printfn "word anchored-objects: %d passed, %d failed" pass fail
if fail > 0 then exit 1
