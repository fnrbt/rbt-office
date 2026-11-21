// Text boxes and shapes: where their anchor puts them (on the page of the
// block they are anchored in - by an offset or an alignment, relative to the
// page, the margins, the column or that paragraph), their text inside their
// insets (wps:bodyPr lIns / tIns / rIns / bIns) from the top, the middle or
// the bottom of the shape, a shape as high as its text (a:spAutoFit), the
// text clipped to the shape, a turned shape, text as large as its shape (a
// watermark), and how the body's text goes around them (square, top and
// bottom, none, behind the text) - also after an edit of the shape alone.
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"

open Rbt.Word.Document
open Rbt.Word.Layout
module TB = Rbt.Word.TextBuffer
module Paint = Rbt.Word.Paint

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
      MarginRight = 96.0; MarginBottom = 96.0; InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 3.5; DefaultTabStop = 48.0; CompatibilityMode = 12; HangingPunctuation = Rbt.Word.Layout.NoHanging; SpaceShrink = 0.0; FirstPageExtraRoom = 0.0; NoteColumns = [] }
let measureText (text: string) (style: TB.TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = rangeMeasurerFromTextMeasurer measureText
let para id (text: string) : ParagraphBlock = { Id = id; Buffer = TB.ofString text; Metadata = { defaultMetadata with SpacingBefore = Some 0.0; SpacingAfter = Some 0.0 } }
let words n = String.concat " " [ for i in 1 .. n -> sprintf "w%d" i ]
let pitch = settings.PageHeight + pageGap
let pageSetup = { defaultPageSetup with PageWidth = 816.0; PageHeight = 1056.0; MarginLeft = 96.0; MarginRight = 96.0; MarginTop = 96.0; MarginBottom = 96.0 }
let docOf (blocks: Block list) : Document = { empty with Blocks = blocks; PageSetup = pageSetup; Styles = []; Selection = None }

let box (shape: TextBoxBlock -> TextBoxBlock) : Block =
    TextBox (shape { defaultTextBox with Id = "box"; Content = [ para "box-1" "boxed text" ]; Width = 200.0; Height = 100.0; FillColor = Some "#ABCDEF"; BorderColor = None
                                         AnchorBlockId = Some "a"; Wrap = Some "none"; X = Some 0.0; Y = Some 0.0 })
// More than a page of text before the anchor's paragraph: it stands on the second page.
let filler = [ for i in 1 .. 100 -> Paragraph (para (sprintf "f%d" i) "filler line") ]
let commandsOf (doc: Document) =
    let layouts, pages = layoutDocumentPages settings measureText measureRange doc
    let _, commands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange doc layouts pages
    layouts, pages, commands
/// The shape's fill: left, top, width, height.
let fillOf (commands: Paint.PaintCommand list) =
    Paint.flattenedCommands commands |> List.tryPick (function Paint.FillRect ("#ABCDEF", x, y, w, h) -> Some (x, y, w, h) | _ -> None) |> Option.defaultValue (0.0, 0.0, 0.0, 0.0)
/// Where the shape's text is drawn: its left edge and its baseline.
let textOf (commands: Paint.PaintCommand list) =
    Paint.flattenedCommands commands |> List.tryPick (function Paint.DrawText (text, _, x, baseline, _) when text.StartsWith "boxed" -> Some (x, baseline) | _ -> None) |> Option.defaultValue (0.0, 0.0)
let placed (shape: TextBoxBlock -> TextBoxBlock) =
    let layouts, pages, commands = commandsOf (docOf (filler @ [ Paragraph (para "a" (words 30)); box shape ]))
    let anchor = layouts |> List.find (fun layout -> layout.BlockId = "a")
    let page = pages.[pageIndexForYInPages pages anchor.Y]
    let x, y, _, _ = fillOf commands
    x, y - page.Top, anchor.Y - page.Top, pageIndexForYInPages pages anchor.Y

// ---------------------------------------------------------------------------
// Where the anchor puts the shape.
// ---------------------------------------------------------------------------
let _, _, anchorTop, anchorPage = placed id
check "the anchor's paragraph is on the second page" anchorPage 1
let at shape = (let x, y, _, _ = placed shape in (x, y))
let isAt (x: float, y: float) (ex: float, ey: float) = near x ex && near y ey
check "from the column and the paragraph, by offsets" (isAt (at (fun b -> { b with X = Some 30.0; Y = Some 12.0 })) (96.0 + 30.0, anchorTop + 12.0)) true
check "from the page, by offsets" (isAt (at (fun b -> { b with PositionHRelativeFrom = Some "page"; PositionVRelativeFrom = Some "page"; X = Some 40.0; Y = Some 50.0 })) (40.0, 50.0)) true
check "from the margins, by offsets" (isAt (at (fun b -> { b with PositionHRelativeFrom = Some "margin"; PositionVRelativeFrom = Some "margin"; X = Some 10.0; Y = Some 20.0 })) (106.0, 116.0)) true
check "centred on the page" (isAt (at (fun b -> { b with PositionHRelativeFrom = Some "page"; PositionVRelativeFrom = Some "page"; PositionHAlign = Some "center"; PositionVAlign = Some "center" })) ((816.0 - 200.0) / 2.0, (1056.0 - 100.0) / 2.0)) true
check "right and bottom in the margins" (isAt (at (fun b -> { b with PositionHRelativeFrom = Some "margin"; PositionVRelativeFrom = Some "margin"; PositionHAlign = Some "right"; PositionVAlign = Some "bottom" })) (816.0 - 96.0 - 200.0, 1056.0 - 96.0 - 100.0)) true

// ---------------------------------------------------------------------------
// The text in the shape.
// ---------------------------------------------------------------------------
let inBox (shape: TextBoxBlock -> TextBoxBlock) =
    let _, _, commands = commandsOf (docOf [ Paragraph (para "a" (words 30)); box shape ])
    let x, y, w, h = fillOf commands
    let textX, baseline = textOf commands
    (x, y, w, h), (textX - x, baseline - y), commands
let _, (defaultInsetX, defaultBaseline), _ = inBox id
check "the default insets: a tenth of an inch at the left" (near defaultInsetX 9.6) true
let _, (insetX, insetBaseline), _ = inBox (fun b -> { b with InsetLeft = Some 30.0; InsetTop = Some 24.8 })
check "the left inset moves the text" (near insetX 30.0) true
check "the top inset moves it down" (near (insetBaseline - defaultBaseline) 20.0) true
let _, (_, middleBaseline), _ = inBox (fun b -> { b with VerticalAnchor = Some "ctr" })
let _, (_, bottomBaseline), _ = inBox (fun b -> { b with VerticalAnchor = Some "b" })
let lineHeight = (let layouts, _, _ = commandsOf (docOf [ Paragraph (para "a" "boxed text") ]) in layouts.Head.Height)
check "anchored in the middle" (near (middleBaseline - defaultBaseline) ((100.0 - 9.6 - lineHeight) / 2.0)) true
check "anchored at the bottom" (near (bottomBaseline - defaultBaseline) (100.0 - 9.6 - lineHeight)) true
// A shape as high as its text.
let (_, _, _, fittedHeight), _, _ = inBox (fun b -> { b with AutoFit = true; Content = [ para "box-1" "boxed text"; para "box-2" "second"; para "box-3" "third" ] })
check "as high as its text and insets" (near fittedHeight (lineHeight * 3.0 + 9.6)) true
// The text is clipped to a shape with a height of its own; not to one that follows its text.
let groupsOf (commands: Paint.PaintCommand list) =
    let rec walk (commands: Paint.PaintCommand list) =
        commands |> List.collect (function
            | Paint.DrawGroup (clip, rotation, cx, cy, opacity, bbox, inner) -> (clip, rotation, (cx, cy), opacity, bbox) :: walk inner
            | _ -> [])
    walk commands
let (bx, by, bw, bh), _, fixedCommands = inBox id
check "the text is clipped to the shape" (groupsOf fixedCommands |> List.map (fun (clip, rotation, _, _, _) -> clip, rotation)) [ Some (bx, by, bw, bh), 0.0 ]
let _, _, fittedCommands = inBox (fun b -> { b with AutoFit = true })
check "a shape that follows its text clips nothing" (groupsOf fittedCommands) []
// A turned shape: all of it turned about its centre, in a box that holds it.
let _, _, turnedCommands = inBox (fun b -> { b with Rotation = 30.0 })
match groupsOf turnedCommands with
| (None, rotation, (cx, cy), _, (gx, gy, gw, gh)) :: _ ->
    check "turned about its centre" (rotation, near cx (bx + bw / 2.0), near cy (by + bh / 2.0)) (30.0, true, true)
    let radius = sqrt (bw * bw + bh * bh) / 2.0
    check "in a box that holds the turned shape" (gx <= cx - radius, gy <= cy - radius, gx + gw >= cx + radius, gy + gh >= cy + radius) (true, true, true, true)
| other -> check "a turned shape is drawn in a group" other []
// Text as large as its shape (a watermark): one line, as wide as the shape.
let _, _, watermarkCommands = inBox (fun b -> { b with FitText = true; Content = [ para "box-1" "DRAFT" ]; Width = 400.0; Height = 300.0; Opacity = 0.5 })
match Paint.flattenedCommands watermarkCommands |> List.tryPick (function Paint.DrawText ("DRAFT", style, x, _, width) -> Some (style.FontSize, x, width) | _ -> None) with
| Some (size, x, width) ->
    check "as wide as the shape" (near width 400.0, near size 160.0) (true, true)
| None -> check "the watermark's text is drawn" false true
check "at its opacity" (groupsOf watermarkCommands |> List.map (fun (_, _, _, opacity, _) -> opacity)) [ 0.5 ]

// A line (prstGeom "line", a straight connector): from one corner of its box to the other, not a rectangle.
let _, _, lineCommands = inBox (fun b -> { b with Geometry = GeoPreset "line"; Content = []; BorderColor = Some "#000000"; Width = 120.0; Height = 40.0 })
check "a line is a line across its box"
    (lineCommands |> List.tryPick (function Paint.DrawShape (_, Some "#000000", _, (x, y, _, _), segs) -> Some (segs = [ PMove (x, y); PLine (x + 120.0, y + 40.0) ]) | _ -> None))
    (Some true)

// A flipped shape (a:xfrm flipH / flipV): its outline mirrored in its box, its text not; a linear
// gradient's direction mirrored with it. (Flips were not read: the shape was drawn as if unflipped.)
let lineSegs (shape: TextBoxBlock -> TextBoxBlock) =
    let _, _, commands = inBox (fun b -> shape { b with Geometry = GeoPreset "line"; Content = []; BorderColor = Some "#000000"; Width = 120.0; Height = 40.0 })
    commands |> List.tryPick (function Paint.DrawShape (_, Some "#000000", _, (x, y, _, _), segs) -> Some (x, y, segs) | _ -> None)
match lineSegs (fun b -> { b with FlipH = true }), lineSegs (fun b -> { b with FlipV = true }), lineSegs (fun b -> { b with FlipH = true; FlipV = true }) with
| Some (x, y, across), Some (_, _, upsideDown), Some (_, _, both) ->
    check "flipped across: the line from the top right to the bottom left" across [ PMove (x + 120.0, y); PLine (x, y + 40.0) ]
    check "flipped upside down: from the bottom left to the top right" upsideDown [ PMove (x, y + 40.0); PLine (x + 120.0, y) ]
    check "flipped both ways: from the bottom right to the top left" both [ PMove (x + 120.0, y + 40.0); PLine (x, y) ]
| other -> check "a flipped line is drawn" (sprintf "%A" other) ""
let triangleText (shape: TextBoxBlock -> TextBoxBlock) =
    let _, _, commands = inBox (fun b -> shape { b with Geometry = GeoPreset "rtTriangle"; BorderColor = Some "#000000" })
    textOf commands
check "the text of a flipped shape stands where it stands in the shape unflipped" (triangleText (fun b -> { b with FlipH = true; FlipV = true })) (triangleText id)
let gradientAngle (shape: TextBoxBlock -> TextBoxBlock) =
    let _, _, commands = inBox (fun b -> shape { b with Gradient = Some { Stops = [ { Offset = 0.0; Color = "#FF0000" }; { Offset = 1.0; Color = "#0000FF" } ]; Angle = 30.0; Radial = false } })
    commands |> List.tryPick (function Paint.DrawShape (Paint.FillGradient g, _, _, _, _) -> Some g.Angle | _ -> None)
check "a linear gradient flipped across / upside down / both" (gradientAngle (fun b -> { b with FlipH = true }), gradientAngle (fun b -> { b with FlipV = true }), gradientAngle (fun b -> { b with FlipH = true; FlipV = true })) (Some 150.0, Some 330.0, Some 210.0)

// A flipped picture: drawn mirrored about its middle (the picture itself is all the group holds).
let pictureCommands (flipH: bool) (flipV: bool) =
    let buffer = TB.insertImage 0 "data:image/png;base64,AA==" (Some 40.0) (Some 20.0) None (TB.ofString "")
    let buffer = { buffer with Runs = buffer.Runs |> List.map (function TB.Image image -> TB.Image { image with FlipH = flipH; FlipV = flipV } | run -> run) }
    let _, _, commands = commandsOf (docOf [ Paragraph { para "pic" "" with Buffer = buffer } ])
    commands
match pictureCommands true false |> List.tryPick (function Paint.DrawMirrored (h, v, cx, cy, _, inner) -> Some (h, v, cx, cy, inner) | _ -> None) with
| Some (h, v, cx, cy, [ Paint.DrawImage (_, x, y, w, ht, _, _, _, _) ]) ->
    check "a picture flipped across: mirrored about its middle" (h, v, near cx (x + w / 2.0), near cy (y + ht / 2.0), w, ht) (true, false, true, true, 40.0, 20.0)
| other -> check "a flipped picture is drawn mirrored" (sprintf "%A" other) ""
check "a picture flipped upside down" (pictureCommands false true |> List.exists (function Paint.DrawMirrored (false, true, _, _, _, _) -> true | _ -> false)) true
check "a picture not flipped is drawn as it is" (pictureCommands false false |> List.exists (function Paint.DrawMirrored _ -> true | _ -> false), pictureCommands false false |> List.exists (function Paint.DrawImage _ -> true | _ -> false)) (false, true)

// ---------------------------------------------------------------------------
// The body's text around the shape.
// ---------------------------------------------------------------------------
let linesOf (layouts: BlockLayout list) (id: string) = layouts |> List.filter (fun layout -> layout.BlockId = id) |> List.collect (fun layout -> layout.Lines)
let startOf (line: MeasuredLine) = line.Segments |> List.choose (function TextSegment (_, _, x, _, _) -> Some x | _ -> None) |> List.min
let around (shape: TextBoxBlock -> TextBoxBlock) =
    let layouts, _ = layoutDocumentPages settings measureText measureRange (docOf [ Paragraph (para "before" (words 20)); Paragraph (para "a" (words 200)); Paragraph (para "after" (words 20)); box shape ])
    let anchor = layouts |> List.find (fun layout -> layout.BlockId = "a")
    anchor, linesOf layouts "a"
let plainAnchor, plainLines = around id
check "no wrapping: the text as without the shape" (plainLines |> List.forall (fun line -> startOf line < 1.0)) true
let squareAnchor, squareLines = around (fun b -> { b with Wrap = Some "square"; Y = Some 10.0 })
let beside = squareLines |> List.filter (fun line -> line.Y + line.Height > squareAnchor.Y + 10.0 && line.Y < squareAnchor.Y + 110.0)
check "square: the lines beside the shape start to its right" (beside.Length >= 3, beside |> List.forall (fun line -> startOf line >= 200.0 - 0.5)) (true, true)
check "square: the lines below it start at the margin" (squareLines |> List.filter (fun line -> line.Y >= squareAnchor.Y + 111.0) |> List.forall (fun line -> startOf line < 1.0)) true
let _, distanceLines = around (fun b -> { b with Wrap = Some "square"; Y = Some 10.0; DistanceRight = Some 24.0 })
check "square: with its distance from text" (distanceLines |> List.exists (fun line -> near (startOf line) 224.0 || startOf line >= 224.0 - 0.5 && startOf line < 260.0)) true
let topBottomAnchor, topBottomLines = around (fun b -> { b with Wrap = Some "topAndBottom"; Y = Some 0.0 })
check "top and bottom: no line beside the shape" (topBottomLines |> List.forall (fun line -> line.Y + line.Height <= topBottomAnchor.Y + 0.5 || line.Y >= topBottomAnchor.Y + 100.0 - 0.5)) true
let _, behindLines = around (fun b -> { b with Wrap = Some "square"; BehindDoc = Some true })
check "behind the text: the text as without the shape" (behindLines |> List.map startOf) (plainLines |> List.map startOf)
// The same from an incremental layout, and after an edit of the shape alone (it is the last block; its anchor is not).
let wrapDoc (width: float) = docOf [ Paragraph (para "before" (words 20)); Paragraph (para "a" (words 200)); Paragraph (para "after" (words 20)); box (fun b -> { b with Wrap = Some "square"; Width = width }) ]
let state = createIncrementalLayoutState ()
let first = wrapDoc 200.0
check "the same incrementally" (layoutDocumentPagesIncremental state settings measureText measureRange first = layoutDocumentPages settings measureText measureRange first) true
let widened = { first with Blocks = first.Blocks |> List.map (function TextBox b -> TextBox { b with Width = 320.0 } | block -> block) }
check "a wider shape: its anchor's paragraph is laid out again" (layoutDocumentPagesIncremental state settings measureText measureRange widened = layoutDocumentPages settings measureText measureRange widened) true
check "and its lines start further right" ((layoutDocumentPagesIncremental state settings measureText measureRange widened |> fst |> fun layouts -> linesOf layouts "a") |> List.exists (fun line -> startOf line >= 320.0 - 0.5)) true
let removed = { widened with Blocks = widened.Blocks |> List.filter (function TextBox _ -> false | _ -> true) }
check "the shape removed: the text takes its place" (layoutDocumentPagesIncremental state settings measureText measureRange removed = layoutDocumentPages settings measureText measureRange removed) true

// Moving through the body does not go into a box: the text of a box is a story of its own.
// (A box in the middle of the blocks and one after the last paragraph, as a file read has them.)
let navDoc =
    docOf [ Paragraph (para "n1" "one"); box (fun b -> { b with Id = "mid"; Content = [ para "mid-1" "middle box" ]; AnchorBlockId = Some "n1" })
            Paragraph (para "n2" "two"); box (fun b -> { b with Id = "last"; Content = [ para "last-1" "last box"; para "last-2" "its second" ]; AnchorBlockId = Some "n2" }) ]
let caretAt id offset doc = setSelection { BlockId = id; Offset = offset } { BlockId = id; Offset = offset } doc
let focusOf (doc: Document) = doc.Selection |> Option.map (fun s -> s.Focus.BlockId, s.Focus.Offset)
let everything = (selectAll navDoc).Selection.Value
check "select everything: from the start of the body to the end of its last paragraph" ((everything.Anchor.BlockId, everything.Anchor.Offset), (everything.Focus.BlockId, everything.Focus.Offset)) (("n1", 0), ("n2", 3))
check "right at the end of a paragraph: the next paragraph of the body, over the box" (navDoc |> caretAt "n1" 3 |> moveCursorRight |> focusOf) (Some ("n2", 0))
check "right at the end of the body stays there" (navDoc |> caretAt "n2" 3 |> moveCursorRight |> focusOf) (Some ("n2", 3))
check "left at the start of a paragraph: the one before it in the body" (navDoc |> caretAt "n2" 0 |> moveCursorLeft |> focusOf) (Some ("n1", 3))
check "extending right over the end of a paragraph" ((navDoc |> caretAt "n1" 3 |> extendSelectionRight).Selection |> Option.map (fun s -> s.Focus.BlockId)) (Some "n2")
let _, _, selectedIds = getNormalizedRange everything navDoc
check "the range of everything is the body's paragraphs" selectedIds [ "n1"; "n2" ]
check "and so are the paragraphs a selection touches" (selectedParagraphRanges (selectAll navDoc) |> List.map (fun (p, a, b) -> p.Id, a, b)) [ "n1", 0, 3; "n2", 0, 3 ]
let boldAll = selectAll navDoc |> toggleBold
let boldOf id = (paragraphsOfBlocks boldAll.Blocks |> List.find (fun p -> p.Id = id)).Buffer.Runs |> List.exists (function TB.Text run -> run.Style.Bold | _ -> false)
check "everything made bold: the body, not the text of the boxes" (boldOf "n1", boldOf "n2", boldOf "mid-1", boldOf "last-1") (true, true, false, false)
// (A position a host sets in a box that is not entered moves within the box.)
check "in a box: right at the end of its paragraph goes to its next, and no further" (navDoc |> caretAt "last-1" 8 |> moveCursorRight |> focusOf, navDoc |> caretAt "last-2" 10 |> moveCursorRight |> focusOf, navDoc |> caretAt "mid-1" 10 |> moveCursorRight |> focusOf) (Some ("last-2", 0), Some ("last-2", 10), Some ("mid-1", 10))
check "the box of a paragraph" (textBoxOfParagraph navDoc "last-2" |> Option.map (fun b -> b.Id), textBoxOfParagraph navDoc "n2" |> Option.map (fun b -> b.Id)) (Some "last", None)

// A selection of the body that is deleted (or typed or pasted over) is made of the body's
// paragraphs. A text box stands among the blocks next to the paragraph it is anchored in: it goes
// with the start of that paragraph (where its anchor is), and stays when that start is not selected.
do
    let anchored id anchor : Block = TextBox { defaultTextBox with Id = id; Content = [ para (id + "-1") ("text of " + id) ]; Width = 100.0; Height = 50.0; AnchorBlockId = Some anchor; Wrap = Some "none" }
    let body : Document = docOf [ Paragraph (para "p1" "first paragraph"); anchored "b1" "p1"; Paragraph (para "p2" "second paragraph"); anchored "b2" "p2"; Paragraph (para "p3" "third paragraph"); anchored "b3" "p3"; Paragraph (para "p4" "fourth") ]
    let select a ao f fo (doc: Document) = { doc with Selection = Some { Anchor = { BlockId = a; Offset = ao }; Focus = { BlockId = f; Offset = fo } } }
    let boxesOf (doc: Document) = doc.Blocks |> List.choose (function TextBox b -> Some (b.Id, b.AnchorBlockId |> Option.defaultValue "") | _ -> None)
    let textsOf (doc: Document) = doc.Blocks |> List.choose (function Paragraph p -> Some (TB.toString p.Buffer) | _ -> None)
    let midToMid = body |> select "p1" 6 "p3" 6 |> deleteSelection
    check "from the middle of the first to the middle of the third paragraph: the first's box stays, the second's and the third's go" (boxesOf midToMid, textsOf midToMid) ([ ("b1", "p1") ], [ "first paragraph"; "fourth" ])
    let fromStart = body |> select "p1" 0 "p2" 6 |> deleteSelection
    check "from the start of the first paragraph: its box goes with it" (boxesOf fromStart, textsOf fromStart) ([ ("b3", "p3") ], [ " paragraph"; "third paragraph"; "fourth" ])
    let toStart = body |> select "p1" 6 "p3" 0 |> deleteSelection
    check "to the start of the third paragraph: its box stays, anchored in the paragraph it is part of now" (boxesOf toStart, textsOf toStart) ([ ("b1", "p1"); ("b3", "p1") ], [ "first third paragraph"; "fourth" ])
    let typedOver = body |> select "p2" 7 "p3" 5 |> insertText "X"
    check "typed over a selection: the box of the paragraph it starts in stays, its text untouched" (boxesOf typedOver, textsOf typedOver, typedOver.Blocks |> List.pick (function TextBox b when b.Id = "b2" -> Some (TB.toString b.Content.Head.Buffer) | _ -> None)) ([ ("b1", "p1"); ("b2", "p2") ], [ "first paragraph"; "second X paragraph"; "fourth" ], "text of b2")
    // The same selections deleted with changes tracked: the body's text is marked as deleted and stays, and so
    // does every box, its text untouched - a range of the body is made of the body's paragraphs (nothing of a
    // box is deleted without a mark of it).
    let deletedIn (doc: Document) = doc.Blocks |> List.choose (function TextBox b when b.Content |> List.exists (fun p -> p.Buffer.Runs |> List.exists (function TB.Text run -> run.Style.RevisionType.IsSome | _ -> false)) -> Some b.Id | _ -> None)
    let markedIn (doc: Document) = doc.Blocks |> List.choose (function Paragraph p when p.Buffer.Runs |> List.exists (function TB.Text run -> run.Style.RevisionType = Some "delete" | _ -> false) -> Some p.Id | _ -> None)
    let trackedMid = body |> select "p1" 6 "p3" 6 |> trackDeleteSelection "A" "2024-01-01T00:00:00Z"
    check "tracked, from the middle of the first to the middle of the third paragraph: the text is marked, every box stays as it is" (markedIn trackedMid, boxesOf trackedMid |> List.map fst, deletedIn trackedMid) ([ "p1"; "p2"; "p3" ], [ "b1"; "b2"; "b3" ], [])
    let trackedTyped = body |> select "p2" 7 "p3" 5 |> insertTrackedText "A" "2024-01-01T00:00:00Z" "X"
    check "tracked, typed over a selection: every box stays as it is" (boxesOf trackedTyped |> List.map fst, deletedIn trackedTyped) ([ "b1"; "b2"; "b3" ], [])
    // A shape moved, resized and taken out (Word.moveShape, resizeShape, deleteShape).
    let boxOf id (doc: Document) = doc.Blocks |> List.pick (function TextBox b when b.Id = id -> Some b | _ -> None)
    let aligned = { body with Blocks = body.Blocks |> List.map (function TextBox b when b.Id = "b1" -> TextBox { b with X = Some 10.0; Y = Some 20.0; PositionHAlign = Some "center"; PositionHRelativeFrom = Some "margin"; PositionVRelativeFrom = Some "paragraph" } | block -> block) }
    let movedBox = moveTextBox "b1" (Some (Some "page", 300.0)) (Some (None, 45.0)) aligned |> Option.map (boxOf "b1")
    check "a shape moved: by the offsets given, from the frame given (the one it had where none is), aligned no more" (movedBox |> Option.map (fun b -> b.X, b.Y, b.PositionHAlign, b.PositionHRelativeFrom, b.PositionVRelativeFrom)) (Some (Some 300.0, Some 45.0, None, Some "page", Some "paragraph"))
    check "a shape resized" (resizeTextBox "b2" 222.0 111.0 body |> Option.map (fun doc -> (boxOf "b2" doc).Width, (boxOf "b2" doc).Height, boxesOf doc |> List.length)) (Some (222.0, 111.0, 3))
    check "to the size it has: the document it was" (resizeTextBox "b2" 100.0 50.0 body |> Option.map (fun doc -> obj.ReferenceEquals(doc, body))) (Some true)
    let withCaretInBox = body |> select "b2-1" 3 "b2-1" 3
    let removed = deleteTextBox "b2" withCaretInBox
    check "a shape taken out: its text with it, the caret in the block it was anchored in" (removed |> Option.map (fun doc -> boxesOf doc |> List.map fst, textsOf doc |> List.length, doc.Selection |> Option.map (fun sel -> sel.Focus.BlockId, sel.Focus.Offset))) (Some ([ "b1"; "b3" ], 4, Some ("p2", 0)))
    check "no shape of that id" ((moveTextBox "none" None None body).IsNone, (resizeTextBox "none" 1.0 1.0 body).IsNone, (deleteTextBox "none" body).IsNone) (true, true, true)
    // A shape of a header: moved, resized and taken out in its part; the body and the footer are the objects they were.
    let headerBox : Block = TextBox { defaultTextBox with Id = "h1"; Content = [ para "h1-1" "in the header" ]; Width = 100.0; Height = 50.0; X = Some 10.0; Y = Some 5.0; Wrap = Some "none" }
    let withHeader =
        { body with
            HeaderParts = [ { Kind = "default"; RelationshipId = None; Blocks = [ Paragraph (para "hp" "header"); headerBox ] } ]
            FooterParts = [ { Kind = "default"; RelationshipId = None; Blocks = [ Paragraph (para "fp" "footer") ] } ] }
    let headerBoxOf (doc: Document) = doc.HeaderParts.Head.Blocks |> List.tryPick (function TextBox b when b.Id = "h1" -> Some (b.X, b.Y, b.Width, b.Height) | _ -> None)
    let movedHeader = moveTextBox "h1" (Some (None, 22.0)) (Some (None, 12.0)) withHeader
    check "a header's shape moved: in its part" (movedHeader |> Option.map (fun doc -> headerBoxOf doc, obj.ReferenceEquals(doc.Blocks, withHeader.Blocks), obj.ReferenceEquals(doc.FooterParts.Head, withHeader.FooterParts.Head))) (Some (Some (Some 22.0, Some 12.0, 100.0, 50.0), true, true))
    check "a header's shape resized" (resizeTextBox "h1" 120.0 60.0 withHeader |> Option.bind headerBoxOf) (Some (Some 10.0, Some 5.0, 120.0, 60.0))
    check "a header's shape taken out" (deleteTextBox "h1" withHeader |> Option.map (fun doc -> headerBoxOf doc, doc.HeaderParts.Head.Blocks.Length, obj.ReferenceEquals(doc.Blocks, withHeader.Blocks))) (Some (None, 1, true))
    check "to the place it has: the document it was" (moveTextBox "h1" None None withHeader |> Option.map (fun doc -> obj.ReferenceEquals(doc, withHeader))) (Some true)
    // A shape resized by a handle: the side (or corner) across stays where it is - also when the shape is turned.
    let rounded (x: float, y: float, w: float, h: float) = (System.Math.Round(x, 2), System.Math.Round(y, 2), System.Math.Round(w, 2), System.Math.Round(h, 2))
    check "the east handle: wider, where it was" (resizedShapeRect 0.0 10.0 20.0 100.0 50.0 "e" 15.0 7.0 |> rounded) (10.0, 20.0, 115.0, 50.0)
    check "the north-west handle: left and top go, right and bottom stay" (resizedShapeRect 0.0 10.0 20.0 100.0 50.0 "nw" -10.0 -6.0 |> rounded) (0.0, 14.0, 110.0, 56.0)
    check "the south handle" (resizedShapeRect 0.0 10.0 20.0 100.0 50.0 "s" 9.0 12.0 |> rounded) (10.0, 20.0, 100.0, 62.0)
    check "not narrower than four pixels: the right edge still stays" (resizedShapeRect 0.0 10.0 20.0 100.0 50.0 "w" 500.0 0.0 |> rounded) (106.0, 20.0, 4.0, 50.0)
    // (Turned a quarter clockwise the shape's east side points down the page: dragged ten pixels down it is ten wider, and
    // its west side - at the top, 50 above the centre (60, 45) - is where it was: the centre is five further down.)
    check "turned a quarter: the east handle dragged down the page" (resizedShapeRect 90.0 10.0 20.0 100.0 50.0 "e" 0.0 10.0 |> rounded) (5.0, 25.0, 110.0, 50.0)
    check "turned a quarter: a drag across the handle's direction changes nothing" (resizedShapeRect 90.0 10.0 20.0 100.0 50.0 "e" 10.0 0.0 |> rounded) (10.0, 20.0, 100.0, 50.0)
    check "place and size in one change" (placeTextBox "b1" (Some (Some "page", 30.0)) None (Some (120.0, 60.0)) body |> Option.map (fun doc -> (boxOf "b1" doc).X, (boxOf "b1" doc).Width, (boxOf "b1" doc).Height)) (Some (Some 30.0, 120.0, 60.0))
    // Everything selected and deleted: the body's paragraphs and what is anchored in them.
    let all = body |> selectAll |> deleteSelection
    check "everything selected and deleted: nothing is left" (boxesOf all, textsOf all) ([], [ "" ])
    // Find and replace go by all the text: the body's and the boxes'.
    let replaced, count = replaceAllText "paragraph" "PARA" true (docOf [ Paragraph (para "p1" "first paragraph"); anchored "b1" "p1" ] |> fun doc -> { doc with Blocks = doc.Blocks |> List.map (function TextBox b -> TextBox { b with Content = [ para "b1-1" "a paragraph in a box" ] } | block -> block) })
    check "replace all: in the body and in the text of a box" (count, textsOf replaced, replaced.Blocks |> List.pick (function TextBox b -> Some (TB.toString b.Content.Head.Buffer) | _ -> None)) (2, [ "first PARA" ], "a PARA in a box")
    check "find: the same" ((findText "paragraph" true (docOf [ Paragraph (para "p1" "first paragraph"); anchored "b1" "p1" ])) |> List.map (fun m -> m.BlockId)) [ "p1" ]

// ---------------------------------------------------------------------------
// allowOverlap="0": a shape that may not overlap is moved down below a shape
// placed before it that it would overlap and that may not overlap either, as
// Word moves it - where it is painted, where the painter says it is, and the
// text goes around it there. (The flag was read and not used: the shapes were
// drawn over each other, and the text went around them where they were.)
// ---------------------------------------------------------------------------
let overlapping (first: TextBoxBlock -> TextBoxBlock) (second: TextBoxBlock -> TextBoxBlock) =
    let shape (shapeId: string) (x: float) (y: float) (colour: string) (change: TextBoxBlock -> TextBoxBlock) =
        let made = { defaultTextBox with Id = shapeId; Content = [ para (shapeId + "-1") shapeId ]; Width = 200.0; Height = 100.0; FillColor = Some colour; BorderColor = None; AnchorBlockId = Some "a"; Wrap = Some "square"; X = Some x; Y = Some y }
        TextBox (change made)
    let doc = docOf [ Paragraph (para "before" (words 20)); Paragraph (para "a" (words 300)); Paragraph (para "after" (words 20)); shape "o1" 0.0 10.0 "#ABCDEF" first; shape "o2" 100.0 40.0 "#FEDCBA" second ]
    let layouts, pages = layoutDocumentPages settings measureText measureRange doc
    let anchor = layouts |> List.find (fun layout -> layout.BlockId = "a")
    let rects = Paint.bodyShapeRects settings measureText measureRange doc layouts pages |> List.map (fun (textBox, x, y, _, _) -> textBox.Id, (x, System.Math.Round(y - anchor.Y, 6)))
    let _, commands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange doc layouts pages
    let painted colour = Paint.flattenedCommands commands |> List.tryPick (function Paint.FillRect (c, x, y, _, _) when c = colour -> Some (x, System.Math.Round(y - anchor.Y, 6)) | _ -> None)
    // The text of the lines between `top` and `bottom` (from the top of the anchor's paragraph): their segments, across.
    let textBetween top bottom =
        linesOf layouts "a"
        |> List.filter (fun line -> line.Y + line.Height > anchor.Y + top && line.Y < anchor.Y + bottom)
        |> List.collect (fun line -> line.Segments |> List.choose (function TextSegment (text, _, x, width, _) when text.Trim() <> "" -> Some (x, x + width) | _ -> None))
    // (Rounded: the anchor's paragraph stands at fractions of a pixel.)
    rects, painted "#ABCDEF", painted "#FEDCBA", textBetween
let keptClear = fun (b: TextBoxBlock) -> { b with AllowOverlap = Some false }
let rects, firstFill, secondFill, textBetween = overlapping keptClear keptClear
check "allowOverlap=0 on both: the first stays where its anchor puts it" (List.tryFind (fun (id, _) -> id = "o1") rects |> Option.map snd) (Some (96.0, 10.0))
check "allowOverlap=0 on both: the second is moved down below the first, across where it was" (List.tryFind (fun (id, _) -> id = "o2") rects |> Option.map snd) (Some (196.0, 110.0))
check "painted where the painter says they are" (firstFill, secondFill) (Some (96.0, 10.0), Some (196.0, 110.0))
let band = textBetween 120.0 200.0
check "the text goes around the second where it was moved to" (band.Length > 0, band |> List.forall (fun (left, right) -> right <= 100.5 || left >= 299.5)) (true, true)
let notMoved (first: TextBoxBlock -> TextBoxBlock) (second: TextBoxBlock -> TextBoxBlock) =
    let rects, _, fill, _ = overlapping first second
    List.tryFind (fun (id, _) -> id = "o2") rects |> Option.map snd, fill
check "allowOverlap=0 on the second alone: it overlaps the first" (notMoved id keptClear) (Some (196.0, 40.0), Some (196.0, 40.0))
check "allowOverlap=0 on the first alone: the second overlaps it" (notMoved keptClear (fun b -> { b with AllowOverlap = Some true })) (Some (196.0, 40.0), Some (196.0, 40.0))
// (Shapes that do not overlap stay where they are.)
let apart = overlapping keptClear (fun b -> { b with AllowOverlap = Some false; X = Some 250.0 }) |> fun (rects, _, _, _) -> List.tryFind (fun (id, _) -> id = "o2") rects |> Option.map snd
check "allowOverlap=0 on both, apart: not moved" apart (Some (346.0, 40.0))

printfn "word text-boxes: %d passed, %d failed" pass fail
if fail > 0 then exit 1
