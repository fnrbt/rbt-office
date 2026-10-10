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
      MarginRight = 96.0; MarginBottom = 96.0; InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 3.5 }
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

printfn "word text-boxes: %d passed, %d failed" pass fail
if fail > 0 then exit 1
