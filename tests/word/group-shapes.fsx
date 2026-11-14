// Groups and drawing canvases (wpg:wgp, wpc:wpc, v:group): one shape whose
// shapes (TextBoxBlock.Children) are drawn where the group's coordinates put
// them - in pixels from the group's top-left corner at the size it was read
// at (ChildSpace), scaled with the group when it is drawn at another size; a
// group in a group scales its own shapes from its own; a flipped group
// mirrors its shapes' places and flips them; a canvas (no ChildSpace) keeps
// its shapes where they are, over its background. (They were not drawn: the
// group was an empty shape.)
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

let settings : LayoutSettings =
    { Width = 816.0; PageHeight = 1056.0; FontSize = 16.0; LineHeight = 24.0; PaddingLeft = 96.0; PaddingTop = 96.0
      MarginRight = 96.0; MarginBottom = 96.0; InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 3.5; DefaultTabStop = 48.0; CompatibilityMode = 12 }
let measureText (text: string) (style: TB.TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = rangeMeasurerFromTextMeasurer measureText
let para id (text: string) : ParagraphBlock = { Id = id; Buffer = TB.ofString text; Metadata = { defaultMetadata with SpacingBefore = Some 0.0; SpacingAfter = Some 0.0 } }
let pageSetup = { defaultPageSetup with PageWidth = 816.0; PageHeight = 1056.0; MarginLeft = 96.0; MarginRight = 96.0; MarginTop = 96.0; MarginBottom = 96.0 }
let docOf (blocks: Block list) : Document = { empty with Blocks = blocks; PageSetup = pageSetup; Styles = []; Selection = None }

/// A shape of a group: its place in the group, its size, its fill.
let child (id: string) (x: float) (y: float) (width: float) (height: float) (fill: string) : TextBoxBlock =
    { defaultTextBox with Id = id; X = Some x; Y = Some y; Width = width; Height = height; FillColor = Some fill; BorderColor = None }
/// A group anchored in paragraph "a" at the column's left edge and the paragraph's top.
let group (shapes: TextBoxBlock list) (change: TextBoxBlock -> TextBoxBlock) : TextBoxBlock =
    change { defaultTextBox with Id = "group"; Width = 200.0; Height = 100.0; FillColor = None; BorderColor = None
                                 AnchorBlockId = Some "a"; Wrap = Some "none"; X = Some 0.0; Y = Some 0.0; Children = shapes; ChildSpace = Some (200.0, 100.0) }

/// The group painted: where its box is (left, top), and the commands.
let painted (shape: TextBoxBlock) =
    let doc = docOf [ Paragraph (para "a" "anchor text"); TextBox shape ]
    let layouts, pages = layoutDocumentPages settings measureText measureRange doc
    let rects = Paint.bodyShapeRects settings measureText measureRange doc layouts pages
    let _, commands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange doc layouts pages
    let x, y = match rects with [ (_, x, y, _, _) ] -> x, y | _ -> nan, nan
    (x, y), Paint.flattenedCommands commands, commands
/// Where a fill of a colour is drawn, from the group's top-left corner.
let fillAt (colour: string) (shape: TextBoxBlock) =
    let (x, y), flat, _ = painted shape
    flat |> List.tryPick (function
        | Paint.FillRect (c, left, top, width, height) when c = colour -> Some (left - x, top - y, width, height)
        | Paint.DrawShape (Paint.FillSolid c, _, _, (left, top, width, height), _) when c = colour -> Some (left - x, top - y, width, height)
        | _ -> None)

let red = child "red" 20.0 10.0 50.0 30.0 "#FF0000"
check "a shape of a group, where the group puts it" (fillAt "#FF0000" (group [ red ] id)) (Some (20.0, 10.0, 50.0, 30.0))
check "the group drawn twice as large: its shape scales with it" (fillAt "#FF0000" (group [ red ] (fun g -> { g with Width = 400.0; Height = 200.0 }))) (Some (40.0, 20.0, 100.0, 60.0))
check "the group drawn half as wide only" (fillAt "#FF0000" (group [ red ] (fun g -> { g with Width = 100.0 }))) (Some (10.0, 10.0, 25.0, 30.0))
check "the group's place moves its shapes" (fillAt "#FF0000" (group [ red ] (fun g -> { g with X = Some 30.0; Y = Some 12.0 }))) (Some (20.0, 10.0, 50.0, 30.0))
let (groupX, groupY), _, _ = painted (group [ red ] (fun g -> { g with X = Some 30.0; Y = Some 12.0 }))
let (baseX, baseY), _, _ = painted (group [ red ] id)
check "the group moved by its offsets" (groupX - baseX, groupY - baseY) (30.0, 12.0)

// A group in a group: placed in the outer's coordinates, its own shapes in its own.
let nested =
    { child "inner" 100.0 0.0 100.0 50.0 "#000001" with FillColor = None; ChildSpace = Some (100.0, 50.0); Children = [ child "green" 50.0 0.0 50.0 50.0 "#00FF00" ] }
check "a shape of a group in a group" (fillAt "#00FF00" (group [ red; nested ] id)) (Some (150.0, 0.0, 50.0, 50.0))
check "a group in a group scales with the outer group" (fillAt "#00FF00" (group [ red; nested ] (fun g -> { g with Width = 400.0; Height = 200.0 }))) (Some (300.0, 0.0, 100.0, 100.0))
check "a group in a group drawn smaller than it was read: its shapes scale with it"
    (fillAt "#00FF00" (group [ { nested with Width = 50.0 } ] id)) (Some (125.0, 0.0, 25.0, 50.0))

// A flipped group: its shapes' places mirrored, each flipped the other way, a turn the other way round.
check "a group flipped across: the shape's place mirrored" (fillAt "#FF0000" (group [ red ] (fun g -> { g with FlipH = true }))) (Some (200.0 - 20.0 - 50.0, 10.0, 50.0, 30.0))
check "a group flipped upside down" (fillAt "#FF0000" (group [ red ] (fun g -> { g with FlipV = true }))) (Some (20.0, 100.0 - 10.0 - 30.0, 50.0, 30.0))
check "a group flipped in a group flipped: back where it was" (fillAt "#00FF00" (group [ { nested with FlipH = true } ] (fun g -> { g with FlipH = true }))) (Some (0.0 + 50.0, 0.0, 50.0, 50.0))
let flips (change: TextBoxBlock -> TextBoxBlock) =
    Paint.groupChildBoxes (group [ { red with FlipH = true; Rotation = 30.0 } ] change) 100.0 |> List.map (fun (shape, _) -> shape.FlipH, shape.FlipV, shape.Rotation)
check "the shape of a group flipped across is flipped back and turned the other way" (flips (fun g -> { g with FlipH = true })) [ (false, false, -30.0) ]
check "the shape of a group flipped both ways: both flips, the turn as it was" (flips (fun g -> { g with FlipH = true; FlipV = true })) [ (false, true, 30.0) ]
// (A triangle's path, mirrored in a group flipped across: the right angle on the other side.)
let triangle = { child "tri" 0.0 0.0 100.0 100.0 "#123456" with Geometry = GeoPreset "rtTriangle" }
let trianglePath (change: TextBoxBlock -> TextBoxBlock) =
    let (x, y), flat, _ = painted (group [ triangle ] change)
    flat |> List.tryPick (function Paint.DrawShape (Paint.FillSolid "#123456", _, _, _, segs) -> Some (segs |> List.map (function PMove (px, py) | PLine (px, py) -> Some (px - x, py - y) | _ -> None)) | _ -> None)
check "a right triangle in a group: drawn" (trianglePath id |> Option.isSome) true
check "a right triangle in a group flipped across is mirrored"
    (trianglePath (fun g -> { g with FlipH = true; Width = 100.0; ChildSpace = Some (100.0, 100.0) }) |> Option.map (List.choose id >> List.map fst >> List.sort))
    (trianglePath (fun g -> { g with Width = 100.0; ChildSpace = Some (100.0, 100.0) }) |> Option.map (List.choose id >> List.map (fun px -> 100.0 - fst px) >> List.sort))

// The text of a shape of a group is drawn in that shape.
let withText = { red with Content = [ para "red-1" "hi" ]; Width = 100.0; Height = 50.0 }
let textAt (shape: TextBoxBlock) =
    let (x, y), flat, _ = painted shape
    flat |> List.tryPick (function Paint.DrawText ("hi", _, left, baseline, _) -> Some (left - x, baseline - y) | _ -> None)
match textAt (group [ withText ] id) with
| Some (left, baseline) -> check "the text of a shape of a group: inside that shape" (left > 20.0 && left < 40.0 && baseline > 10.0 && baseline < 60.0) true
| None -> check "the text of a shape of a group: drawn" false true
match textAt (group [ withText ] (fun g -> { g with Width = 400.0; Height = 200.0 })) with
| Some (left, _) -> check "the text of a shape of a group drawn twice as large: in the shape where it is now" (left > 40.0 && left < 60.0) true
| None -> check "the text of a shape of a group drawn twice as large: drawn" false true

// A picture of a group: drawn in its box, flipped as it says.
let picture = { child "pic" 0.0 50.0 50.0 50.0 "#000002" with FillColor = None; ImageSource = Some "data:image/png;base64,AAAA" }
let (picX, picY), pictureFlat, pictureCommands = painted (group [ picture ] id)
check "a picture of a group, where the group puts it"
    (pictureFlat |> List.tryPick (function Paint.DrawImage ("data:image/png;base64,AAAA", left, top, width, height, _, _, _, _) -> Some (left - picX, top - picY, width, height) | _ -> None)) (Some (0.0, 50.0, 50.0, 50.0))
let rec mirrored (commands: Paint.PaintCommand list) =
    commands |> List.exists (function
        | Paint.DrawMirrored (_, true, _, _, _, inner) -> inner |> List.exists (function Paint.DrawImage _ -> true | _ -> false)
        | Paint.DrawGroup (_, _, _, _, _, _, inner) | Paint.DrawMirrored (_, _, _, _, _, inner) -> mirrored inner
        | _ -> false)
let _, _, flippedPictureCommands = painted (group [ { picture with FlipV = true } ] id)
check "a picture of a group flipped upside down is drawn mirrored" (mirrored flippedPictureCommands, mirrored pictureCommands) (true, false)

// A canvas: its background in its box, its shapes where they are whatever its size.
let canvas (change: TextBoxBlock -> TextBoxBlock) =
    group [ child "blue" 100.0 50.0 100.0 50.0 "#0000FF" ] (fun g -> change { g with ChildSpace = None; FillColor = Some "#FFFF00" })
check "a canvas's background fills its box" (fillAt "#FFFF00" (canvas id)) (Some (0.0, 0.0, 200.0, 100.0))
check "a shape of a canvas, where it is in it" (fillAt "#0000FF" (canvas id)) (Some (100.0, 50.0, 100.0, 50.0))
check "a canvas drawn larger: its shape is not scaled" (fillAt "#0000FF" (canvas (fun g -> { g with Width = 400.0; Height = 200.0 }))) (Some (100.0, 50.0, 100.0, 50.0))
let _, canvasFlat, _ = painted (canvas id)
let order = canvasFlat |> List.choose (function Paint.FillRect (c, _, _, _, _) when c = "#FFFF00" || c = "#0000FF" -> Some c | _ -> None)
check "the background is drawn below the canvas's shapes" order [ "#FFFF00"; "#0000FF" ]

// A turned group turns its shapes with it (one turn about the group's middle).
let _, _, turnedCommands = painted (group [ red ] (fun g -> { g with Rotation = 30.0 }))
let turned = turnedCommands |> List.tryPick (function Paint.DrawGroup (_, rotation, _, _, _, _, inner) when abs (rotation - 30.0) < 0.01 -> Some (Paint.flattenedCommands inner |> List.exists (function Paint.FillRect ("#FF0000", _, _, _, _) -> true | _ -> false)) | _ -> None)
check "a turned group: its shapes turned with it" turned (Some true)

// The group's box is the group's: the painter's rectangles list it once, as large as it is drawn.
let doc = docOf [ Paragraph (para "a" "anchor text"); TextBox (group [ red; nested ] (fun g -> { g with Width = 300.0 })) ]
let layouts, pages = layoutDocumentPages settings measureText measureRange doc
check "the group is one shape of the painter's rectangles, as large as it is drawn"
    (Paint.bodyShapeRects settings measureText measureRange doc layouts pages |> List.map (fun (shape, _, _, width, height) -> shape.Id, width, height)) [ ("group", 300.0, 100.0) ]

printfn "word group-shapes: %d passed, %d failed" pass fail
if fail > 0 then exit 1
