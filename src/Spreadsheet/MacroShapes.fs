/// The shapes a macro draws (Shapes.AddShape, AddLine, AddTextbox, Duplicate, Shapes.Range(...).Group):
/// a tree in points, as Excel's object model sees them, written as the drawing XML (`xdr:sp`, `xdr:cxnSp`,
/// `xdr:grpSp`) that the grid draws and a save keeps.
///
/// A group's frame is the box round its members. Rotating or scaling a group moves and turns its members
/// (Excel keeps a rotation on the group; drawn, it is the same), so the XML has no rotation on a group.
module Rbt.Spreadsheet.MacroShapes

open System

/// A shape: "autoshape", "textbox", "line" or "group" (Members). Positions and sizes in points from the
/// sheet's top left; colours as VBA's RGB numbers (blue * 65536 + green * 256 + red).
type ShapeNode(id: int, name: string, kind: string) =
    member val Id = id with get, set
    member val Name = name with get, set
    member _.Kind = kind
    /// The DrawingML preset ("rect", "ellipse", "pentagon"...) and the MsoAutoShapeType it came from.
    member val Preset = "rect" with get, set
    member val AutoShapeType = 1 with get, set
    member val Left = 0.0 with get, set
    member val Top = 0.0 with get, set
    member val Width = 0.0 with get, set
    member val Height = 0.0 with get, set
    member val Rotation = 0.0 with get, set
    member val FlipH = false with get, set
    member val FlipV = false with get, set
    member val Visible = true with get, set
    member val FillColor = 0xC47244 with get, set
    member val FillVisible = true with get, set
    member val FillTransparency = 0.0 with get, set
    member val LineColor = 0x8F522F with get, set
    member val LineWeight = 1.0 with get, set
    member val LineVisible = true with get, set
    member val Text = "" with get, set
    member val TextColor = 0xFFFFFF with get, set
    member val TextSize = 11.0 with get, set
    member val TextBold = false with get, set
    member val TextItalic = false with get, set
    /// "l" | "ctr" | "r" and "t" | "ctr" | "b".
    member val TextAlign = "ctr" with get, set
    member val TextAnchor = "ctr" with get, set
    member val OnAction = "" with get, set
    member val Members = ResizeArray<ShapeNode>() with get
    member val Parent : ShapeNode option = None with get, set

/// MsoAutoShapeType: the preset it draws and the name Excel gives a new one ("Oval 3").
let presetOf (autoShapeType: int) : string * string =
    match autoShapeType with
    | 1 -> "rect", "Rectangle"
    | 2 -> "parallelogram", "Parallelogram"
    | 3 -> "trapezoid", "Trapezoid"
    | 4 -> "diamond", "Diamond"
    | 5 -> "roundRect", "Rectangle: Rounded Corners"
    | 6 -> "octagon", "Octagon"
    | 7 -> "triangle", "Isosceles Triangle"
    | 8 -> "rtTriangle", "Right Triangle"
    | 9 -> "ellipse", "Oval"
    | 10 -> "hexagon", "Hexagon"
    | 11 -> "plus", "Cross"
    | 12 -> "pentagon", "Regular Pentagon"
    | 13 -> "can", "Cylinder"
    | 14 -> "cube", "Cube"
    | 15 -> "bevel", "Bevel"
    | 16 -> "foldedCorner", "Folded Corner"
    | 17 -> "smileyFace", "Smiley Face"
    | 18 -> "donut", "Donut"
    | 19 -> "noSmoking", "No Symbol"
    | 20 -> "blockArc", "Block Arc"
    | 21 -> "heart", "Heart"
    | 22 -> "lightningBolt", "Lightning Bolt"
    | 23 -> "sun", "Sun"
    | 24 -> "moon", "Moon"
    | 25 -> "arc", "Arc"
    | 33 -> "rightArrow", "Arrow: Right"
    | 34 -> "leftArrow", "Arrow: Left"
    | 35 -> "upArrow", "Arrow: Up"
    | 36 -> "downArrow", "Arrow: Down"
    | 37 -> "leftRightArrow", "Arrow: Left-Right"
    | 38 -> "upDownArrow", "Arrow: Up-Down"
    | 51 -> "homePlate", "Arrow: Pentagon"
    | 52 -> "chevron", "Arrow: Chevron"
    | 61 -> "flowChartProcess", "Flowchart: Process"
    | 91 -> "star4", "Star: 4 Points"
    | 92 -> "star5", "Star: 5 Points"
    | 93 -> "star8", "Star: 8 Points"
    | 94 -> "star16", "Star: 16 Points"
    | 95 -> "star24", "Star: 24 Points"
    | 96 -> "star32", "Star: 32 Points"
    | _ -> "rect", "Rectangle"

/// The MsoAutoShapeType constants a macro names.
let constants : (string * int) list =
    [ "msoShapeRectangle", 1; "msoShapeParallelogram", 2; "msoShapeTrapezoid", 3; "msoShapeDiamond", 4
      "msoShapeRoundedRectangle", 5; "msoShapeOctagon", 6; "msoShapeIsoscelesTriangle", 7; "msoShapeRightTriangle", 8
      "msoShapeOval", 9; "msoShapeHexagon", 10; "msoShapeCross", 11; "msoShapeRegularPentagon", 12; "msoShapeCan", 13
      "msoShapeCube", 14; "msoShapeBevel", 15; "msoShapeFoldedCorner", 16; "msoShapeSmileyFace", 17; "msoShapeDonut", 18
      "msoShapeNoSymbol", 19; "msoShapeBlockArc", 20; "msoShapeHeart", 21; "msoShapeLightningBolt", 22; "msoShapeSun", 23
      "msoShapeMoon", 24; "msoShapeArc", 25; "msoShapeRightArrow", 33; "msoShapeLeftArrow", 34; "msoShapeUpArrow", 35
      "msoShapeDownArrow", 36; "msoShapeLeftRightArrow", 37; "msoShapeUpDownArrow", 38; "msoShapePentagon", 51
      "msoShapeChevron", 52; "msoShapeFlowchartProcess", 61; "msoShape4pointStar", 91; "msoShape5pointStar", 92
      "msoShape8pointStar", 93; "msoShape16pointStar", 94; "msoShape24pointStar", 95; "msoShape32pointStar", 96
      // MsoShapeType (Shape.Type), MsoTriState, MsoScaleFrom, MsoTextOrientation, MsoFlipCmd, MsoZOrderCmd.
      "msoAutoShape", 1; "msoChart", 3; "msoGroup", 6; "msoLine", 9; "msoPicture", 13; "msoTextBox", 17
      "msoTrue", -1; "msoFalse", 0; "msoCTrue", 1; "msoTriStateToggle", -3; "msoTriStateMixed", -2
      "msoScaleFromTopLeft", 0; "msoScaleFromMiddle", 1; "msoScaleFromBottomRight", 2
      "msoTextOrientationHorizontal", 1; "msoTextOrientationUpward", 2; "msoTextOrientationDownward", 3
      "msoFlipHorizontal", 0; "msoFlipVertical", 1
      "msoBringToFront", 0; "msoSendToBack", 1; "msoBringForward", 2; "msoSendBackward", 3
      "msoAnchorTop", 1; "msoAnchorMiddle", 3; "msoAnchorBottom", 4
      "msoAlignLeft", 1; "msoAlignCenter", 2; "msoAlignRight", 3 ]

// ---- the frame of a shape, and what moves it --------------------------------------------------------------

/// Left, top, width, height: a group's is the box round its members.
let rec frame (node: ShapeNode) : float * float * float * float =
    if node.Kind <> "group" || node.Members.Count = 0 then node.Left, node.Top, node.Width, node.Height
    else
        let mutable left, top, right, bottom = Double.MaxValue, Double.MaxValue, Double.MinValue, Double.MinValue
        for item in node.Members do
            let l, t, w, h = frame item
            left <- min left l
            top <- min top t
            right <- max right (l + w)
            bottom <- max bottom (t + h)
        left, top, right - left, bottom - top

let rec translate (node: ShapeNode) (dx: float) (dy: float) =
    if node.Kind = "group" then for item in node.Members do translate item dx dy
    else
        node.Left <- node.Left + dx
        node.Top <- node.Top + dy

/// Stretch by sx, sy about the point (ox, oy).
let rec scaleAbout (node: ShapeNode) (ox: float) (oy: float) (sx: float) (sy: float) =
    if node.Kind = "group" then for item in node.Members do scaleAbout item ox oy sx sy
    else
        node.Left <- ox + (node.Left - ox) * sx
        node.Top <- oy + (node.Top - oy) * sy
        node.Width <- node.Width * sx
        node.Height <- node.Height * sy

let private normalAngle (degrees: float) =
    let a = degrees % 360.0
    if a < 0.0 then a + 360.0 else a

/// Turn by `degrees` (clockwise, as Excel's Rotation) about the point (cx, cy).
let rec rotateAbout (node: ShapeNode) (cx: float) (cy: float) (degrees: float) =
    node.Rotation <- normalAngle (node.Rotation + degrees)
    if node.Kind = "group" then for item in node.Members do rotateAbout item cx cy degrees
    else
        let angle = degrees * Math.PI / 180.0
        let mx, my = node.Left + node.Width / 2.0, node.Top + node.Height / 2.0
        let dx, dy = mx - cx, my - cy
        let nx = cx + dx * Math.Cos angle - dy * Math.Sin angle
        let ny = cy + dx * Math.Sin angle + dy * Math.Cos angle
        node.Left <- nx - node.Width / 2.0
        node.Top <- ny - node.Height / 2.0

let setLeft (node: ShapeNode) (value: float) = let l, _, _, _ = frame node in translate node (value - l) 0.0
let setTop (node: ShapeNode) (value: float) = let _, t, _, _ = frame node in translate node 0.0 (value - t)

/// A new width or height, the top left staying (Excel's Width / Height; ScaleWidth from the top left).
let setSize (node: ShapeNode) (width: float option) (height: float option) =
    let l, t, w, h = frame node
    let sx = match width with Some value when w > 0.0 -> value / w | _ -> 1.0
    let sy = match height with Some value when h > 0.0 -> value / h | _ -> 1.0
    if node.Kind = "group" then scaleAbout node l t sx sy
    else
        width |> Option.iter (fun value -> node.Width <- value)
        height |> Option.iter (fun value -> node.Height <- value)

/// ScaleWidth / ScaleHeight: by a factor, from the top left (0), the middle (1) or the bottom right (2).
let scale (node: ShapeNode) (factor: float) (fromWhere: int) (horizontal: bool) =
    let l, t, w, h = frame node
    let ox = match fromWhere with 1 -> l + w / 2.0 | 2 -> l + w | _ -> l
    let oy = match fromWhere with 1 -> t + h / 2.0 | 2 -> t + h | _ -> t
    if horizontal then scaleAbout node ox oy factor 1.0 else scaleAbout node ox oy 1.0 factor

let setRotation (node: ShapeNode) (degrees: float) =
    let target = normalAngle degrees
    if node.Kind = "group" then
        let l, t, w, h = frame node
        rotateAbout node (l + w / 2.0) (t + h / 2.0) (target - node.Rotation)
    else node.Rotation <- target

/// A copy (with its members), its ids from `nextId`.
let rec copy (nextId: unit -> int) (node: ShapeNode) : ShapeNode =
    let c = ShapeNode(nextId (), node.Name, node.Kind)
    c.Preset <- node.Preset; c.AutoShapeType <- node.AutoShapeType
    c.Left <- node.Left; c.Top <- node.Top; c.Width <- node.Width; c.Height <- node.Height
    c.Rotation <- node.Rotation; c.FlipH <- node.FlipH; c.FlipV <- node.FlipV; c.Visible <- node.Visible
    c.FillColor <- node.FillColor; c.FillVisible <- node.FillVisible; c.FillTransparency <- node.FillTransparency
    c.LineColor <- node.LineColor; c.LineWeight <- node.LineWeight; c.LineVisible <- node.LineVisible
    c.Text <- node.Text; c.TextColor <- node.TextColor; c.TextSize <- node.TextSize; c.TextBold <- node.TextBold
    c.TextItalic <- node.TextItalic; c.TextAlign <- node.TextAlign; c.TextAnchor <- node.TextAnchor; c.OnAction <- node.OnAction
    for item in node.Members do
        let m = copy nextId item
        m.Parent <- Some c
        c.Members.Add m
    c

/// Every shape of a tree, the group first.
let rec all (node: ShapeNode) : ShapeNode seq =
    seq {
        yield node
        for item in node.Members do yield! all item }

// ---- XML ----------------------------------------------------------------------------------------------------

let private emu (points: float) = int64 (Math.Round(points * 12700.0))

let private escape (text: string) =
    text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;")

/// VBA's RGB number as RRGGBB.
let hexOf (color: int) =
    let r, g, b = color &&& 0xFF, (color >>> 8) &&& 0xFF, (color >>> 16) &&& 0xFF
    sprintf "%02X%02X%02X" r g b

let private solid (color: int) (transparency: float) =
    if transparency > 0.0 then
        sprintf "<a:solidFill><a:srgbClr val=\"%s\"><a:alpha val=\"%d\"/></a:srgbClr></a:solidFill>" (hexOf color) (int (Math.Round((1.0 - transparency) * 100000.0)))
    else sprintf "<a:solidFill><a:srgbClr val=\"%s\"/></a:solidFill>" (hexOf color)

let private transformXml (node: ShapeNode) =
    let rot = if node.Rotation <> 0.0 then sprintf " rot=\"%d\"" (int64 (Math.Round(node.Rotation * 60000.0))) else ""
    let flips = (if node.FlipH then " flipH=\"1\"" else "") + (if node.FlipV then " flipV=\"1\"" else "")
    sprintf "<a:xfrm%s%s><a:off x=\"%d\" y=\"%d\"/><a:ext cx=\"%d\" cy=\"%d\"/></a:xfrm>"
        rot flips (emu node.Left) (emu node.Top) (emu (max 0.0 node.Width)) (emu (max 0.0 node.Height))

let private lineXml (node: ShapeNode) =
    if node.LineVisible then sprintf "<a:ln w=\"%d\">%s</a:ln>" (emu node.LineWeight) (solid node.LineColor 0.0)
    else "<a:ln><a:noFill/></a:ln>"

let private textXml (node: ShapeNode) =
    if node.Text = "" then ""
    else
        let paragraphs = node.Text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')
        let run (text: string) =
            if text = "" then sprintf "<a:endParaRPr lang=\"en-US\" sz=\"%d\"/>" (int (node.TextSize * 100.0))
            else
                sprintf "<a:r><a:rPr lang=\"en-US\" sz=\"%d\"%s%s>%s</a:rPr><a:t>%s</a:t></a:r>"
                    (int (node.TextSize * 100.0)) (if node.TextBold then " b=\"1\"" else "") (if node.TextItalic then " i=\"1\"" else "")
                    (solid node.TextColor 0.0) (escape text)
        sprintf "<xdr:txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\" rtlCol=\"0\" anchor=\"%s\"/><a:lstStyle/>%s</xdr:txBody>"
            node.TextAnchor (paragraphs |> Array.map (fun p -> sprintf "<a:p><a:pPr algn=\"%s\"/>%s</a:p>" node.TextAlign (run p)) |> String.concat "")

let private nameXml (node: ShapeNode) =
    sprintf "<xdr:cNvPr id=\"%d\" name=\"%s\"%s/>" node.Id (escape node.Name) (if node.Visible then "" else " hidden=\"1\"")

/// The element of a shape (`xdr:sp`, `xdr:cxnSp`, `xdr:grpSp`).
let rec elementXml (node: ShapeNode) : string =
    match node.Kind with
    | "group" ->
        let l, t, w, h = frame node
        let box = sprintf "<a:off x=\"%d\" y=\"%d\"/><a:ext cx=\"%d\" cy=\"%d\"/>" (emu l) (emu t) (emu w) (emu h)
        sprintf "<xdr:grpSp><xdr:nvGrpSpPr>%s<xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm>%s%s</a:xfrm></xdr:grpSpPr>%s</xdr:grpSp>"
            (nameXml node) box (box.Replace("a:off", "a:chOff").Replace("a:ext", "a:chExt"))
            (node.Members |> Seq.filter (fun m -> m.Visible) |> Seq.map elementXml |> String.concat "")
    | "line" ->
        sprintf "<xdr:cxnSp macro=\"%s\"><xdr:nvCxnSpPr>%s<xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr>%s<a:prstGeom prst=\"line\"><a:avLst/></a:prstGeom>%s</xdr:spPr></xdr:cxnSp>"
            (escape node.OnAction) (nameXml node) (transformXml node) (lineXml node)
    | kind ->
        let fill = if node.FillVisible then solid node.FillColor node.FillTransparency else "<a:noFill/>"
        sprintf "<xdr:sp macro=\"%s\" textlink=\"\"><xdr:nvSpPr>%s<xdr:cNvSpPr%s/></xdr:nvSpPr><xdr:spPr>%s<a:prstGeom prst=\"%s\"><a:avLst/></a:prstGeom>%s%s</xdr:spPr>%s</xdr:sp>"
            (escape node.OnAction) (nameXml node) (if kind = "textbox" then " txBox=\"1\"" else "") (transformXml node) node.Preset fill (lineXml node) (textXml node)

/// The anchor of a top-level shape, at a cell and an offset into it (pixels), for the drawing it is.
let anchorXml (node: ShapeNode) (row: int) (col: int) (offsetX: float) (offsetY: float) =
    let _, _, w, h = frame node
    let pixelsToEmu (pixels: float) = int64 (Math.Round(pixels * 9525.0))
    sprintf "<xdr:oneCellAnchor xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><xdr:from><xdr:col>%d</xdr:col><xdr:colOff>%d</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>%d</xdr:rowOff></xdr:from><xdr:ext cx=\"%d\" cy=\"%d\"/>%s<xdr:clientData/></xdr:oneCellAnchor>"
        col (pixelsToEmu offsetX) row (pixelsToEmu offsetY) (emu w) (emu h) (elementXml node)
