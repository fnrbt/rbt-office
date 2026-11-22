/// Form controls (check boxes, option buttons, combo and list boxes, spinners, scroll bars, group boxes,
/// labels, buttons) drawn as Excel 2019 draws them on a sheet at 100 %, and what a click on one does.
///
/// Measured on Excel's pictures of synthetic-form-controls.xlsx (tools/excel-parity/build-form-controls.vbs,
/// every kind flat and with 3-D shading, `shoot-sheets.ps1`):
///
/// - A control is painted over its anchor's rectangle grown by a pixel at the left and at the top: a combo
///   box's frame, a group box's, a button's edges are on the pixel left of / above the anchor and on its last
///   pixel at the right and bottom. Here the SVG of a control is that rectangle, (w + 1) x (h + 1).
/// - Check boxes and option buttons: a 13-px box or circle 4 px right of the anchor's left edge, centred down
///   (rounded up), the check mark or dot a 7-px picture 3 px in; flat: black lines (a mixed box chequered),
///   3-D: the sunken edges of Windows' classic controls (#A0A0A0 and #646464 above / left, #F0F0F0 inside
///   right / below); the caption 22 px right of the SVG's left, its baseline 11 px below the top of a 15-px line
///   centred down (rounded up).
/// - Group box: a frame on the rectangle's edges (flat black; 3-D an etched line, #A0A0A0 with white inside),
///   open at the top for the caption from 8 px to its end + 7 (GDI widths), the caption 13 px in, its baseline 7 px down.
/// - Combo box: a frame (flat #828790, 3-D sunken), the selected item in Tahoma 8 at 4 px, and a raised button
///   `dx` + 1 px wide at the right with an 11 x 6 arrow; list box: a frame (flat black, 3-D sunken), items 13 px
///   apart from 3 px down, the selected one on #0078D7 in white, a scroll bar of 21 px when they do not fit.
/// - Spinner: two raised buttons (#F0F0F0 / white above and left, #A0A0A0 / black below and right) with black
///   arrows. Scroll bar: the flat look of Windows 10 whatever the shading - a white line along its top (left),
///   track #F0F0F0, chevrons #606060 in square buttons of its thickness, a thumb #CDCDCD (thickness - 13 px
///   long) at value / (max - min) of the track left.
/// - Button: #F0F0F0 with #A0A0A0 / #646464 edges at the right and bottom, its caption centred.
///
/// What a control shows is read from its linked cell (Excel reads the cell, not the state saved with the
/// control): a check box is checked by TRUE or a number not 0, unchecked by FALSE or 0, mixed by #N/A; an
/// option group, a combo or list box selects the item the number gives (whole part; 0 or past the option
/// buttons none, past the items the last item of a combo box); a scroll bar's value is the number kept to
/// its min .. max. Text, an empty cell or no link: the state saved with the control.
module Rbt.Spreadsheet.FormControls

open System
open Rbt.Spreadsheet.Model

/// What a control shows.
type Shown = {
    /// 0 unchecked, 1 checked, 2 mixed.
    Checked: int
    /// The items of a combo / list box (its input range's cells as they show).
    Items: string array
    /// The selected item (from 1; 0 none).
    Selected: int
    /// A spinner's or scroll bar's value.
    Value: int
}

/// What a click on a control does.
type Hit =
    | NoHit
    /// A check box is turned over.
    | Toggle
    /// An option button is chosen.
    | Choose
    /// An item of a list box (from 1).
    | PickItem of int
    /// A combo box's list is opened.
    | OpenList
    /// A spinner's or scroll bar's value moves by this much.
    | Step of int
    /// A button is pressed.
    | Press

let private numberOf (value: CellValue option) : float option =
    match value |> Option.map unwrapFormulaValue with
    | Some (Number n) -> Some n
    | Some (Boolean b) -> Some (if b then 1.0 else 0.0)
    | _ -> None

/// A check box's state from its linked cell's value (None: no linked cell).
let checkedOf (control: FormControl) (linked: CellValue option) : int =
    match linked |> Option.map unwrapFormulaValue with
    | Some (Boolean b) -> if b then 1 else 0
    | Some (Number n) -> if n <> 0.0 then 1 else 0
    | Some (Error "#N/A") -> 2
    | _ -> control.Checked

/// The item a combo / list box or option group selects from its linked cell's value (from 1, 0 none), of
/// `count` items (a combo box past its items shows the last).
let selectedOf (control: FormControl) (linked: CellValue option) (count: int) : int =
    match numberOf linked with
    | Some n ->
        let index = int (floor n)
        if index < 1 then 0
        elif index > count then (if control.Kind = "Drop" then count else 0)
        else index
    | None -> control.Selected

/// A spinner's or scroll bar's value from its linked cell's value.
let valueOf (control: FormControl) (linked: CellValue option) : int =
    let low, high = min control.Min control.Max, max control.Min control.Max
    match numberOf linked with
    | Some n -> max low (min high (int (floor n)))
    | None -> max low (min high control.Value)

/// The option groups of a sheet's option buttons: each button's group (the group box it lies in - the
/// smallest - or none: the sheet's) and its place in the group (from 1, in the order of the file), by shape
/// id. `rect` gives a control's rectangle (x, y, width, height).
let optionGroups (controls: FormControl list) (rect: FormControl -> float * float * float * float) : Map<int, int * int> =
    let boxes = controls |> List.filter (fun c -> c.Kind = "GBox") |> List.map (fun c -> c, rect c)
    let inside (x, y, w, h) (bx, by, bw, bh) = x >= bx && y >= by && x + w <= bx + bw && y + h <= by + bh
    let groupOf (control: FormControl) =
        let r = rect control
        boxes
        |> List.filter (fun (_, box) -> inside r box)
        |> List.sortBy (fun (_, (_, _, bw, bh)) -> bw * bh)
        |> List.tryHead
        |> Option.map (fun (box, _) -> box.ShapeId)
        |> Option.defaultValue 0
    controls
    |> List.filter (fun c -> c.Kind = "Radio")
    |> List.map (fun c -> c, groupOf c)
    |> List.groupBy snd
    |> List.collect (fun (group, members) -> members |> List.mapi (fun i (c, _) -> c.ShapeId, (group, i + 1)))
    |> Map.ofList

/// The linked cell of an option group: the first of its buttons' that has one.
let groupLink (controls: FormControl list) (groups: Map<int, int * int>) (control: FormControl) : string option =
    match Map.tryFind control.ShapeId groups with
    | Some (group, _) ->
        controls
        |> List.tryPick (fun c ->
            match Map.tryFind c.ShapeId groups with
            | Some (g, _) when g = group -> c.LinkedCell |> Option.filter (fun l -> l <> "")
            | _ -> None)
    | None -> control.LinkedCell

// ---- drawing ---------------------------------------------------------------------

let private num (v: float) = (Math.Round(v, 2)).ToString(Globalization.CultureInfo.InvariantCulture)

let private esc (text: string) =
    text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;")

/// A filled rectangle of whole pixels.
let private box (sb: Text.StringBuilder) (x: int) (y: int) (w: int) (h: int) (color: string) =
    if w > 0 && h > 0 then sb.AppendFormat("<rect x=\"{0}\" y=\"{1}\" width=\"{2}\" height=\"{3}\" fill=\"{4}\"/>", x, y, w, h, color) |> ignore

/// A picture of pixels: a row a string, a character a colour of `colors` ('.' none), at (x, y).
let private mask (sb: Text.StringBuilder) (x: int) (y: int) (rows: string list) (colors: Map<char, string>) =
    rows
    |> List.iteri (fun j row ->
        let mutable i = 0
        while i < row.Length do
            let c = row.[i]
            if c = '.' then i <- i + 1
            else
                let mutable k = i
                while k < row.Length && row.[k] = c do k <- k + 1
                match Map.tryFind c colors with
                | Some color -> box sb (x + i) (y + j) (k - i) 1 color
                | None -> ()
                i <- k)

let private black = Map.ofList [ 'x', "#000000" ]

let private checkMark = [ "......x"; ".....xx"; "x...xxx"; "xx.xxx."; "xxxxx.."; ".xxx..."; "..x...." ]

let private flatRing =
    [ ".....xxx....."; "...xx...xx..."; "..x.......x.."; ".x.........x."; ".x.........x."; "x...........x"; "x...........x"
      "x...........x"; ".x.........x."; ".x.........x."; "..x.......x.."; "...xx...xx..."; ".....xxx....." ]

let private flatDot = [ "..xxx.."; ".xxxxx."; "xxxxxxx"; "xxxxxxx"; "xxxxxxx"; ".xxxxx."; "..xxx.." ]

let private shadedRing =
    [ "....aaaa...."; "..adddddda.."; ".ada.....da."; ".da.......a."; "ad.........a"; "ad.........a"; "ad.........a"; "ad.........a"
      ".d.........a"; ".ad.......a."; "..aa.....al."; "....aaaaa..." ]

let private shadedDot = [ "..aaa.."; ".xxxxx."; "axxxxxa"; "axxxxxa"; "axxxxxa"; ".xxxxx."; "..aaa.." ]

let private shades = Map.ofList [ 'a', "#A0A0A0"; 'd', "#646464"; 'l', "#F0F0F0"; 'x', "#000000" ]

let private chevronUp = [ "....x...."; "...xxx..."; "..xxxxx.."; ".xxx.xxx."; "xxx...xxx"; "xx.....xx"; "x.......x" ]
let private chevronLeft = [ "....xxx"; "...xxx."; "..xxx.."; ".xxx..."; "xxx...."; ".xxx..."; "..xxx.."; "...xxx."; "....xxx" ]
let private flip (rows: string list) = List.rev rows
let private mirror (rows: string list) = rows |> List.map (fun row -> System.String(row.ToCharArray() |> Array.rev))

/// A solid arrow `width` wide (odd) at (cx, top), pointing up or down.
let private arrow (sb: Text.StringBuilder) (cx: int) (top: int) (width: int) (up: bool) =
    let rows = (width + 1) / 2
    for j in 0 .. rows - 1 do
        let half = if up then j else rows - 1 - j
        box sb (cx - half) (top + j) (2 * half + 1) 1 "#000000"

/// A raised button of Windows' classic controls over (x, y, w, h).
let private raised (sb: Text.StringBuilder) (x: int) (y: int) (w: int) (h: int) =
    box sb x y w h "#F0F0F0"
    box sb (x + 1) (y + 1) (w - 2) 1 "#FFFFFF"
    box sb (x + 1) (y + 1) 1 (h - 2) "#FFFFFF"
    box sb (x + 1) (y + h - 2) (w - 2) 1 "#A0A0A0"
    box sb (x + w - 2) (y + 1) 1 (h - 2) "#A0A0A0"
    box sb x (y + h - 1) w 1 "#000000"
    box sb (x + w - 1) y 1 h "#000000"

/// A frame on the edges of (x, y, w, h): flat in one colour, or sunken (3-D) two pixels deep.
let private frame (sb: Text.StringBuilder) (x: int) (y: int) (w: int) (h: int) (threeD: bool) (flatColor: string) =
    if threeD then
        box sb x y w 1 "#A0A0A0"
        box sb x y 1 h "#A0A0A0"
        box sb (x + 1) (y + 1) (w - 2) 1 "#646464"
        box sb (x + 1) (y + 1) 1 (h - 2) "#646464"
        box sb (x + 1) (y + h - 2) (w - 2) 1 "#F0F0F0"
        box sb (x + w - 2) (y + 1) 1 (h - 2) "#F0F0F0"
        box sb x (y + h - 1) w 1 "#FFFFFF"
        box sb (x + w - 1) y 1 h "#FFFFFF"
    else
        box sb x y w 1 flatColor
        box sb x y 1 h flatColor
        box sb x (y + h - 1) w 1 flatColor
        box sb (x + w - 1) y 1 h flatColor

let private fontPx (points: float) = int (Math.Round(points * 4.0 / 3.0, MidpointRounding.AwayFromZero))

let private familyOf (face: string) =
    "'" + face.Replace("'", "") + "', 'Segoe UI', Tahoma, sans-serif"

/// A line of text with its baseline at (x, y), as wide as GDI draws it (`width`: Excel's advances at the pixel size,
/// `ExcelText` - Tahoma 8 "North" 27 px where the browser's Tahoma is 25.6; 0: the browser's own).
let private text (sb: Text.StringBuilder) (x: float) (y: float) (face: string) (px: int) (bold: bool) (italic: bool) (color: string) (anchor: string) (width: float) (line: string) =
    if line <> "" then
        sb.AppendFormat("<text x=\"{0}\" y=\"{1}\" font-family=\"{2}\" font-size=\"{3}px\" fill=\"{4}\"{5}{6}{7}{9} xml:space=\"preserve\">{8}</text>",
                        num x, num y, esc (familyOf face), px, color,
                        (if bold then " font-weight=\"bold\"" else ""), (if italic then " font-style=\"italic\"" else ""),
                        (if anchor = "start" then "" else " text-anchor=\"" + anchor + "\""), esc line,
                        (if width > 0.0 && line.Length > 1 then " textLength=\"" + num width + "\" lengthAdjust=\"spacing\"" else "")) |> ignore

/// The thumb of a scroll bar `thickness` across: thickness - 13 px long (9 at 22 px, 10 at 23).
let private thumbLength (thickness: int) = max 6 (thickness - 13)

/// The parts of a scroll bar along its length `length`, `thickness` across: the arrow buttons' length, and the
/// thumb's start and length (from the bar's start) for `value` of min .. max.
let scrollParts (length: int) (thickness: int) (value: int) (low: int) (high: int) =
    let button = min thickness (length / 2)
    let track = max 0 (length - 2 * button)
    let thumb = min track (thumbLength thickness)
    let span = high - low
    let at = if span <= 0 then 0 else int (floor (float (value - low) / float span * float (track - thumb)))
    button, button + at, thumb

/// The rows a list box shows: how many fit in its height `h`, 13 px each from 3 px down.
let listRows (h: int) = max 1 ((h - 3) / 13)

/// The SVG of a control whose anchor is `w` x `h` pixels: it covers the anchor grown by a pixel at the left and
/// the top, (w + 1) x (h + 1), its (0, 0) a pixel left of and above the anchor's corner. `measure` gives the
/// width of a text in a face, its size in pixels, bold, italic.
let svg (control: FormControl) (shown: Shown) (w: float) (h: float) (measure: string -> int -> bool -> bool -> string -> float) : string =
    let w = int (Math.Round w)
    let h = int (Math.Round h)
    let sb = Text.StringBuilder()
    let face, points, color, bold, italic = control.Font
    let ink = defaultArg color "#000000"
    let lines = control.Text.Split('\n')
    let captionPx = fontPx points
    match control.Kind with
    | "CheckBox" | "Radio" ->
        let x0 = 5
        let y0 = 1 + (h - 13) / 2
        if control.Kind = "CheckBox" then
            box sb x0 y0 13 13 "#FFFFFF"
            if control.ThreeD then
                frame sb x0 y0 13 13 true ""
                match shown.Checked with
                | 1 -> mask sb (x0 + 3) (y0 + 3) checkMark black
                | 2 -> mask sb (x0 + 3) (y0 + 3) [ "x.x.x.x"; ".x.x.x."; "x.x.x.x"; ".x.x.x."; "x.x.x.x"; ".x.x.x."; "x.x.x.x" ] black
                | _ -> ()
            else
                frame sb x0 y0 13 13 false "#000000"
                match shown.Checked with
                | 1 -> mask sb (x0 + 3) (y0 + 3) checkMark black
                | 2 -> mask sb (x0 + 1) (y0 + 1) [ for j in 0 .. 10 -> System.String(Array.init 11 (fun i -> if (i + j) % 2 = 0 then 'x' else '.')) ] black
                | _ -> ()
        else
            if control.ThreeD then
                mask sb x0 y0 shadedRing shades
                if shown.Checked = 1 then mask sb (x0 + 3) (y0 + 3) shadedDot shades
            else
                mask sb x0 y0 flatRing black
                if shown.Checked = 1 then mask sb (x0 + 3) (y0 + 3) flatDot black
        let lineHeight = captionPx + 4
        let top = 1 + int (Math.Ceiling (float (h - lineHeight * lines.Length) / 2.0))
        lines |> Array.iteri (fun i line -> text sb 22.0 (float (top + lineHeight * i + captionPx)) face captionPx bold italic ink "start" (measure face captionPx bold italic line) line)
    | "GBox" ->
        let caption = lines |> Array.tryHead |> Option.defaultValue ""
        let gapEnd = if caption = "" then 8 else 12 + int (Math.Round (measure face captionPx bold italic caption)) + 7
        let line color dx =
            box sb (dx) (dx) 1 (h + 1 - dx) color
            box sb (w - 1 + dx) (dx) 1 (h + 1 - dx) color
            box sb (dx) (h - 1 + dx) (w + 1 - dx) 1 color
            box sb (dx) (dx) (min 8 (w + 1)) 1 color
            if gapEnd < w then box sb (gapEnd + dx) (dx) (w + 1 - gapEnd - dx) 1 color
        if control.ThreeD then
            line "#A0A0A0" 0
            line "#FFFFFF" 1
        else line "#000000" 0
        text sb 13.0 7.0 face captionPx bold italic ink "start" (measure face captionPx bold italic caption) caption
    | "Drop" ->
        let threeD = control.ThreeD
        box sb 0 0 (w + 1) (h + 1) "#FFFFFF"
        frame sb 0 0 (w + 1) (h + 1) threeD "#828790"
        let border = if threeD then 2 else 1
        let buttonW = 23
        let bx = w + 1 - border - buttonW
        raised sb bx border buttonW (h + 1 - 2 * border)
        let cy = (border + (h - border)) / 2
        arrow sb (bx + buttonW / 2) (cy - 2) 11 false
        if shown.Selected >= 1 && shown.Selected <= shown.Items.Length then
            text sb 5.0 (float ((h + 1) / 2 + 5)) "Tahoma" 11 false false "#000000" "start" (measure "Tahoma" 11 false false shown.Items.[shown.Selected - 1]) shown.Items.[shown.Selected - 1]
    | "List" ->
        let threeD = control.ThreeD
        box sb 0 0 (w + 1) (h + 1) "#FFFFFF"
        let border = if threeD then 2 else 1
        let rows = listRows h
        let scrolled = shown.Items.Length > rows
        let barW = if scrolled then 21 else 0
        let right = w + 1 - border - barW
        // (The items are cut at the frame: drawn in a box of the list's inside.)
        sb.AppendFormat("<svg x=\"0\" y=\"{0}\" width=\"{1}\" height=\"{2}\" overflow=\"hidden\"><g transform=\"translate(0,-{0})\">", border, right, max 0 (h + 1 - 2 * border)) |> ignore
        shown.Items
        |> Array.iteri (fun i item ->
            let top = 3 + 13 * i
            if top < h then
                let selected = i + 1 = shown.Selected
                if selected then box sb border top (right - border) 13 "#0078D7"
                text sb 5.0 (float (top + 10)) "Tahoma" 11 false false (if selected then "#FFFFFF" else "#000000") "start" (measure "Tahoma" 11 false false item) item)
        sb.Append("</g></svg>") |> ignore
        if scrolled then
            let bx = right
            let length = h + 1 - 2 * border
            box sb bx border barW length "#F0F0F0"
            let button = min barW (length / 2)
            mask sb (bx + barW / 2 - 4) (border + button / 2 - 3) chevronUp (Map.ofList [ 'x', "#606060" ])
            mask sb (bx + barW / 2 - 4) (border + length - button + button / 2 - 3) (flip chevronUp) (Map.ofList [ 'x', "#606060" ])
            let track = length - 2 * button
            let thumb = max 8 (track * rows / shown.Items.Length)
            box sb (bx + 1) (border + button) (barW - 2) thumb "#CDCDCD"
        frame sb 0 0 (w + 1) (h + 1) threeD "#000000"
    | "Spin" ->
        let half = (h + 1) / 2
        raised sb 0 0 (w + 1) half
        raised sb 0 half (w + 1) (h + 1 - half)
        let width = 2 * (min (w + 1) half / 4) + 1
        let cx = (w + 1) / 2
        arrow sb cx ((half - 1) / 2 - 3) width true
        arrow sb cx (half + (h - half) / 2 - 2) width false
    | "Scroll" ->
        let horizontal = control.Horizontal
        let length = if horizontal then w + 1 else h + 1
        let thickness = if horizontal then h else w
        box sb 0 0 (w + 1) (h + 1) "#F0F0F0"
        if horizontal then box sb 0 0 (w + 1) 1 "#FFFFFF" else box sb 0 0 1 (h + 1) "#FFFFFF"
        let button, thumbAt, thumb = scrollParts length thickness shown.Value control.Min control.Max
        let gray = Map.ofList [ 'x', "#606060" ]
        if horizontal then
            let cy = 1 + thickness / 2
            mask sb (button / 2 - 4) (cy - 4) chevronLeft gray
            mask sb (length - button + (button + 1) / 2 - 4) (cy - 4) (mirror chevronLeft) gray
            box sb thumbAt 1 thumb (thickness - 1) "#CDCDCD"
        else
            let cx = 1 + thickness / 2
            mask sb (cx - 4) (button / 2 - 3) chevronUp gray
            mask sb (cx - 4) (length - button + button / 2 - 3) (flip chevronUp) gray
            box sb 1 thumbAt (thickness - 1) thumb "#CDCDCD"
    | "Label" ->
        let lineHeight = captionPx + 4
        lines |> Array.iteri (fun i line -> text sb 3.0 (float (1 + 2 + lineHeight * i + captionPx)) face captionPx bold italic ink "start" (measure face captionPx bold italic line) line)
    | "Button" ->
        box sb 0 0 (w + 1) (h + 1) "#F0F0F0"
        box sb 0 0 (w + 1) 1 "#FFFFFF"
        box sb 0 0 1 (h + 1) "#FFFFFF"
        box sb (w - 1) 1 1 h "#A0A0A0"
        box sb 1 (h - 1) (w - 1) 1 "#A0A0A0"
        box sb w 0 1 (h + 1) "#646464"
        box sb 0 h (w + 1) 1 "#646464"
        let lineHeight = int (Math.Round (float captionPx * 1.25))
        let top = (h + 1 - lineHeight * lines.Length) / 2
        let anchor, x =
            match control.TextHAlign with
            | "Left" -> "start", 4.0
            | "Right" -> "end", float (w - 4)
            | _ -> "middle", float (w + 1) / 2.0
        lines |> Array.iteri (fun i line -> text sb x (float (top + lineHeight * i) + float captionPx * 0.95) face captionPx bold italic ink anchor (measure face captionPx bold italic line) line)
    | "EditBox" ->
        box sb 0 0 (w + 1) (h + 1) "#FFFFFF"
        frame sb 0 0 (w + 1) (h + 1) control.ThreeD "#000000"
    | _ -> ()
    if sb.Length = 0 then ""
    else
        sprintf "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"%d\" height=\"%d\" viewBox=\"0 0 %d %d\" shape-rendering=\"crispEdges\" style=\"display:block;overflow:visible\">%s</svg>"
            (w + 1) (h + 1) (w + 1) (h + 1) (sb.ToString())

/// What a click at (x, y) of a control's SVG (its coordinates: (0, 0) a pixel left of and above the anchor)
/// does, for a control whose anchor is `w` x `h`.
let hit (control: FormControl) (shown: Shown) (w: float) (h: float) (x: float) (y: float) : Hit =
    let w = int (Math.Round w)
    let h = int (Math.Round h)
    if x < 0.0 || y < 0.0 || x > float (w + 1) || y > float (h + 1) then NoHit
    else
        match control.Kind with
        | "CheckBox" -> Toggle
        | "Radio" -> Choose
        | "Drop" -> OpenList
        | "List" ->
            let border = if control.ThreeD then 2 else 1
            let rows = listRows h
            let barW = if shown.Items.Length > rows then 21 else 0
            if x >= float (w + 1 - border - barW) then NoHit
            else
                let index = int (floor ((y - 3.0) / 13.0)) + 1
                if index >= 1 && index <= shown.Items.Length && index <= rows then PickItem index else NoHit
        | "Spin" -> if y < float ((h + 1) / 2) then Step control.Inc else Step (-control.Inc)
        | "Scroll" ->
            let along = if control.Horizontal then x else y
            let length = if control.Horizontal then w + 1 else h + 1
            let thickness = if control.Horizontal then h else w
            let button, thumbAt, thumb = scrollParts length thickness shown.Value control.Min control.Max
            if along < float button then Step (-control.Inc)
            elif along >= float (length - button) then Step control.Inc
            elif along < float thumbAt then Step (-control.Page)
            elif along >= float (thumbAt + thumb) then Step control.Page
            else NoHit
        | "Button" -> Press
        | _ -> NoHit
