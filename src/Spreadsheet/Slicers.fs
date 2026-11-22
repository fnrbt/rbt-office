/// Slicers and timelines of a workbook drawn where its drawing anchors them, as Excel shows them.
///
/// What a slicer or timeline is comes from the reader (`XlsxReader.SlicerParts`: the slicer, its cache,
/// the items of the pivot cache's field or the values of the table's column, which are selected and
/// which have no data); this module draws one for the rectangle of its anchor as `Charts` primitives:
///
/// - a slicer: the whole slicer's fill and frame, its caption over a line (the header), and a button an
///   item in `columnCount` columns of buttons `rowHeight` high, in the look the style gives a selected or
///   unselected item with or without data; the items that do not fit are cut off at the bottom with a
///   scroll bar at the right; the first shown is `startItem`. The look is a custom style of the workbook
///   (`<x14:slicerStyle>` for the buttons, the `<tableStyle>` of the same name for the whole slicer and
///   its header) or a built-in one (SlicerStyleLight1 when none is named) in the workbook's theme.
/// - a timeline: its caption, the selected period as text, the time level, and a strip of the periods of
///   its level from the date it is scrolled to - the larger period's label and each period's own above
///   its block, the selected blocks filled - and its scroll bar (simplified: no arrows, no handles).
///
/// Built-in slicer styles Light1-6 are Excel's own definitions as Excel copies them into a file when one
/// is duplicated (the e-commerce dashboard has four: borders, fills, tints and greys); Other1-2, Dark1-6
/// and the timeline styles are approximations of Excel's look. Unchecked against Excel: the margins,
/// the header's height and the gaps between buttons (Excel's are not in the file); the header's two
/// buttons (multi-select, clear filter) are not drawn.
module Rbt.Spreadsheet.Slicers

open System
open XlsxReader
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Charts

type Dxf = XlsxReader.TableStyleDxf

/// How each part of a slicer looks.
type SlicerLook = {
    Whole: Dxf
    Header: Dxf
    SelectedWithData: Dxf
    SelectedWithNoData: Dxf
    UnselectedWithData: Dxf
    UnselectedWithNoData: Dxf
}

/// How each part of a timeline looks.
type TimelineLook = {
    Whole: Dxf
    Header: Dxf
    SelectionLabel: Dxf
    TimeLevel: Dxf
    PeriodLabel1: Dxf
    PeriodLabel2: Dxf
    SelectedTimeBlock: Dxf
    UnselectedTimeBlock: Dxf
}

let private plain : Dxf =
    { Fill = None; FontColor = None; Bold = None; Italic = None
      Top = None; Bottom = None; Left = None; Right = None; Horizontal = None; Vertical = None }

let private thin (color: string) : XlsxReader.TableStyleLine option = Some { Style = "thin"; Color = Some color }

let private boxed (color: string) (dxf: Dxf) : Dxf =
    { dxf with Top = thin color; Bottom = thin color; Left = thin color; Right = thin color }

let private tint (color: string) (amount: float) = XlsxReader.Xlsx.applyColorTint color amount

/// The theme's colours in the order `<color theme="N">` counts them (lt1, dk1, lt2, dk2, accent1-6 ...).
let private themeColor (theme: string array) (index: int) (fallback: string) =
    if index >= 0 && index < theme.Length then theme.[index] else fallback

/// A built-in slicer style in a theme; None for a name that is not one.
let builtInSlicerLook (theme: string array) (name: string) : SlicerLook option =
    let numbered (prefix: string) =
        if name.StartsWith prefix then
            match Int32.TryParse (name.Substring prefix.Length) with
            | true, number -> Some number
            | _ -> None
        else None
    let dark = themeColor theme 1 "000000"
    let light (accent: string) : SlicerLook =
        // (Excel's definitions: SlicerStyleLight1 2 .. Light6 2 of the e-commerce dashboard.)
        { Whole = boxed accent { plain with FontColor = Some dark }
          Header = { plain with FontColor = Some dark; Bold = Some true; Bottom = thin accent }
          SelectedWithData = boxed "999999" { plain with Fill = Some (tint accent 0.59999389629810485); FontColor = Some "000000" }
          SelectedWithNoData = boxed "CCCCCC" { plain with Fill = Some (tint accent 0.79998168889431442); FontColor = Some "828282" }
          UnselectedWithData = boxed "CCCCCC" { plain with Fill = Some "FFFFFF"; FontColor = Some "000000" }
          UnselectedWithNoData = boxed "E0E0E0" { plain with Fill = Some "FFFFFF"; FontColor = Some "828282" } }
    let darkLook (accent: string) : SlicerLook =
        // (Approximated: a dark slicer, its selected buttons in the accent with white text.)
        let back = tint accent -0.5
        { Whole = boxed back { plain with Fill = Some back; FontColor = Some "FFFFFF" }
          Header = { plain with FontColor = Some "FFFFFF"; Bold = Some true; Bottom = thin accent }
          SelectedWithData = boxed (tint accent -0.25) { plain with Fill = Some accent; FontColor = Some "FFFFFF" }
          SelectedWithNoData = boxed (tint accent -0.25) { plain with Fill = Some (tint accent -0.25); FontColor = Some "BFBFBF" }
          UnselectedWithData = boxed (tint accent -0.25) { plain with Fill = Some (tint accent 0.6); FontColor = Some "000000" }
          UnselectedWithNoData = boxed (tint accent -0.25) { plain with Fill = Some (tint accent 0.8); FontColor = Some "828282" } }
    match numbered "SlicerStyleLight", numbered "SlicerStyleDark", numbered "SlicerStyleOther" with
    | Some n, _, _ when n >= 1 && n <= 6 -> Some (light (themeColor theme (3 + n) "4472C4"))
    | _, Some n, _ when n >= 1 && n <= 6 -> Some (darkLook (themeColor theme (3 + n) "4472C4"))
    | _, _, Some 1 -> Some (light (themeColor theme 3 "44546A"))
    | _, _, Some 2 -> Some (light (tint (themeColor theme 2 "E7E6E6") -0.5))
    | _ -> None

/// The look of a slicer style of a workbook: its own (a custom style) or a built-in one; SlicerStyleLight1
/// for a name it does not know.
let slicerLook (workbook: XlsxReader.Workbook) (name: string option) : SlicerLook =
    let theme = workbook.ThemeColors
    let name = name |> Option.orElse workbook.SlicerStyles.DefaultSlicer |> Option.defaultValue "SlicerStyleLight1"
    let custom = workbook.SlicerStyles.Slicers |> List.tryFind (fun style -> style.Name = name)
    let table = workbook.TableStyles |> List.tryFind (fun style -> style.Name = name)
    match custom, table with
    | None, None -> builtInSlicerLook theme name |> Option.orElse (builtInSlicerLook theme "SlicerStyleLight1") |> Option.get
    | _ ->
        let element (kind: string) = custom |> Option.bind (fun style -> style.Elements |> List.tryPick (fun (known, format) -> if known = kind then Some format else None)) |> Option.defaultValue plain
        let tableElement (kind: string) = table |> Option.bind (fun style -> style.Elements |> List.tryPick (fun known -> if known.Type = kind then Some known.Format else None)) |> Option.defaultValue plain
        ({ Whole = tableElement "wholeTable"
           Header = tableElement "headerRow"
           SelectedWithData = element "selectedItemWithData"
           SelectedWithNoData = element "selectedItemWithNoData"
           UnselectedWithData = element "unselectedItemWithData"
           UnselectedWithNoData = element "unselectedItemWithNoData" } : SlicerLook)

/// The look of a timeline style: a custom one of the workbook (its `<x15:timelineStyle>` elements, the
/// `<tableStyle>` of the same name for the whole timeline and its header; an element it does not have is
/// plain, as in Excel), or a built-in one - TimeSlicerStyleLight1-6 as Excel copies them into a file when
/// one is duplicated, Dark1-6 approximated.
let timelineLook (workbook: XlsxReader.Workbook) (name: string option) : TimelineLook =
    let theme = workbook.ThemeColors
    let name = name |> Option.orElse workbook.SlicerStyles.DefaultTimeline |> Option.defaultValue "TimeSlicerStyleLight1"
    let numberAfter (prefix: string) =
        if name.StartsWith prefix then
            match Int32.TryParse (name.Substring prefix.Length) with
            | true, number when number >= 1 && number <= 6 -> Some number
            | _ -> None
        else None
    let dark = themeColor theme 1 "000000"
    let builtIn : TimelineLook =
        match numberAfter "TimeSlicerStyleDark" with
        | Some n ->
            let accent = themeColor theme (3 + n) "4472C4"
            let back = tint accent -0.5
            { Whole = boxed back { plain with Fill = Some back; FontColor = Some "FFFFFF" }
              Header = { plain with FontColor = Some "FFFFFF" }
              SelectionLabel = { plain with FontColor = Some "F2F2F2" }
              TimeLevel = { plain with FontColor = Some (tint accent 0.6) }
              PeriodLabel1 = { plain with FontColor = Some "F2F2F2" }
              PeriodLabel2 = { plain with FontColor = Some "F2F2F2" }
              SelectedTimeBlock = { plain with Fill = Some accent }
              UnselectedTimeBlock = { plain with Fill = Some (tint accent -0.25) } }
        | None ->
            // (Excel's TimeSlicerStyleLight1, accent n for LightN: the frame in the accent, the selected
            // periods from the accent lighter 60% down to the accent, the others light grey, the labels grey.)
            let accent = themeColor theme (3 + (numberAfter "TimeSlicerStyleLight" |> Option.defaultValue 1)) "4472C4"
            let grey = tint dark 0.499984740745262
            { Whole = boxed accent { plain with Fill = Some "FFFFFF"; FontColor = Some dark }
              Header = { plain with FontColor = Some dark; Bold = Some true }
              SelectionLabel = { plain with FontColor = Some (tint accent -0.249977111117893) }
              TimeLevel = { plain with FontColor = Some grey }
              PeriodLabel1 = { plain with FontColor = Some grey }
              PeriodLabel2 = { plain with FontColor = Some grey }
              SelectedTimeBlock = { plain with Fill = Some ("gradient(90;0 #" + tint accent 0.59999389629810485 + " 1;1 #" + accent + " 1)") }
              UnselectedTimeBlock = { plain with Fill = Some (tint (themeColor theme 0 "FFFFFF") -0.14999847407452621) } }
    match workbook.SlicerStyles.Timelines |> List.tryFind (fun style -> style.Name = name) with
    | None -> builtIn
    | Some style ->
        let element (kind: string) = style.Elements |> List.tryPick (fun (known, format) -> if known = kind then Some format else None) |> Option.defaultValue plain
        let table = workbook.TableStyles |> List.tryFind (fun known -> known.Name = name)
        let tableElement (kind: string) = table |> Option.bind (fun known -> known.Elements |> List.tryPick (fun item -> if item.Type = kind then Some item.Format else None)) |> Option.defaultValue plain
        { Whole = tableElement "wholeTable"
          Header = tableElement "headerRow"
          SelectionLabel = element "selectionLabel"
          TimeLevel = element "timeLevel"
          PeriodLabel1 = element "periodLabel1"
          PeriodLabel2 = element "periodLabel2"
          SelectedTimeBlock = element "selectedTimeBlock"
          UnselectedTimeBlock = element "unselectedTimeBlock" }

// ---------------------------------------------------------------------------
// Drawing
// ---------------------------------------------------------------------------

/// 11 points, the size of a slicer's text, in pixels.
let private textSize = 11.0 * 96.0 / 72.0
let private hash (color: string) = "#" + color

/// A text cut to a width, with an ellipsis where it is cut.
let private fitted (size: float) (bold: bool) (width: float) (text: string) : string =
    if runWidth size bold text <= width then text
    else
        let mutable kept = text
        while kept.Length > 0 && runWidth size bold (kept + "…") > width do
            kept <- kept.Substring(0, kept.Length - 1)
        if kept = "" then "" else kept + "…"

/// The width of a slicer's text as Excel 2019 draws it: GDI's advances of the face at 11 pt (15 px), whole pixels.
let private gdiWidth (face: string) (bold: bool) (text: string) : float =
    ExcelText.textWidth (fun piece -> runWidth 15.0 bold piece) face bold false 15 false text

/// A slicer's caption or button text in the room Excel 2019 gives it, cut as GDI's end ellipsis cuts it
/// (build-slicer-cut.vbs: 400 slicers 1 px apart in Calibri and Tw Cen MT): the whole text where it is no wider
/// than the room; else the most of its characters (spaces kept) followed by three full stops "..." that is
/// narrower than the room; where even "..." is not, as many of its stops as are; none.
let slicerFitted (face: string) (bold: bool) (room: float) (text: string) : string =
    let width = gdiWidth face bold
    if width text <= room then text
    else
        let mutable kept = text
        while kept.Length > 0 && width (kept + "...") >= room do
            kept <- kept.Substring(0, kept.Length - 1)
        if kept <> "" then kept + "..."
        else [ "..."; ".."; "." ] |> List.tryFind (fun dots -> width dots < room) |> Option.defaultValue ""

/// The lines of a format around a rectangle (each edge in its colour), as primitives.
let private edges (dxf: Dxf) (x: float) (y: float) (width: float) (height: float) : Primitive list =
    let edge (line: XlsxReader.TableStyleLine option) x1 y1 x2 y2 =
        match line with
        | Some line -> [ PLine (x1, y1, x2, y2, hash (line.Color |> Option.defaultValue "000000"), 1.0, None) ]
        | None -> []
    let right, bottom = x + width - 0.5, y + height - 0.5
    edge dxf.Top x (y + 0.5) (x + width) (y + 0.5)
    @ edge dxf.Bottom x bottom (x + width) bottom
    @ edge dxf.Left (x + 0.5) y (x + 0.5) (y + height)
    @ edge dxf.Right right y right (y + height)

/// A rectangle in a format: its fill, its lines.
let private box (dxf: Dxf) (x: float) (y: float) (width: float) (height: float) : Primitive list =
    (match dxf.Fill with Some fill -> [ PRect (x, y, width, height, Some (hash fill), 1.0, None, 0.0, 0.0) ] | None -> [])
    @ edges dxf x y width height

/// Margins of a slicer, pixels (Excel's are not in the file). Measured on Excel 2019's own pictures of
/// synthetic-slicers.xlsx, synthetic-slicer-looks.xlsx and synthetic-slicer-cut.xlsx at 100%: the buttons and the
/// header's line 9 px in from the frame at the left (the slicer drawn a pixel wider than its anchor, `objectRect`)
/// and 9.67 at the right (between 9.5 and 9.83 by slicers at a third of a pixel and at whole pixels: 8 or 9 px
/// between the last button and the frame), the line 27 px below the top, the first row of buttons 32 px below it,
/// columns of buttons 3 px apart, a button's text 8 px into it, its corners rounded by 2 px.
let private margin = 9.0
let private rightMargin = 9.67
let private bottomMargin = 6.0
let private headerHeight = 30.0
let private buttonPadding = 8.0
let private columnGap = 3.0
let private rowGap = 3.0
let private scrollWidth = 10.0

/// The buttons of a slicer in the order it shows them.
let slicerButtons (data: ChartData) (info: XlsxReader.SlicerInfo) : XlsxReader.SlicerItem list =
    let items =
        match info.Source with
        | XlsxReader.CachedItems items -> items
        | XlsxReader.TableColumn (sheet, firstRow, lastRow, column, selected) ->
            // (The column's values, each once, in order: numbers by value, then texts; blanks last.)
            let cells =
                if lastRow < firstRow then [||]
                else data.Cells sheet firstRow column lastRow column |> Option.defaultValue [||] |> Array.map (fun row -> if row.Length > 0 then row.[0] else None)
            let values =
                cells
                |> Array.map (fun cell ->
                    match cell with
                    | Some cell ->
                        match unwrapFormulaValue cell.Value with
                        | CellValue.Empty -> None
                        | CellValue.Number number -> Some (0, number, data.Display cell)
                        | _ -> Some (1, 0.0, data.Display cell)
                    | None -> None)
            let known = values |> Array.choose id |> Array.distinctBy (fun (_, _, label) -> label)
            let sorted = known |> Array.sortWith (fun (kindA, numberA, labelA) (kindB, numberB, labelB) ->
                            if kindA <> kindB then compare kindA kindB
                            elif kindA = 0 then compare numberA numberB
                            else String.Compare (labelA, labelB, StringComparison.OrdinalIgnoreCase))
            let ordered = if info.Descending then Array.rev sorted else sorted
            let blanks = if values |> Array.exists Option.isNone then [ "(blank)" ] else []
            [ for _, _, label in ordered -> label ] @ blanks
            |> List.map (fun label ->
                { Label = label
                  Selected = (match selected with None -> true | Some chosen -> List.contains label chosen)
                  HasData = true } : XlsxReader.SlicerItem)
    let items = if info.HideItemsWithNoData then items |> List.filter (fun item -> item.HasData) else items
    match info.CrossFilter with
    | "none" -> items |> List.map (fun item -> { item with HasData = true })
    | "showItemsWithNoData" -> items
    | _ -> (items |> List.filter (fun item -> item.HasData)) @ (items |> List.filter (fun item -> not item.HasData))

/// The two buttons at the right of a slicer's header, as Excel 2019 draws them (16 px pictures 60 and 28 px left
/// of the slicer's right edge, 6 / 7 px down): Multi-Select (three ticks, each before a line) and Clear Filter
/// (a funnel with a red cross while the slicer filters - some item not selected -, all in light grey while it
/// does not). Their pixels are Excel's, without the anti-aliasing.
/// The Clear Filter button of a slicer's or timeline's header, its funnel's top left at x, y: a grey funnel with a red
/// cross while the object filters, all light grey while it does not (Excel 2019).
let private clearFilterButton (x: float) (y: float) (filtering: bool) : Primitive list =
    let funnel = [ "xxxxxxxxxxx"; "xxxxxxxxxxx"; ".xxxxxxxxx."; "..xxxxxxx.."; "...xxxxx..."; "....xxx...."; "....xx....."; "....x......"; "....x......"; "....x......"; "....xxx...." ]
    let gray, red = (if filtering then "#727272", "#D65532" else "#B0ADB0", "#B0ADB0")
    (funnel
     |> List.mapi (fun j row ->
        row |> Seq.mapi (fun i c -> i, c) |> Seq.filter (fun (_, c) -> c = 'x')
        |> Seq.map (fun (i, _) -> PRect (x + float i, y + float j, 1.0, 1.0, Some gray, 1.0, None, 0.0, 0.0)) |> List.ofSeq)
     |> List.concat)
    @ [ PLine (x + 7.0, y + 7.5, x + 13.5, y + 14.0, red, 1.6, None)
        PLine (x + 13.5, y + 7.5, x + 7.0, y + 14.0, red, 1.6, None) ]

let private headerButtons (width: float) (filtering: bool) : Primitive list =
    let pixels (x0: float) (y0: float) (rows: string list) (color: string) =
        rows
        |> List.mapi (fun j row ->
            row |> Seq.mapi (fun i c -> i, c) |> Seq.filter (fun (_, c) -> c = 'x')
            |> Seq.map (fun (i, _) -> PRect (x0 + float i, y0 + float j, 1.0, 1.0, Some color, 1.0, None, 0.0, 0.0)) |> List.ofSeq)
        |> List.concat
    let tick = [ ".....x"; "....x."; ".x.xx."; ".xxx.."; "..x..." ]
    let multi = width - 60.0
    let ticks =
        [ for k in 0 .. 2 do
            yield! pixels multi (6.0 + 5.0 * float k) tick "#727272"
            yield PRect (multi + 7.0, 8.0 + 5.0 * float k, 9.0, 1.0, Some "#727272", 1.0, None, 0.0, 0.0) ]
    ticks @ clearFilterButton (width - 28.0) 7.0 filtering

/// A slicer for a rectangle of that width and height (pixels), its origin at the top left, its text in a typeface
/// (the theme's body font, as Excel draws it).
let slicerItems (face: string) (look: SlicerLook) (data: ChartData) (info: XlsxReader.SlicerInfo) (width: float) (height: float) : Primitive list =
    let buttons = slicerButtons data info
    let header =
        if not info.ShowCaption then []
        else
            let bold = look.Header.Bold |> Option.defaultValue false
            let color = look.Header.FontColor |> Option.orElse look.Whole.FontColor |> Option.defaultValue "000000"
            (match look.Header.Fill with Some fill -> [ PRect (0.0, 0.0, width, headerHeight, Some (hash fill), 1.0, None, 0.0, 0.0) ] | None -> [])
            // (The caption stops before the header's buttons: cut where it is wider than the slicer less 72 px.)
            @ [ PText (margin, headerHeight / 2.0, slicerFitted face bold (width - 72.0) info.Caption, textSize, hash color, "start", bold, false, 0.0) ]
            @ edges { plain with Bottom = look.Header.Bottom } margin 0.0 (Math.Ceiling (width - rightMargin - 0.5) - margin) (headerHeight - 2.0)
            @ headerButtons width (buttons |> List.exists (fun item -> not item.Selected))
    let top = if info.ShowCaption then headerHeight + 2.0 else bottomMargin
    // (Whole pixels: 241300 EMU, 25.33 px, is 25 in Excel - buttons 28 px apart.)
    let buttonHeight = max 4.0 (Math.Round (info.RowHeight / 9525.0))
    let columns = max 1 info.ColumnCount
    let rows = (buttons.Length + columns - 1) / columns
    let fit = max 1 (int (floor ((height - top - bottomMargin + rowGap) / (buttonHeight + rowGap))))
    let firstRow = min (info.StartItem / columns) (max 0 (rows - fit))
    let scrolls = rows > fit
    let inner = width - margin - rightMargin - (if scrolls then scrollWidth + columnGap else 0.0)
    let buttonWidth = max 4.0 ((inner - float (columns - 1) * columnGap) / float columns)
    // (Each button on whole pixels, its edges rounded half down: 108.25-px buttons 108, 108, 109, 108 px wide.)
    let snap (value: float) = Math.Ceiling (value - 0.5)
    let buttonItems =
        buttons
        |> List.mapi (fun index item -> index, item)
        |> List.filter (fun (index, _) -> index / columns >= firstRow && index / columns < firstRow + fit + 1)
        |> List.collect (fun (index, item) ->
            let row, column = index / columns - firstRow, index % columns
            let x0 = float column * (buttonWidth + columnGap)
            let x = margin + snap x0
            let buttonWidth = snap (x0 + buttonWidth) - snap x0
            let y = top + float row * (buttonHeight + rowGap)
            let format =
                match item.Selected, item.HasData with
                | true, true -> look.SelectedWithData
                | true, false -> look.SelectedWithNoData
                | false, true -> look.UnselectedWithData
                | false, false -> look.UnselectedWithNoData
            let bold = format.Bold |> Option.defaultValue false
            let color = format.FontColor |> Option.orElse look.Whole.FontColor |> Option.defaultValue "000000"
            // (A button framed alike on its four sides has its corners rounded by 2 px, as Excel draws it.)
            let rounded =
                match format.Top, format.Bottom, format.Left, format.Right with
                | Some top, Some bottom, Some left, Some right when top = bottom && top = left && top = right && top.Color.IsSome ->
                    Some [ PRect (x, y, buttonWidth, buttonHeight, format.Fill |> Option.map hash, 1.0, None, 0.0, 2.0)
                           PRect (x + 0.5, y + 0.5, buttonWidth - 1.0, buttonHeight - 1.0, None, 1.0, Some (hash top.Color.Value), 1.0, 2.0) ]
                | _ -> None
            (match rounded with Some button -> button | None -> box format x y buttonWidth buttonHeight)
            // (The text at the button's left, as Excel sets it.)
            // (Cut where it is wider than the button less 18 px.)
            @ [ PText (x + buttonPadding, y + buttonHeight / 2.0, slicerFitted face bold (buttonWidth - 18.0) item.Label, textSize, hash color, "start", bold, format.Italic |> Option.defaultValue false, 0.0) ])
    let scrollBar =
        if not scrolls then []
        else
            let x = width - margin - scrollWidth
            let track = height - top - bottomMargin
            let thumb = max 12.0 (track * float fit / float rows)
            let thumbTop = top + (track - thumb) * float firstRow / float (max 1 (rows - fit))
            [ PRect (x, top, scrollWidth, track, Some "#F0F0F0", 1.0, None, 0.0, 0.0)
              PRect (x + 2.0, thumbTop, scrollWidth - 4.0, thumb, Some "#C1C1C1", 1.0, None, 0.0, 2.0) ]
    (match look.Whole.Fill with Some fill -> [ PRect (0.0, 0.0, width, height, Some (hash fill), 1.0, None, 0.0, 0.0) ] | None -> [ PRect (0.0, 0.0, width, height, Some "#FFFFFF", 1.0, None, 0.0, 0.0) ])
    @ header
    // (The buttons stop at the bottom margin; what is below is cut off.)
    @ [ PGroup (0.0, 0.0, width, height - bottomMargin + 1.0, buttonItems) ]
    @ scrollBar
    @ edges look.Whole 0.0 0.0 width height

// ---------------------------------------------------------------------------
// Timelines
// ---------------------------------------------------------------------------

let private monthNames = [| "Jan"; "Feb"; "Mar"; "Apr"; "May"; "Jun"; "Jul"; "Aug"; "Sep"; "Oct"; "Nov"; "Dec" |]

/// A period of a timeline's level: (its first day, its last day, its label, the label of the larger
/// period it starts - "" when it starts none).
let private periodAt (level: int) (start: int * int * int) (offset: int) : (int * int * int) * (int * int * int) * string * string =
    let year, month, day = start
    match level with
    | 0 ->
        let y = year + offset
        (y, 1, 1), (y, 12, 31), string y, ""
    | 1 ->
        let index = year * 4 + (month - 1) / 3 + offset
        let y, q = index / 4, index % 4
        let lastMonth = q * 3 + 3
        (y, q * 3 + 1, 1), (y, lastMonth, DateTime.DaysInMonth (y, lastMonth)), sprintf "Q%d" (q + 1), (if q = 0 || offset = 0 then string y else "")
    | 3 ->
        let date = DateTime(year, month, day).AddDays(float offset)
        (date.Year, date.Month, date.Day), (date.Year, date.Month, date.Day), string date.Day,
        (if date.Day = 1 || offset = 0 then sprintf "%s %d" (monthNames.[date.Month - 1].ToUpperInvariant()) date.Year else "")
    | _ ->
        let index = year * 12 + (month - 1) + offset
        let y, m = index / 12, index % 12
        (y, m + 1, 1), (y, m + 1, DateTime.DaysInMonth (y, m + 1)), monthNames.[m].ToUpperInvariant(), (if m = 0 || offset = 0 then string y else "")

let private dateKey ((year, month, day): int * int * int) = year * 10000 + month * 100 + day

/// What a timeline's selection says in words at its level, as Excel 2019 words it (make-timeline-variants.py,
/// synthetic-slicers.xlsx's timeline at each level in Excel): what the two ends share is said once - "Mar 5, 2012",
/// "Mar 5 - 20 2012" (no comma), "Mar 28 - Apr 3, 2012", "Dec 30, 2012 - Jan 2, 2013"; "Jun 2012", "Mar - Dec 2012",
/// "Nov 2012 - Feb 2013"; "Q2 2012", "Q2 - Q3 2012", "Q4 2012 - Q1 2013"; "2012", "2012 - 2013"; "All Periods".
let selectionText (info: XlsxReader.TimelineInfo) : string =
    match info.Selection with
    | None -> "All Periods"
    | Some ((y0, m0, d0), (y1, m1, d1)) ->
        let month m = monthNames.[m - 1]
        match info.Level with
        | 0 -> if y0 = y1 then string y0 else sprintf "%d - %d" y0 y1
        | 1 ->
            let q0, q1 = (m0 - 1) / 3 + 1, (m1 - 1) / 3 + 1
            if y0 = y1 && q0 = q1 then sprintf "Q%d %d" q0 y0
            elif y0 = y1 then sprintf "Q%d - Q%d %d" q0 q1 y0
            else sprintf "Q%d %d - Q%d %d" q0 y0 q1 y1
        | 3 ->
            if (y0, m0, d0) = (y1, m1, d1) then sprintf "%s %d, %d" (month m0) d0 y0
            elif y0 = y1 && m0 = m1 then sprintf "%s %d - %d %d" (month m0) d0 d1 y0
            elif y0 = y1 then sprintf "%s %d - %s %d, %d" (month m0) d0 (month m1) d1 y0
            else sprintf "%s %d, %d - %s %d, %d" (month m0) d0 y0 (month m1) d1 y1
        | _ ->
            if y0 = y1 && m0 = m1 then sprintf "%s %d" (month m0) y0
            elif y0 = y1 then sprintf "%s - %s %d" (month m0) (month m1) y0
            else sprintf "%s %d - %s %d" (month m0) y0 (month m1) y1

/// Where a timeline's parts are, pixels from its top left - measured on Excel 2019's own pictures of
/// synthetic-slicers.xlsx and its variants (make-timeline-variants.py) at 100 %: the caption (Calibri 11), the
/// selection label (Calibri 10) and the level (Calibri 9), the larger periods' and the periods' labels (Calibri 9)
/// with the tops of their capitals at 8, 36, 59 and 80 px; the blocks 17 px high from 95; the scroll bar 16 px
/// high from 119.
let private timelinePad = 8.0
let private captionY = 13.0
let private selectionY = 41.0
let private largerY = 64.0
let private periodY = 85.0
let private blockY = 95.0
let private blockHeight = 17.0
let private timelineScrollY = 119.0

/// How far a period's block starts from where the timeline is scrolled to (pixels; Excel 2019): months by their
/// days, 1.6 px a day (a month of 30 days 48 px, of 31 days 49 or 50: blocks at 7 + floor(days x 1.6)); quarters by
/// their days too, 0.3771 px a day (34 or 35 px); a year 54 px; a day 30 px. From 7 px in.
let private periodLeft (level: int) (first: DateTime) (starts: DateTime) (offset: int) : float =
    match level with
    | 0 -> 7.0 + 54.0 * float offset
    | 3 -> 7.0 + 30.0 * float offset
    | 1 -> 7.0 + Math.Floor ((starts - first).TotalDays * 0.3771 + 1e-9)
    | _ -> 7.0 + Math.Floor ((starts - first).TotalDays * 1.6 + 1e-9)

/// Pixels a day at a timeline's level (`periodLeft`).
let private pixelsPerDay (level: int) =
    match level with
    | 0 -> 54.0 / 365.25
    | 3 -> 30.0
    | 1 -> 0.3771
    | _ -> 1.6

/// A timeline for a rectangle (see the module's notes for what is simplified).
let timelineItems (look: TimelineLook) (info: XlsxReader.TimelineInfo) (width: float) (height: float) : Primitive list =
    let color (dxf: Dxf) = hash (dxf.FontColor |> Option.orElse look.Whole.FontColor |> Option.defaultValue "000000")
    // (The parts below the header and the labels move up when those are hidden.)
    let lift = (if info.ShowHeader then 0.0 else 26.0) + (if info.ShowSelectionLabel || info.ShowTimeLevel then 0.0 else 24.0)
    let header =
        if not info.ShowHeader then []
        else
            [ PText (timelinePad, captionY, fitted textSize false (width - 2.0 * timelinePad - 30.0) info.Caption, textSize, color look.Header, "start", look.Header.Bold |> Option.defaultValue false, false, 0.0) ]
            // The Clear Filter button at the header's right, 25 px in: a red cross while a period is picked (Excel 2019).
            @ clearFilterButton (width - 25.0) 5.0 info.Selection.IsSome
    let labels =
        if not (info.ShowSelectionLabel || info.ShowTimeLevel) then []
        else
            let y = selectionY - (if info.ShowHeader then 0.0 else 26.0)
            (if info.ShowSelectionLabel then [ PText (timelinePad, y, selectionText info, 10.0 * 96.0 / 72.0, color look.SelectionLabel, "start", false, false, 0.0) ] else [])
            @ (if info.ShowTimeLevel then
                // (The level's name ends 26 px in from the right; its arrow is a light grey 7-px triangle after it.)
                let top = y - 3.0
                [ PText (width - 26.0, y, (match info.Level with 0 -> "YEARS" | 1 -> "QUARTERS" | 3 -> "DAYS" | _ -> "MONTHS"), 12.0, color look.TimeLevel, "end", false, false, 0.0)
                  PPolyline ([ width - 19.0, top; width - 12.0, top; width - 15.5, top + 3.5 ], Some "#B2B2B2", 1.0, None, 0.0, None) ]
               else [])
    let first =
        info.ScrollPosition
        |> Option.orElse (info.Bounds |> Option.map fst)
        |> Option.orElse (info.Selection |> Option.map fst)
        |> Option.defaultValue (DateTime.Today.Year, 1, 1)
    // A selection that starts out of sight of where the timeline is scrolled to is shown in the middle of the strip:
    // the strip starts round((periods shown - periods selected) / 2) periods before it (Excel 2019: Mar 5 - 20 2012 at
    // the day level from Mar 4, Mar 28 - Apr 3 from Mar 22, a day from 9 days before).
    let first =
        match info.Selection with
        | Some (selFrom, selUntil) ->
            let date ((y: int), (m: int), (d: int)) = DateTime(y, m, d)
            let shownDays = (width - 16.0) / pixelsPerDay info.Level
            if (date selFrom - date first).TotalDays >= 0.0 && (date selFrom - date first).TotalDays < shownDays then first
            else
                let periodOf (y, m, d) = match info.Level with 0 -> y | 1 -> y * 4 + (m - 1) / 3 | 3 -> int (date (y, m, d) - DateTime(2000, 1, 1)).TotalDays | _ -> y * 12 + m - 1
                let perPeriod = match info.Level with 0 -> 54.0 | 1 -> 34.4 | 3 -> 30.0 | _ -> 48.7
                let selected = periodOf selUntil - periodOf selFrom + 1
                let back = max 0 (int (Math.Round (((width - 16.0) / perPeriod - float selected) / 2.0, MidpointRounding.AwayFromZero)))
                let s0, _, _, _ = periodAt info.Level selFrom 0
                let start, _, _, _ = periodAt info.Level s0 -back
                start
        | None -> first
    let firstDate = let y, m, d = first in DateTime(y, m, d)
    let upperY, lowerY, bandY = largerY - lift, periodY - lift, blockY - lift
    // (The periods that start before the strip's right end, 10 px in.)
    let periods =
        [ let mutable offset = 0
          let mutable go = true
          while go && offset < 1000 do
            let starts, ends, label, larger = periodAt info.Level first offset
            let startDate = let y, m, d = starts in DateTime(y, m, d)
            let x = periodLeft info.Level firstDate startDate offset
            // (Only the periods of the cache's bounds: a timeline of 2012 - 2013 at the quarter level shows 8.)
            let pastBounds = match info.Bounds with Some (_, (y1, m1, d1)) -> startDate >= DateTime(y1, m1, d1) | None -> false
            if x >= width - 10.0 || pastBounds then go <- false
            else
                let nextStart, _, _, _ = periodAt info.Level first (offset + 1)
                let next = periodLeft info.Level firstDate (let y, m, d = nextStart in DateTime(y, m, d)) (offset + 1)
                let selected =
                    match info.Selection with
                    | Some (from, until) -> dateKey starts <= dateKey until && dateKey ends >= dateKey from
                    | None -> true
                let block = if selected then look.SelectedTimeBlock else look.UnselectedTimeBlock
                // (A fill is a colour or a gradient: a colour of a format is written without its #. A pixel
                // between blocks.)
                let fill = block.Fill |> Option.map (fun fill -> if fill.StartsWith "gradient(" then fill else hash fill)
                yield PRect (x, bandY, next - x - 1.0, blockHeight, fill, 1.0, None, 0.0, 0.0)
                yield PText (x + 1.0, lowerY, label, 12.0, color look.PeriodLabel2, "start", false, false, 0.0)
                if larger <> "" then yield PText (x + 1.0, upperY, larger, 12.0, color look.PeriodLabel1, "start", false, false, 0.0)
                offset <- offset + 1 ]
    // The scroll bar (Excel 2019): no track; a white button with a grey arrow at either end (15 px wide, 7 px in from
    // the left and 9 from the right) and a white thumb between them, framed #C4C4C4, as long as the part of the
    // bounds the strip shows, where it is scrolled to.
    let scrollBar =
        if not info.ShowHorizontalScrollbar then []
        else
            let y = timelineScrollY - lift
            // (In from the anchor's right side: the timeline is drawn a pixel wider than its anchor, `objectRect`.)
            let width = width - 1.0
            let frame (x: float) (w: float) = PRect (x + 0.5, y + 0.5, w - 1.0, 15.0, Some "#FFFFFF", 1.0, Some "#C4C4C4", 1.0, 0.0)
            let trackStart, trackEnd = 23.0, width - 25.0
            let track = max 1.0 (trackEnd - trackStart + 1.0)
            let span, at, mark =
                match info.Bounds with
                | Some (b0, b1) ->
                    let d0 = let y, m, d = b0 in DateTime(y, m, d)
                    let d1 = let y, m, d = b1 in DateTime(y, m, d)
                    let days = max 1.0 (d1 - d0).TotalDays
                    let shown = (width - 16.0) / pixelsPerDay info.Level
                    let place (date: DateTime) = trackStart + Math.Floor (track * max 0.0 (min 1.0 ((date - d0).TotalDays / days)))
                    // The selected days on the track, 7 px high in its middle, in the selected blocks' colour 80 % of
                    // the way to white (#70AD47: #E3F0DB) - under the thumb, seen beside it.
                    let mark =
                        match info.Selection, look.SelectedTimeBlock.Fill with
                        | Some ((y0, m0, dd0), (y1, m1, dd1)), Some fill ->
                            let hex = let found = System.Text.RegularExpressions.Regex.Match(fill, "[0-9A-Fa-f]{6}") in if found.Success then found.Value else "70AD47"
                            let channel (i: int) = let v = Convert.ToInt32 (hex.Substring (i, 2), 16) in int (Math.Ceiling (float v + float (255 - v) * 0.8))
                            let light = sprintf "#%02X%02X%02X" (channel 0) (channel 2) (channel 4)
                            let x0 = place (DateTime(y0, m0, dd0))
                            let x1 = place (DateTime(y1, m1, dd1).AddDays 1.0)
                            if x1 > x0 then [ PRect (x0, y + 4.0, x1 - x0 + 1.0, 7.0, Some light, 1.0, None, 0.0, 0.0) ] else []
                        | _ -> []
                    min 1.0 (shown / days), max 0.0 (min 1.0 ((firstDate - d0).TotalDays / days)), mark
                | None -> 1.0, 0.0, []
            // (No thumb where the strip shows all of the bounds.)
            let thumb = min track (Math.Floor (track * span) + 2.0)
            let thumbX = min (trackEnd + 1.0 - thumb) (trackStart + Math.Floor (track * at))
            mark
            @ [ frame 7.0 15.0
                PPolyline ([ 17.0, y + 4.0; 17.0, y + 12.0; 12.5, y + 8.0 ], Some "#787775", 1.0, None, 0.0, None)
                frame (width - 23.0) 15.0
                PPolyline ([ width - 18.0, y + 4.0; width - 18.0, y + 12.0; width - 13.5, y + 8.0 ], Some "#787775", 1.0, None, 0.0, None)
                if span < 1.0 then frame thumbX thumb ]
    // (A style without a fill of its own - a custom one - leaves the cells under the timeline showing.)
    (match look.Whole.Fill with Some fill -> [ PRect (0.0, 0.0, width, height, Some (hash fill), 1.0, None, 0.0, 0.0) ] | None -> [])
    @ header
    @ labels
    @ [ PGroup (0.0, 0.0, width - 10.0, height, periods) ]
    @ scrollBar
    @ edges look.Whole 0.0 0.0 width height

/// What a drawing's `DrawingParts.Slicer` is for a workbook: a slicer or timeline of it by kind and name,
/// drawn; the workbook's slicer parts are read when the first one is drawn.
let drawer (workbook: XlsxReader.Workbook) : string -> string -> ChartData -> float -> float -> Primitive list option =
    let book = lazy (try XlsxReader.SlicerParts.read workbook with _ -> XlsxReader.SlicerParts.empty)
    // (A slicer's text is in the theme's body font: Tw Cen MT in the e-commerce dashboard, as Excel draws it.)
    let face =
        lazy (
            workbook.Preserved.ThemePath
            |> Option.bind (fun path -> workbook.Preserved.Parts |> List.tryPick (fun (known, data) -> if known = path then Some data else None))
            |> Option.bind (fun data -> Map.tryFind "#minorFont" (parseTheme (Text.Encoding.UTF8.GetString data)))
            |> Option.defaultValue "Calibri")
    fun kind name data width height ->
        match kind with
        | "slicer" ->
            Map.tryFind name book.Value.Slicers
            |> Option.map (fun info ->
                let items = slicerItems face.Value (slicerLook workbook info.Style) data info width height
                if face.Value = "Calibri" then items else [ PFace (face.Value, items) ])
        | "timeslicer" ->
            Map.tryFind name book.Value.Timelines
            |> Option.map (fun info -> timelineItems (timelineLook workbook info.Style) info width height)
        | _ -> None
