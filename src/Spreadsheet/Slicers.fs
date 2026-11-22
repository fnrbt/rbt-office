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
/// synthetic-slicers.xlsx and synthetic-slicer-looks.xlsx at 100%: the buttons and the header's line 9 px in
/// from the frame at the sides (the slicer drawn a pixel wider than its anchor, `objectRect`), the line 27 px
/// below the top, the first row of buttons 32 px below it, a button's text 8 px into it, its corners rounded
/// by 2 px.
let private margin = 9.0
let private bottomMargin = 6.0
let private headerHeight = 30.0
let private buttonPadding = 8.0
let private columnGap = 4.0
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
    let clear = width - 28.0
    let funnel = [ "xxxxxxxxxxx"; "xxxxxxxxxxx"; ".xxxxxxxxx."; "..xxxxxxx.."; "...xxxxx..."; "....xxx...."; "....xx....."; "....x......"; "....x......"; "....x......"; "....xxx...." ]
    let gray, red = (if filtering then "#727272", "#D65532" else "#B0ADB0", "#B0ADB0")
    ticks
    @ pixels clear 7.0 funnel gray
    @ [ PLine (clear + 7.0, 14.5, clear + 13.5, 21.0, red, 1.6, None)
        PLine (clear + 13.5, 14.5, clear + 7.0, 21.0, red, 1.6, None) ]

/// A slicer for a rectangle of that width and height (pixels), its origin at the top left.
let slicerItems (look: SlicerLook) (data: ChartData) (info: XlsxReader.SlicerInfo) (width: float) (height: float) : Primitive list =
    let buttons = slicerButtons data info
    let header =
        if not info.ShowCaption then []
        else
            let bold = look.Header.Bold |> Option.defaultValue false
            let color = look.Header.FontColor |> Option.orElse look.Whole.FontColor |> Option.defaultValue "000000"
            (match look.Header.Fill with Some fill -> [ PRect (0.0, 0.0, width, headerHeight, Some (hash fill), 1.0, None, 0.0, 0.0) ] | None -> [])
            // (The caption stops before the header's buttons.)
            @ [ PText (margin, headerHeight / 2.0, fitted textSize bold (width - margin - 64.0) info.Caption, textSize, hash color, "start", bold, false, 0.0) ]
            @ edges { plain with Bottom = look.Header.Bottom } margin 0.0 (width - 2.0 * margin) (headerHeight - 2.0)
            @ headerButtons width (buttons |> List.exists (fun item -> not item.Selected))
    let top = if info.ShowCaption then headerHeight + 2.0 else bottomMargin
    // (Whole pixels: 241300 EMU, 25.33 px, is 25 in Excel - buttons 28 px apart.)
    let buttonHeight = max 4.0 (Math.Round (info.RowHeight / 9525.0))
    let columns = max 1 info.ColumnCount
    let rows = (buttons.Length + columns - 1) / columns
    let fit = max 1 (int (floor ((height - top - bottomMargin + rowGap) / (buttonHeight + rowGap))))
    let firstRow = min (info.StartItem / columns) (max 0 (rows - fit))
    let scrolls = rows > fit
    let inner = width - 2.0 * margin - (if scrolls then scrollWidth + columnGap else 0.0)
    let buttonWidth = max 4.0 ((inner - float (columns - 1) * columnGap) / float columns)
    let buttonItems =
        buttons
        |> List.mapi (fun index item -> index, item)
        |> List.filter (fun (index, _) -> index / columns >= firstRow && index / columns < firstRow + fit + 1)
        |> List.collect (fun (index, item) ->
            let row, column = index / columns - firstRow, index % columns
            let x = margin + float column * (buttonWidth + columnGap)
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
            @ [ PText (x + buttonPadding, y + buttonHeight / 2.0, fitted textSize bold (buttonWidth - buttonPadding - 4.0) item.Label, textSize, hash color, "start", bold, format.Italic |> Option.defaultValue false, 0.0) ])
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

/// Where a timeline's parts are, pixels from its top left - measured on Excel 2019's own picture of
/// synthetic-slicers.xlsx at 100% (a month 48.5 px wide; its label at its left, above its block).
let private timelinePad = 8.0
let private captionY = 12.0
let private selectionY = 40.0
let private largerY = 63.0
let private periodY = 84.0
let private blockY = 96.0
let private blockHeight = 16.0
let private timelineScrollY = 118.0

/// A timeline for a rectangle (see the module's notes for what is simplified).
let timelineItems (look: TimelineLook) (info: XlsxReader.TimelineInfo) (width: float) (height: float) : Primitive list =
    let color (dxf: Dxf) = hash (dxf.FontColor |> Option.orElse look.Whole.FontColor |> Option.defaultValue "000000")
    // (The parts below the header and the labels move up when those are hidden.)
    let lift = (if info.ShowHeader then 0.0 else 26.0) + (if info.ShowSelectionLabel || info.ShowTimeLevel then 0.0 else 24.0)
    let header =
        if not info.ShowHeader then []
        else [ PText (timelinePad, captionY, fitted textSize false (width - 2.0 * timelinePad) info.Caption, textSize, color look.Header, "start", look.Header.Bold |> Option.defaultValue false, false, 0.0) ]
    let labels =
        if not (info.ShowSelectionLabel || info.ShowTimeLevel) then []
        else
            let y = selectionY - (if info.ShowHeader then 0.0 else 26.0)
            (if info.ShowSelectionLabel then [ PText (timelinePad, y, selectionText info, 10.0 * 96.0 / 72.0, color look.SelectionLabel, "start", false, false, 0.0) ] else [])
            @ (if info.ShowTimeLevel then
                [ PText (width - timelinePad - 2.0, y, (match info.Level with 0 -> "YEARS" | 1 -> "QUARTERS" | 3 -> "DAYS" | _ -> "MONTHS") + " ▾", 12.0, color look.TimeLevel, "end", false, false, 0.0) ]
               else [])
    let periodWidth = match info.Level with 0 -> 64.0 | 1 -> 64.0 | 3 -> 24.0 | _ -> 48.5
    let first =
        info.ScrollPosition
        |> Option.orElse (info.Bounds |> Option.map fst)
        |> Option.orElse (info.Selection |> Option.map fst)
        |> Option.defaultValue (DateTime.Today.Year, 1, 1)
    let count = max 1 (int (ceil ((width - 2.0 * timelinePad) / periodWidth)))
    let upperY, lowerY, bandY = largerY - lift, periodY - lift, blockY - lift
    let periods =
        [ for offset in 0 .. count - 1 do
            let starts, ends, label, larger = periodAt info.Level first offset
            let x = timelinePad + float offset * periodWidth
            let selected =
                match info.Selection with
                | Some (from, until) -> dateKey starts <= dateKey until && dateKey ends >= dateKey from
                | None -> true
            let block = if selected then look.SelectedTimeBlock else look.UnselectedTimeBlock
            // (A fill is a colour or a gradient: a colour of a format is written without its #.)
            let fill = block.Fill |> Option.map (fun fill -> if fill.StartsWith "gradient(" then fill else hash fill)
            yield PRect (x, bandY, periodWidth - 1.0, blockHeight, fill, 1.0, None, 0.0, 0.0)
            yield PText (x + 1.0, lowerY, label, 12.0, color look.PeriodLabel2, "start", false, false, 0.0)
            if larger <> "" then yield PText (x + 1.0, upperY, larger, 12.0, color look.PeriodLabel1, "start", false, false, 0.0) ]
    let scrollBar =
        if not info.ShowHorizontalScrollbar then []
        else
            let y = timelineScrollY - lift
            let total =
                match info.Bounds with
                | Some ((y0, m0, _), (y1, m1, _)) ->
                    match info.Level with
                    | 0 -> y1 - y0 + 1
                    | 1 -> (y1 * 4 + (m1 - 1) / 3) - (y0 * 4 + (m0 - 1) / 3) + 1
                    | 3 -> (y1 - y0 + 1) * 366
                    | _ -> (y1 * 12 + m1) - (y0 * 12 + m0) + 1
                | None -> count
            let track = width - 2.0 * timelinePad
            let share = min 1.0 (float count / float (max 1 total))
            [ PRect (timelinePad, y, track, 12.0, Some "#F0F0F0", 1.0, Some "#CDCDCD", 1.0, 0.0)
              PRect (timelinePad + 14.0, y + 2.0, max 12.0 ((track - 28.0) * share), 8.0, Some "#C1C1C1", 1.0, None, 0.0, 0.0) ]
    (match look.Whole.Fill with Some fill -> [ PRect (0.0, 0.0, width, height, Some (hash fill), 1.0, None, 0.0, 0.0) ] | None -> [ PRect (0.0, 0.0, width, height, Some "#FFFFFF", 1.0, None, 0.0, 0.0) ])
    @ header
    @ labels
    @ [ PGroup (0.0, 0.0, width, height, periods @ scrollBar) ]
    @ edges look.Whole 0.0 0.0 width height

/// What a drawing's `DrawingParts.Slicer` is for a workbook: a slicer or timeline of it by kind and name,
/// drawn; the workbook's slicer parts are read when the first one is drawn.
let drawer (workbook: XlsxReader.Workbook) : string -> string -> ChartData -> float -> float -> Primitive list option =
    let book = lazy (try XlsxReader.SlicerParts.read workbook with _ -> XlsxReader.SlicerParts.empty)
    fun kind name data width height ->
        match kind with
        | "slicer" ->
            Map.tryFind name book.Value.Slicers
            |> Option.map (fun info -> slicerItems (slicerLook workbook info.Style) data info width height)
        | "timeslicer" ->
            Map.tryFind name book.Value.Timelines
            |> Option.map (fun info -> timelineItems (timelineLook workbook info.Style) info width height)
        | _ -> None
