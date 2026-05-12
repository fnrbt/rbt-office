module RedBat.Spreadsheet.Paint

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Layout

type Color = string

type TextAlign =
    | AlignLeft
    | AlignCenter
    | AlignRight

type PaintText = {
    Text: string
    Rect: Rect
    Color: Color
    FontSize: float
    Bold: bool
    Italic: bool
    Align: TextAlign
}

type PaintItem =
    | FillRect of Rect * Color
    | StrokeRect of Rect * Color * float
    | DrawText of PaintText
    | GridLine of float * float * float * float * Color * float

type ViewportPaintItem = {
    Zone: ViewportZone
    Item: PaintItem
}

/// Excel's General horizontal alignment puts numbers/booleans on the right,
/// strings on the left, and errors in the center. Explicit alignment on
/// CellFormat overrides this default.
let rec private generalAlign (value: CellValue): TextAlign =
    match value with
    | Number _ -> AlignRight
    | Boolean _ -> AlignCenter
    | Error _ -> AlignCenter
    | Formula (_, cached) -> generalAlign cached
    | CellValue.Empty -> AlignLeft
    | _ -> AlignLeft

let private alignFromFormat (format: CellFormat) (value: CellValue): TextAlign =
    match format.HorizontalAlign with
    | Some "center" -> AlignCenter
    | Some "right" -> AlignRight
    | Some "left" -> AlignLeft
    | _ -> generalAlign value

/// Interpret common Excel date/time format codes (m/d/yyyy, h:mm, etc.) and
/// render the Excel serial number as a date string. Excel serial 1 ==
/// 1900-01-01; the 1900 leap-year bug means dates after 1900-02-28 need an
/// extra day subtracted, but for typical post-1901 fixtures the simple
/// `baseDate.AddDays(serial - 2)` covers it.
/// Excel format codes use `;` to separate positive/negative/zero/text
/// sections. For numeric rendering we only care about the positive section,
/// and the `;@` text-fallback suffix shouldn't leak into the displayed value.
let private firstSection (code: string) =
    let idx = code.IndexOf(';')
    if idx >= 0 then code.Substring(0, idx) else code

let private isDateFormat (code: string) =
    let lower = (firstSection code).ToLowerInvariant()
    lower.Contains "y" || lower.Contains "d" || lower.Contains "m/" || lower.Contains "mmm"

let private isTimeFormat (code: string) =
    let lower = (firstSection code).ToLowerInvariant()
    lower.Contains "h" || (lower.Contains "s" && lower.Contains "m")

let private excelSerialToDate (serial: float) : System.DateTime =
    // Excel's 1900 calendar: serial 1 is 1900-01-01 but Excel also treats
    // 1900 as a leap year (it isn't). For serials >= 60 we subtract one to
    // line up with the real Gregorian calendar.
    let baseDate = System.DateTime(1899, 12, 30)
    let days = if serial >= 60.0 then serial else serial + 1.0
    baseDate.AddDays(days)

/// Tokenize an Excel format code into .NET DateTime.ToString format. Tracks
/// context (after `h`/`s` → minute; after month-only marker → month) for the
/// shared `m` token. Recognizes fractional-second suffixes `.0`/`.00`/`.000`
/// (Excel) and emits the .NET equivalents `.f`/`.ff`/`.fff`. Quoted/escaped
/// literals are emitted unchanged. Unrecognized tokens are wrapped in single
/// quotes so .NET treats them as literals rather than throwing.
let private excelToNetFormat (code: string) : string =
    let lower = (firstSection code).ToLowerInvariant()
    let sb = System.Text.StringBuilder()
    let mutable i = 0
    // mInTimeContext is true after we just emitted h/H or are about to emit
    // s/S — Excel's `m` token resolves to minute in those positions.
    let mutable mInTimeContext = false
    let has12Hour = lower.Contains "am/pm" || lower.Contains "a/p"

    let runOf c (s: string) =
        let mutable n = 0
        while i + n < s.Length && s.[i + n] = c do n <- n + 1
        n

    while i < lower.Length do
        match lower.[i] with
        | 'y' ->
            let n = runOf 'y' lower
            sb.Append(if n >= 4 then "yyyy" elif n >= 2 then "yy" else "yyyy") |> ignore
            i <- i + n
            mInTimeContext <- false
        | 'd' ->
            let n = runOf 'd' lower
            sb.Append(System.String('d', min 4 n)) |> ignore
            i <- i + n
            mInTimeContext <- false
        | 'h' ->
            let n = runOf 'h' lower
            let token = if has12Hour then (if n >= 2 then "hh" else "h") else (if n >= 2 then "HH" else "H")
            sb.Append(token: string) |> ignore
            i <- i + n
            mInTimeContext <- true
        | 's' ->
            let n = runOf 's' lower
            sb.Append(System.String('s', min 2 n)) |> ignore
            i <- i + n
            mInTimeContext <- true
        | 'm' ->
            let n = runOf 'm' lower
            // 3+ m's are always month (mmm/mmmm). 1-2 m's: minute if adjacent
            // to a time token, otherwise month.
            let isMinute =
                n <= 2 && (
                    mInTimeContext ||
                    (let j = i + n
                     j < lower.Length && (lower.[j] = ':' || lower.[j] = 'h' || lower.[j] = 's')))
            if isMinute then
                sb.Append(System.String('m', min 2 n)) |> ignore
                mInTimeContext <- true
            else
                let upper =
                    match n with
                    | 1 -> "M"
                    | 2 -> "MM"
                    | 3 -> "MMM"
                    | 4 -> "MMMM"
                    | _ -> "MMMMM"
                sb.Append(upper) |> ignore
                mInTimeContext <- false
            i <- i + n
        | 'a' when i + 4 < lower.Length && lower.Substring(i, 5) = "am/pm" ->
            sb.Append("tt") |> ignore
            i <- i + 5
        | 'a' when i + 2 < lower.Length && lower.Substring(i, 3) = "a/p" ->
            sb.Append("t") |> ignore
            i <- i + 3
        | '.' when i + 1 < lower.Length && lower.[i + 1] = '0' ->
            let mutable n = 0
            while i + 1 + n < lower.Length && lower.[i + 1 + n] = '0' do n <- n + 1
            sb.Append('.') |> ignore
            sb.Append(System.String('f', min 7 n)) |> ignore
            i <- i + 1 + n
        | '\\' when i + 1 < lower.Length ->
            // Escaped literal char.
            sb.Append('\'').Append(lower.[i + 1]).Append('\'') |> ignore
            i <- i + 2
        | '"' ->
            // Quoted literal.
            let endQuote = lower.IndexOf('"', i + 1)
            if endQuote < 0 then i <- lower.Length
            else
                sb.Append('\'').Append(lower.Substring(i + 1, endQuote - i - 1)).Append('\'') |> ignore
                i <- endQuote + 1
        | '[' ->
            // Elapsed-time bracket [h]/[m]/[s] — .NET can't express elapsed
            // hours past 24 with a stock format; emit a normal token and
            // accept the wraparound for now.
            let endBracket = lower.IndexOf(']', i + 1)
            if endBracket < 0 then i <- i + 1
            else i <- endBracket + 1
        | ':' | '/' | '-' | ' ' | ',' as c ->
            sb.Append(c) |> ignore
            i <- i + 1
        | c ->
            // Unrecognized: wrap as single-quote literal so .NET doesn't
            // throw on it.
            sb.Append('\'').Append(c).Append('\'') |> ignore
            i <- i + 1
    sb.ToString()

let private formatExcelDate (code: string) (serial: float) : string =
    // Excel allows huge serials (the "testdateoverflow" fixture has values
    // that exceed DateTime.MaxValue). Catch and fall back to the raw float
    // so we never crash a render on a single bad cell.
    let dt =
        try Some (excelSerialToDate serial)
        with _ -> None
    let inv = System.Globalization.CultureInfo.InvariantCulture
    match dt with
    | None -> serial.ToString("G", inv)
    | Some dt ->
        let netFormat = excelToNetFormat code
        try dt.ToString(netFormat, inv)
        with _ -> dt.ToString("MM/dd/yyyy", inv)

/// Width-aware Excel "General" number formatting. Picks the longest .NET
/// "G<n>" representation that fits the cell width. If even the shortest
/// representation overflows, returns `#####` like Excel does.
let private generalNumberDisplay (cellWidthPx: float) (n: float): string =
    // Calibri 11 max-digit-width ≈ 7 px at 96 DPI. Leave 4 px padding on
    // each side (matches the cell-text inset in buildCellPaintItems).
    let usableWidth = max 0.0 (cellWidthPx - 8.0)
    let maxChars = max 1 (int (usableWidth / 7.0))

    let inv = System.Globalization.CultureInfo.InvariantCulture
    let candidates = [ "G10"; "G8"; "G6"; "G5"; "G4"; "G3"; "G2"; "G1" ]
    let fits =
        candidates
        |> List.tryPick (fun fmt ->
            let s = n.ToString(fmt, inv)
            if s.Length <= maxChars then Some s else None)
    match fits with
    | Some s -> s
    | None -> System.String('#', maxChars)

/// Apply an Excel non-date number-format code (`0.000`, `#,##0.00`, `0%`,
/// `0.00E+00`, etc.) using .NET's custom-format-string support. Excel's
/// number-format grammar is largely compatible with .NET's; we strip the
/// section separator first and pass the positive section to `Double.ToString`,
/// falling back to `G` if .NET rejects it.
let private formatExcelNumber (code: string) (n: float) : string =
    let inv = System.Globalization.CultureInfo.InvariantCulture
    let pos = firstSection code
    let cleaned =
        // Replace literal-backslash-escapes with .NET single-quote literals.
        let sb = System.Text.StringBuilder(pos.Length)
        let mutable i = 0
        while i < pos.Length do
            if pos.[i] = '\\' && i + 1 < pos.Length then
                sb.Append('\'').Append(pos.[i + 1]).Append('\'') |> ignore
                i <- i + 2
            else
                sb.Append(pos.[i]) |> ignore
                i <- i + 1
        sb.ToString()
    try n.ToString(cleaned, inv)
    with _ ->
        try n.ToString("G", inv)
        with _ -> string n

/// Excel shows ##### when a number's display string is wider than the cell.
let private fitOrOverflow (cellWidth: float) (text: string) : string =
    let usableWidth = max 0.0 (cellWidth - 8.0)
    let maxChars = max 1 (int (usableWidth / 7.0))
    if text.Length <= maxChars then text
    else System.String('#', maxChars)

let rec private cellDisplayText (cellWidth: float) (format: CellFormat) (value: CellValue): string =
    match value with
    | CellValue.Empty -> ""
    | Number n ->
        match format.NumberFormat with
        | Some code when isDateFormat code || isTimeFormat code ->
            fitOrOverflow cellWidth (formatExcelDate code n)
        | Some "currency"
        | Some "percentage"
        | Some "date"
        | Some "decimal2"
        | Some "number" ->
            // Legacy named format codes from the Fable web UI's formatting
            // commands. Keep these on the original Model.formatNumber path.
            fitOrOverflow cellWidth (formatNumber format.NumberFormat n)
        | Some "General"
        | Some "" -> generalNumberDisplay cellWidth n
        | Some code ->
            // Real Excel format code (0.000, #,##0.00, 0%, 0.00E+00, etc.).
            fitOrOverflow cellWidth (formatExcelNumber code n)
        | None -> generalNumberDisplay cellWidth n
    | String s -> s
    | Boolean b -> if b then "TRUE" else "FALSE"
    | Error e -> sprintf "#%s" e
    | Formula (_, cached) -> cellDisplayText cellWidth format cached

let private isSelected (selection: Selection) row col =
    match selection with
    | SingleCell (r, c) -> r = row && c = col
    | Range (sr, sc, er, ec) ->
        row >= sr && row <= er && col >= sc && col <= ec

let private buildHeaderPaintItems (model: SpreadsheetModel) (viewport: Viewport): PaintItem list =
    [
        for header in headers model viewport do
            match header with
            | CornerHeader rect ->
                yield FillRect (rect, "#e8e8e8")
                yield StrokeRect (rect, "#a6a6a6", 1.0)
            | ColumnHeader (col, rect) ->
                yield FillRect (rect, "#f2f2f2")
                yield StrokeRect (rect, "#a6a6a6", 1.0)
                yield DrawText {
                    Text = columnToLetter col
                    Rect = rect
                    Color = "#444444"
                    FontSize = 11.0
                    Bold = false
                    Italic = false
                    Align = AlignCenter
                }
            | RowHeader (row, rect) ->
                yield FillRect (rect, "#f2f2f2")
                yield StrokeRect (rect, "#a6a6a6", 1.0)
                yield DrawText {
                    Text = string (row + 1)
                    Rect = rect
                    Color = "#444444"
                    FontSize = 11.0
                    Bold = false
                    Italic = false
                    Align = AlignCenter
                }
    ]

let private buildCellPaintItems (model: SpreadsheetModel) (viewport: Viewport): PaintItem list =
    [
        for cell in visibleCells model viewport do
            let addr = cell.Address
            let cellData = getCell addr model
            let value = getCellValue addr model
            let format = cellData |> Option.map (fun c -> c.Format) |> Option.defaultValue defaultFormat
            let selected = isSelected model.Selection addr.Row addr.Col
            let background =
                if selected then "#E8F5E9"
                else format.BgColor |> Option.defaultValue "#ffffff"

            yield FillRect (cell.Rect, background)
            // Excel print gridlines render as a medium gray; matches our column-
            // header stroke #a6a6a6 so cells and headers form a continuous grid.
            yield StrokeRect (cell.Rect, (if selected then "#217346" else "#a6a6a6"), (if selected then 2.0 else 1.0))

            let displayText = cellDisplayText cell.Rect.Width format value

            if displayText <> "" then
                yield DrawText {
                    Text = displayText
                    Rect = { cell.Rect with X = cell.Rect.X + 4.0; Width = max 0.0 (cell.Rect.Width - 8.0) }
                    Color = format.TextColor |> Option.defaultValue "#000000"
                    FontSize = format.FontSize |> Option.map float |> Option.defaultValue 11.0
                    Bold = format.Bold
                    Italic = format.Italic
                    Align = alignFromFormat format value
                }
    ]

let buildPaintItems (model: SpreadsheetModel) (viewport: Viewport): PaintItem list =
    [
        yield FillRect ({ X = 0.0; Y = 0.0; Width = viewport.Width; Height = viewport.Height }, "#ffffff")
        yield! buildHeaderPaintItems model viewport
        yield! buildCellPaintItems model viewport
    ]

let private translatePaintItem viewport zone item =
    let tx = toViewportRect viewport zone
    match item with
    | FillRect (rect, color) -> FillRect (tx rect, color)
    | StrokeRect (rect, color, width) -> StrokeRect (tx rect, color, width)
    | DrawText text -> DrawText { text with Rect = tx text.Rect }
    | GridLine (x1, y1, x2, y2, color, width) ->
        let p1 = tx { X = x1; Y = y1; Width = 0.0; Height = 0.0 }
        let p2 = tx { X = x2; Y = y2; Width = 0.0; Height = 0.0 }
        GridLine (p1.X, p1.Y, p2.X, p2.Y, color, width)

let buildViewportPaintItems (model: SpreadsheetModel) (viewport: Viewport): ViewportPaintItem list =
    [
        yield { Zone = CornerZone
                Item = FillRect ({ X = 0.0; Y = 0.0; Width = viewport.Width; Height = viewport.Height }, "#ffffff") }

        for header in headers model viewport do
            match header with
            | CornerHeader rect ->
                yield { Zone = CornerZone; Item = FillRect (rect, "#e8e8e8") }
                yield { Zone = CornerZone; Item = StrokeRect (rect, "#a6a6a6", 1.0) }
            | ColumnHeader (col, rect) ->
                let zone = ColumnHeaderZone
                yield { Zone = zone; Item = translatePaintItem viewport zone (FillRect (rect, "#f2f2f2")) }
                yield { Zone = zone; Item = translatePaintItem viewport zone (StrokeRect (rect, "#a6a6a6", 1.0)) }
                yield {
                    Zone = zone
                    Item =
                        translatePaintItem viewport zone (DrawText {
                            Text = columnToLetter col
                            Rect = rect
                            Color = "#444444"
                            FontSize = 11.0
                            Bold = false
                            Italic = false
                            Align = AlignCenter
                        })
                }
            | RowHeader (row, rect) ->
                let zone = RowHeaderZone
                yield { Zone = zone; Item = translatePaintItem viewport zone (FillRect (rect, "#f2f2f2")) }
                yield { Zone = zone; Item = translatePaintItem viewport zone (StrokeRect (rect, "#a6a6a6", 1.0)) }
                yield {
                    Zone = zone
                    Item =
                        translatePaintItem viewport zone (DrawText {
                            Text = string (row + 1)
                            Rect = rect
                            Color = "#444444"
                            FontSize = 11.0
                            Bold = false
                            Italic = false
                            Align = AlignCenter
                        })
                }

        for item in buildCellPaintItems model viewport do
            yield { Zone = BodyZone; Item = translatePaintItem viewport BodyZone item }
    ]
