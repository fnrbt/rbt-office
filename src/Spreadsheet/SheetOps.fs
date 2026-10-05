/// Operations on a sheet model behind the component's host API: what a host
/// asks of a workbook in A1 terms. Everything here is plain model code (no
/// DOM, no JS values), so it runs and is tested on .NET as well.
module Rbt.Spreadsheet.SheetOps

open System
open Rbt.Spreadsheet.Model

// ---------------------------------------------------------------------------
// A1 references
// ---------------------------------------------------------------------------

/// The last row and column a sheet can have (Excel's limits). A whole-column
/// range runs from row 0 to `maxSheetRow`, a whole-row range from column 0
/// to `maxSheetCol`, whatever the size of the model.
[<Literal>]
let maxSheetRow = 1048575
[<Literal>]
let maxSheetCol = 16383

let private isAsciiLetter (c: char) = (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z')
let private isAsciiDigit (c: char) = c >= '0' && c <= '9'

/// Split a reference into its sheet qualifier and the rest: "Sheet2!A1" ->
/// (Some "Sheet2", "A1"), "'My ''old'' sheet'!B2:C3" -> (Some "My 'old'
/// sheet", "B2:C3"), "A1" -> (None, "A1"). The split is at the last `!`
/// outside quotes.
let splitSheetPrefix (reference: string) : string option * string =
    let text = if isNull reference then "" else reference.Trim()
    let mutable quoted = false
    let mutable bang = -1
    for i in 0 .. text.Length - 1 do
        let c = text.[i]
        if c = '\'' then quoted <- not quoted
        elif c = '!' && not quoted then bang <- i
    if bang < 0 then None, text
    else
        let name = text.Substring(0, bang)
        let name =
            if name.Length >= 2 && name.[0] = '\'' && name.[name.Length - 1] = '\'' then
                name.Substring(1, name.Length - 2).Replace("''", "'")
            else name
        Some name, text.Substring(bang + 1)

/// A sheet name as it is written in front of a reference: quoted when it
/// holds anything but letters, digits, `_` and `.`, or starts with a digit,
/// or could be read as a cell reference.
let quoteSheetName (name: string) : string =
    let plain =
        name.Length > 0
        && not (isAsciiDigit name.[0])
        && name |> Seq.forall (fun c -> Char.IsLetterOrDigit c || c = '_' || c = '.')
        && (parseCellRef name).IsNone
    if plain then name else "'" + name.Replace("'", "''") + "'"

/// One corner of an area: (row, column), each present or not ("B" has no
/// row, "7" no column). None when the text is not a corner or is out of the
/// sheet's bounds.
let private parseCorner (part: string) : (int option * int option) option =
    let mutable i = 0
    let mutable col = 0
    let mutable valid = true
    if i < part.Length && part.[i] = '$' then i <- i + 1
    let lettersFrom = i
    while i < part.Length && isAsciiLetter part.[i] do
        if col <= maxSheetCol + 1 then col <- col * 26 + (int (Char.ToUpperInvariant part.[i]) - int 'A' + 1)
        i <- i + 1
    let letters = i - lettersFrom
    // (A `$` with no letters before it belongs to the row: "$3".)
    if i < part.Length && part.[i] = '$' then
        // ("$$3" and "B$" are not references.)
        if (letters = 0 && i > 0) || i = part.Length - 1 then valid <- false
        i <- i + 1
    let digitsFrom = i
    let mutable row = 0
    while i < part.Length && isAsciiDigit part.[i] do
        if row <= maxSheetRow + 1 then row <- row * 10 + (int part.[i] - int '0')
        i <- i + 1
    let digits = i - digitsFrom
    if not valid || i <> part.Length || (letters = 0 && digits = 0) then None
    elif letters > 0 && (col < 1 || col > maxSheetCol + 1) then None
    elif digits > 0 && (row < 1 || row > maxSheetRow + 1) then None
    else Some ((if digits > 0 then Some (row - 1) else None), (if letters > 0 then Some (col - 1) else None))

/// Read an A1 range without a sheet qualifier: a cell ("B2"), a block
/// ("B2:D5", corners in any order), whole columns ("A:C") or whole rows
/// ("3:5"); `$` signs are allowed and ignored. Whole columns and rows keep
/// their full extent (see `maxSheetRow`). None for anything else - also for
/// a mix such as "A:5" and for references beyond Excel's sheet limits.
let tryParseRange (text: string) : CellRange option =
    if isNull text then None
    else
        let parts = text.Trim().Split(':')
        if parts.Length < 1 || parts.Length > 2 then None
        else
            match parseCorner parts.[0], parseCorner parts.[parts.Length - 1] with
            | Some (Some r0, Some c0), Some (Some r1, Some c1) ->
                Some { From = { Row = min r0 r1; Col = min c0 c1 }; To = { Row = max r0 r1; Col = max c0 c1 } }
            | Some (None, Some c0), Some (None, Some c1) when parts.Length = 2 ->
                Some { From = { Row = 0; Col = min c0 c1 }; To = { Row = maxSheetRow; Col = max c0 c1 } }
            | Some (Some r0, None), Some (Some r1, None) when parts.Length = 2 ->
                Some { From = { Row = min r0 r1; Col = 0 }; To = { Row = max r0 r1; Col = maxSheetCol } }
            | _ -> None

/// `tryParseRange`, or an error that names the member the host called.
let parseRange (memberName: string) (text: string) : CellRange =
    match tryParseRange text with
    | Some range -> range
    | None -> failwith (sprintf "%s: \"%s\" is not a range (expected A1 notation such as \"B2\", \"B2:D5\", \"A:C\" or \"3:5\")" memberName (if isNull text then "" else text))

/// Read one cell reference ("B2", "$B$2"); None for anything else.
let tryParseCell (text: string) : CellAddress option =
    match tryParseRange text with
    | Some range when range.From = range.To -> Some range.From
    | _ -> None

/// Whether a range is one of whole columns / whole rows.
let isWholeColumns (range: CellRange) = range.From.Row = 0 && range.To.Row >= maxSheetRow
let isWholeRows (range: CellRange) = range.From.Col = 0 && range.To.Col >= maxSheetCol

/// A range as A1 text: "B2", "B2:D5", "A:C" for whole columns, "3:5" for
/// whole rows (the whole sheet is "1:1048576").
let rangeText (range: CellRange) : string =
    if isWholeRows range then string (range.From.Row + 1) + ":" + string (range.To.Row + 1)
    elif isWholeColumns range then columnToLetter range.From.Col + ":" + columnToLetter range.To.Col
    elif range.From = range.To then formatCellRef range.From.Row range.From.Col
    else formatCellRef range.From.Row range.From.Col + ":" + formatCellRef range.To.Row range.To.Col

/// The range of a block of cells given by its corners.
let rangeOf (top: int) (left: int) (bottom: int) (right: int) : CellRange =
    { From = { Row = min top bottom; Col = min left right }; To = { Row = max top bottom; Col = max left right } }

/// The part of a range inside a sheet of `rowCount` x `colCount` cells; None
/// when it lies outside.
let clipRange (rowCount: int) (colCount: int) (range: CellRange) : CellRange option =
    let bottom = min range.To.Row (rowCount - 1)
    let right = min range.To.Col (colCount - 1)
    if range.From.Row > bottom || range.From.Col > right then None
    else Some { From = range.From; To = { Row = bottom; Col = right } }

/// The cells of the sheet that hold something (a value, a format, a merged
/// cell): from A1 to the last such row and column. None for an empty sheet.
let usedRange (model: SpreadsheetModel) : CellRange option =
    usedBounds model |> Option.map (fun (row, col) -> { From = { Row = 0; Col = 0 }; To = { Row = row; Col = col } })

/// A range cut down to the used range of the sheet (a range that starts
/// beyond it keeps its first cell): what reading "A:A" walks.
let clipToUsed (model: SpreadsheetModel) (range: CellRange) : CellRange =
    let usedRow, usedCol = usedBounds model |> Option.defaultValue (0, 0)
    { From = range.From
      To = { Row = max range.From.Row (min range.To.Row usedRow); Col = max range.From.Col (min range.To.Col usedCol) } }

// ---------------------------------------------------------------------------
// Formats
// ---------------------------------------------------------------------------

/// The format a cell shows with: its own, else the default style of its
/// column (`ColumnStyles`), else the default format. (Conditional formats
/// are not part of it: they are an overlay, see ConditionalFormatting.)
let effectiveFormat (model: SpreadsheetModel) (addr: CellAddress) : CellFormat =
    match Map.tryFind addr model.Cells with
    | Some cell -> cell.Format
    | None ->
        match model.ColumnStyles with
        | [] -> defaultFormat
        | styles ->
            match styles |> List.tryFind (fun style -> addr.Col >= style.First && addr.Col <= style.Last) with
            | Some style -> style.Format
            | None -> defaultFormat

/// The properties of a format a host works with, as bits (for `FormatSummary.Mixed`).
module FormatField =
    [<Literal>]
    let Bold = 1
    [<Literal>]
    let Italic = 2
    [<Literal>]
    let Underline = 4
    [<Literal>]
    let Strikethrough = 8
    [<Literal>]
    let FontName = 16
    [<Literal>]
    let FontSize = 32
    [<Literal>]
    let TextColor = 64
    [<Literal>]
    let Fill = 128
    [<Literal>]
    let HorizontalAlign = 256
    [<Literal>]
    let VerticalAlign = 512
    [<Literal>]
    let WrapText = 1024
    [<Literal>]
    let NumberFormat = 2048
    [<Literal>]
    let BorderTop = 4096
    [<Literal>]
    let BorderBottom = 8192
    [<Literal>]
    let BorderLeft = 16384
    [<Literal>]
    let BorderRight = 32768
    [<Literal>]
    let Indent = 65536
    [<Literal>]
    let TextRotation = 131072
    [<Literal>]
    let ShrinkToFit = 262144
    [<Literal>]
    let UnderlineDouble = 524288
    [<Literal>]
    let BorderDiagonal = 1048576
    /// Every field.
    [<Literal>]
    let All = 2097151

/// "General" and no number format are the same format.
let private sameNumberFormat (a: string option) (b: string option) =
    let general (code: string option) =
        match code with
        | None -> true
        | Some code -> code = "" || code.Equals("General", StringComparison.OrdinalIgnoreCase)
    a = b || (general a && general b)

/// "middle" (set in the editor) and "center" (from a file) are the same alignment.
let private sameVertical (a: string option) (b: string option) =
    let middle (align: string option) = align = Some "middle" || align = Some "center"
    a = b || (middle a && middle b)

/// The fields (FormatField bits) in which two formats differ.
let formatDifference (a: CellFormat) (b: CellFormat) : int =
    if obj.ReferenceEquals(a, b) then 0
    else
        (if a.Bold <> b.Bold then FormatField.Bold else 0)
        ||| (if a.Italic <> b.Italic then FormatField.Italic else 0)
        ||| (if a.Underline <> b.Underline then FormatField.Underline else 0)
        ||| (if a.UnderlineDouble <> b.UnderlineDouble then FormatField.UnderlineDouble else 0)
        ||| (if a.Strikethrough <> b.Strikethrough then FormatField.Strikethrough else 0)
        ||| (if a.FontName <> b.FontName then FormatField.FontName else 0)
        ||| (if a.FontSize <> b.FontSize then FormatField.FontSize else 0)
        ||| (if a.TextColor <> b.TextColor then FormatField.TextColor else 0)
        ||| (if a.BgColor <> b.BgColor then FormatField.Fill else 0)
        ||| (if a.HorizontalAlign <> b.HorizontalAlign then FormatField.HorizontalAlign else 0)
        ||| (if a.VerticalAlign <> b.VerticalAlign && not (sameVertical a.VerticalAlign b.VerticalAlign) then FormatField.VerticalAlign else 0)
        ||| (if a.WrapText <> b.WrapText then FormatField.WrapText else 0)
        ||| (if a.Indent <> b.Indent then FormatField.Indent else 0)
        ||| (if a.TextRotation <> b.TextRotation then FormatField.TextRotation else 0)
        ||| (if a.ShrinkToFit <> b.ShrinkToFit then FormatField.ShrinkToFit else 0)
        ||| (if not (sameNumberFormat a.NumberFormat b.NumberFormat) then FormatField.NumberFormat else 0)
        ||| (if a.BorderTop <> b.BorderTop then FormatField.BorderTop else 0)
        ||| (if a.BorderBottom <> b.BorderBottom then FormatField.BorderBottom else 0)
        ||| (if a.BorderLeft <> b.BorderLeft then FormatField.BorderLeft else 0)
        ||| (if a.BorderRight <> b.BorderRight then FormatField.BorderRight else 0)
        ||| (if a.BorderDiagonal <> b.BorderDiagonal || a.DiagonalUp <> b.DiagonalUp || a.DiagonalDown <> b.DiagonalDown then FormatField.BorderDiagonal else 0)

/// What the cells of a block have in common: the format of one of them, and
/// the fields (FormatField bits) in which some cell differs from it.
type FormatSummary = {
    Format: CellFormat
    Mixed: int
}

/// Blocks up to this many cells are read cell by cell; larger ones (whole
/// rows, columns, the sheet) through the cells that exist.
[<Literal>]
let private cellByCellLimit = 20000.0

/// The common format of the cells of a block (corners inclusive), see
/// `FormatSummary`. The format the summary starts from is the one of the
/// block's first cell. A large block costs one pass over the sheet's cells,
/// not one step per cell of the block: cells without a format of their own
/// count once per default they fall back to (a column style, the default
/// format).
let summarizeFormats (model: SpreadsheetModel) (top: int, left: int, bottom: int, right: int) : FormatSummary =
    let first = effectiveFormat model { Row = top; Col = left }
    let mutable mixed = 0
    // Formats already taken in, by identity: cells of one style share the
    // format they were given, and comparing a record again costs far more
    // than finding it here.
    let seen = ResizeArray<CellFormat>()
    seen.Add first
    let mutable last = first
    let add (format: CellFormat) =
        if not (obj.ReferenceEquals(format, last)) then
            last <- format
            let mutable known = false
            let mutable i = 0
            while not known && i < seen.Count do
                if obj.ReferenceEquals(seen.[i], format) then known <- true
                i <- i + 1
            if not known then
                if seen.Count < 48 then seen.Add format
                mixed <- mixed ||| formatDifference first format
    let area = float (bottom - top + 1) * float (right - left + 1)
    if area <= cellByCellLimit then
        for row in top .. bottom do
            for col in left .. right do
                if mixed <> FormatField.All then add (effectiveFormat model { Row = row; Col = col })
    else
        let mutable inside = 0.0
        for KeyValue(addr, cell) in model.Cells do
            if addr.Row >= top && addr.Row <= bottom && addr.Col >= left && addr.Col <= right then
                inside <- inside + 1.0
                if mixed <> FormatField.All then add cell.Format
        if inside < area then
            // Cells that do not exist show a column's default style or the default format.
            let mutable covered = 0
            for style in model.ColumnStyles do
                let lo = max left style.First
                let hi = min right style.Last
                if lo <= hi then
                    covered <- covered + (hi - lo + 1)
                    add style.Format
            if covered < right - left + 1 then add defaultFormat
    { Format = first; Mixed = mixed }

/// The merged cell a cell is part of, as a range.
let mergeRangeAt (model: SpreadsheetModel) (addr: CellAddress) : CellRange option =
    findMergedCell addr model
    |> Option.map (fun merge -> { From = { Row = merge.StartRow; Col = merge.StartCol }; To = { Row = merge.EndRow; Col = merge.EndCol } })

/// The hyperlink on a cell: the last one defined whose range holds it.
let hyperlinkAt (model: SpreadsheetModel) (addr: CellAddress) : SheetHyperlink option =
    model.Hyperlinks |> List.tryFindBack (fun link -> cellRangeContains addr link.Range)

// ---------------------------------------------------------------------------
// Number format codes
// ---------------------------------------------------------------------------

/// The sections of a format code (split at `;` outside quotes, brackets and escapes).
let private formatSections (code: string) : string list =
    let sections = ResizeArray<string>()
    let sb = System.Text.StringBuilder()
    let mutable i = 0
    let mutable quoted = false
    let mutable bracket = false
    while i < code.Length do
        let c = code.[i]
        if not quoted && not bracket && (c = '\\' || c = '_' || c = '*') && i + 1 < code.Length then
            sb.Append(c).Append(code.[i + 1]) |> ignore
            i <- i + 2
        else
            if c = '"' && not bracket then quoted <- not quoted
            elif c = '[' && not quoted then bracket <- true
            elif c = ']' && not quoted then bracket <- false
            if c = ';' && not quoted && not bracket then
                sections.Add (sb.ToString())
                sb.Clear() |> ignore
            else sb.Append c |> ignore
            i <- i + 1
    sections.Add (sb.ToString())
    List.ofSeq sections

/// What is wrong with a number format code as a host hands it over, or None:
/// empty, longer than Excel takes (255), an open quote or bracket, a
/// dangling escape, more than four sections, or a code the formatter fails on.
let numberFormatProblem (code: string) : string option =
    if isNull code || code.Trim() = "" then Some "the format code is empty"
    elif code.Length > 255 then Some "the format code is longer than 255 characters"
    else
        let mutable quoted = false
        let mutable bracket = false
        let mutable dangling = false
        let mutable strayClose = false
        let mutable i = 0
        while i < code.Length do
            let c = code.[i]
            if not quoted && not bracket && (c = '\\' || c = '_' || c = '*') then
                if i + 1 >= code.Length then dangling <- true
                i <- i + 2
            else
                if c = '"' && not bracket then quoted <- not quoted
                elif c = '[' && not quoted then bracket <- true
                elif c = ']' && not quoted then (if bracket then bracket <- false else strayClose <- true)
                i <- i + 1
        if quoted then Some "a quote (\") is not closed"
        elif bracket then Some "a bracket ([) is not closed"
        elif strayClose then Some "a bracket (]) has no opening bracket"
        elif dangling then Some "the code ends in \\, _ or * with nothing after it"
        elif (formatSections code).Length > 4 then Some "a format code has at most four sections (positive;negative;zero;text)"
        else
            try
                Rbt.Spreadsheet.CellText.formatNumberWithCode code 1234.5 |> ignore
                Rbt.Spreadsheet.CellText.formatNumberWithCode code -0.25 |> ignore
                None
            with _ -> Some "the formatter cannot read the code"

/// One section of a number format with a decimal place more (`delta` > 0)
/// or fewer; unchanged when it has no number in it (text, a fraction).
let private sectionWithDecimals (delta: int) (section: string) : string =
    // Where the placeholders of the number are: outside quotes, brackets and
    // escapes, before an exponent.
    let plain = Array.create section.Length false
    let mutable i = 0
    let mutable quoted = false
    let mutable bracket = false
    let mutable exponentAt = -1
    let mutable fraction = false
    while i < section.Length do
        let c = section.[i]
        if not quoted && not bracket && (c = '\\' || c = '_' || c = '*') then i <- i + 2
        else
            if c = '"' && not bracket then quoted <- not quoted
            elif c = '[' && not quoted then bracket <- true
            elif c = ']' && not quoted then bracket <- false
            elif not quoted && not bracket then
                plain.[i] <- true
                if c = '/' then fraction <- true
                if exponentAt < 0 && (c = 'E' || c = 'e') && i + 1 < section.Length && (section.[i + 1] = '+' || section.[i + 1] = '-') then exponentAt <- i
            i <- i + 1
    let limit = if exponentAt >= 0 then exponentAt else section.Length
    let isPlaceholder (index: int) = index < limit && plain.[index] && (section.[index] = '0' || section.[index] = '#' || section.[index] = '?')
    let mutable lastDigit = -1
    let mutable point = -1
    for index in 0 .. limit - 1 do
        if isPlaceholder index then lastDigit <- index
        elif plain.[index] && section.[index] = '.' && point < 0 then point <- index
    let isGeneral = section.Trim().Equals("General", System.StringComparison.OrdinalIgnoreCase)
    if fraction || lastDigit < 0 || isGeneral then section
    else
        // The decimals: the placeholders right after the point.
        let mutable decimals = 0
        if point >= 0 then
            while isPlaceholder (point + 1 + decimals) do decimals <- decimals + 1
        if delta > 0 then
            if decimals >= 30 then section
            elif point >= 0 && decimals > 0 then section.Substring(0, point + 1 + decimals) + "0" + section.Substring(point + 1 + decimals)
            elif point >= 0 then section.Substring(0, point + 1) + "0" + section.Substring(point + 1)
            else section.Substring(0, lastDigit + 1) + ".0" + section.Substring(lastDigit + 1)
        else
            if point < 0 || decimals = 0 then section
            elif decimals = 1 then section.Substring(0, point) + section.Substring(point + 2)
            else section.Substring(0, point + decimals) + section.Substring(point + decimals + 1)

/// The Excel format code behind one of the editor's older named formats.
let private namedFormatCode (name: string) : string option =
    match name with
    | "currency" -> Some "$#,##0.00"
    | "percentage" | "percent" -> Some "0.00%"
    | "decimal2" -> Some "0.00"
    | "number" -> Some "#,##0.00"
    | _ -> None

/// The number format a cell gets from "increase decimal" (`delta` = 1) or
/// "decrease decimal" (-1), given the format it has and the value it shows:
/// every section of its code gains or loses a decimal place; a cell in the
/// General format starts from the decimals its value shows (12.5 -> "0.00"
/// / "0"). None when nothing changes: a date or time format, a fraction,
/// text, no decimals left to take away, or a cell that holds no number.
let changeDecimals (delta: int) (format: string option) (value: CellValue) : string option =
    let current =
        match format with
        | Some code when code.Trim() <> "" && not (code.Trim().Equals("General", System.StringComparison.OrdinalIgnoreCase)) ->
            (match namedFormatCode code with Some known -> Some known | None -> Some code)
        | _ -> None
    match current with
    | Some "date" -> None
    | Some code when Rbt.Spreadsheet.CellText.isDateTimeNumberFormat code -> None
    | Some code ->
        let changed = formatSections code |> List.map (sectionWithDecimals delta) |> String.concat ";"
        if changed = code && format = Some code then None else Some changed
    | None ->
        match unwrapFormulaValue value with
        | CellValue.Number number ->
            let shown = Rbt.Spreadsheet.CellText.formatNumberWithCode "General" number
            if shown.IndexOf 'E' >= 0 then None
            else
                let point = shown.IndexOf '.'
                let decimals = if point < 0 then 0 else shown.Length - point - 1
                let wanted = max 0 (min 30 (decimals + delta))
                // (Fewer decimals than none: nothing to take away.)
                if decimals + delta < 0 then None
                else Some (if wanted = 0 then "0" else "0." + System.String('0', wanted))
        | _ -> None

// ---------------------------------------------------------------------------
// Clearing cells
// ---------------------------------------------------------------------------

/// What "Clear" takes out of cells.
type ClearWhat =
    /// Values and formulas (and rich-text runs); formats and comments stay.
    | ClearContents
    /// Formats: the cells show with the default format again.
    | ClearFormats
    /// Contents, formats and cell comments: the cells are empty cells again.
    | ClearEverything

/// A format with what the cell's look says cleared: what the importer derived
/// from other things (filter button, comment marker, dropdown) stays.
let private clearedFormat (format: CellFormat) : CellFormat =
    if not format.FilterButton && not format.HasComment && not format.HasListDropdown then defaultFormat
    else { defaultFormat with FilterButton = format.FilterButton; HasComment = format.HasComment; HasListDropdown = format.HasListDropdown }

/// Clear the cells of a block. Only cells that exist are touched (a whole
/// column costs one pass over the sheet's cells); a cell left with nothing -
/// no value, the default format, no comment - is dropped from the map. The
/// model is returned as it is when nothing changed.
let clearCells (what: ClearWhat) (top: int, left: int, bottom: int, right: int) (model: SpreadsheetModel) : SpreadsheetModel =
    let mutable cells = model.Cells
    let mutable runs = model.RichRuns
    let change (addr: CellAddress) (data: CellData) =
        let updated =
            match what with
            | ClearContents -> { data with Value = CellValue.Empty }
            | ClearFormats -> { data with Format = clearedFormat data.Format }
            | ClearEverything -> { Value = CellValue.Empty; Format = defaultFormat; Comment = None }
        let nothingLeft = updated.Value = CellValue.Empty && updated.Comment.IsNone && updated.Format = defaultFormat
        if nothingLeft then cells <- Map.remove addr cells
        elif updated <> data then cells <- Map.add addr updated cells
        if what <> ClearFormats && not (Map.isEmpty runs) && Map.containsKey addr runs then runs <- Map.remove addr runs
    let area = float (bottom - top + 1) * float (right - left + 1)
    if area <= cellByCellLimit then
        for row in top .. bottom do
            for col in left .. right do
                let addr = { Row = row; Col = col }
                match Map.tryFind addr model.Cells with
                | Some data -> change addr data
                | None -> ()
    else
        for KeyValue(addr, data) in model.Cells do
            if addr.Row >= top && addr.Row <= bottom && addr.Col >= left && addr.Col <= right then change addr data
    if obj.ReferenceEquals(cells, model.Cells) && obj.ReferenceEquals(runs, model.RichRuns) then model
    else { model with Cells = cells; RichRuns = runs } |> invalidateFilterProjection

// ---------------------------------------------------------------------------
// Several areas in one reference
// ---------------------------------------------------------------------------

/// Read a reference of one or more areas separated by spaces or commas
/// ("A1:A5 C1:C5", "A:A,C3"). None when a part is not a range, or there is none.
let tryParseRanges (text: string) : CellRange list option =
    if isNull text then None
    else
        let parts = text.Split([| ' '; ','; ';' |], StringSplitOptions.RemoveEmptyEntries)
        let ranges = parts |> Array.choose tryParseRange
        if parts.Length = 0 || ranges.Length <> parts.Length then None else Some (List.ofArray ranges)

/// `tryParseRanges`, or an error that names the member the host called.
let parseRanges (memberName: string) (text: string) : CellRange list =
    match tryParseRanges text with
    | Some ranges -> ranges
    | None -> failwith (sprintf "%s: \"%s\" is not a range (expected A1 notation such as \"B2:D5\", or several areas: \"A1:A5 C1:C5\")" memberName (if isNull text then "" else text))

/// Some areas as one reference, separated by spaces ("A1:A5 C1:C5").
let rangesText (ranges: CellRange list) : string =
    ranges |> List.map rangeText |> String.concat " "

/// What is left of a range when the cells of `hole` are taken out of it: up
/// to four ranges (above, left, right, below), the range itself when the two
/// do not meet, nothing when the hole covers it.
let subtractRange (hole: CellRange) (range: CellRange) : CellRange list = subtractCellRange hole range

// ---------------------------------------------------------------------------
// Conditional formats
// ---------------------------------------------------------------------------
// A host describes a rule in plain data (`RuleDescription`: the fields of the
// `rule` object of the host API); `ruleOfDescription` checks it and gives the
// model's rule, `describeRule` goes the other way. The rules of a sheet are
// addressed by their `Id`.

/// Where a colour-scale stop, a data-bar end or an icon-set step sits:
/// `Kind` is "min" | "max" | "number" | "percent" | "percentile" | "formula";
/// `Value` the number / percentage / formula (none for min and max).
type ThresholdDescription = {
    Kind: string
    Value: string option
}

/// A conditional-format rule in plain data. `Type` says which fields count:
///   "cellIs"            Operator, Formula1 (and Formula2 for between / notBetween)
///   "expression"        Formula1
///   "containsText" | "notContainsText" | "beginsWith" | "endsWith"    Text
///   "top"               Rank, Percent, Bottom
///   "aboveAverage"      Above, OrEqual, StdDev
///   "duplicateValues" | "uniqueValues" | "blanks" | "noBlanks" | "errors" | "noErrors"
///   "dateOccurring"     Period ("today", "last7Days", "nextMonth" ...: `ConditionalFormatting.timePeriods`)
///   "colorScale"        Stops (2 or 3, each with its colour)
///   "dataBar"           Low, High, Color
///   "iconSet"           IconSet, Thresholds, Reverse
/// `Style` is the format of a matching cell (not for the last three, which
/// draw by themselves); None leaves the choice to the caller.
type RuleDescription = {
    Type: string
    Operator: string option
    Formula1: string option
    Formula2: string option
    Text: string option
    Rank: int
    Percent: bool
    Bottom: bool
    Above: bool
    OrEqual: bool
    StdDev: int
    Stops: (ThresholdDescription * string) list
    Low: ThresholdDescription option
    High: ThresholdDescription option
    Color: string option
    IconSet: string option
    Thresholds: ThresholdDescription list
    Reverse: bool
    /// The period of a "dateOccurring" rule.
    Period: string option
    Style: CellFormat option
}

let emptyRuleDescription (kind: string) : RuleDescription =
    { Type = kind; Operator = None; Formula1 = None; Formula2 = None; Text = None; Rank = 10; Percent = false; Bottom = false
      Above = true; OrEqual = false; StdDev = 0; Stops = []; Low = None; High = None; Color = None; IconSet = None
      Thresholds = []; Reverse = false; Period = None; Style = None }

/// The rule types a host can name.
let ruleTypes =
    [ "cellIs"; "expression"; "containsText"; "notContainsText"; "beginsWith"; "endsWith"; "top"; "aboveAverage"
      "duplicateValues"; "uniqueValues"; "blanks"; "noBlanks"; "errors"; "noErrors"; "dateOccurring"; "colorScale"; "dataBar"; "iconSet" ]

/// The icon sets a workbook file can name in a rule.
let iconSetNames =
    [ "3Arrows"; "3ArrowsGray"; "3Flags"; "3TrafficLights1"; "3TrafficLights2"; "3Signs"; "3Symbols"; "3Symbols2"
      "4Arrows"; "4ArrowsGray"; "4RedToBlack"; "4Rating"; "4TrafficLights"; "5Arrows"; "5ArrowsGray"; "5Rating"; "5Quarters" ]

/// Whether a rule draws by itself (colour scale, data bar, icon set) instead
/// of applying a format.
let isVisualRule (rule: ConditionalFormatRule) : bool =
    match rule with
    | ColorScale _ | DataBar _ | IconSet _ -> true
    | _ -> false

/// The properties of a format a conditional format applies (and a file
/// keeps): text colour, fill, bold, italic, underline, strikethrough.
let ruleStyleOf (format: CellFormat) : CellFormat =
    { defaultFormat with
        TextColor = format.TextColor; BgColor = format.BgColor; Bold = format.Bold; Italic = format.Italic
        Underline = format.Underline; Strikethrough = format.Strikethrough }

let private thresholdKindText (kind: CfThresholdKind) : string =
    match kind with
    | CfvoMin -> "min"
    | CfvoMax -> "max"
    | CfvoNumber -> "number"
    | CfvoPercent -> "percent"
    | CfvoPercentile -> "percentile"
    | CfvoFormula -> "formula"

let private describeThreshold (threshold: CfThreshold) : ThresholdDescription =
    { Kind = thresholdKindText threshold.Kind
      Value = (match threshold.Kind with CfvoMin | CfvoMax -> None | _ -> threshold.Value) }

/// A formula as a host reads and writes it: without the leading "=".
let private formulaBody (text: string) : string =
    let text = text.Trim()
    if text.StartsWith "=" then text.Substring(1).Trim() else text

let private numberOperand (value: float) : string =
    if value = Math.Floor value && abs value < 1e15 then string (int64 value)
    else Rbt.Spreadsheet.NumberText.numberToString "G" value

/// A model rule (with its style) in plain data. The literal rules the old
/// dialog makes are described as the general ones they are equal to.
let describeRule (format: ConditionalFormat) : RuleDescription =
    let styled (description: RuleDescription) = { description with Style = Some (ruleStyleOf format.Style) }
    let cellIs (op: CompareOperator) (first: string) (second: string option) =
        styled { emptyRuleDescription "cellIs" with
                    Operator = Some (Rbt.Spreadsheet.ConditionalFormatting.operatorText op)
                    Formula1 = Some (formulaBody first); Formula2 = second |> Option.map formulaBody }
    let text (kind: string) (value: string) = styled { emptyRuleDescription kind with Text = Some value }
    match format.Rule with
    | GreaterThan n -> cellIs OpGreaterThan (numberOperand n) None
    | LessThan n -> cellIs OpLessThan (numberOperand n) None
    | Equal n -> cellIs OpEqual (numberOperand n) None
    | Between (lo, hi) -> cellIs OpBetween (numberOperand lo) (Some (numberOperand hi))
    | CellIs (op, first, second) -> cellIs op first second
    | Expression formula -> styled { emptyRuleDescription "expression" with Formula1 = Some (formulaBody formula) }
    | ContainsText value | TextRule (TextContains, value) -> text "containsText" value
    | TextRule (TextNotContains, value) -> text "notContainsText" value
    | TextRule (TextBeginsWith, value) -> text "beginsWith" value
    | TextRule (TextEndsWith, value) -> text "endsWith" value
    | TopBottom (rank, percent, bottom) -> styled { emptyRuleDescription "top" with Rank = rank; Percent = percent; Bottom = bottom }
    | AboveAverage (above, orEqual, stdDev) -> styled { emptyRuleDescription "aboveAverage" with Above = above; OrEqual = orEqual; StdDev = stdDev }
    | DuplicateValues -> styled (emptyRuleDescription "duplicateValues")
    | UniqueValues -> styled (emptyRuleDescription "uniqueValues")
    | Blanks true -> styled (emptyRuleDescription "blanks")
    | Blanks false -> styled (emptyRuleDescription "noBlanks")
    | Errors true -> styled (emptyRuleDescription "errors")
    | Errors false -> styled (emptyRuleDescription "noErrors")
    | TimePeriod period -> styled { emptyRuleDescription "dateOccurring" with Period = Some period }
    | ColorScale stops ->
        { emptyRuleDescription "colorScale" with Stops = stops |> List.map (fun (threshold, color) -> describeThreshold threshold, color) }
    | DataBar (low, high, color) ->
        { emptyRuleDescription "dataBar" with Low = Some (describeThreshold low); High = Some (describeThreshold high); Color = Some color }
    | IconSet (family, thresholds, reverse) ->
        { emptyRuleDescription "iconSet" with IconSet = Some family; Thresholds = thresholds |> List.map describeThreshold; Reverse = reverse }

let private isNumberText (text: string) =
    match Double.TryParse(text.Trim(), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
    | true, number -> not (Double.IsNaN number || Double.IsInfinity number)
    | _ -> false

let private isHexColor (text: string) =
    not (isNull text) && text.Length = 7 && text.[0] = '#'
    && text.Substring 1 |> Seq.forall (fun c -> isAsciiDigit c || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F'))

/// The percentages icon sets start with: three icons at 0 / 33 / 67, four at
/// 0 / 25 / 50 / 75, five at 0 / 20 / 40 / 60 / 80.
let defaultIconThresholds (count: int) : ThresholdDescription list =
    let steps =
        match count with
        | 4 -> [ 0; 25; 50; 75 ]
        | 5 -> [ 0; 20; 40; 60; 80 ]
        | _ -> [ 0; 33; 67 ]
    steps |> List.map (fun percent -> { Kind = "percent"; Value = Some (string percent) })

/// The model rule a description stands for, checked: an error naming
/// `memberName` says what is missing or wrong. `fallbackStyle` is the format
/// of a rule whose description has none.
let ruleOfDescription (memberName: string) (fallbackStyle: CellFormat) (description: RuleDescription) : ConditionalFormat =
    let fail (message: string) = failwith (memberName + ": " + message)
    let kind = description.Type
    let threshold (what: string) (allowed: string list) (source: ThresholdDescription) : CfThreshold =
        if not (List.contains source.Kind allowed) then
            fail (sprintf "%s.type must be one of %s, got \"%s\"" what (allowed |> List.map (fun name -> "\"" + name + "\"") |> String.concat ", ") source.Kind)
        let value = source.Value |> Option.map (fun text -> text.Trim()) |> Option.filter (fun text -> text <> "")
        match source.Kind, value with
        | "min", _ -> { Kind = CfvoMin; Value = None }
        | "max", _ -> { Kind = CfvoMax; Value = None }
        | "formula", Some text -> { Kind = CfvoFormula; Value = Some (formulaBody text) }
        | "formula", None -> fail (sprintf "%s.value is missing (a formula)" what)
        | _, None -> fail (sprintf "%s.value is missing (a number)" what)
        | other, Some text ->
            if not (isNumberText text) then fail (sprintf "%s.value must be a number, got \"%s\"" what text)
            let number = Double.Parse(text, Globalization.CultureInfo.InvariantCulture)
            if other <> "number" && (number < 0.0 || number > 100.0) then fail (sprintf "%s.value must be between 0 and 100, got %s" what text)
            { Kind = (match other with "percent" -> CfvoPercent | "percentile" -> CfvoPercentile | _ -> CfvoNumber); Value = Some text }
    let color (what: string) (value: string) : string =
        if isHexColor value then value.ToUpperInvariant() else fail (sprintf "%s must be a colour like \"#RRGGBB\", got \"%s\"" what value)
    let formula (what: string) (value: string option) : string =
        match value |> Option.map formulaBody with
        | Some text when text <> "" -> text
        | _ -> fail (sprintf "a \"%s\" rule needs %s" kind what)
    let anyKind = [ "min"; "max"; "number"; "percent"; "percentile"; "formula" ]
    let rule =
        match kind with
        | "cellIs" ->
            let op =
                match description.Operator |> Option.bind Rbt.Spreadsheet.ConditionalFormatting.parseOperator with
                | Some op -> op
                | None ->
                    fail (sprintf "rule.operator must be one of \"greaterThan\", \"lessThan\", \"greaterThanOrEqual\", \"lessThanOrEqual\", \"equal\", \"notEqual\", \"between\", \"notBetween\", got %s"
                            (match description.Operator with Some text -> "\"" + text + "\"" | None -> "nothing"))
            let first = formula "rule.formula1 (a number, a quoted text or a formula)" description.Formula1
            (match op with
             | OpBetween | OpNotBetween -> CellIs (op, first, Some (formula "rule.formula2 (the other bound)" description.Formula2))
             | _ ->
                if description.Formula2 |> Option.exists (fun text -> formulaBody text <> "") then
                    fail (sprintf "rule.formula2 is only for \"between\" and \"notBetween\" (operator is \"%s\")" (Rbt.Spreadsheet.ConditionalFormatting.operatorText op))
                CellIs (op, first, None))
        | "expression" -> Expression (formula "rule.formula" description.Formula1)
        | "containsText" | "notContainsText" | "beginsWith" | "endsWith" ->
            let text =
                match description.Text with
                | Some text when text <> "" -> text
                | _ -> fail (sprintf "a \"%s\" rule needs rule.text" kind)
            TextRule ((match kind with
                       | "containsText" -> TextContains
                       | "notContainsText" -> TextNotContains
                       | "beginsWith" -> TextBeginsWith
                       | _ -> TextEndsWith), text)
        | "top" ->
            let limit = if description.Percent then 100 else 1000
            if description.Rank < 1 || description.Rank > limit then
                fail (sprintf "rule.rank must be between 1 and %d, got %d" limit description.Rank)
            TopBottom (description.Rank, description.Percent, description.Bottom)
        | "aboveAverage" ->
            if description.StdDev < 0 || description.StdDev > 3 then fail (sprintf "rule.stdDev must be between 0 and 3, got %d" description.StdDev)
            AboveAverage (description.Above, description.OrEqual, description.StdDev)
        | "duplicateValues" -> DuplicateValues
        | "uniqueValues" -> UniqueValues
        | "blanks" -> Blanks true
        | "noBlanks" -> Blanks false
        | "errors" -> Errors true
        | "noErrors" -> Errors false
        | "dateOccurring" ->
            let periods = Rbt.Spreadsheet.ConditionalFormatting.timePeriods
            (match description.Period with
             | Some period when List.contains period periods -> TimePeriod period
             | other ->
                fail (sprintf "rule.period must be one of %s, got %s" (periods |> List.map (fun name -> "\"" + name + "\"") |> String.concat ", ")
                        (match other with Some text -> "\"" + text + "\"" | None -> "nothing")))
        | "colorScale" ->
            let count = description.Stops.Length
            if count <> 2 && count <> 3 then fail (sprintf "a colour scale has 2 or 3 stops, got %d" count)
            ColorScale (description.Stops |> List.mapi (fun index (stop, stopColor) ->
                let what = sprintf "rule.stops[%d]" index
                threshold what anyKind stop, color (what + ".color") stopColor))
        | "dataBar" ->
            let low = description.Low |> Option.defaultValue { Kind = "min"; Value = None }
            let high = description.High |> Option.defaultValue { Kind = "max"; Value = None }
            DataBar (threshold "rule.min" anyKind low, threshold "rule.max" anyKind high,
                     color "rule.color" (description.Color |> Option.defaultValue "#638EC6"))
        | "iconSet" ->
            let family = description.IconSet |> Option.defaultValue "3TrafficLights1"
            if not (List.contains family iconSetNames) then
                fail (sprintf "rule.iconSet must be one of %s, got \"%s\"" (String.concat ", " iconSetNames) family)
            let count = int family.[0] - int '0'
            let thresholds = if description.Thresholds.IsEmpty then defaultIconThresholds count else description.Thresholds
            if thresholds.Length <> count then
                fail (sprintf "icon set \"%s\" needs %d thresholds (one per icon, the first is the floor), got %d" family count thresholds.Length)
            IconSet (family, thresholds |> List.mapi (fun index source -> threshold (sprintf "rule.thresholds[%d]" index) [ "number"; "percent"; "percentile"; "formula" ] source),
                     description.Reverse)
        | other ->
            fail (sprintf "unknown rule type \"%s\" (known: %s)" other (String.concat ", " ruleTypes))
    if isVisualRule rule then
        if description.Style.IsSome then fail (sprintf "a \"%s\" rule draws by itself and takes no format" kind)
        { Rule = rule; Style = defaultFormat }
    else
        { Rule = rule; Style = ruleStyleOf (description.Style |> Option.defaultValue fallbackStyle) }

/// A sheet's rules in order of precedence (the first wins where two set the
/// same property).
let conditionalFormatsInOrder (model: SpreadsheetModel) : ConditionalFormatRange list =
    model.ConditionalFormats |> List.sortBy (fun entry -> entry.Priority)

let tryFindConditionalFormat (id: int) (model: SpreadsheetModel) : ConditionalFormatRange option =
    model.ConditionalFormats |> List.tryFind (fun entry -> entry.Id = id)

let private findConditionalFormat (memberName: string) (id: int) (model: SpreadsheetModel) : ConditionalFormatRange =
    match tryFindConditionalFormat id model with
    | Some entry -> entry
    | None -> failwith (sprintf "%s: the sheet has no conditional format with id %d" memberName id)

/// The highest rule id on a sheet (0 when it has no rule).
let highestConditionalFormatId (model: SpreadsheetModel) : int =
    model.ConditionalFormats |> List.fold (fun highest entry -> max highest entry.Id) 0

/// Add a rule with the given id. Like a rule made in Excel's dialog it takes
/// precedence over the rules already there.
let addConditionalFormatWithId (id: int) (ranges: CellRange list) (stopIfTrue: bool) (format: ConditionalFormat) (model: SpreadsheetModel) : SpreadsheetModel =
    let priority =
        match model.ConditionalFormats with
        | [] -> 1
        | existing -> (existing |> List.map (fun entry -> entry.Priority) |> List.min) - 1
    { model with
        ConditionalFormats = { Id = id; Ranges = ranges; Format = format; Priority = priority; StopIfTrue = stopIfTrue } :: model.ConditionalFormats }

/// Change the rule with an id; an error when the sheet has none.
let updateConditionalFormat (memberName: string) (id: int) (change: ConditionalFormatRange -> ConditionalFormatRange) (model: SpreadsheetModel) : SpreadsheetModel =
    let existing = findConditionalFormat memberName id model
    let changed = { change existing with Id = id }
    if changed = existing then model
    else { model with ConditionalFormats = model.ConditionalFormats |> List.map (fun entry -> if entry.Id = id then changed else entry) }

/// Remove the rule with an id; an error when the sheet has none.
let removeConditionalFormat (memberName: string) (id: int) (model: SpreadsheetModel) : SpreadsheetModel =
    findConditionalFormat memberName id model |> ignore
    { model with ConditionalFormats = model.ConditionalFormats |> List.filter (fun entry -> entry.Id <> id) }

/// Put the rules in the order of `ids` (the first takes precedence): every
/// rule of the sheet is named once. Priorities become 1, 2, 3 ...
let reorderConditionalFormats (memberName: string) (ids: int list) (model: SpreadsheetModel) : SpreadsheetModel =
    let known = model.ConditionalFormats |> List.map (fun entry -> entry.Id)
    for id in ids do
        if not (List.contains id known) then failwith (sprintf "%s: the sheet has no conditional format with id %d" memberName id)
    if (List.distinct ids).Length <> ids.Length then failwith (sprintf "%s: an id is listed twice" memberName)
    if ids.Length <> known.Length then
        failwith (sprintf "%s: every conditional format of the sheet must be listed once (%d listed, the sheet has %d: %s)"
                    memberName ids.Length known.Length (conditionalFormatsInOrder model |> List.map (fun entry -> string entry.Id) |> String.concat ", "))
    let reordered =
        ids |> List.mapi (fun index id ->
            let entry = model.ConditionalFormats |> List.find (fun entry -> entry.Id = id)
            if entry.Priority = index + 1 then entry else { entry with Priority = index + 1 })
    if reordered = conditionalFormatsInOrder model then model else { model with ConditionalFormats = reordered }

/// Move the rule with an id to a place in the order of precedence (1 = first).
let moveConditionalFormat (memberName: string) (id: int) (position: int) (model: SpreadsheetModel) : SpreadsheetModel =
    findConditionalFormat memberName id model |> ignore
    let others = conditionalFormatsInOrder model |> List.map (fun entry -> entry.Id) |> List.filter (fun other -> other <> id)
    if position < 1 || position > others.Length + 1 then
        failwith (sprintf "%s: priority must be between 1 and %d, got %d" memberName (others.Length + 1) position)
    reorderConditionalFormats memberName (List.take (position - 1) others @ [ id ] @ List.skip (position - 1) others) model

/// Take conditional formats off cells: off every cell of the sheet (None),
/// or off the cells of some ranges - a rule that covers more keeps the rest
/// of its range, a rule left with no cell is gone.
let clearConditionalFormats (holes: CellRange list option) (model: SpreadsheetModel) : SpreadsheetModel =
    match holes with
    | None -> if model.ConditionalFormats.IsEmpty then model else { model with ConditionalFormats = [] }
    | Some holes ->
        let remaining =
            model.ConditionalFormats
            |> List.choose (fun entry ->
                let ranges = holes |> List.fold (fun ranges hole -> ranges |> List.collect (subtractRange hole)) entry.Ranges
                match ranges with
                | [] -> None
                | ranges -> Some (if ranges = entry.Ranges then entry else { entry with Ranges = ranges }))
        if remaining = model.ConditionalFormats then model else { model with ConditionalFormats = remaining }

// ---------------------------------------------------------------------------
// Data validation
// ---------------------------------------------------------------------------
// A host describes a validation in plain data (`ValidationDescription`: the
// fields of the `rule` object of the host API); `validationOfDescription`
// checks it and gives the model's validation, `describeValidation` goes the
// other way. A sheet's validations are stored by range (`Model.ValidationRange`)
// and addressed by their `Id`.

module CF = Rbt.Spreadsheet.ConditionalFormatting

/// A validation as plain data.
type ValidationDescription = {
    /// "any" | "list" | "whole" | "decimal" | "date" | "time" | "textLength" | "custom".
    Type: string
    /// Of the comparison types: "between" | "notBetween" | "equal" | "notEqual"
    /// | "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual".
    Operator: string option
    /// The operands of a comparison (a number, or a formula without its `=`;
    /// for "date" also "2024-03-31", for "time" also "14:30"), or the formula
    /// of a "custom" rule.
    Formula1: string option
    Formula2: string option
    /// A list written out in the rule.
    Values: string list option
    /// A list read from cells: a range ("$A$1:$A$9", "Lists!$B$2:$B$20") or
    /// a defined name.
    Source: string option
    /// A list offers its values in the cell.
    Dropdown: bool
    AllowBlank: bool
    ShowInput: bool
    InputTitle: string option
    InputMessage: string option
    ShowError: bool
    /// "stop" | "warning" | "information".
    ErrorStyle: string
    ErrorTitle: string option
    ErrorMessage: string option
}

let emptyValidationDescription (kind: string) : ValidationDescription =
    { Type = kind; Operator = None; Formula1 = None; Formula2 = None; Values = None; Source = None; Dropdown = true
      AllowBlank = true; ShowInput = true; InputTitle = None; InputMessage = None
      ShowError = true; ErrorStyle = "stop"; ErrorTitle = None; ErrorMessage = None }

/// The validation types and operators a host can name.
let validationTypes = [ "any"; "list"; "whole"; "decimal"; "date"; "time"; "textLength"; "custom" ]
let validationOperators =
    [ "between"; "notBetween"; "equal"; "notEqual"; "greaterThan"; "lessThan"; "greaterThanOrEqual"; "lessThanOrEqual" ]

/// A model validation as plain data.
let describeValidation (validation: DataValidation) : ValidationDescription =
    let ooxml = CF.ooxmlOfValidation validation
    let common =
        { emptyValidationDescription (if ooxml.Type = "none" then "any" else ooxml.Type) with
            Dropdown = validation.ShowDropdown
            AllowBlank = validation.AllowBlank
            ShowInput = validation.ShowInput
            InputTitle = validation.InputTitle
            InputMessage = validation.InputMessage
            ShowError = validation.ShowError
            ErrorStyle = (match validation.Style with StopStyle -> "stop" | WarningStyle -> "warning" | InformationStyle -> "information")
            ErrorTitle = validation.ErrorTitle
            ErrorMessage = validation.ErrorMessage }
    match validation.Rule with
    | AnyValue -> common
    | ListValues values -> { common with Values = Some values }
    | ListSource source -> { common with Source = Some (formulaBody source) }
    | CustomFormula _ -> { common with Formula1 = ooxml.Formula1 }
    | _ ->
        // (A comparison without an operator in the file is a "between".)
        { common with Operator = Some (defaultArg ooxml.Operator "between"); Formula1 = ooxml.Formula1; Formula2 = ooxml.Formula2 }

/// "2024-03-31" (optionally with " 14:30[:15]" or "T14:30[:15]") as a date
/// serial, "14:30[:15]" as a fraction of a day; None for anything else.
let private dateOrTimeSerial (text: string) : float option =
    let text = text.Trim()
    let number (part: string) =
        if part.Length > 0 && part |> Seq.forall isAsciiDigit then Some (int part) else None
    let time (part: string) : float option =
        match part.Split(':') |> Array.map number with
        | [| Some h; Some m |] when h < 24 && m < 60 -> Some ((float h * 3600.0 + float m * 60.0) / 86400.0)
        | [| Some h; Some m; Some s |] when h < 24 && m < 60 && s < 60 -> Some ((float h * 3600.0 + float m * 60.0 + float s) / 86400.0)
        | _ -> None
    let date (part: string) : float option =
        match part.Split('-') |> Array.map number with
        | [| Some y; Some m; Some d |] when part.Length = 10 && y >= 1900 && m >= 1 && m <= 12 && d >= 1 && d <= 31 ->
            Some (Rbt.Spreadsheet.Input.dateSerial y m d)
        | _ -> None
    match text.Split([| 'T'; ' ' |]) with
    | [| one |] -> (match date one with Some serial -> Some serial | None -> time one)
    | [| day; clock |] -> (match date day, time clock with Some d, Some t -> Some (d + t) | _ -> None)
    | _ -> None

/// Check a described validation and give the model's validation. Throws
/// (naming `memberName`) for an unknown type or operator, a missing operand,
/// formula, list or source, a list value with a comma in it, and an unknown
/// error style.
let validationOfDescription (memberName: string) (description: ValidationDescription) : DataValidation =
    let fail (message: string) = failwith (sprintf "%s: %s" memberName message)
    let present (text: string option) = match text with Some text when text.Trim() <> "" -> true | _ -> false
    let style =
        match description.ErrorStyle with
        | "stop" -> StopStyle
        | "warning" -> WarningStyle
        | "information" -> InformationStyle
        | other -> fail (sprintf "unknown errorStyle \"%s\" (stop, warning or information)" other)
    let operand (kind: string) (what: string) (text: string option) : string =
        match text with
        | Some text when text.Trim() <> "" ->
            let body = formulaBody (text.Trim())
            (match (if kind = "date" || kind = "time" then dateOrTimeSerial body else None) with
             | Some serial -> numberOperand serial
             | None -> body)
        | _ -> fail (sprintf "a \"%s\" rule needs %s" kind what)
    let rule =
        match description.Type with
        | "any" -> AnyValue
        | "list" ->
            (match description.Values, description.Source with
             | Some _, Some source when source.Trim() <> "" -> fail "a list rule takes either values or a source, not both"
             | Some values, _ ->
                let values = values |> List.map (fun value -> value.Trim()) |> List.filter (fun value -> value <> "")
                if values.IsEmpty then fail "a list rule needs at least one value"
                match values |> List.tryFind (fun value -> value.Contains ",") with
                | Some value -> fail (sprintf "the list value \"%s\" holds a comma; a list with such values has to come from a source range" value)
                | None -> ListValues values
             | None, Some source when source.Trim() <> "" -> ListSource (formulaBody (source.Trim()))
             | _ -> fail "a list rule needs values (an array of texts) or a source (a range or a name)")
        | "custom" ->
            if not (present description.Formula1) then fail "a \"custom\" rule needs a formula"
            CustomFormula (formulaBody (description.Formula1.Value.Trim()))
        | "whole" | "decimal" | "date" | "time" | "textLength" as kind ->
            let operator =
                match description.Operator with
                | Some operator when List.contains operator validationOperators -> operator
                | Some operator -> fail (sprintf "unknown operator \"%s\" (%s)" operator (String.concat ", " validationOperators))
                | None -> fail (sprintf "a \"%s\" rule needs an operator (%s)" kind (String.concat ", " validationOperators))
            let two = operator = "between" || operator = "notBetween"
            let first = operand kind (if two then "formula1 (the lower bound)" else "formula1 (the value to compare with)") description.Formula1
            let second = if two then Some (operand kind "formula2 (the upper bound)" description.Formula2) else None
            let ooxml : CF.OoxmlValidation =
                { Type = kind; Operator = Some operator; Formula1 = Some first; Formula2 = second
                  AllowBlank = true; ShowErrorMessage = true; ShowInputMessage = true; ShowDropDown = true
                  ErrorStyle = None; ErrorTitle = None; Error = None; PromptTitle = None; Prompt = None }
            (match CF.validationRuleOfOoxml ooxml with
             | Some rule -> rule
             | None -> fail (sprintf "the \"%s\" rule could not be read" kind))
        | other -> fail (sprintf "unknown validation type \"%s\" (%s)" other (String.concat ", " validationTypes))
    { Rule = rule
      ErrorMessage = description.ErrorMessage
      InputMessage = description.InputMessage
      ErrorTitle = description.ErrorTitle
      InputTitle = description.InputTitle
      Style = style
      ShowError = description.ShowError
      AllowBlank = description.AllowBlank
      ShowInput = description.ShowInput
      ShowDropdown = description.Dropdown
      Anchor = None }

/// The validation with an id.
let tryFindValidation (id: int) (model: SpreadsheetModel) : ValidationRange option =
    model.Validations |> List.tryFind (fun entry -> entry.Id = id)

/// Reads the source of a list rule in a workbook: a range of the rule's own
/// sheet or of another one (`lookupSheet` gives a sheet by name), or a
/// defined name (`lookupName` gives what it refers to) that is such a range.
let listSourceReader (lookupSheet: string -> SpreadsheetModel option) (lookupName: string -> string option) : ListSourceReader =
    fun model deltaRow deltaCol source ->
        let rec resolve (depth: int) (text: string) : string list option =
            let text = formulaBody (text.Trim())
            match splitSheetPrefix text with
            | Some sheetName, rest -> lookupSheet sheetName |> Option.bind (fun other -> sameSheetListValues other deltaRow deltaCol rest)
            | None, rest ->
                match (if depth < 4 then lookupName rest else None) with
                | Some refersTo -> resolve (depth + 1) refersTo
                | None -> sameSheetListValues model deltaRow deltaCol rest
        resolve 0 source

/// The values the list validation on a cell offers (in order, as texts), or
/// None when the cell has no list validation or its source cannot be read.
let listValuesAt (reader: ListSourceReader) (addr: CellAddress) (model: SpreadsheetModel) : string list option =
    match validationEntryAt addr model with
    | Some entry ->
        (match entry.Validation.Rule with
         | ListValues values -> Some values
         | ListSource source ->
            let anchor = validationAnchor entry
            reader model (addr.Row - anchor.Row) (addr.Col - anchor.Col) source
         | _ -> None)
    | None -> None

// ---- defined names ----------------------------------------------------------------

/// What is wrong with the refers-to text of a defined name (with or without
/// a leading `=`), or None when it reads as a formula.
let refersToProblem (refersTo: string) : string option =
    let text = if isNull refersTo then "" else refersTo.Trim()
    let body = if text.StartsWith "=" then text.Substring 1 else text
    if body.Trim() = "" then Some "a name has to refer to something"
    else
        match Rbt.Spreadsheet.Formula.parseFormula ("=" + body) with
        | Ok _ -> None
        | Result.Error problem -> Some (sprintf "\"%s\" is not a reference or formula (%s)" body problem)

/// The areas of a reference list ("Sheet1!$A$1:$A$3,Sheet1!$C$1" -> two):
/// split at the commas outside quotes and parentheses.
let splitAreas (text: string) : string list =
    let parts = ResizeArray<string>()
    let current = System.Text.StringBuilder()
    let mutable quoted = false
    let mutable depth = 0
    for c in (if isNull text then "" else text) do
        if c = '\'' then quoted <- not quoted
        if not quoted && c = '(' then depth <- depth + 1
        if not quoted && c = ')' then depth <- depth - 1
        if c = ',' && not quoted && depth = 0 then
            parts.Add (current.ToString().Trim())
            current.Clear() |> ignore
        else current.Append c |> ignore
    parts.Add (current.ToString().Trim())
    parts |> Seq.filter (fun part -> part <> "") |> List.ofSeq
