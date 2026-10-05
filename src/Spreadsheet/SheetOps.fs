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
