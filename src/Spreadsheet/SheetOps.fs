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

/// The hyperlinks of a sheet laid out for looking one up by cell: the links
/// on a single cell by that cell, the others in a list (a sheet can have
/// thousands of links, and the grid asks for every cell it draws).
type private HyperlinkIndex = {
    Source: SheetHyperlink list
    Single: System.Collections.Generic.Dictionary<float, SheetHyperlink>
    Wide: SheetHyperlink array
}

let mutable private hyperlinkIndexCache : HyperlinkIndex list = []

let private hyperlinkIndexOf (links: SheetHyperlink list) : HyperlinkIndex =
    match hyperlinkIndexCache |> List.tryFind (fun index -> obj.ReferenceEquals(index.Source, links)) with
    | Some index -> index
    | None ->
        let single = System.Collections.Generic.Dictionary<float, SheetHyperlink>()
        let wide = ResizeArray<SheetHyperlink>()
        // (In order, so that of two links on a cell the later one is found.)
        for link in links do
            if link.Range.From = link.Range.To then single.[float link.Range.From.Row * 16384.0 + float link.Range.From.Col] <- link
            else wide.Add link
        let index = { Source = links; Single = single; Wide = wide.ToArray() }
        hyperlinkIndexCache <- index :: (hyperlinkIndexCache |> List.truncate 3)
        index

/// The hyperlink on a cell: the last one defined whose range holds it.
let hyperlinkAt (model: SpreadsheetModel) (addr: CellAddress) : SheetHyperlink option =
    if model.Hyperlinks.IsEmpty then None
    else
        let index = hyperlinkIndexOf model.Hyperlinks
        // (A link over a range defined after a single-cell one wins only by
        // position in the list; single cells are the common case and are
        // looked at first when no range holds the cell.)
        let fromWide = if index.Wide.Length = 0 then None else index.Wide |> Array.tryFindBack (fun link -> cellRangeContains addr link.Range)
        match index.Single.TryGetValue (float addr.Row * 16384.0 + float addr.Col) with
        | true, single ->
            (match fromWide with
             | Some wide when (model.Hyperlinks |> List.findIndexBack (fun link -> obj.ReferenceEquals(link, wide))) > (model.Hyperlinks |> List.findIndexBack (fun link -> obj.ReferenceEquals(link, single))) -> Some wide
             | _ -> Some single)
        | false, _ -> fromWide

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
    // (The rows an AutoFilter hides are left as they are.)
    if area <= cellByCellLimit then
        for row in top .. bottom do
            if not (isFilterHiddenRow row model) then
                for col in left .. right do
                    let addr = { Row = row; Col = col }
                    match Map.tryFind addr model.Cells with
                    | Some data -> change addr data
                    | None -> ()
    else
        for KeyValue(addr, data) in model.Cells do
            if addr.Row >= top && addr.Row <= bottom && addr.Col >= left && addr.Col <= right && not (isFilterHiddenRow addr.Row model) then change addr data
    // Comment threads are beside the cells: "clear all" takes them too.
    let threads =
        if what <> ClearEverything || Map.isEmpty model.Threads then model.Threads
        else
            let kept = model.Threads |> Map.filter (fun addr _ -> not (addr.Row >= top && addr.Row <= bottom && addr.Col >= left && addr.Col <= right && not (isFilterHiddenRow addr.Row model)))
            if kept.Count = model.Threads.Count then model.Threads else kept
    if obj.ReferenceEquals(cells, model.Cells) && obj.ReferenceEquals(runs, model.RichRuns) && obj.ReferenceEquals(threads, model.Threads) then model
    else { model with Cells = cells; RichRuns = runs; Threads = threads } |> invalidateFilterProjection

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
    /// The options of an Excel 2010+ data bar (`DataBarEx`); None: a plain data bar.
    Bar: DataBarOptions option
    /// Custom icons / hidden values of an icon set (`IconSetEx`); None: the family's icons, values shown.
    Icons: IconSetOptions option
    /// The period of a "dateOccurring" rule.
    Period: string option
    Style: CellFormat option
}

let emptyRuleDescription (kind: string) : RuleDescription =
    { Type = kind; Operator = None; Formula1 = None; Formula2 = None; Text = None; Rank = 10; Percent = false; Bottom = false
      Above = true; OrEqual = false; StdDev = 0; Stops = []; Low = None; High = None; Color = None; IconSet = None
      Thresholds = []; Reverse = false; Bar = None; Icons = None; Period = None; Style = None }

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
    | ColorScale _ | DataBar _ | IconSet _ | DataBarEx _ | IconSetEx _ -> true
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
    | CfvoAutoMin -> "autoMin"
    | CfvoAutoMax -> "autoMax"

let private describeThreshold (threshold: CfThreshold) : ThresholdDescription =
    { Kind = thresholdKindText threshold.Kind
      Value = (match threshold.Kind with CfvoMin | CfvoMax | CfvoAutoMin | CfvoAutoMax -> None | _ -> threshold.Value) }

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
    | DataBarEx (low, high, color, options) ->
        { emptyRuleDescription "dataBar" with Low = Some (describeThreshold low); High = Some (describeThreshold high); Color = Some color; Bar = Some options }
    | IconSetEx (family, thresholds, reverse, options) ->
        { emptyRuleDescription "iconSet" with IconSet = Some family; Thresholds = thresholds |> List.map describeThreshold; Reverse = reverse; Icons = Some options }

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
        | "autoMin", _ -> { Kind = CfvoAutoMin; Value = None }
        | "autoMax", _ -> { Kind = CfvoAutoMax; Value = None }
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
            let barColor = color "rule.color" (description.Color |> Option.defaultValue "#638EC6")
            (match description.Bar with
             | None -> DataBar (threshold "rule.min" anyKind low, threshold "rule.max" anyKind high, barColor)
             | Some options ->
                // The data bar of Excel 2010+: its ends may be automatic.
                let optional (what: string) (value: string option) = value |> Option.map (color what)
                if options.MinLength < 0 || options.MaxLength > 100 || options.MinLength > options.MaxLength then
                    fail (sprintf "rule.minLength and rule.maxLength must be 0..100, the first not above the second, got %d and %d" options.MinLength options.MaxLength)
                DataBarEx (threshold "rule.min" ("autoMin" :: anyKind) low, threshold "rule.max" ("autoMax" :: anyKind) high, barColor,
                           { options with
                                Border = optional "rule.borderColor" options.Border
                                NegativeFill = optional "rule.negativeColor" options.NegativeFill
                                NegativeBorder = optional "rule.negativeBorderColor" options.NegativeBorder
                                AxisColor = optional "rule.axisColor" options.AxisColor }))
        | "iconSet" ->
            let family = description.IconSet |> Option.defaultValue "3TrafficLights1"
            let known = iconSetNames @ Rbt.Spreadsheet.ConditionalFormatting.extendedIconSets
            if not (List.contains family known) then
                fail (sprintf "rule.iconSet must be one of %s, got \"%s\"" (String.concat ", " known) family)
            let count = int family.[0] - int '0'
            let thresholds = if description.Thresholds.IsEmpty then defaultIconThresholds count else description.Thresholds
            if thresholds.Length <> count then
                fail (sprintf "icon set \"%s\" needs %d thresholds (one per icon, the first is the floor), got %d" family count thresholds.Length)
            let steps = thresholds |> List.mapi (fun index source -> threshold (sprintf "rule.thresholds[%d]" index) [ "number"; "percent"; "percentile"; "formula" ] source)
            let options = description.Icons |> Option.defaultValue defaultIconSetOptions
            if not options.Custom.IsEmpty then
                if options.Custom.Length <> count then
                    fail (sprintf "rule.icons needs %d icons (one per step, the lowest first), got %d" count options.Custom.Length)
                if description.Reverse then fail "rule.reverse does not go with rule.icons (custom icons are given in the order they show)"
                options.Custom |> List.iteri (fun index (iconFamily, icon) ->
                    if iconFamily <> "NoIcons" then
                        if not (List.contains iconFamily known) then
                            fail (sprintf "rule.icons[%d].iconSet must be \"NoIcons\" or one of %s, got \"%s\"" index (String.concat ", " known) iconFamily)
                        if icon < 0 || icon >= int iconFamily.[0] - int '0' then
                            fail (sprintf "rule.icons[%d].index must be 0..%d for \"%s\", got %d" index (int iconFamily.[0] - int '0' - 1) iconFamily icon))
            if options = defaultIconSetOptions && not (List.contains family Rbt.Spreadsheet.ConditionalFormatting.extendedIconSets) then
                IconSet (family, steps, description.Reverse)
            else IconSetEx (family, steps, description.Reverse, options)
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
let private dateOrTimeSerial (date1904: bool) (text: string) : float option =
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
            Some (Rbt.Spreadsheet.Input.dateSerialIn date1904 y m d)
        | _ -> None
    match text.Split([| 'T'; ' ' |]) with
    | [| one |] -> (match date one with Some serial -> Some serial | None -> time one)
    | [| day; clock |] -> (match date day, time clock with Some d, Some t -> Some (d + t) | _ -> None)
    | _ -> None

/// Check a described validation and give the model's validation. Throws
/// (naming `memberName`) for an unknown type or operator, a missing operand,
/// formula, list or source, a list value with a comma in it, and an unknown
/// error style.
let validationOfDescriptionIn (date1904: bool) (memberName: string) (description: ValidationDescription) : DataValidation =
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
            (match (if kind = "date" || kind = "time" then dateOrTimeSerial date1904 body else None) with
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

/// `validationOfDescriptionIn` in the 1900 date system.
let validationOfDescription (memberName: string) (description: ValidationDescription) : DataValidation =
    validationOfDescriptionIn false memberName description

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

/// `listSourceReader` for a sheet of a workbook, which also reads a source
/// that is a formula - `OFFSET(...)`, `INDEX(...):INDEX(...)`, `INDIRECT(...)`,
/// or a defined name whose refers-to is one (a dynamic range) or a range of
/// several areas' worth of cells on any sheet: the formula is evaluated
/// with the workbook in reach and the cells (or array items) it gives are
/// the list, blanks and errors left out. `workbook` and `sheet` are asked
/// each time (the workbook as it is now, the id of the rule's sheet).
let listSourceReaderIn (workbook: unit -> Rbt.Spreadsheet.Workbook.Workbook option) (sheet: unit -> Rbt.Spreadsheet.Workbook.SheetId option)
                       (lookupSheet: string -> SpreadsheetModel option) (lookupName: string -> string option) : ListSourceReader =
    let plain = listSourceReader lookupSheet lookupName
    fun model deltaRow deltaCol source ->
        match plain model deltaRow deltaCol source with
        | Some values -> Some values
        | None ->
            let cell = { Row = max 0 deltaRow; Col = max 0 deltaCol }
            match Rbt.Spreadsheet.Formula.evaluateFormulaValuesShifted (workbook ()) (sheet ()) ("=" + formulaBody (source.Trim())) deltaRow deltaCol model cell with
            | None -> None
            | Some values ->
                // (A formula that does not give cells or items - a name that
                // is not there, a reference to nowhere - is one error.)
                match values with
                | [ CellValue.Error _ ] -> None
                | _ ->
                    Some [ for value in values do
                            match value with
                            | CellValue.Empty | CellValue.Error _ -> ()
                            | other ->
                                let shown = listItemText other
                                if shown <> "" then yield shown ]

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

// ---------------------------------------------------------------------------
// Rows and columns: sizes, hiding, auto-fit, outline
// ---------------------------------------------------------------------------
// Every function takes `rows`: true for rows, false for columns. Sizes are
// sheet pixels (what the model holds, at 100 % zoom). A hidden row / column
// has size 0 in the model's size map; `Axes` keeps the size it had.

[<Literal>]
let defaultRowHeightPx = 25.0
[<Literal>]
let defaultColumnWidthPx = 100.0
/// Excel's limits: a row of 409.5 pt, a column of 255 characters.
[<Literal>]
let maxRowHeightPx = 546.0
[<Literal>]
let maxColumnWidthPx = 1785.0
/// The height of a row that holds one line of the default font (11 pt):
/// what auto-fit gives a row with nothing taller in it.
[<Literal>]
let autoRowHeightPx = 20.0

let private axisSizes (rows: bool) (model: SpreadsheetModel) = if rows then model.RowHeights else model.ColumnWidths
let private axisDefault (rows: bool) (model: SpreadsheetModel) = if rows then model.DefaultRowHeight else model.DefaultColumnWidth
let private axisCount (rows: bool) (model: SpreadsheetModel) = if rows then model.RowCount else model.ColCount
let private keptSizes (rows: bool) (model: SpreadsheetModel) = if rows then model.Axes.HiddenRowSizes else model.Axes.HiddenColSizes

let private withAxis (rows: bool) (sizes: Map<int, float>) (kept: Map<int, float>) (model: SpreadsheetModel) : SpreadsheetModel =
    let axes =
        if obj.ReferenceEquals(kept, keptSizes rows model) then model.Axes
        elif rows then { model.Axes with HiddenRowSizes = kept }
        else { model.Axes with HiddenColSizes = kept }
    if rows then
        if obj.ReferenceEquals(sizes, model.RowHeights) && obj.ReferenceEquals(axes, model.Axes) then model
        else { model with RowHeights = sizes; Axes = axes }
    else
        if obj.ReferenceEquals(sizes, model.ColumnWidths) && obj.ReferenceEquals(axes, model.Axes) then model
        else { model with ColumnWidths = sizes; Axes = axes }

/// The size of a row / column as the sheet shows it: 0 when it is hidden.
let axisSize (rows: bool) (index: int) (model: SpreadsheetModel) : float =
    match Map.tryFind index (axisSizes rows model) with
    | Some size -> max 0.0 size
    | None -> axisDefault rows model

let isAxisHidden (rows: bool) (index: int) (model: SpreadsheetModel) : bool =
    axisSize rows index model <= 0.0

/// The size a row / column has, or - hidden - goes back to when it is shown.
let shownAxisSize (rows: bool) (index: int) (model: SpreadsheetModel) : float =
    let size = axisSize rows index model
    if size > 0.0 then size
    else Map.tryFind index (keptSizes rows model) |> Option.defaultValue (axisDefault rows model)

/// The rows / columns `first..last` clipped to the sheet (None when none is on it).
let clipAxisSpan (rows: bool) (first: int) (last: int) (model: SpreadsheetModel) : (int * int) option =
    let first, last = max 0 (min first last), min (axisCount rows model - 1) (max first last)
    if first <= last then Some (first, last) else None

/// Give the rows / columns `first..last` a size. Size 0 hides them (see
/// `hideAxis`); any other size also shows a hidden one.
let rec setAxisSize (rows: bool) (first: int) (last: int) (size: float) (model: SpreadsheetModel) : SpreadsheetModel =
    if size <= 0.0 then hideAxis rows first last model
    else
        match clipAxisSpan rows first last model with
        | None -> model
        | Some (first, last) ->
            let mutable sizes = axisSizes rows model
            let mutable kept = keptSizes rows model
            let isDefault = size = axisDefault rows model
            for index in first .. last do
                match Map.tryFind index sizes with
                | Some existing when existing = size -> ()
                | Some _ when isDefault -> sizes <- Map.remove index sizes
                | None when isDefault -> ()
                | _ -> sizes <- Map.add index size sizes
                if not kept.IsEmpty && Map.containsKey index kept then kept <- Map.remove index kept
            withAxis rows sizes kept model

/// Hide the rows / columns `first..last`: each keeps the size it had for
/// when it is shown again.
and hideAxis (rows: bool) (first: int) (last: int) (model: SpreadsheetModel) : SpreadsheetModel =
    match clipAxisSpan rows first last model with
    | None -> model
    | Some (first, last) ->
        let mutable sizes = axisSizes rows model
        let mutable kept = keptSizes rows model
        for index in first .. last do
            match Map.tryFind index sizes with
            | Some size when size <= 0.0 -> ()
            | Some size ->
                kept <- Map.add index size kept
                sizes <- Map.add index 0.0 sizes
            | None -> sizes <- Map.add index 0.0 sizes
        withAxis rows sizes kept model

/// Show the hidden rows / columns among `first..last` at the size they had.
let unhideAxis (rows: bool) (first: int) (last: int) (model: SpreadsheetModel) : SpreadsheetModel =
    let first, last = min first last, max first last
    let mutable sizes = axisSizes rows model
    let mutable kept = keptSizes rows model
    // (The hidden ones are in the size map: a range of a million rows costs what the map holds.)
    for KeyValue(index, size) in axisSizes rows model do
        if index >= first && index <= last && size <= 0.0 then
            match Map.tryFind index kept with
            | Some previous when previous > 0.0 && previous <> axisDefault rows model -> sizes <- Map.add index previous sizes
            | _ -> sizes <- Map.remove index sizes
            if Map.containsKey index kept then kept <- Map.remove index kept
    withAxis rows sizes kept model

/// The hidden rows / columns of a sheet as runs (first, last), in order.
/// Rows a filter hides are not among them.
let hiddenRuns (rows: bool) (model: SpreadsheetModel) : (int * int) list =
    let count = axisCount rows model
    let runs = ResizeArray<int * int>()
    for KeyValue(index, size) in axisSizes rows model do
        if size <= 0.0 && index >= 0 && index < count then
            if runs.Count > 0 && snd runs.[runs.Count - 1] = index - 1 then runs.[runs.Count - 1] <- (fst runs.[runs.Count - 1], index)
            else runs.Add ((index, index))
    List.ofSeq runs

/// "3:5" / "B:D" for a run of rows / columns.
let axisRunText (rows: bool) (first: int, last: int) : string =
    if rows then sprintf "%d:%d" (first + 1) (last + 1)
    else columnToLetter first + ":" + columnToLetter last

// --- auto-fit -----------------------------------------------------------------

/// Width of a text in pixels: text, font size (points), font family, bold,
/// italic (the shape of `Layout.TextMeasurer`).
type TextMeasure = string -> float -> string option -> bool -> bool -> float

/// Pixels for a font size in points: 96 px to the inch, 72 pt. Excel at 100 %
/// draws an 11 pt font 14.67 px tall.
let pointsToPx (points: float) = points * 96.0 / 72.0
/// The height of a line of text of a font size (points) as Excel auto-fits a
/// row to it: 20 px for 11 pt (15 pt, the default row), 21 for 12, 25 for 14,
/// 28 for 16, 35 for 20, 42 for 24; a row of n lines is n of them (Excel
/// 2019, Calibri; within a pixel a line elsewhere).
let lineHeightPxOfPoints (points: float) = Math.Ceiling (pointsToPx points * 1.3 - 1e-9)
let private cellFontPt (format: CellFormat) = float (defaultArg format.FontSize 11)
/// The height of a line of a cell's text, as the grid draws it.
let cellLineHeightPx (format: CellFormat) = lineHeightPxOfPoints (cellFontPt format)
let private cellIndentPx (format: CellFormat) =
    if format.Indent > 0 then float format.Indent * Rbt.Spreadsheet.CellText.mdwPx (cellFontPt format) * 1.4 else 0.0
/// A cell's text sits 5 px from its left and right edge, 3 px from top and
/// bottom (turned text; a line's height has room above and below it already).
[<Literal>]
let private cellPadX = 10.0
[<Literal>]
let private cellPadY = 6.0

let private shownText (model: SpreadsheetModel) (data: CellData) : string =
    Rbt.Spreadsheet.CellText.cellDisplayTextIn model.Date1904 None 1.0e6 data.Format (viewedValue model data.Value)

let private isTurned (format: CellFormat) = format.TextRotation = 90 || format.TextRotation = 180

/// The number of lines a text takes when it is wrapped to `width` pixels:
/// broken at spaces, a word wider than the width broken where it must be,
/// and a line break in the text always starts a line.
let wrappedLineCount (measure: string -> float) (width: float) (text: string) : int =
    let mutable lines = 0
    for paragraph in (if isNull text then "" else text).Replace("\r\n", "\n").Split('\n') do
        if paragraph = "" || width <= 0.0 || measure paragraph <= width then lines <- lines + 1
        else
            let space = measure " "
            let mutable used = 0.0
            lines <- lines + 1
            for word in paragraph.Split(' ') do
                let wordWidth = measure word
                if used > 0.0 && used + space + wordWidth <= width then used <- used + space + wordWidth
                elif wordWidth <= width then
                    if used > 0.0 then lines <- lines + 1
                    used <- wordWidth
                else
                    // A word longer than a line: as many lines as it takes, the rest goes on.
                    if used > 0.0 then lines <- lines + 1
                    let mutable rest = 0.0
                    for c in word do
                        let w = measure (string c)
                        if rest > 0.0 && rest + w > width then
                            lines <- lines + 1
                            rest <- w
                        else rest <- rest + w
                    used <- rest
    max 1 lines

/// The width each column of `first..last` needs to show the text of its
/// cells on one line: (column, width) for the columns that hold text, in
/// order - a column without any keeps its width. Measured is what a cell
/// displays (number format applied) in its font, plus its indent and the
/// cell's padding. A wrapped cell counts with its longest line break-free
/// line, text turned by 90 degrees with the height of a line, and a cell
/// merged across columns not at all (as in Excel). Hidden rows and columns
/// are left out. Only the cells that exist are looked at.
let autoFitColumnWidths (measure: TextMeasure) (model: SpreadsheetModel) (first: int, last: int) : (int * float) list =
    let needed = Collections.Generic.Dictionary<int, float>()
    let hasMerges = not model.MergedCells.IsEmpty
    let hiddenRows = model.RowHeights |> Map.exists (fun _ size -> size <= 0.0)
    for KeyValue(addr, data) in model.Cells do
        if addr.Col >= first && addr.Col <= last then
            match data.Value with
            | Empty | String "" -> ()
            | _ ->
                let merged =
                    hasMerges
                    && (match findMergedCell addr model with
                        | Some merge -> merge.EndCol > merge.StartCol || merge.StartRow <> addr.Row
                        | None -> false)
                let hidden = (hiddenRows && isAxisHidden true addr.Row model) || isFilterHiddenRow addr.Row model
                if not merged && not hidden then
                    let format = data.Format
                    let text = shownText model data
                    // (No glyph is wider than its font is tall and a fifth: a text that
                    // cannot be wider than what the column needs already is not measured.)
                    let known = match needed.TryGetValue addr.Col with | true, known -> known | false, _ -> 0.0
                    let extra = cellPadX + cellIndentPx format
                    if text <> "" && (isTurned format || float text.Length * pointsToPx (cellFontPt format) * 1.2 + extra > known) then
                        let width =
                            if isTurned format then cellLineHeightPx format
                            else
                                let line (part: string) = measure part (cellFontPt format) format.FontName format.Bold format.Italic
                                if text.IndexOf '\n' >= 0 then text.Replace("\r\n", "\n").Split('\n') |> Array.map line |> Array.max
                                else line text
                        let width = Math.Ceiling width + extra
                        if width > known then needed.[addr.Col] <- width
    [ for KeyValue(col, width) in needed do
        if not (isAxisHidden false col model) then yield col, min maxColumnWidthPx (Math.Ceiling width) ]
    |> List.sortBy fst

/// The height each row of `first..last` needs: (row, height) for every row
/// that is not hidden. A row's height is that of its tallest cell - the
/// lines of its text (one, or as many as wrapped text takes at the width of
/// its column, or of its merged columns) at the line height of its font
/// (`lineHeightPxOfPoints`; turned text: its length plus the cell's padding) -
/// and at least `autoRowHeightPx`. Cells merged
/// across rows do not count. Only the cells that exist are looked at.
let autoFitRowHeights (measure: TextMeasure) (model: SpreadsheetModel) (first: int, last: int) : (int * float) list =
    let first, last = max 0 first, min last (model.RowCount - 1)
    let needed = Collections.Generic.Dictionary<int, float>()
    let hasMerges = not model.MergedCells.IsEmpty
    for KeyValue(addr, data) in model.Cells do
        if addr.Row >= first && addr.Row <= last then
            match data.Value with
            | Empty | String "" -> ()
            | _ ->
                let merge = if hasMerges then findMergedCell addr model else None
                let skip =
                    match merge with
                    | Some merge -> merge.EndRow > merge.StartRow || merge.StartCol <> addr.Col
                    | None -> false
                if not skip && not (isAxisHidden false addr.Col model) then
                    let format = data.Format
                    // (One line of text that is not wrapped or turned: its line height, whatever it says.
                    // Only a text can hold a line break.)
                    let oneLine = not format.WrapText && not (isTurned format) && (match unwrapFormulaValue data.Value with String _ -> false | _ -> true)
                    let text = if oneLine then "1" else shownText model data
                    if text <> "" then
                        let fontPt = cellFontPt format
                        let line (part: string) = measure part fontPt format.FontName format.Bold format.Italic
                        let height =
                            if isTurned format then Math.Ceiling (line text) + cellPadY
                            else
                                let lines =
                                    if format.WrapText then
                                        let lastCol = match merge with Some merge -> merge.EndCol | None -> addr.Col
                                        let mutable width = 0.0
                                        for col in addr.Col .. lastCol do width <- width + axisSize false col model
                                        wrappedLineCount line (width - cellPadX - cellIndentPx format) text
                                    else text.Replace("\r\n", "\n").Split('\n').Length
                                float lines * cellLineHeightPx format
                        match needed.TryGetValue addr.Row with
                        | true, known when known >= height -> ()
                        | _ -> needed.[addr.Row] <- height
    // (A span of a million rows: the rows with a size or a cell of their own; the others have the default already.)
    let rowsToSet =
        if last - first < 20000 then seq { first .. last }
        else
            seq { for KeyValue(row, _) in model.RowHeights do if row >= first && row <= last then yield row
                  yield! needed.Keys }
            |> Seq.distinct |> Seq.sort
    [ for row in rowsToSet do
        if not (isAxisHidden true row model) then
            let height =
                match needed.TryGetValue row with
                | true, height -> max autoRowHeightPx height
                | false, _ -> autoRowHeightPx
            yield row, min maxRowHeightPx height ]

/// Apply sizes (index, size) to rows / columns.
let applyAxisSizes (rows: bool) (sizes: (int * float) list) (model: SpreadsheetModel) : SpreadsheetModel =
    sizes |> List.fold (fun m (index, size) -> setAxisSize rows index index size m) model

// --- outline (Excel's Group / Ungroup) ---------------------------------------------

[<Literal>]
let maxOutlineLevel = 7

let private levelsOf (rows: bool) (model: SpreadsheetModel) = if rows then model.Axes.RowLevels else model.Axes.ColLevels
let private collapsedOf (rows: bool) (model: SpreadsheetModel) = if rows then model.Axes.CollapsedRows else model.Axes.CollapsedCols

// ---------------------------------------------------------------------------
// Spilled arrays
// ---------------------------------------------------------------------------

/// The ranges the dynamic arrays of a sheet fill, by anchor, for the spill
/// map they were last worked out for (one sheet's at a time: what a host
/// asks about is the sheet on display).
let mutable private spillRangesFor : obj = null
let private spillRanges = System.Collections.Generic.Dictionary<CellAddress, CellRange>()

/// The dynamic array a cell is part of: the formula cell that spills (the
/// anchor) and the range it fills, the anchor included. None for a cell
/// that is neither a spilling formula nor one it filled.
let spillRangeAt (addr: CellAddress) (model: SpreadsheetModel) : (CellAddress * CellRange) option =
    if Map.isEmpty model.Spill then None
    else
        if not (obj.ReferenceEquals(spillRangesFor, model.Spill)) then
            spillRanges.Clear()
            for KeyValue (ghost, anchor) in model.Spill do
                let range =
                    match spillRanges.TryGetValue anchor with
                    | true, known -> known
                    | false, _ -> { From = anchor; To = anchor }
                spillRanges.[anchor] <-
                    { From = { Row = min range.From.Row ghost.Row; Col = min range.From.Col ghost.Col }
                      To = { Row = max range.To.Row ghost.Row; Col = max range.To.Col ghost.Col } }
            spillRangesFor <- box model.Spill
        let anchor = match Map.tryFind addr model.Spill with Some anchor -> anchor | None -> addr
        match spillRanges.TryGetValue anchor with
        | true, range -> Some (anchor, range)
        | false, _ -> None

/// Whether the summary of a group comes after it (rows: below, columns: to
/// the right; Excel's default) or before it: the sheet's `outlinePr`.
let summaryAfter (rows: bool) (model: SpreadsheetModel) : bool =
    if rows then model.Axes.SummaryBelow else model.Axes.SummaryRight

/// The summary row / column of the group `first..last`: the one that carries
/// the group's collapsed mark and its +/- button. (-1 for a group at the
/// first row whose summary is above: it has none on the sheet.)
let summaryIndex (rows: bool) (first: int) (last: int) (model: SpreadsheetModel) : int =
    if summaryAfter rows model then last + 1 else first - 1

/// The outline level of a row / column: 0 when it is in no group.
let outlineLevel (rows: bool) (index: int) (model: SpreadsheetModel) : int =
    Map.tryFind index (levelsOf rows model) |> Option.defaultValue 0

/// Raise (`delta` 1: Group) or lower (-1: Ungroup) the outline level of the
/// rows / columns `first..last`, between 0 and 7.
let changeOutlineLevel (rows: bool) (first: int) (last: int) (delta: int) (model: SpreadsheetModel) : SpreadsheetModel =
    match clipAxisSpan rows first last model with
    | None -> model
    | Some (first, last) ->
        let mutable levels = levelsOf rows model
        if delta > 0 then
            for index in first .. last do
                let level = min maxOutlineLevel (outlineLevel rows index model + delta)
                levels <- Map.add index level levels
        else
            for KeyValue(index, level) in levelsOf rows model do
                if index >= first && index <= last then
                    let lowered = level + delta
                    levels <- if lowered <= 0 then Map.remove index levels else Map.add index lowered levels
        if levels = levelsOf rows model then model
        else
            // A summary row / column whose group is gone is no longer one.
            let after = summaryAfter rows model
            let hasGroupBefore index = if after then index > 0 && Map.containsKey (index - 1) levels else Map.containsKey (index + 1) levels
            let collapsed = collapsedOf rows model |> Set.filter hasGroupBefore
            if rows then { model with Axes = { model.Axes with RowLevels = levels; CollapsedRows = collapsed } }
            else { model with Axes = { model.Axes with ColLevels = levels; CollapsedCols = collapsed } }

/// The group a row / column is in at its own level: the run of neighbours
/// at that level or deeper, as (first, last, level). None outside any group.
let outlineGroupAt (rows: bool) (index: int) (model: SpreadsheetModel) : (int * int * int) option =
    let level = outlineLevel rows index model
    if level <= 0 then None
    else
        let mutable first = index
        while first > 0 && outlineLevel rows (first - 1) model >= level do first <- first - 1
        let mutable last = index
        while outlineLevel rows (last + 1) model >= level do last <- last + 1
        Some (first, last, level)

/// Collapse (hide) or expand (show) the group the row / column `index` is
/// in. The group's summary row / column (see `summaryIndex`) carries the
/// collapsed mark. Expanding leaves the
/// groups inside that are themselves collapsed as they are.
let setGroupCollapsed (memberName: string) (rows: bool) (index: int) (collapsed: bool) (model: SpreadsheetModel) : SpreadsheetModel =
    match outlineGroupAt rows index model with
    | None -> failwith (sprintf "%s: %s %d is not in a group" memberName (if rows then "row" else "column") index)
    | Some (first, last, level) ->
        let marks = collapsedOf rows model
        let withMarks (marks: Set<int>) (m: SpreadsheetModel) =
            if marks = collapsedOf rows m then m
            elif rows then { m with Axes = { m.Axes with CollapsedRows = marks } }
            else { m with Axes = { m.Axes with CollapsedCols = marks } }
        if collapsed then
            hideAxis rows first last model |> withMarks (Set.add (summaryIndex rows first last model) marks)
        else
            let marks = Set.remove (summaryIndex rows first last model) marks
            let mutable result = model
            let rec expand (first: int) (last: int) (level: int) =
                let mutable at = first
                while at <= last do
                    if outlineLevel rows at model <= level then
                        result <- unhideAxis rows at at result
                        at <- at + 1
                    else
                        let mutable stop = at
                        while stop < last && outlineLevel rows (stop + 1) model > level do stop <- stop + 1
                        if not (Set.contains (summaryIndex rows at stop model) marks) then expand at stop (level + 1)
                        at <- stop + 1
            expand first last level
            withMarks marks result

/// Every group of the rows / columns: (first, last, level, collapsed), outer
/// groups before the ones inside them.
let outlineGroups (rows: bool) (model: SpreadsheetModel) : (int * int * int * bool) list =
    let levels = levelsOf rows model
    if levels.IsEmpty then []
    else
        let marks = collapsedOf rows model
        let indices = levels |> Map.toArray
        let groups = ResizeArray<int * int * int * bool>()
        let deepest = indices |> Array.map snd |> Array.max
        for level in 1 .. deepest do
            let mutable start = -1
            let mutable previous = -2
            let close () = if start >= 0 then groups.Add ((start, previous, level, Set.contains (summaryIndex rows start previous model) marks))
            for index, own in indices do
                if own >= level then
                    if start < 0 || index <> previous + 1 then
                        close ()
                        start <- index
                    previous <- index
            close ()
        groups |> Seq.sortBy (fun (first, _, level, _) -> first, level) |> List.ofSeq

/// The deepest outline level of the rows / columns (0: no groups).
let outlineDepth (rows: bool) (model: SpreadsheetModel) : int =
    levelsOf rows model |> Map.fold (fun deepest _ level -> max deepest level) 0

/// Excel's outline level buttons (1, 2, 3 ...): show the rows / columns up
/// to `level` - 1 and collapse every group from `level` on. Level 1 shows
/// only what is in no group; one more than the deepest level shows all.
let showOutlineLevel (rows: bool) (level: int) (model: SpreadsheetModel) : SpreadsheetModel =
    let levels = levelsOf rows model
    if levels.IsEmpty then model
    else
        let groups = outlineGroups rows model
        let marks = groups |> List.choose (fun (first, last, own, _) -> if own >= level then Some (summaryIndex rows first last model) else None) |> Set.ofList
        // Runs of rows to hide (at `level` or deeper) and to show (above it).
        let runsWhere (wanted: int -> bool) =
            let runs = ResizeArray<int * int>()
            let mutable start = -1
            let mutable previous = -2
            for KeyValue(index, own) in levels do
                if wanted own then
                    if start < 0 || index <> previous + 1 then
                        if start >= 0 then runs.Add ((start, previous))
                        start <- index
                    previous <- index
            if start >= 0 then runs.Add ((start, previous))
            List.ofSeq runs
        let mutable result = model
        for (first, last) in runsWhere (fun own -> own < level) do result <- unhideAxis rows first last result
        for (first, last) in runsWhere (fun own -> own >= level) do result <- hideAxis rows first last result
        if marks = collapsedOf rows result then result
        elif rows then { result with Axes = { result.Axes with CollapsedRows = marks } }
        else { result with Axes = { result.Axes with CollapsedCols = marks } }

/// Say where the summary rows / columns of the outline are (the sheet's
/// `outlinePr`). The collapsed groups stay collapsed: their marks move to
/// the new summary row / column.
let setOutlineSummary (below: bool) (right: bool) (model: SpreadsheetModel) : SpreadsheetModel =
    if model.Axes.SummaryBelow = below && model.Axes.SummaryRight = right then model
    else
        let collapsedGroups (rows: bool) = outlineGroups rows model |> List.filter (fun (_, _, _, collapsed) -> collapsed)
        let rowGroups = collapsedGroups true
        let colGroups = collapsedGroups false
        let moved = { model with Axes = { model.Axes with SummaryBelow = below; SummaryRight = right } }
        let marksOf (rows: bool) (groups: (int * int * int * bool) list) =
            groups |> List.map (fun (first, last, _, _) -> summaryIndex rows first last moved) |> Set.ofList
        { moved with Axes = { moved.Axes with CollapsedRows = marksOf true rowGroups; CollapsedCols = marksOf false colGroups } }

/// Show or hide the outline's symbols on the sheet (the sheet view's `showOutlineSymbols`).
let setOutlineSymbols (show: bool) (model: SpreadsheetModel) : SpreadsheetModel =
    if model.Axes.ShowSymbols = show then model else { model with Axes = { model.Axes with ShowSymbols = show } }

// ---------------------------------------------------------------------------
// Sorting a range
// ---------------------------------------------------------------------------

/// What a sort key orders the rows by.
type SortBy =
    | SortByValue
    | SortByCellColor
    | SortByFontColor

/// One key of a sort: a sheet column (0-based) of the sorted range.
type SortKey = {
    Column: int
    Descending: bool
    By: SortBy
    /// For the colour sorts: the rows with this fill / text colour come
    /// first (last with `Descending`); "" = no fill / the automatic colour.
    Color: string
}

type SortOptions = {
    /// The first row of the range is a header row and stays where it is.
    HasHeader: bool
    /// Text that differs only in case is ordered too (lower case first).
    CaseSensitive: bool
}

/// The order of two cell values in an ascending sort, blanks apart: numbers
/// before text before FALSE before TRUE before errors.
let private compareSortValues (caseSensitive: bool) (a: CellValue) (b: CellValue) : int =
    let rank (value: CellValue) =
        match value with
        | Number _ -> 0
        | String _ -> 1
        | Boolean _ -> 2
        | Error _ -> 3
        | _ -> 4
    match a, b with
    | Number x, Number y -> compare x y
    | String x, String y ->
        let order = String.Compare(x, y, StringComparison.OrdinalIgnoreCase)
        if order <> 0 || not caseSensitive then order else -(String.CompareOrdinal(x, y))
    | Boolean x, Boolean y -> compare x y
    | _ -> compare (rank a) (rank b)

/// The rows of `top` .. `bottom` that take part in a sort: those the
/// AutoFilter does not hide (a hidden row stays where it is, as in Excel).
let private sortableRows (top: int) (bottom: int) (model: SpreadsheetModel) : int array =
    filterShownRows top bottom model |> Array.ofList

/// Why `range` cannot be sorted by `keys`, or None: no key, a key outside
/// the range, merged cells among the rows to sort, or a dynamic array (a
/// spilling formula or the cells it fills) among them.
let sortProblem (range: CellRange) (keys: SortKey list) (options: SortOptions) (model: SpreadsheetModel) : string option =
    let range = clipToUsed model range
    let top = if options.HasHeader then range.From.Row + 1 else range.From.Row
    let inRows (row: int) = row >= top && row <= range.To.Row
    let inCols (col: int) = col >= range.From.Col && col <= range.To.Col
    if keys.IsEmpty then Some "no sort key was given"
    else
        match keys |> List.tryFind (fun key -> not (inCols key.Column)) with
        | Some key -> Some (sprintf "the key column %s is outside the range %s" (columnToLetter key.Column) (rangeText range))
        | None ->
            let merged =
                model.MergedCells
                |> List.exists (fun merge -> merge.StartRow <= range.To.Row && merge.EndRow >= top && merge.StartCol <= range.To.Col && merge.EndCol >= range.From.Col)
            let spills =
                model.Spill
                |> Map.exists (fun ghost anchor -> (inRows ghost.Row && inCols ghost.Col) || (inRows anchor.Row && inCols anchor.Col))
            if merged then Some "the range holds merged cells"
            elif spills then Some "the range holds a dynamic array (a formula that spills, or the cells it fills)"
            else None

/// Sort the rows of `range` by `keys` (the first key decides, the next ones
/// break ties, rows that tie on every key keep their order). Only the cells
/// of the range move - whole rows of it, each with its value, format,
/// comment, rich text, threaded comment and single-cell hyperlink; cells
/// beside the range, row heights, validations and conditional formats stay.
/// A formula that moves is re-pointed as a copy of it would be: its relative
/// references shift by the rows it moved, `$`-locked rows stay. Formulas
/// elsewhere are not re-pointed. Blank cells sort last in either direction.
/// Rows the AutoFilter hides stay where they are. Throws (naming
/// `memberName`) when `sortProblem` finds one. Cached values of formulas are
/// stale afterwards: recalculate.
let sortRange (memberName: string) (range: CellRange) (keys: SortKey list) (options: SortOptions) (model: SpreadsheetModel) : SpreadsheetModel =
    match sortProblem range keys options model with
    | Some problem -> failwith (sprintf "%s: %s" memberName problem)
    | None ->
        let range = clipToUsed model range
        let top = if options.HasHeader then range.From.Row + 1 else range.From.Row
        let left, right = range.From.Col, range.To.Col
        let rows = sortableRows top range.To.Row model
        let sortState =
            match keys with
            | key :: _ -> Some { Column = key.Column; Ascending = not key.Descending }
            | [] -> model.SortState
        if rows.Length < 2 then { model with SortState = sortState }
        else
            let valueAt (row: int) (col: int) =
                match getCellAt row col model with
                | Some data -> unwrapFormulaValue data.Value
                | None -> CellValue.Empty
            let colorAt (font: bool) (row: int) (col: int) =
                match getCellAt row col model with
                | Some data -> ((if font then data.Format.TextColor else data.Format.BgColor) |> Option.defaultValue "").ToUpperInvariant()
                | None -> ""
            let isBlank (value: CellValue) = match value with CellValue.Empty | String "" -> true | _ -> false
            // What each key compares, read once per row (rows.[i] at i), not at every comparison.
            let keyColumns =
                keys
                |> List.map (fun key ->
                    match key.By with
                    | SortByValue ->
                        let values = rows |> Array.map (fun row -> valueAt row key.Column)
                        values, Array.map isBlank values, [||]
                    | SortByCellColor | SortByFontColor ->
                        let wanted = key.Color.ToUpperInvariant()
                        [||], [||], rows |> Array.map (fun row -> if colorAt (key.By = SortByFontColor) row key.Column = wanted then 0 else 1))
                |> Array.ofList
            let keyArray = Array.ofList keys
            let compareBy (index: int) (a: int) (b: int) : int =
                let key = keyArray.[index]
                let values, blanks, colors = keyColumns.[index]
                match key.By with
                | SortByValue ->
                    let va, vb = values.[a], values.[b]
                    match blanks.[a], blanks.[b] with
                    | true, true -> 0
                    | true, false -> 1
                    | false, true -> -1
                    | false, false ->
                        let order = compareSortValues options.CaseSensitive va vb
                        if key.Descending then -order else order
                | SortByCellColor | SortByFontColor ->
                    let order = compare colors.[a] colors.[b]
                    if key.Descending then -order else order
            // (Positions in `rows`; rows that tie on every key keep their order.)
            let comparePositions (a: int) (b: int) : int =
                let mutable order = 0
                let mutable index = 0
                while order = 0 && index < keyArray.Length do
                    order <- compareBy index a b
                    index <- index + 1
                if order <> 0 then order else compare a b
            let ordered = Array.init rows.Length id |> Array.sortWith comparePositions |> Array.map (fun position -> rows.[position])
            // ordered.[i] is the row whose cells land on rows.[i].
            let target = System.Collections.Generic.Dictionary<int, int>()
            let sourceOf = System.Collections.Generic.Dictionary<int, int>()
            let mutable moved = false
            for i in 0 .. rows.Length - 1 do
                target.[ordered.[i]] <- rows.[i]
                sourceOf.[rows.[i]] <- ordered.[i]
                if ordered.[i] <> rows.[i] then moved <- true
            if not moved then { model with SortState = sortState }
            else
                let moves (addr: CellAddress) = addr.Col >= left && addr.Col <= right && target.ContainsKey addr.Row
                // (The maps by address are made anew in one pass: `moveRowBlocks`.)
                let moveMap (items: Map<CellAddress, 'T>) (onMove: int -> 'T -> 'T) : Map<CellAddress, 'T> =
                    moveRowBlocks left right sourceOf onMove items
                let cells =
                    moveMap model.Cells (fun delta data ->
                        match data.Value with
                        | Formula (formula, cached) when delta <> 0 -> { data with Value = Formula (shiftFormulaReferences delta 0 formula, cached) }
                        | _ -> data)
                // (The formula cells of the sorted map, from those of the map before: no walk of the cells.)
                match tryFormulaIndex model.Cells with
                | Some index ->
                    let moved = index |> Set.map (fun addr -> if moves addr then { addr with Row = target.[addr.Row] } else addr)
                    registerFormulaIndex cells moved
                | None -> ()
                // What moved, for the recalculation: the rows that took another's cells.
                let sorted =
                    let pairs = [| for i in 0 .. rows.Length - 1 do if ordered.[i] <> rows.[i] then yield rows.[i], ordered.[i] |]
                    RowsSorted(left, right, Array.map fst pairs, Array.map snd pairs, model.Cells, cells)
                let hyperlinks =
                    if model.Hyperlinks.IsEmpty then model.Hyperlinks
                    else
                        model.Hyperlinks
                        |> List.map (fun link ->
                            if link.Range.From = link.Range.To && moves link.Range.From then
                                let addr = { link.Range.From with Row = target.[link.Range.From.Row] }
                                { link with Range = { From = addr; To = addr } }
                            else link)
                { model with
                    Cells = cells
                    RichRuns = moveMap model.RichRuns (fun _ runs -> runs)
                    ArrayMarks = followArrayMarks cells (moveMap model.ArrayMarks (fun _ mark -> mark))
                    Threads = moveMap model.Threads (fun _ thread -> thread)
                    Hyperlinks = hyperlinks
                    SortState = sortState
                    RowsSorted = Some sorted }
                |> invalidateFilterProjection

/// What a quick sort (Sort A to Z / Z to A with one cell selected) acts on:
/// the block of data around `active` (`currentRegion`; a table's totals row
/// left out) and whether its first row is a header (`regionHasHeader`).
/// None when there is no data around the cell.
let quickSortTarget (active: CellAddress) (model: SpreadsheetModel) : (CellRange * SortOptions) option =
    match currentRegion active model with
    | Some region when active.Col >= region.From.Col && active.Col <= region.To.Col ->
        let bottom =
            model.Tables
            |> List.tryPick (fun t ->
                if t.HasTotals && t.EndRow - 1 = region.To.Row && t.EndRow - 1 > region.From.Row
                   && active.Col >= t.StartCol - 1 && active.Col <= t.EndCol - 1 then Some (t.EndRow - 2)
                else None)
            |> Option.defaultValue region.To.Row
        Some ({ region with To = { region.To with Row = bottom } }, { HasHeader = regionHasHeader region model; CaseSensitive = false })
    | _ -> None

// ---------------------------------------------------------------------------
// AutoFilter criteria as plain data
// ---------------------------------------------------------------------------

/// One comparison of a "conditions" filter.
type FilterConditionDescription = {
    /// "equals" | "notEquals" | "greaterThan" | "greaterOrEqual" | "lessThan"
    /// | "lessOrEqual" | "contains" | "notContains" | "beginsWith" | "endsWith".
    Operator: string
    /// A number or text. With "equals" / "notEquals" text may hold the
    /// wildcards `*` and `?` (`~` before one means the character itself);
    /// the four text operators take their text literally.
    Value: string
}

/// The filter of a column as plain data; `Kind` says which fields count.
type FilterDescription = {
    /// "values" | "conditions" | "top" | "color" | "dynamic" | "kept".
    Kind: string
    Values: string list
    Blanks: bool
    Conditions: FilterConditionDescription list
    And: bool
    Count: int
    Percent: bool
    Bottom: bool
    Color: string
    /// "cellColor" | "fontColor".
    By: string
    /// "aboveAverage" | "belowAverage".
    Dynamic: string
}

let emptyFilterDescription (kind: string) : FilterDescription =
    { Kind = kind; Values = []; Blanks = false; Conditions = []; And = true; Count = 10; Percent = false; Bottom = false
      Color = ""; By = "cellColor"; Dynamic = "" }

let filterOperators =
    [ "equals"; "notEquals"; "greaterThan"; "greaterOrEqual"; "lessThan"; "lessOrEqual"; "contains"; "notContains"; "beginsWith"; "endsWith" ]

/// A pattern with its `~` escapes taken out, when it holds no wildcard.
let private literalOfPattern (pattern: string) : string option =
    let sb = System.Text.StringBuilder()
    let mutable i = 0
    let mutable literal = true
    while literal && i < pattern.Length do
        let c = pattern.[i]
        if c = '~' && i + 1 < pattern.Length && (pattern.[i + 1] = '*' || pattern.[i + 1] = '?' || pattern.[i + 1] = '~') then
            sb.Append(pattern.[i + 1]) |> ignore
            i <- i + 2
        elif c = '*' || c = '?' then literal <- false
        else
            sb.Append(c) |> ignore
            i <- i + 1
    if literal then Some (sb.ToString()) else None

let private describeCondition (op: CompareOperator, operand: string) : FilterConditionDescription =
    let plain name = { Operator = name; Value = operand }
    match op with
    | OpGreaterThan -> plain "greaterThan"
    | OpGreaterOrEqual -> plain "greaterOrEqual"
    | OpLessThan -> plain "lessThan"
    | OpLessOrEqual -> plain "lessOrEqual"
    | OpEqual | OpNotEqual ->
        let negative = op = OpNotEqual
        let starts = operand.Length >= 2 && operand.StartsWith "*"
        // (A closing `*` that is escaped is not a wildcard.)
        let ends = operand.Length >= 2 && operand.EndsWith "*" && not (operand.EndsWith "~*")
        let inner (text: string) = literalOfPattern text |> Option.filter (fun literal -> literal <> "")
        let text =
            if starts && ends && operand.Length >= 3 then inner (operand.Substring(1, operand.Length - 2)) |> Option.map (fun literal -> (if negative then "notContains" else "contains"), literal)
            elif ends && not negative then inner (operand.Substring(0, operand.Length - 1)) |> Option.map (fun literal -> "beginsWith", literal)
            elif starts && not negative then inner (operand.Substring 1) |> Option.map (fun literal -> "endsWith", literal)
            else None
        match text with
        | Some (name, literal) -> { Operator = name; Value = literal }
        | None -> plain (if negative then "notEquals" else "equals")
    | _ -> plain "equals"

/// A column filter as plain data.
let describeFilter (criteria: FilterCriteria) : FilterDescription =
    match criteria with
    | FilterValues (values, blanks) -> { emptyFilterDescription "values" with Values = Set.toList values; Blanks = blanks }
    | FilterCustom (conditions, matchAll) -> { emptyFilterDescription "conditions" with Conditions = conditions |> List.map describeCondition; And = matchAll }
    | FilterTop (count, percent, bottom) -> { emptyFilterDescription "top" with Count = count; Percent = percent; Bottom = bottom }
    | FilterColor (color, font) -> { emptyFilterDescription "color" with Color = color; By = (if font then "fontColor" else "cellColor") }
    | FilterDynamic kind -> { emptyFilterDescription "dynamic" with Dynamic = kind }
    | FilterKept _ -> emptyFilterDescription "kept"

/// Check a described filter and give the model's criteria. Throws (naming
/// `memberName`) for an unknown kind or operator, no or more than two
/// conditions, a count below 1 (or a percentage above 100), a colour that is
/// not "#RRGGBB" (or ""), and an unknown dynamic filter. A date written
/// "2024-03-31" (or a time "14:30") in a condition becomes its serial number.
let filterOfDescriptionIn (date1904: bool) (memberName: string) (description: FilterDescription) : FilterCriteria =
    let fail message = failwith (sprintf "%s: %s" memberName message)
    match description.Kind with
    | "values" ->
        let values = description.Values |> List.filter (fun value -> value <> "") |> Set.ofList
        let blanks = description.Blanks || List.contains "" description.Values
        FilterValues (values, blanks)
    | "conditions" ->
        if description.Conditions.IsEmpty then fail "a conditions filter needs at least one condition"
        elif description.Conditions.Length > 2 then fail "a conditions filter takes at most two conditions"
        else
            let conditions =
                description.Conditions
                |> List.map (fun condition ->
                    let value =
                        match dateOrTimeSerial date1904 condition.Value with
                        | Some serial when condition.Value.Contains "-" || condition.Value.Contains ":" -> numberOperand serial
                        | _ -> condition.Value
                    match condition.Operator with
                    | "equals" -> OpEqual, value
                    | "notEquals" -> OpNotEqual, value
                    | "greaterThan" -> OpGreaterThan, value
                    | "greaterOrEqual" -> OpGreaterOrEqual, value
                    | "lessThan" -> OpLessThan, value
                    | "lessOrEqual" -> OpLessOrEqual, value
                    | "contains" | "notContains" | "beginsWith" | "endsWith" when condition.Value = "" ->
                        fail (sprintf "the operator \"%s\" needs a text" condition.Operator)
                    | "contains" -> OpEqual, "*" + escapeWildcards condition.Value + "*"
                    | "notContains" -> OpNotEqual, "*" + escapeWildcards condition.Value + "*"
                    | "beginsWith" -> OpEqual, escapeWildcards condition.Value + "*"
                    | "endsWith" -> OpEqual, "*" + escapeWildcards condition.Value
                    | other -> fail (sprintf "unknown filter operator \"%s\" (one of %s)" other (String.concat ", " filterOperators)))
            FilterCustom (conditions, description.And || conditions.Length = 1)
    | "top" ->
        if description.Count < 1 then fail "the count of a top / bottom filter must be at least 1"
        elif description.Percent && description.Count > 100 then fail "a percentage cannot be above 100"
        else FilterTop (description.Count, description.Percent, description.Bottom)
    | "color" ->
        if description.Color <> "" && not (isHexColor description.Color) then fail (sprintf "\"%s\" is not a colour (\"#RRGGBB\", or \"\" for none)" description.Color)
        else
            match description.By with
            | "cellColor" -> FilterColor (description.Color.ToUpperInvariant(), false)
            | "fontColor" -> FilterColor (description.Color.ToUpperInvariant(), true)
            | other -> fail (sprintf "unknown colour filter \"%s\" (\"cellColor\" or \"fontColor\")" other)
    | "dynamic" ->
        (match description.Dynamic with
         | "aboveAverage" | "belowAverage" -> FilterDynamic description.Dynamic
         | other -> fail (sprintf "unknown dynamic filter \"%s\" (\"aboveAverage\" or \"belowAverage\")" other))
    | other -> fail (sprintf "unknown filter kind \"%s\"" other)

/// `filterOfDescriptionIn` in the 1900 date system.
let filterOfDescription (memberName: string) (description: FilterDescription) : FilterCriteria =
    filterOfDescriptionIn false memberName description

/// Check a range for an AutoFilter: clipped to the sheet, and - a whole
/// column or row being no place for a header - at least one cell.
let autoFilterRange (memberName: string) (range: CellRange) (model: SpreadsheetModel) : CellRange =
    match clipRange model.RowCount model.ColCount range with
    | Some clipped -> clipped
    | None -> failwith (sprintf "%s: the range %s is outside the sheet" memberName (rangeText range))

// --- comments ------------------------------------------------------------------
// A cell has a note (`CellData.Comment`: a text and an author), a comment
// thread (`SpreadsheetModel.Threads`: a first comment with replies), or
// neither. A thread wins where a cell has both: a file keeps a note beside
// every thread for the applications that cannot show threads.

/// The comment of a cell, as a host sees it.
type CommentInfo = {
    Cell: CellAddress
    /// The note's text, or the text of the thread's first comment.
    Text: string
    Author: string option
    /// The runs of a note whose text is formatted in runs; empty otherwise.
    Runs: (string * RunOverride) list
    Thread: CommentThread option
}

/// The comment on a cell: its thread when it has one, else its note.
let commentAt (model: SpreadsheetModel) (addr: CellAddress) : CommentInfo option =
    match (if Map.isEmpty model.Threads then None else Map.tryFind addr model.Threads) with
    | Some thread -> Some { Cell = addr; Text = thread.Text; Author = thread.Author; Runs = []; Thread = Some thread }
    | None ->
        match Map.tryFind addr model.Cells with
        | Some { Comment = Some note } -> Some { Cell = addr; Text = note.Text; Author = note.Author; Runs = note.Runs; Thread = None }
        | _ -> None

/// Every comment of a sheet, row by row.
let allComments (model: SpreadsheetModel) : CommentInfo list =
    let cells = System.Collections.Generic.HashSet<CellAddress>()
    for KeyValue(addr, _) in model.Threads do cells.Add addr |> ignore
    for KeyValue(addr, data) in model.Cells do
        if data.Comment.IsSome then cells.Add addr |> ignore
    cells
    |> Seq.sortBy (fun addr -> addr.Row, addr.Col)
    |> Seq.choose (commentAt model)
    |> List.ofSeq

/// The cell at `addr` with its note set or taken away, and the marker flag
/// of its format saying whether the cell has a comment of either kind.
let private withNote (addr: CellAddress) (note: CellComment option) (marked: bool) (model: SpreadsheetModel) : SpreadsheetModel =
    match Map.tryFind addr model.Cells with
    | Some data ->
        if data.Comment = note && data.Format.HasComment = marked then model
        else
            let format = if data.Format.HasComment = marked then data.Format else { data.Format with HasComment = marked }
            let changed = { data with Comment = note; Format = format }
            // (A cell that was only there for its comment goes with it.)
            if changed.Value = CellValue.Empty && changed.Comment.IsNone && changed.Format = defaultFormat then { model with Cells = Map.remove addr model.Cells }
            else { model with Cells = Map.add addr changed model.Cells }
    | None ->
        if note.IsNone && not marked then model
        else { model with Cells = Map.add addr { Value = CellValue.Empty; Format = { defaultFormat with HasComment = marked }; Comment = note } model.Cells }

/// The author a note set on a cell has: the one given, else the one the
/// cell's note has now (an edit of a note keeps its author).
let private noteAuthor (addr: CellAddress) (author: string option) (model: SpreadsheetModel) : string option =
    match author with
    | Some _ -> author
    | None -> Map.tryFind addr model.Cells |> Option.bind (fun cell -> cell.Comment) |> Option.bind (fun note -> note.Author)

/// Put a comment on a cell. A cell without one, or with a note, gets the
/// note (a note that is there keeps its author unless one is given); on a
/// cell with a thread the text (and, when given, the author) of the
/// thread's first comment changes and the replies stay.
let setComment (addr: CellAddress) (text: string) (author: string option) (model: SpreadsheetModel) : SpreadsheetModel =
    match Map.tryFind addr model.Threads with
    | Some thread ->
        let edited = { thread with Text = text; Author = (match author with Some _ -> author | None -> thread.Author); PersonId = (if author.IsSome && author <> thread.Author then None else thread.PersonId) }
        if edited = thread then model else { model with Threads = Map.add addr edited model.Threads }
    | None -> withNote addr (Some (plainNote text (noteAuthor addr author model))) true model

/// Put a note whose text is in runs on a cell (runs without text do not
/// count; with no formatting in any run it is a plain note). A thread on the
/// cell gets the joined text, as `setComment` gives it.
let setCommentRuns (addr: CellAddress) (runs: (string * RunOverride) list) (author: string option) (model: SpreadsheetModel) : SpreadsheetModel =
    let runs = runs |> List.filter (fun (text, _) -> text <> "")
    let text = runs |> List.map fst |> String.concat ""
    if Map.containsKey addr model.Threads then setComment addr text author model
    else
        let formatted = runs |> List.exists (fun (_, run) -> run <> emptyRunOverride)
        withNote addr (Some { Text = text; Author = noteAuthor addr author model; Runs = (if formatted then runs else []) }) true model

/// Take the comment off a cell: its note and its thread. The same model when it had neither.
let deleteComment (addr: CellAddress) (model: SpreadsheetModel) : SpreadsheetModel =
    let cleared = withNote addr None false model
    if Map.containsKey addr cleared.Threads then { cleared with Threads = Map.remove addr cleared.Threads } else cleared

/// An id for a new comment of a thread: the form the file wants (a GUID in
/// braces), not used by any comment of the sheet.
let freshCommentId (model: SpreadsheetModel) : string =
    let used = System.Collections.Generic.HashSet<string>()
    for KeyValue(_, thread) in model.Threads do
        used.Add (thread.Id.ToUpperInvariant()) |> ignore
        for reply in thread.Replies do used.Add (reply.Id.ToUpperInvariant()) |> ignore
    let mutable n = used.Count + 1
    let make (n: int) = sprintf "{5C0FFEE5-0000-4000-8000-%012d}" n
    while used.Contains (make n) do n <- n + 1
    make n

/// Add a reply to the comment of a cell; the changed model and the reply's
/// id. A note becomes the first comment of a thread (as Excel's "convert to
/// comments" does) and the reply follows it. `date` is the time of the
/// reply ("2024-03-31T14:30:00.00"). An Error when the cell has no comment.
let addCommentReply (addr: CellAddress) (text: string) (author: string option) (date: string option) (model: SpreadsheetModel) : Result<SpreadsheetModel * string, string> =
    match commentAt model addr with
    | None -> Result.Error (sprintf "the cell %s has no comment to reply to" (formatCellRef addr.Row addr.Col))
    | Some info ->
        let model, thread =
            match info.Thread with
            | Some thread -> model, thread
            | None ->
                let root = { Id = freshCommentId model; Author = info.Author; PersonId = None; Date = date; Text = info.Text; Replies = []; Done = false }
                // (The note goes: the file gets one written for the thread.)
                let converted = withNote addr None true model
                { converted with Threads = Map.add addr root converted.Threads }, root
        let id = freshCommentId model
        let reply = { Id = id; Author = author; PersonId = None; Date = date; Text = text; Replies = []; Done = false }
        Ok ({ model with Threads = Map.add addr { thread with Replies = thread.Replies @ [ reply ] } model.Threads }, id)

/// Take one reply out of the thread of a cell. An Error when the cell has
/// no thread or the thread no such reply.
let deleteCommentReply (addr: CellAddress) (replyId: string) (model: SpreadsheetModel) : Result<SpreadsheetModel, string> =
    match Map.tryFind addr model.Threads with
    | None -> Result.Error (sprintf "the cell %s has no comment thread" (formatCellRef addr.Row addr.Col))
    | Some thread ->
        let same (reply: CommentThread) = System.String.Equals(reply.Id, replyId, System.StringComparison.OrdinalIgnoreCase)
        if not (thread.Replies |> List.exists same) then Result.Error (sprintf "the comment on %s has no reply with the id %s" (formatCellRef addr.Row addr.Col) replyId)
        else Ok { model with Threads = Map.add addr { thread with Replies = thread.Replies |> List.filter (same >> not) } model.Threads }

/// Mark the thread of a cell resolved or open again. An Error when the cell has no thread.
let setCommentResolved (addr: CellAddress) (resolved: bool) (model: SpreadsheetModel) : Result<SpreadsheetModel, string> =
    match Map.tryFind addr model.Threads with
    | None -> Result.Error (sprintf "the cell %s has no comment thread (a note cannot be resolved)" (formatCellRef addr.Row addr.Col))
    | Some thread -> Ok (if thread.Done = resolved then model else { model with Threads = Map.add addr { thread with Done = resolved } model.Threads })

/// The note a file keeps beside a thread, for applications that only show notes.
let threadPlaceholderNote (thread: CommentThread) : CellComment =
    { Text = "[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    " + thread.Text
               + (thread.Replies |> List.map (fun reply -> "\nReply:\n    " + reply.Text) |> String.concat "")
      Author = Some ("tc=" + thread.Id)
      Runs = [] }

// --- hyperlinks -----------------------------------------------------------------
// A link is on a range of cells and points out of the workbook (an address)
// or at a place in it (a reference or a defined name). The editor never
// opens an address itself; a host does, and only for the schemes below.

/// The scheme of an address ("https" of "https://example.com"), lower case; "" when it has none.
let hyperlinkScheme (target: string) : string =
    let target = if isNull target then "" else target.Trim()
    let colon = target.IndexOf ':'
    if colon <= 0 then ""
    else
        let scheme = target.Substring(0, colon)
        let letter (ch: char) = (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z')
        if letter scheme.[0] && scheme |> Seq.forall (fun ch -> letter ch || (ch >= '0' && ch <= '9') || ch = '+' || ch = '.' || ch = '-') then scheme.ToLowerInvariant() else ""

/// Whether a host may open the address: http, https, mailto and tel only.
/// A place in the workbook is not an address to open.
let hyperlinkIsSafe (link: SheetHyperlink) : bool =
    not link.IsInternal
    && (match hyperlinkScheme link.Target with
        | "http" | "https" | "mailto" | "tel" -> true
        | _ -> false)

/// Read a target as a host gives it: (target as stored, is a place in the
/// workbook). "#Sheet2!A1" and anything that reads as a reference, with or
/// without a sheet, or that `isName` knows as a defined name, is a place in
/// the workbook; everything else is an address.
let classifyHyperlinkTarget (isName: string -> bool) (target: string) : string * bool =
    let target = target.Trim()
    if target.StartsWith "#" then target.Substring 1, true
    elif hyperlinkScheme target <> "" && (tryParseRange target).IsNone then target, false
    else
        let _, reference = splitSheetPrefix target
        if (tryParseRange reference).IsSome || isName target then target, true
        else target, false

let private linkColor = "#0563C1"

/// The hyperlinks without the cells of `ranges`: a link that covered more
/// keeps the rest (as up to four links).
let private hyperlinksWithout (ranges: CellRange list) (links: SheetHyperlink list) : SheetHyperlink list =
    let remaining =
        links
        |> List.collect (fun link ->
            let parts = ranges |> List.fold (fun parts hole -> parts |> List.collect (subtractCellRange hole)) [ link.Range ]
            if parts = [ link.Range ] then [ link ] else parts |> List.map (fun part -> { link with Range = part }))
    if remaining = links then links else remaining

/// The cells of some ranges that exist or are few enough to be made (a link
/// on a whole column changes the look of the cells it has).
let private linkCells (ranges: CellRange list) (model: SpreadsheetModel) : CellAddress list =
    [ for range in ranges do
        match clipRange model.RowCount model.ColCount range with
        | None -> ()
        | Some range ->
            let area = float (range.To.Row - range.From.Row + 1) * float (range.To.Col - range.From.Col + 1)
            if area <= 4096.0 then
                for row in range.From.Row .. range.To.Row do
                    for col in range.From.Col .. range.To.Col -> { Row = row; Col = col }
            else
                for KeyValue(addr, _) in model.Cells do
                    if cellRangeContains addr range then yield addr ]

/// Put a hyperlink on ranges, in place of whatever links their cells had.
/// `display`, when given, becomes the text of the first cell; without it an
/// empty first cell shows the target, as in Excel. The cells get the look of
/// a link (blue, underlined) unless they have a text colour of their own.
let setHyperlink (ranges: CellRange list) (target: string) (isInternal: bool) (display: string option) (tooltip: string option) (model: SpreadsheetModel) : SpreadsheetModel =
    let links = hyperlinksWithout ranges model.Hyperlinks @ (ranges |> List.map (fun range -> { Range = range; Target = target; IsInternal = isInternal; Tooltip = tooltip; Display = display }))
    let mutable cells = model.Cells
    let first = match ranges with range :: _ -> Some range.From | [] -> None
    for addr in linkCells ranges model do
        let existing = Map.tryFind addr cells
        let isFirst = (Some addr = first)
        let data = existing |> Option.defaultValue { Value = CellValue.Empty; Format = defaultFormat; Comment = None }
        let value =
            if not isFirst then data.Value
            else
                match display with
                | Some text -> CellValue.String text
                | None -> (match data.Value with CellValue.Empty | CellValue.String "" -> CellValue.String target | value -> value)
        // (Only cells that hold something, and the first one, are given the look.)
        if existing.IsSome || isFirst then
            let format =
                if data.Format.TextColor.IsSome && data.Format.TextColor <> Some linkColor then data.Format
                elif data.Format.TextColor = Some linkColor && data.Format.Underline then data.Format
                else { data.Format with TextColor = Some linkColor; Underline = true }
            let changed = { data with Value = value; Format = format }
            if existing <> Some changed then cells <- Map.add addr changed cells
    let withCells = if obj.ReferenceEquals(cells, model.Cells) then model else { model with Cells = cells } |> invalidateFilterProjection
    { withCells with Hyperlinks = links }

/// Take the hyperlinks off the cells of ranges, and with them the look of a
/// link (Excel's "Remove Hyperlink"); the texts stay. The changed model and
/// how many links were removed or cut. The same model when there was none.
let removeHyperlinks (ranges: CellRange list) (model: SpreadsheetModel) : SpreadsheetModel * int =
    let touched = model.Hyperlinks |> List.filter (fun link -> ranges |> List.exists (fun range -> (subtractCellRange range link.Range) <> [ link.Range ]))
    if touched.IsEmpty then model, 0
    else
        let mutable cells = model.Cells
        let linked (addr: CellAddress) = touched |> List.exists (fun link -> cellRangeContains addr link.Range)
        for addr in linkCells ranges model do
            match Map.tryFind addr cells with
            | Some data when linked addr && data.Format.TextColor = Some linkColor && data.Format.Underline ->
                let plain = { data with Format = { data.Format with TextColor = None; Underline = false } }
                if plain.Value = CellValue.Empty && plain.Comment.IsNone && plain.Format = defaultFormat then cells <- Map.remove addr cells
                else cells <- Map.add addr plain cells
            | _ -> ()
        let withCells = if obj.ReferenceEquals(cells, model.Cells) then model else { model with Cells = cells } |> invalidateFilterProjection
        { withCells with Hyperlinks = hyperlinksWithout ranges model.Hyperlinks }, touched.Length

// --- pictures -------------------------------------------------------------------
// A picture is a drawing object of the sheet (`SpreadsheetModel.Drawings`,
// back to front) known by its `Id`. Charts and the other objects a file
// brings are kept as they are: only pictures are placed through here.

/// What the first bytes of an image say it is.
type ImageInfo = {
    /// "image/png" | "image/jpeg" | "image/gif"
    ContentType: string
    /// The extension of its part in a file: "png" | "jpeg" | "gif".
    Extension: string
    /// Its size in pixels.
    Width: int
    Height: int
}

/// Read the kind and the pixel size of a PNG, JPEG or GIF off its header;
/// None for anything else (and for a file cut short of its size).
let imageInfo (data: byte[]) : ImageInfo option =
    let at (index: int) = int data.[index]
    let n = if isNull data then 0 else data.Length
    if n >= 24 && at 0 = 0x89 && at 1 = 0x50 && at 2 = 0x4E && at 3 = 0x47 && at 4 = 0x0D && at 5 = 0x0A && at 6 = 0x1A && at 7 = 0x0A then
        // The IHDR chunk comes first: width and height, big-endian.
        let be (index: int) = (at index <<< 24) ||| (at (index + 1) <<< 16) ||| (at (index + 2) <<< 8) ||| at (index + 3)
        let width, height = be 16, be 20
        if width > 0 && height > 0 then Some { ContentType = "image/png"; Extension = "png"; Width = width; Height = height } else None
    elif n >= 10 && at 0 = 0x47 && at 1 = 0x49 && at 2 = 0x46 && at 3 = 0x38 && (at 4 = 0x37 || at 4 = 0x39) && at 5 = 0x61 then
        let width, height = at 6 ||| (at 7 <<< 8), at 8 ||| (at 9 <<< 8)
        if width > 0 && height > 0 then Some { ContentType = "image/gif"; Extension = "gif"; Width = width; Height = height } else None
    elif n >= 4 && at 0 = 0xFF && at 1 = 0xD8 then
        // Segments until a start-of-frame one, which has the size.
        let mutable index = 2
        let mutable found = None
        let mutable going = true
        while going && index + 3 < n do
            if at index <> 0xFF then going <- false
            else
                let marker = at (index + 1)
                if marker = 0xFF then index <- index + 1                       // padding
                elif marker = 0xD8 || (marker >= 0xD0 && marker <= 0xD7) || marker = 0x01 then index <- index + 2   // no length
                elif marker >= 0xC0 && marker <= 0xCF && marker <> 0xC4 && marker <> 0xC8 && marker <> 0xCC then
                    if index + 8 < n then
                        let height, width = (at (index + 5) <<< 8) ||| at (index + 6), (at (index + 7) <<< 8) ||| at (index + 8)
                        if width > 0 && height > 0 then found <- Some { ContentType = "image/jpeg"; Extension = "jpeg"; Width = width; Height = height }
                    going <- false
                else index <- index + 2 + ((at (index + 2) <<< 8) ||| at (index + 3))
        found
    else None

/// The number for the next drawing object of a sheet.
let nextDrawingId (model: SpreadsheetModel) : int =
    (model.Drawings |> List.fold (fun highest drawing -> max highest drawing.Id) 0) + 1

/// The drawing object with an id, or None.
let drawingById (id: int) (model: SpreadsheetModel) : SheetDrawing option =
    if id <= 0 then None else model.Drawings |> List.tryFind (fun drawing -> drawing.Id = id)

let private isPicture (drawing: SheetDrawing) = match drawing.Content with Picture _ -> true | KeptObject _ -> false

/// The picture with an id; an Error that says why when there is none, or
/// the object is a chart or another object the editor only keeps.
let pictureById (id: int) (model: SpreadsheetModel) : Result<SheetDrawing, string> =
    match drawingById id model with
    | None -> Result.Error (sprintf "the sheet has no picture with the id %d" id)
    | Some drawing when not (isPicture drawing) -> Result.Error (sprintf "the object %d is not a picture (charts and other kept objects cannot be changed)" id)
    | Some drawing -> Ok drawing

/// An anchor at a cell with a size of its own: the object moves with the
/// cell and keeps its size.
let pictureAnchor (cell: CellAddress) (offsetX: float) (offsetY: float) (width: float) (height: float) : DrawingAnchor =
    { Kind = "oneCell"; FromRow = cell.Row; FromCol = cell.Col; FromOffsetX = offsetX; FromOffsetY = offsetY
      ToRow = None; ToCol = None; ToOffsetX = 0.0; ToOffsetY = 0.0; X = 0.0; Y = 0.0
      Width = Some width; Height = Some height; EditAs = None }

/// Put a picture in front of the sheet's other objects, at a cell, in the
/// size given. `id` is its number (`nextDrawingId`).
let insertPicture (id: int) (info: ImageInfo) (data: byte[]) (cell: CellAddress) (width: float) (height: float) (name: string option) (model: SpreadsheetModel) : SpreadsheetModel =
    let name = match name with Some _ -> name | None -> Some (sprintf "Picture %d" id)
    let drawing = { Id = id; Anchor = pictureAnchor cell 0.0 0.0 width height; Name = name; Description = None; Content = Picture (info.ContentType, info.Extension, data) }
    { model with Drawings = model.Drawings @ [ drawing ] }

let private changePicture (id: int) (change: SheetDrawing -> SheetDrawing) (model: SpreadsheetModel) : Result<SpreadsheetModel, string> =
    match pictureById id model with
    | Result.Error message -> Result.Error message
    | Ok _ ->
        let drawings = model.Drawings |> List.map (fun drawing -> if drawing.Id = id then change drawing else drawing)
        Ok (if drawings = model.Drawings then model else { model with Drawings = drawings })

/// Place a picture: at `cell`, `offsetX` / `offsetY` pixels into it, in the
/// size given. (Whatever anchor it had, it now moves with that cell and
/// keeps its size.)
let placePicture (id: int) (cell: CellAddress) (offsetX: float) (offsetY: float) (width: float) (height: float) (model: SpreadsheetModel) : Result<SpreadsheetModel, string> =
    changePicture id (fun drawing -> { drawing with Anchor = pictureAnchor cell offsetX offsetY width height }) model

/// Take a picture off the sheet.
let deletePicture (id: int) (model: SpreadsheetModel) : Result<SpreadsheetModel, string> =
    match pictureById id model with
    | Result.Error message -> Result.Error message
    | Ok _ -> Ok { model with Drawings = model.Drawings |> List.filter (fun drawing -> drawing.Id <> id) }

/// Set the name and the description (alternative text) of a picture; None leaves one as it is.
let describePicture (id: int) (name: string option) (description: string option) (model: SpreadsheetModel) : Result<SpreadsheetModel, string> =
    changePicture id (fun drawing ->
        { drawing with
            Name = (match name with Some _ -> name | None -> drawing.Name)
            Description = (match description with Some "" -> None | Some _ -> description | None -> drawing.Description) }) model

// ---------------------------------------------------------------------------
// Find and replace
// ---------------------------------------------------------------------------

/// How a search reads the cells.
type FindOptions = {
    /// Upper and lower case are different letters.
    MatchCase: bool
    /// The whole text of the cell has to match, not a part of it.
    WholeCell: bool
    /// The text searched for is a regular expression.
    Regex: bool
    /// Search what was entered (the formula of a formula cell, the entry of
    /// any other) instead of what the cell shows.
    InFormulas: bool
}

let defaultFindOptions = { MatchCase = false; WholeCell = false; Regex = false; InFormulas = false }

/// A search text made ready: whether a text matches, and the text with
/// every match replaced.
type Finder = {
    IsMatch: string -> bool
    Replace: string -> string -> string
}

/// The finder of a search text. Throws (naming `memberName`) for an empty
/// text and for a regular expression that cannot be read.
let finderOf (memberName: string) (text: string) (options: FindOptions) : Finder =
    if String.IsNullOrEmpty text then failwith (sprintf "%s: there is no text to search for" memberName)
    if options.Regex then
        let flags =
            if options.MatchCase then Text.RegularExpressions.RegexOptions.None
            else Text.RegularExpressions.RegexOptions.IgnoreCase
        let regex =
            try
                // (The whole cell: the pattern from the first character to the last.)
                let made = Text.RegularExpressions.Regex((if options.WholeCell then "^(?:" + text + ")$" else text), flags)
                // (A pattern is read when it is first used in the browser.)
                made.IsMatch "" |> ignore
                made
            with error -> failwith (sprintf "%s: \"%s\" is not a regular expression (%s)" memberName text error.Message)
        { IsMatch = (fun candidate -> regex.IsMatch candidate)
          Replace = (fun candidate replacement -> regex.Replace(candidate, replacement)) }
    else
        let needle = if options.MatchCase then text else text.ToLowerInvariant()
        let fold (candidate: string) = if options.MatchCase then candidate else candidate.ToLowerInvariant()
        let isMatch (candidate: string) =
            if options.WholeCell then fold candidate = needle else (fold candidate).Contains needle
        let replace (candidate: string) (replacement: string) =
            if options.WholeCell then (if fold candidate = needle then replacement else candidate)
            else
                let folded = fold candidate
                // (Folding keeps the length of every letter the grid can hold a match in.)
                if folded.Length <> candidate.Length then candidate.Replace(text, replacement)
                else
                    let sb = Text.StringBuilder()
                    let mutable at = 0
                    let mutable next = folded.IndexOf(needle, 0, StringComparison.Ordinal)
                    while next >= 0 do
                        sb.Append(candidate.Substring(at, next - at)).Append(replacement) |> ignore
                        at <- next + needle.Length
                        next <- if at <= folded.Length - needle.Length then folded.IndexOf(needle, at, StringComparison.Ordinal) else -1
                    sb.Append(candidate.Substring at).ToString()
        { IsMatch = isMatch; Replace = replace }

/// The text of a cell a search reads: what it shows, or (in formulas) what
/// was entered - the formula with its `=`, a date as a date, text as it is.
/// (`date1904`: the date system of the cell's workbook.)
let searchTextIn (date1904: bool) (options: FindOptions) (data: CellData) : string =
    if options.InFormulas then
        match data.Value with
        | String text -> text
        | _ -> Rbt.Spreadsheet.Input.editTextFor date1904 (Some data)
    else Rbt.Spreadsheet.CellText.cellDisplayTextIn date1904 None 1.0e6 data.Format data.Value

/// `searchTextIn` in the 1900 date system.
let searchText (options: FindOptions) (data: CellData) : string = searchTextIn false options data

/// The cells of a sheet (of `range` when given) that match, row by row,
/// each with the text that matched.
let findInSheet (finder: Finder) (options: FindOptions) (range: CellRange option) (model: SpreadsheetModel) : (CellAddress * string) list =
    let inside (addr: CellAddress) =
        match range with
        | Some area -> addr.Row >= area.From.Row && addr.Row <= area.To.Row && addr.Col >= area.From.Col && addr.Col <= area.To.Col
        | None -> true
    let found = ResizeArray<CellAddress * string>()
    // (The cell map is in row order, columns in order within a row.)
    for KeyValue(addr, data) in model.Cells do
        match data.Value with
        | Empty -> ()
        | _ ->
            if inside addr then
                let text = searchTextIn model.Date1904 options data
                if text <> "" && finder.IsMatch text then found.Add((addr, text))
    List.ofSeq found

/// A cell with every match replaced, read again the way typed input is
/// (the cell keeps its format); None when the cell does not match or cannot
/// take the replacement. Searching what cells show, a formula cell is left
/// alone (its value is not its own to change) and the replacement is made
/// in the text shown; searching formulas it is made in the formula text.
let replaceInCellIn (date1904: bool) (finder: Finder) (options: FindOptions) (replacement: string) (currentYear: int) (data: CellData) : CellData option =
    match data.Value with
    | Empty -> None
    | Formula _ when not options.InFormulas -> None
    | _ ->
        let text = searchTextIn date1904 options data
        if text = "" || not (finder.IsMatch text) then None
        else
            let replaced = finder.Replace text replacement
            if replaced = text then None
            elif replaced = "" then Some { data with Value = Empty }
            else Some (Rbt.Spreadsheet.Input.cellFromInputFor date1904 currentYear (Some data) replaced)

/// `replaceInCellIn` in the 1900 date system.
let replaceInCell (finder: Finder) (options: FindOptions) (replacement: string) (currentYear: int) (data: CellData) : CellData option =
    replaceInCellIn false finder options replacement currentYear data

/// Every cell of a sheet (of `range` when given) a replace-all changes:
/// (address, the cell before, the cell after). The cells a dynamic array
/// spilled into are not changed (they are their formula's).
let replaceInSheet (finder: Finder) (options: FindOptions) (replacement: string) (currentYear: int) (range: CellRange option) (model: SpreadsheetModel) : (CellAddress * CellData * CellData) list =
    findInSheet finder options range model
    |> List.choose (fun (addr, _) ->
        match Map.tryFind addr model.Cells with
        | Some _ when Map.containsKey addr model.Spill -> None
        | Some data -> replaceInCellIn model.Date1904 finder options replacement currentYear data |> Option.map (fun changed -> addr, data, changed)
        | None -> None)

/// The match a step from a place leads to. `matches` are in search order,
/// each with its place as (sheet position in the search order, row,
/// column); `from` is the place the step starts at. Forward: the first
/// match after it, else the first of all; backward: the last before it,
/// else the last of all. None when there is no match.
let stepMatch (forward: bool) (from: int * int * int) (places: (int * int * int) array) : int option =
    if places.Length = 0 then None
    elif forward then
        match places |> Array.tryFindIndex (fun place -> place > from) with
        | Some index -> Some index
        | None -> Some 0
    else
        match places |> Array.tryFindIndexBack (fun place -> place < from) with
        | Some index -> Some index
        | None -> Some (places.Length - 1)

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------
// A table is kept twice: on its sheet (`SpreadsheetModel.Tables`: where it
// is, its style and filter buttons - what is drawn and written to a file)
// and in the workbook (`Workbook.Tables`: its columns by name - what
// structured references read). Every operation here changes both.

module Wb = Rbt.Spreadsheet.Workbook

/// A table as hosts see it.
type TableInfo = {
    Name: string
    SheetId: Wb.SheetId
    /// Header row, data rows and total row.
    Range: CellRange
    HasHeaders: bool
    HasTotals: bool
    Columns: string list
    Style: TableStyle option
    HasAutoFilter: bool
}

/// What a table operation did: the workbook after it and the cells it
/// changed (for the undo step, which puts those cells back).
type TableChange = {
    Book: Wb.Workbook
    Cells: (Wb.SheetId * CellAddress) list
}

let private sameName (a: string) (b: string) = String.Equals(a, b, StringComparison.OrdinalIgnoreCase)

/// The table of a sheet's own list with a name.
let private sheetTableNamed (name: string) (model: SpreadsheetModel) : Table option =
    model.Tables |> List.tryFind (fun table -> sameName table.Name name)

let private tryFindTable (name: string) (workbook: Wb.Workbook) : Wb.Table option =
    workbook.Tables |> Map.tryPick (fun _ table -> if sameName table.Name name then Some table else None)

let private findTable (memberName: string) (name: string) (workbook: Wb.Workbook) : Wb.Table =
    match tryFindTable name workbook with
    | Some table -> table
    | None -> failwith (sprintf "%s: there is no table \"%s\"" memberName name)

let private sheetModelOf (id: Wb.SheetId) (workbook: Wb.Workbook) : SpreadsheetModel =
    match Map.tryFind id workbook.Sheets with
    | Some sheet -> sheet.Model
    | None -> failwith (sprintf "the workbook has no sheet \"%s\"" id)

let private withSheetModel (id: Wb.SheetId) (model: SpreadsheetModel) (workbook: Wb.Workbook) : Wb.Workbook =
    match Map.tryFind id workbook.Sheets with
    | Some sheet when not (obj.ReferenceEquals(sheet.Model, model)) -> { workbook with Sheets = Map.add id { sheet with Model = model } workbook.Sheets }
    | _ -> workbook

/// The tables of a workbook, sheet by sheet in tab order, each sheet's from
/// the top left.
let tablesOf (workbook: Wb.Workbook) : TableInfo list =
    [ for id in workbook.SheetOrder do
        match Map.tryFind id workbook.Sheets with
        | None -> ()
        | Some sheet ->
            let onSheet =
                workbook.Tables |> Map.toList |> List.map snd
                |> List.filter (fun table -> table.SheetId = id)
                |> List.sortBy (fun table -> table.StartRow, table.StartCol)
            for table in onSheet do
                let own = sheetTableNamed table.Name sheet.Model
                yield
                    { Name = table.Name
                      SheetId = id
                      Range = rangeOf table.StartRow table.StartCol table.EndRow table.EndCol
                      HasHeaders = table.HasHeaders
                      HasTotals = table.HasTotals
                      Columns = table.Columns
                      Style = own |> Option.bind (fun t -> t.Style)
                      HasAutoFilter = (match own with Some t -> t.AutoFilter.IsSome | None -> false) } ]

/// The header and band fills of a built-in table style: the styles cycle
/// through black and the six accent colours of the Office theme.
let tableStyleFills (styleName: string) : string * string =
    let accents = [| "4472C4"; "ED7D31"; "A5A5A5"; "FFC000"; "5B9BD5"; "70AD47" |]
    let mutable n = 0
    let mutable any = false
    for ch in styleName do
        if Char.IsDigit ch && n < 100000 then
            any <- true
            n <- n * 10 + (int ch - int '0')
    let slot = (max 1 (if any then n else 2) - 1) % 7
    if slot = 0 then "#000000", "#D9D9D9"
    // (The band is the accent lighter 80%, as Excel shows that tint.)
    else "#" + accents.[slot - 1], "#" + XlsxReader.Xlsx.applyColorTint accents.[slot - 1] 0.79998168889431442

/// The style record of a built-in table style ("TableStyleMedium2") with
/// banded rows, as Excel applies one: Excel's definition of the style in the
/// colours of the Office theme (no elements for a name Excel does not have).
let tableStyleNamed (name: string) : TableStyle =
    { Name = name; ShowFirstColumn = false; ShowLastColumn = false; ShowRowStripes = true; ShowColumnStripes = false
      Elements = TableStyles.builtIn TableStyles.officeTheme name |> Option.defaultValue [] }

let defaultTableStyleName = "TableStyleMedium2"

/// Column names as Excel makes them from header cells: a blank one is
/// "ColumnN" (N its place), and a name that was used already gets the next
/// free number appended.
let uniqueColumnNames (raw: string list) : string list =
    let taken = Collections.Generic.HashSet<string>()
    raw
    |> List.mapi (fun index text ->
        let wanted = if String.IsNullOrWhiteSpace text then sprintf "Column%d" (index + 1) else text
        let mutable name = wanted
        let mutable next = 2
        while taken.Contains (name.ToUpperInvariant()) do
            name <- wanted + string next
            next <- next + 1
        taken.Add (name.ToUpperInvariant()) |> ignore
        name)

/// What is wrong with a name for a (new) table, or None: it follows the
/// rules of defined names and is not the name of another table or of a
/// defined name.
let tableNameProblem (name: string) (workbook: Wb.Workbook) : string option =
    match Wb.definedNameProblem name with
    | Some problem -> Some problem
    | None ->
        if (tryFindTable name workbook).IsSome then Some (sprintf "there is already a table named \"%s\"" name)
        elif workbook.Names |> Map.exists (fun (_, known) _ -> known = name.ToUpperInvariant()) then Some (sprintf "\"%s\" is a defined name" name)
        else None

/// The first of Table1, Table2 ... that nothing is called yet.
let freeTableName (workbook: Wb.Workbook) : string =
    let mutable n = 1
    while (tableNameProblem (sprintf "Table%d" n) workbook).IsSome do n <- n + 1
    sprintf "Table%d" n

let private headerNames (model: SpreadsheetModel) (row: int) (left: int) (right: int) : string list =
    [ for col in left .. right -> displayValue (getCellValue { Row = row; Col = col } model) ]

/// Write the column names into the header row where the cells do not hold
/// them already (as text); gives the model and the cells written.
let private writeHeaders (row: int) (left: int) (names: string list) (model: SpreadsheetModel) : SpreadsheetModel * CellAddress list =
    let mutable m = model
    let touched = ResizeArray<CellAddress>()
    names |> List.iteri (fun index name ->
        let addr = { Row = row; Col = left + index }
        let existing = getCell addr m
        match existing with
        | Some data when data.Value = String name -> ()
        | _ ->
            let data = existing |> Option.defaultValue { Value = Empty; Format = defaultFormat; Comment = None }
            m <- setCell addr { data with Value = String name } m
            touched.Add addr)
    m, List.ofSeq touched

let private tableRangeProblem (ignoreTable: string option) (sheetId: Wb.SheetId) (range: CellRange) (workbook: Wb.Workbook) : string option =
    let model = sheetModelOf sheetId workbook
    if range.To.Row >= model.RowCount || range.To.Col >= model.ColCount then
        Some (sprintf "%s is outside the sheet" (rangeText range))
    else
        let overlaps (table: Wb.Table) =
            table.SheetId = sheetId
            && (match ignoreTable with Some name -> not (sameName name table.Name) | None -> true)
            && table.StartRow <= range.To.Row && table.EndRow >= range.From.Row && table.StartCol <= range.To.Col && table.EndCol >= range.From.Col
        match workbook.Tables |> Map.tryPick (fun _ table -> if overlaps table then Some table else None) with
        | Some other -> Some (sprintf "%s overlaps the table \"%s\"" (rangeText range) other.Name)
        | None ->
            if not (mergedCellsIntersectingRange range.From.Row (range.To.Row + 1) range.From.Col (range.To.Col + 1) model).IsEmpty then
                Some (sprintf "%s holds merged cells (a table cannot)" (rangeText range))
            else None

/// Make a table of a range of a sheet. With headers the first row names
/// the columns (blank and repeated header cells are given Excel's names,
/// and every header cell becomes text); without, the columns are Column1,
/// Column2 ... and every row is data. `style` None is a table without a
/// style. Throws (naming `memberName`) for a bad name or range.
let defineTable (memberName: string) (sheetId: Wb.SheetId) (range: CellRange) (name: string option) (hasHeaders: bool)
                (style: TableStyle option) (workbook: Wb.Workbook) : TableChange * string =
    let fail (message: string) = failwith (sprintf "%s: %s" memberName message)
    let name =
        match name with
        | Some given ->
            (match tableNameProblem given workbook with
             | Some problem -> fail problem
             | None -> given)
        | None -> freeTableName workbook
    if isWholeColumns range || isWholeRows range then fail "a table cannot be whole rows or columns"
    (match tableRangeProblem None sheetId range workbook with
     | Some problem -> fail problem
     | None -> ())
    if hasHeaders && range.To.Row = range.From.Row then fail "a table with a header row needs at least one more row"
    let model = sheetModelOf sheetId workbook
    let top, left, bottom, right = range.From.Row, range.From.Col, range.To.Row, range.To.Col
    let columns =
        if hasHeaders then uniqueColumnNames (headerNames model top left right)
        else [ for index in 1 .. right - left + 1 -> sprintf "Column%d" index ]
    let model, touched = if hasHeaders then writeHeaders top left columns model else model, []
    let own : Table =
        { Name = name; StartRow = top + 1; StartCol = left + 1; EndRow = bottom + 1; EndCol = right + 1
          HeaderRowCount = (if hasHeaders then 1 else 0)
          Style = style
          AutoFilter = (if hasHeaders then Some (left + 1, right + 1) else None)
          HasTotals = false }
    let table : Wb.Table =
        { Name = name; SheetId = sheetId; StartRow = top; StartCol = left; EndRow = bottom; EndCol = right
          Columns = columns; HasHeaders = hasHeaders; HasTotals = false }
    let book =
        { withSheetModel sheetId { model with Tables = model.Tables @ [ own ] } workbook with
            Tables = Map.add name table workbook.Tables }
    { Book = book; Cells = touched |> List.map (fun addr -> sheetId, addr) }, name

/// Replace a table's records (the sheet's and the workbook's).
let private putTable (before: Wb.Table) (after: Wb.Table) (own: Table -> Table) (model: SpreadsheetModel) (workbook: Wb.Workbook) : Wb.Workbook =
    let model = { model with Tables = model.Tables |> List.map (fun t -> if sameName t.Name before.Name then own t else t) }
    let tables = workbook.Tables |> Map.filter (fun _ table -> not (sameName table.Name before.Name)) |> Map.add after.Name after
    { withSheetModel before.SheetId model workbook with Tables = tables }

let private rowIsEmpty (model: SpreadsheetModel) (row: int) (left: int) (right: int) : bool =
    [ left .. right ] |> List.forall (fun col ->
        match getCellValue { Row = row; Col = col } model with
        | Empty | String "" -> true
        | _ -> false)

/// Show or remove the total row of a table. Shown, it is the row below the
/// data (which has to be empty): "Total" in the first column and, in the
/// last, the sum of the column when it holds numbers, else the count of
/// its entries (SUBTOTAL, so filtered rows are left out). Removed, its
/// cells are emptied.
let setTableTotalRow (memberName: string) (name: string) (show: bool) (workbook: Wb.Workbook) : TableChange =
    let fail (message: string) = failwith (sprintf "%s: %s" memberName message)
    let table = findTable memberName name workbook
    let model = sheetModelOf table.SheetId workbook
    if show = table.HasTotals then { Book = workbook; Cells = [] }
    elif show then
        let row = table.EndRow + 1
        if row >= model.RowCount then fail "there is no row below the table for the totals"
        if not (rowIsEmpty model row table.StartCol table.EndCol) then
            fail (sprintf "the row below the table (%s) is not empty" (rangeText (rangeOf row table.StartCol row table.EndCol)))
        if tableRangeProblem (Some table.Name) table.SheetId (rangeOf table.StartRow table.StartCol row table.EndCol) workbook |> Option.isSome then
            fail "the row below the table belongs to another table or holds merged cells"
        let put (addr: CellAddress) (value: CellValue) (m: SpreadsheetModel) =
            let data = getCell addr m |> Option.defaultValue { Value = Empty; Format = defaultFormat; Comment = None }
            setCell addr { data with Value = value } m
        let first = { Row = row; Col = table.StartCol }
        let last = { Row = row; Col = table.EndCol }
        let mutable m = put first (String "Total") model
        let touched = ResizeArray<CellAddress>([ first ])
        if table.EndCol > table.StartCol && not table.Columns.IsEmpty then
            let dataStart = table.StartRow + (if table.HasHeaders then 1 else 0)
            let numeric =
                [ dataStart .. table.EndRow ] |> List.exists (fun r ->
                    match unwrapFormulaValue (getCellValue { Row = r; Col = table.EndCol } model) with
                    | Number _ -> true
                    | _ -> false)
            let column = List.item (table.EndCol - table.StartCol) table.Columns
            let formula =
                Rbt.Spreadsheet.Formula.exprToFormula
                    (Rbt.Spreadsheet.Formula.Function ("SUBTOTAL",
                        [ Rbt.Spreadsheet.Formula.Number (if numeric then 109.0 else 103.0)
                          Rbt.Spreadsheet.Formula.TableRef { Table = table.Name; Item = Rbt.Spreadsheet.Formula.TblData; ColStart = Some column; ColEnd = None } ]))
            m <- put last (Formula ("=" + formula, Empty)) m
            touched.Add last
        let after = { table with EndRow = row; HasTotals = true }
        { Book = putTable table after (fun t -> { t with EndRow = t.EndRow + 1; HasTotals = true }) m workbook
          Cells = touched |> Seq.map (fun addr -> table.SheetId, addr) |> List.ofSeq }
    else
        let row = table.EndRow
        let mutable m = model
        let touched = ResizeArray<CellAddress>()
        for col in table.StartCol .. table.EndCol do
            let addr = { Row = row; Col = col }
            match getCell addr m with
            | Some data when data.Value <> Empty ->
                m <- setCell addr { data with Value = Empty } m
                touched.Add addr
            | _ -> ()
        let after = { table with EndRow = row - 1; HasTotals = false }
        { Book = putTable table after (fun t -> { t with EndRow = t.EndRow - 1; HasTotals = false }) m workbook
          Cells = touched |> Seq.map (fun addr -> table.SheetId, addr) |> List.ofSeq }

/// Give a table another range: the header row stays in its row, and the
/// new range has to overlap the old one (Excel's rule). Columns that come
/// in are named by their header cells; a total row moves to the new last
/// row (`range` includes it, as the table's range does).
let resizeTable (memberName: string) (name: string) (range: CellRange) (workbook: Wb.Workbook) : TableChange =
    let fail (message: string) = failwith (sprintf "%s: %s" memberName message)
    let table = findTable memberName name workbook
    if isWholeColumns range || isWholeRows range then fail "a table cannot be whole rows or columns"
    if range.From.Row <> table.StartRow then fail (sprintf "the table has to keep its first row (row %d)" (table.StartRow + 1))
    if range.From.Col > table.EndCol || range.To.Col < table.StartCol then fail "the new range has to overlap the table"
    let needed = (if table.HasHeaders then 1 else 0) + (if table.HasTotals then 1 else 0)
    if range.To.Row - range.From.Row + 1 <= needed && needed > 0 then fail "the table needs at least one row of data"
    (match tableRangeProblem (Some table.Name) table.SheetId range workbook with
     | Some problem -> fail problem
     | None -> ())
    if rangeOf table.StartRow table.StartCol table.EndRow table.EndCol = range then { Book = workbook; Cells = [] }
    else
        // The total row is taken off, the table resized, and the total row put back.
        let without = if table.HasTotals then setTableTotalRow memberName name false workbook else { Book = workbook; Cells = [] }
        let plain = findTable memberName name without.Book
        let model = sheetModelOf plain.SheetId without.Book
        let bottom = if table.HasTotals then range.To.Row - 1 else range.To.Row
        let columns =
            if plain.HasHeaders then
                // Columns that stay keep their names; new ones are read off the header row.
                let raw =
                    [ for col in range.From.Col .. range.To.Col ->
                        if col >= plain.StartCol && col <= plain.EndCol && col - plain.StartCol < plain.Columns.Length then List.item (col - plain.StartCol) plain.Columns
                        else displayValue (getCellValue { Row = plain.StartRow; Col = col } model) ]
                uniqueColumnNames raw
            else [ for index in 1 .. range.To.Col - range.From.Col + 1 -> sprintf "Column%d" index ]
        let model, touched = if plain.HasHeaders then writeHeaders plain.StartRow range.From.Col columns model else model, []
        let after = { plain with StartCol = range.From.Col; EndCol = range.To.Col; EndRow = bottom; Columns = columns }
        let resized =
            putTable plain after (fun t ->
                { t with StartCol = range.From.Col + 1; EndCol = range.To.Col + 1; EndRow = bottom + 1
                         AutoFilter = t.AutoFilter |> Option.map (fun _ -> range.From.Col + 1, range.To.Col + 1) }) model without.Book
        let cells = without.Cells @ (touched |> List.map (fun addr -> plain.SheetId, addr))
        if table.HasTotals then
            let withTotals = setTableTotalRow memberName name true resized
            { withTotals with Cells = List.distinct (cells @ withTotals.Cells) }
        else { Book = resized; Cells = cells }

/// Apply `rewrite` (given the sheet and the cell a formula stands in; no
/// cell for the formulas of rules and names) to every formula of a
/// workbook: cells, conditional formats, validations and defined names.
let private rewriteBookFormulas (mentions: string -> bool) (rewrite: Wb.SheetId -> CellAddress option -> string -> string) (workbook: Wb.Workbook) : TableChange =
    let touched = ResizeArray<Wb.SheetId * CellAddress>()
    let sheets =
        workbook.Sheets
        |> Map.map (fun id sheet ->
            let model = sheet.Model
            let mutable cells = model.Cells
            for KeyValue(addr, data) in model.Cells do
                match data.Value with
                | Formula (formula, cached) when mentions formula ->
                    let rewritten = rewrite id (Some addr) formula
                    if rewritten <> formula then
                        cells <- Map.add addr { data with Value = Formula (rewritten, cached) } cells
                        touched.Add((id, addr))
                | _ -> ()
            // (The rules of the sheet, through the walk that knows where their formulas are.)
            let rules = mapModelFormulas (fun formula -> if mentions formula then rewrite id None formula else formula) { model with Cells = Map.empty }
            let changed =
                not (obj.ReferenceEquals(cells, model.Cells))
                || rules.ConditionalFormats <> model.ConditionalFormats || rules.Validations <> model.Validations
            if not changed then sheet
            else
                { sheet with
                    Model =
                        { model with
                            Cells = cells
                            ConditionalFormats = (if rules.ConditionalFormats = model.ConditionalFormats then model.ConditionalFormats else rules.ConditionalFormats)
                            Validations = (if rules.Validations = model.Validations then model.Validations else rules.Validations) }
                        |> invalidateFilterProjection })
    let names =
        workbook.Names |> Map.map (fun (scope, _) refersTo ->
            if mentions refersTo then rewrite (defaultArg scope "") None refersTo else refersTo)
    { Book = { workbook with Sheets = sheets; Names = (if names = workbook.Names then workbook.Names else names) }; Cells = List.ofSeq touched }

/// Rename a table: structured references in every formula of the workbook
/// follow.
let renameTable (memberName: string) (name: string) (newName: string) (workbook: Wb.Workbook) : TableChange =
    let table = findTable memberName name workbook
    if table.Name = newName then { Book = workbook; Cells = [] }
    else
        (match (if sameName table.Name newName then Wb.definedNameProblem newName else tableNameProblem newName workbook) with
         | Some problem -> failwith (sprintf "%s: %s" memberName problem)
         | None -> ())
        let change = RenameTableRefs (table.Name, newName)
        let rewritten =
            rewriteBookFormulas
                (fun formula -> formula.IndexOf(table.Name, StringComparison.OrdinalIgnoreCase) >= 0)
                (fun _ _ formula -> rewriteFormulaForSheets change formula) workbook
        let model = sheetModelOf table.SheetId rewritten.Book
        { rewritten with Book = putTable table { table with Name = newName } (fun t -> { t with Name = newName }) model rewritten.Book }

/// Whether an entry typed at a cell of a sheet may concern one of its tables
/// (`tableEntry`): the cell is in a table, or directly below or right of one.
let entryConcernsTable (addr: CellAddress) (model: SpreadsheetModel) : bool =
    not model.Tables.IsEmpty
    && model.Tables |> List.exists (fun table ->
        let r0, c0, r1, c1 = table.StartRow - 1, table.StartCol - 1, table.EndRow - 1, table.EndCol - 1
        addr.Row >= r0 && addr.Row <= r1 + 1 && addr.Col >= c0 && addr.Col <= c1 + 1)

/// What an entry typed into a cell does when it concerns a table, as Excel
/// does it - the entry and everything that follows from it as one change:
/// - in a header cell: the column takes the text as its name (a blank one
///   "ColumnN", a name another column has gets a number appended), the cell
///   holds that name as text, and structured references to the column in
///   every formula of the workbook follow;
/// - directly below the last row of a table without a total row, or directly
///   right of its last column: the table grows by that row / column - unless
///   a cell of the new row / column already holds something, or the table
///   would run into another table or a merged cell. Calculated columns are
///   filled into a new row; a new column gets its name in the header row;
/// - a formula in a data cell of a column whose other data cells are empty
///   (or hold that same calculated column): the formula fills the column.
/// `entry` None is a cleared cell. None when the entry concerns no table:
/// it is an ordinary entry then.
let tableEntry (sheetId: Wb.SheetId) (addr: CellAddress) (entry: CellData option) (workbook: Wb.Workbook) : TableChange option =
    let onSheet = workbook.Tables |> Map.toList |> List.map snd |> List.filter (fun table -> table.SheetId = sheetId)
    let inside (table: Wb.Table) = addr.Row >= table.StartRow && addr.Row <= table.EndRow && addr.Col >= table.StartCol && addr.Col <= table.EndCol
    let below (table: Wb.Table) = not table.HasTotals && addr.Row = table.EndRow + 1 && addr.Col >= table.StartCol && addr.Col <= table.EndCol
    let beside (table: Wb.Table) = addr.Col = table.EndCol + 1 && addr.Row >= table.StartRow && addr.Row <= table.EndRow
    let isBlank (value: CellValue) = match value with Empty | String "" -> true | _ -> false
    let hasEntry = match entry with Some data -> not (isBlank data.Value) | None -> false
    match onSheet |> List.tryFind inside |> Option.orElse (if hasEntry then onSheet |> List.tryFind (fun table -> below table || beside table) else None) with
    | None -> None
    | Some table ->
        let model = sheetModelOf sheetId workbook
        let touched = ResizeArray<Wb.SheetId * CellAddress>()
        let write (target: CellAddress) (data: CellData option) (m: SpreadsheetModel) =
            touched.Add((sheetId, target))
            match data with
            | Some data -> setCell target data m
            | None -> clearCell target m
        let blankAt (m: SpreadsheetModel) (row: int) (col: int) = isBlank (getCellValue { Row = row; Col = col } m)
        // 1. The table grows to take the entry in (the cells it would cover have to be empty).
        let grown : (CellRange * Wb.Workbook) option =
            if inside table then None
            else
                let range =
                    if below table then rangeOf table.StartRow table.StartCol (table.EndRow + 1) table.EndCol
                    else rangeOf table.StartRow table.StartCol table.EndRow (table.EndCol + 1)
                let free =
                    if below table then [ table.StartCol .. table.EndCol ] |> List.forall (fun col -> blankAt model addr.Row col)
                    else [ table.StartRow .. table.EndRow ] |> List.forall (fun row -> blankAt model row addr.Col)
                if free && (tableRangeProblem (Some table.Name) sheetId range workbook).IsNone then Some (range, workbook) else None
        if not (inside table) && grown.IsNone then None
        else
            // The entry itself.
            let book = withSheetModel sheetId (write addr entry model) workbook
            let book =
                match grown with
                | Some (range, _) ->
                    let resized = resizeTable "tableEntry" table.Name range book
                    touched.AddRange resized.Cells
                    resized.Book
                | None -> book
            let table = (tryFindTable table.Name book).Value
            let firstData = table.StartRow + (if table.HasHeaders then 1 else 0)
            let lastData = if table.HasTotals then table.EndRow - 1 else table.EndRow
            let formulaAt (m: SpreadsheetModel) (row: int) (col: int) =
                match getCellValue { Row = row; Col = col } m with
                | Formula (formula, _) -> Some formula
                | _ -> None
            /// The formula of a calculated column as it reads in `row`, when the
            /// rows of `rows` all hold it (each the one above moved down).
            let calculated (m: SpreadsheetModel) (col: int) (rows: int list) (row: int) : string option =
                match rows with
                | [] -> None
                | first :: _ ->
                    match formulaAt m first col with
                    | Some formula when rows |> List.forall (fun r -> formulaAt m r col = Some (shiftFormulaReferences (r - first) 0 formula)) ->
                        Some (shiftFormulaReferences (row - first) 0 formula)
                    | _ -> None
            let mutable m = sheetModelOf sheetId book
            let fill (row: int) (col: int) (formula: string) =
                let target = { Row = row; Col = col }
                let existing = getCell target m |> Option.defaultValue { Value = Empty; Format = defaultFormat; Comment = None }
                if existing.Value <> Formula (formula, Empty) then
                    (match existing.Value with
                     | Formula (current, _) when current = formula -> ()
                     | _ -> m <- write target (Some { existing with Value = Formula (formula, Empty) }) m)
            // 2. A new row takes the calculated columns of the table.
            if grown.IsSome && addr.Row = lastData && addr.Row > firstData then
                for col in table.StartCol .. table.EndCol do
                    if col <> addr.Col then
                        match calculated m col [ firstData .. lastData - 1 ] lastData with
                        | Some formula -> fill lastData col formula
                        | None -> ()
            // 3. A formula in a data cell of a column that is otherwise empty (or the same
            //    calculated column): the whole column takes it.
            if addr.Row >= firstData && addr.Row <= lastData then
                match entry with
                | Some { Value = Formula (formula, _) } ->
                    let others = [ firstData .. lastData ] |> List.filter (fun row -> row <> addr.Row)
                    let allBlank = others |> List.forall (fun row -> blankAt m row addr.Col)
                    let sameColumn = not others.IsEmpty && (calculated m addr.Col others addr.Row).IsSome
                    if not others.IsEmpty && (allBlank || sameColumn) then
                        for row in others do fill row addr.Col (shiftFormulaReferences (row - addr.Row) 0 formula)
                | _ -> ()
            let book = withSheetModel sheetId m book
            // 4. A header cell: the column's name is the text.
            if table.HasHeaders && addr.Row = table.StartRow && addr.Col >= table.StartCol && addr.Col <= table.EndCol && addr.Col - table.StartCol < table.Columns.Length then
                let index = addr.Col - table.StartCol
                let oldName = List.item index table.Columns
                let typed = match entry with Some data -> displayValue data.Value | None -> ""
                let wanted = if String.IsNullOrWhiteSpace typed then sprintf "Column%d" (index + 1) else typed
                let taken = table.Columns |> List.mapi (fun at name -> at, name) |> List.filter (fun (at, _) -> at <> index) |> List.map (fun (_, name) -> name.ToUpperInvariant()) |> Set.ofList
                let mutable name = wanted
                let mutable next = 2
                while Set.contains (name.ToUpperInvariant()) taken do
                    name <- wanted + string next
                    next <- next + 1
                let model = sheetModelOf sheetId book
                let header = getCell addr model |> Option.defaultValue { Value = Empty; Format = defaultFormat; Comment = None }
                let model = if header.Value = String name then model else write addr (Some { header with Value = String name }) model
                let columns = table.Columns |> List.mapi (fun at existing -> if at = index then name else existing)
                let renamedBook = putTable table { table with Columns = columns } id model book
                if name = oldName then Some { Book = renamedBook; Cells = List.distinct (List.ofSeq touched) }
                else
                    // Structured references follow: by the table's name anywhere, without it in the table's own cells.
                    let rewritten =
                        rewriteBookFormulas
                            (fun formula -> formula.IndexOf(oldName, StringComparison.OrdinalIgnoreCase) >= 0 && formula.IndexOf('[') >= 0)
                            (fun formulaSheet cell formula ->
                                let ownCell =
                                    match cell with
                                    | Some at -> formulaSheet = sheetId && at.Row >= table.StartRow && at.Row <= table.EndRow && at.Col >= table.StartCol && at.Col <= table.EndCol
                                    | None -> false
                                rewriteFormulaForSheets (RenameTableColumnRefs (table.Name, oldName, name, ownCell)) formula)
                            renamedBook
                    Some { Book = rewritten.Book; Cells = List.distinct (List.ofSeq touched @ rewritten.Cells) }
            else Some { Book = book; Cells = List.distinct (List.ofSeq touched) }

/// Remove a table. Keeping the data (Excel's "convert to range") the cells
/// stay as they are and every structured reference to the table becomes
/// the plain reference it stood for; otherwise the table's cells are
/// emptied as well (and references to it are #REF!).
let deleteTable (memberName: string) (name: string) (keepData: bool) (workbook: Wb.Workbook) : TableChange =
    let table = findTable memberName name workbook
    let inTable (id: Wb.SheetId) (cell: CellAddress option) =
        match cell with
        | Some at -> id = table.SheetId && at.Row >= table.StartRow && at.Row <= table.EndRow && at.Col >= table.StartCol && at.Col <= table.EndCol
        | None -> false
    let rewritten =
        if keepData then
            rewriteBookFormulas
                (fun formula -> formula.IndexOf '[' >= 0)
                (fun id cell formula ->
                    if formula.IndexOf(table.Name, StringComparison.OrdinalIgnoreCase) >= 0 || inTable id cell then
                        Rbt.Spreadsheet.Formula.tableRefsToRanges workbook table.Name id cell formula
                    else formula) workbook
        else { Book = workbook; Cells = [] }
    let model = sheetModelOf table.SheetId rewritten.Book
    let mutable m = { model with Tables = model.Tables |> List.filter (fun t -> not (sameName t.Name table.Name)) }
    let touched = ResizeArray<Wb.SheetId * CellAddress>(rewritten.Cells)
    if not keepData then
        for KeyValue(addr, data) in model.Cells do
            if inTable table.SheetId (Some addr) && data.Value <> Empty then
                m <- setCell addr { data with Value = Empty } m
                touched.Add((table.SheetId, addr))
    let book = withSheetModel table.SheetId m rewritten.Book
    { Book = { book with Tables = book.Tables |> Map.filter (fun _ known -> not (sameName known.Name table.Name)) }
      Cells = List.distinct (List.ofSeq touched) }

/// Give a table another style (None: no style).
let setTableStyle (memberName: string) (name: string) (style: TableStyle option) (workbook: Wb.Workbook) : TableChange =
    let table = findTable memberName name workbook
    let model = sheetModelOf table.SheetId workbook
    match sheetTableNamed table.Name model with
    | Some own when own.Style <> style -> { Book = putTable table table (fun t -> { t with Style = style }) model workbook; Cells = [] }
    | Some _ -> { Book = workbook; Cells = [] }
    | None ->
        // A table the workbook alone knows (defined the old way) gets its record on the sheet.
        let own : Table =
            { Name = table.Name; StartRow = table.StartRow + 1; StartCol = table.StartCol + 1; EndRow = table.EndRow + 1; EndCol = table.EndCol + 1
              HeaderRowCount = (if table.HasHeaders then 1 else 0); Style = style
              AutoFilter = (if table.HasHeaders then Some (table.StartCol + 1, table.EndCol + 1) else None); HasTotals = table.HasTotals }
        { Book = withSheetModel table.SheetId { model with Tables = model.Tables @ [ own ] } workbook; Cells = [] }

// ---------------------------------------------------------------------------
// The status bar: what a selection adds up to
// ---------------------------------------------------------------------------

/// What a status bar says of a selection, as Excel's does: `Count` cells
/// that hold something, `NumericCount` of them numbers (a formula counts by
/// its result; text, booleans and errors are not numbers), and the sum,
/// smallest and largest of the numbers (0 / None when there is none). Rows
/// an AutoFilter hides are left out.
type SelectionStatus = {
    Count: int
    NumericCount: int
    Sum: float
    Min: float option
    Max: float option
}

/// The status of a block of cells. A block larger than the sheet's cells
/// costs one pass over the cells that exist.
let selectionStatus (model: SpreadsheetModel) (top: int, left: int, bottom: int, right: int) : SelectionStatus =
    let mutable count = 0
    let mutable numeric = 0
    let mutable sum = 0.0
    let mutable lowest = System.Double.PositiveInfinity
    let mutable highest = System.Double.NegativeInfinity
    let filtered = model.FilterHiddenRows.Length > 0 && not (Map.isEmpty model.Filters)
    let add (row: int) (cell: CellData) =
        match unwrapFormulaValue cell.Value with
        | CellValue.Empty -> ()
        | CellValue.String "" -> ()
        | value ->
            if not (filtered && isFilterHiddenRow row model) then
                count <- count + 1
                match value with
                | CellValue.Number number ->
                    numeric <- numeric + 1
                    sum <- sum + number
                    if number < lowest then lowest <- number
                    if number > highest then highest <- number
                | _ -> ()
    let area = float (bottom - top + 1) * float (right - left + 1)
    if area <= cellByCellLimit then
        for row in top .. bottom do
            for col in left .. right do
                match Map.tryFind { Row = row; Col = col } model.Cells with
                | Some cell -> add row cell
                | None -> ()
    else
        for KeyValue(addr, cell) in model.Cells do
            if addr.Row >= top && addr.Row <= bottom && addr.Col >= left && addr.Col <= right then add addr.Row cell
    { Count = count; NumericCount = numeric; Sum = sum
      Min = (if numeric > 0 then Some lowest else None)
      Max = (if numeric > 0 then Some highest else None) }

/// The smallest range that holds every cell of the sheet with a value, a
/// format or a comment and every merged cell (Excel's UsedRange: it need not
/// start at A1). None for an empty sheet.
let usedExtent (model: SpreadsheetModel) : CellRange option =
    let mutable any = false
    let mutable top = System.Int32.MaxValue
    let mutable left = System.Int32.MaxValue
    let mutable bottom = -1
    let mutable right = -1
    let take (r0: int) (c0: int) (r1: int) (c1: int) =
        any <- true
        if r0 < top then top <- r0
        if c0 < left then left <- c0
        if r1 > bottom then bottom <- r1
        if c1 > right then right <- c1
    for KeyValue(addr, _) in model.Cells do take addr.Row addr.Col addr.Row addr.Col
    for merge in model.MergedCells do take merge.StartRow merge.StartCol merge.EndRow merge.EndCol
    if any then Some { From = { Row = top; Col = left }; To = { Row = bottom; Col = right } } else None

// ---------------------------------------------------------------------------
// Page setup
// ---------------------------------------------------------------------------

/// The three sections of a header or footer. Each is text in Excel's code
/// syntax: `&P` page number, `&N` pages, `&D` date, `&T` time, `&F` file,
/// `&A` sheet name, `&&` an ampersand, `&"Font,Style"`, `&12` a size ...
type HeaderFooterSections = {
    Left: string
    Center: string
    Right: string
}

/// Split a header / footer string (`&LLeft&CCentre&RRight`) into its
/// sections. Text before the first section marker belongs to the centre, as
/// in Excel; a section named twice is the two texts joined.
let splitHeaderFooter (raw: string) : HeaderFooterSections =
    let left = System.Text.StringBuilder()
    let center = System.Text.StringBuilder()
    let right = System.Text.StringBuilder()
    let mutable current = center
    let text = if isNull raw then "" else raw
    let mutable i = 0
    while i < text.Length do
        let ch = text.[i]
        if ch = '&' && i + 1 < text.Length then
            match text.[i + 1] with
            | 'L' | 'l' -> current <- left; i <- i + 2
            | 'C' | 'c' -> current <- center; i <- i + 2
            | 'R' | 'r' -> current <- right; i <- i + 2
            | '"' ->
                // A font name: up to the closing quote, whatever letters it holds.
                let close = text.IndexOf('"', i + 2)
                let stop = if close < 0 then text.Length else close + 1
                current.Append(text.Substring(i, stop - i)) |> ignore
                i <- stop
            | next ->
                current.Append(ch).Append(next) |> ignore
                i <- i + 2
        else
            current.Append(ch) |> ignore
            i <- i + 1
    { Left = left.ToString(); Center = center.ToString(); Right = right.ToString() }

/// The header / footer string of three sections; None when all are empty.
let joinHeaderFooter (sections: HeaderFooterSections) : string option =
    let part (marker: string) (text: string) = if System.String.IsNullOrEmpty text then "" else marker + text
    match part "&L" sections.Left + part "&C" sections.Center + part "&R" sections.Right with
    | "" -> None
    | joined -> Some joined

/// The papers a host names; any other is given by its OOXML code.
let paperSizeNames : (string * int) list =
    [ "letter", 1; "tabloid", 3; "legal", 5; "executive", 7; "a3", 8; "a4", 9; "a5", 11; "b5", 13 ]

/// The name of a paper-size code ("a4"), when it has one. No code is Letter.
let paperSizeName (code: int option) : string option =
    match code with
    | None -> Some "letter"
    | Some code -> paperSizeNames |> List.tryFind (fun (_, known) -> known = code) |> Option.map fst

let paperSizeCode (name: string) : int option =
    let wanted = (if isNull name then "" else name).Trim().ToLowerInvariant()
    paperSizeNames |> List.tryFind (fun (known, _) -> known = wanted) |> Option.map snd

/// One change of a sheet's page setup, as a host asks for it.
type PageSetupChange =
    | SetOrientation of PrintOrientation
    /// An OOXML paper-size code.
    | SetPaper of int
    /// "top" | "bottom" | "left" | "right" | "header" | "footer", in inches.
    | SetMargin of side: string * inches: float
    /// A scale in percent; the sheet is no longer fitted to pages.
    | SetScale of int
    /// Fit to this many pages wide / tall (0: as many as it takes). A side
    /// that is not given keeps its number.
    | SetFitTo of width: int option * height: int option
    | SetPrintArea of CellRange option
    /// A print area of several areas, each printed as its own run of pages ([]: none).
    | SetPrintAreas of CellRange list
    | SetRepeatRows of (int * int) option
    | SetRepeatCols of (int * int) option
    /// "header" | "footer" | "firstHeader" | "firstFooter" | "evenHeader" | "evenFooter".
    | SetHeaderFooterText of slot: string * raw: string option
    | SetDifferentFirst of bool
    | SetDifferentOddEven of bool
    | SetScaleWithDocument of bool
    | SetPrintGridLines of bool
    | SetPrintHeadings of bool
    | SetCenterHorizontally of bool
    | SetCenterVertically of bool
    | SetPageOrder of PageOrder
    | SetBlackAndWhite of bool
    /// None: numbered from 1 (automatic).
    | SetFirstPageNumber of int option
    | SetCommentsPrint of CellCommentsPrint

/// Apply one change to a page setup. Throws (naming `memberName`) for a
/// value Excel does not take.
let applyPageSetupChange (memberName: string) (setup: PageSetup) (change: PageSetupChange) : PageSetup =
    let fail (message: string) = failwith (memberName + ": " + message)
    match change with
    | SetOrientation orientation -> { setup with Orientation = orientation }
    | SetPaper code ->
        if code < 1 || code > 32767 then fail (sprintf "paperSize code %d is not a paper size" code)
        let width, height = paperSizePoints code
        { setup with PaperSize = Some code; PageWidth = width; PageHeight = height }
    | SetMargin (side, inches) ->
        if System.Double.IsNaN inches || inches < 0.0 || inches > 49.0 then fail (sprintf "margins.%s must be between 0 and 49 inches" side)
        let margins = setup.Margins
        let changed =
            match side with
            | "top" -> { margins with Top = inches }
            | "bottom" -> { margins with Bottom = inches }
            | "left" -> { margins with Left = inches }
            | "right" -> { margins with Right = inches }
            | "header" -> { margins with Header = inches }
            | "footer" -> { margins with Footer = inches }
            | other -> fail (sprintf "margins has no side \"%s\"" other)
        { setup with Margins = changed }
    | SetScale percent ->
        if percent < 10 || percent > 400 then fail "scale must be between 10 and 400 (percent)"
        { setup with Scale = percent; FitToPage = false }
    | SetFitTo (width, height) ->
        let check (what: string) (pages: int option) =
            match pages with
            | Some pages when pages < 0 || pages > 32767 -> fail (sprintf "%s must be between 0 and 32767 (pages; 0 = as many as it takes)" what)
            | _ -> ()
        check "fitToWidth" width
        check "fitToHeight" height
        { setup with FitToPage = true; FitToWidth = defaultArg width setup.FitToWidth; FitToHeight = defaultArg height setup.FitToHeight }
    | SetPrintArea area ->
        match area with
        | Some range when range.From.Row < 0 || range.From.Col < 0 || range.To.Row > maxSheetRow || range.To.Col > maxSheetCol || range.To.Row < range.From.Row || range.To.Col < range.From.Col ->
            fail "printArea is not a range of the sheet"
        | _ -> { setup with PrintArea = area; ExtraPrintAreas = [] }
    | SetPrintAreas areas ->
        if areas |> List.exists (fun range -> range.From.Row < 0 || range.From.Col < 0 || range.To.Row > maxSheetRow || range.To.Col > maxSheetCol || range.To.Row < range.From.Row || range.To.Col < range.From.Col) then
            fail "printArea is not a range of the sheet"
        match areas with
        | [] -> { setup with PrintArea = None; ExtraPrintAreas = [] }
        | first :: rest -> { setup with PrintArea = Some first; ExtraPrintAreas = rest }
    | SetRepeatRows rows ->
        match rows with
        | Some (first, last) when first < 0 || last < first || last > maxSheetRow -> fail "printTitles.rows is not a run of rows"
        | _ -> { setup with RepeatRows = rows }
    | SetRepeatCols cols ->
        match cols with
        | Some (first, last) when first < 0 || last < first || last > maxSheetCol -> fail "printTitles.columns is not a run of columns"
        | _ -> { setup with RepeatCols = cols }
    | SetHeaderFooterText (slot, raw) ->
        let raw = raw |> Option.filter (fun text -> text <> "")
        match raw with
        | Some text when text.Length > 255 -> fail (sprintf "%s is longer than 255 characters (with its codes)" slot)
        | _ -> ()
        let hf = setup.HeaderFooter
        let changed =
            match slot with
            | "header" -> { hf with OddHeader = raw }
            | "footer" -> { hf with OddFooter = raw }
            | "firstHeader" -> { hf with FirstHeader = raw }
            | "firstFooter" -> { hf with FirstFooter = raw }
            | "evenHeader" -> { hf with EvenHeader = raw }
            | "evenFooter" -> { hf with EvenFooter = raw }
            | other -> fail (sprintf "there is no header or footer \"%s\"" other)
        { setup with HeaderFooter = changed }
    | SetDifferentFirst on -> { setup with HeaderFooter = { setup.HeaderFooter with DifferentFirst = on } }
    | SetDifferentOddEven on -> { setup with HeaderFooter = { setup.HeaderFooter with DifferentOddEven = on } }
    | SetScaleWithDocument on -> { setup with HeaderFooter = { setup.HeaderFooter with ScaleWithDoc = on } }
    | SetPrintGridLines on -> { setup with Options = { setup.Options with GridLines = on } }
    | SetPrintHeadings on -> { setup with Options = { setup.Options with Headings = on } }
    | SetCenterHorizontally on -> { setup with Options = { setup.Options with HorizontalCentered = on } }
    | SetCenterVertically on -> { setup with Options = { setup.Options with VerticalCentered = on } }
    | SetPageOrder order -> { setup with PageOrder = order }
    | SetBlackAndWhite on -> { setup with BlackAndWhite = on }
    | SetFirstPageNumber number ->
        match number with
        | Some number when number < -32765 || number > 32767 -> fail "firstPageNumber must be between -32765 and 32767"
        | Some number -> { setup with FirstPageNumber = number; UseFirstPageNumber = true }
        | None -> { setup with FirstPageNumber = 1; UseFirstPageNumber = false }
    | SetCommentsPrint comments -> { setup with CellComments = comments }

/// Apply the changes in order (a later one wins).
let applyPageSetupChanges (memberName: string) (changes: PageSetupChange list) (setup: PageSetup) : PageSetup =
    changes |> List.fold (applyPageSetupChange memberName) setup

/// Manual page breaks as hosts see them: the 0-based rows / columns BEFORE
/// which a break sits (a break before row 20 puts row 20 on a new page).
/// The model - like the file - keeps the 1-based number of the row the break
/// comes after, which is the same number.
let pageBreaksOf (setup: PageSetup) : int list * int list =
    List.sort setup.ManualRowBreaks, List.sort setup.ManualColBreaks

/// Replace the manual page breaks. Throws for a break before the first row
/// / column or outside the sheet.
let setPageBreaks (memberName: string) (rows: int list) (cols: int list) (setup: PageSetup) : PageSetup =
    for row in rows do
        if row < 1 || row > maxSheetRow then failwith (sprintf "%s: there can be no page break before row %d (rows 1 to %d, 0-based)" memberName row maxSheetRow)
    for col in cols do
        if col < 1 || col > maxSheetCol then failwith (sprintf "%s: there can be no page break before column %d (columns 1 to %d, 0-based)" memberName col maxSheetCol)
    let rows = rows |> List.distinct |> List.sort
    let cols = cols |> List.distinct |> List.sort
    if rows = setup.ManualRowBreaks && cols = setup.ManualColBreaks then setup
    else { setup with ManualRowBreaks = rows; ManualColBreaks = cols }

// ---------------------------------------------------------------------------
// The number format a formula's cell takes from what the formula reads
// ---------------------------------------------------------------------------
// Excel: a formula entered into a cell whose format is General shows its
// result in the format of what it reads - `=A1+1` under a date is a date,
// `=SUM(B2:B9)` under amounts an amount, `=TODAY()` a date. The cell gets
// that format when the formula is entered (it is then the cell's own: it
// does not follow the cells later), and only a cell that is General does.

/// Whether a number format says something (General, an empty code and the
/// text format say nothing a result could take over).
let private inheritableFormat (code: string option) : string option =
    match code with
    | Some code when code <> "" && code <> "@" && not (code.Equals("General", StringComparison.OrdinalIgnoreCase)) -> Some code
    | _ -> None

let private isPercentCode (code: string) = code.IndexOf '%' >= 0 && code.IndexOf '"' < 0
/// A format that shows a day (a date, or a date with a time): not a time of day alone, not a duration.
let private isDateCode (code: string) =
    CellText.isDateTimeNumberFormat code
    && (let lower = code.ToLowerInvariant()
        lower.Contains "d" || lower.Contains "y")

/// The number format the result of `formula` (with or without its `=`) is
/// shown in, by Excel's rules, or None for General. `formatAt sheet row col`
/// is the number format of a cell (`sheet`: the sheet a reference names, None
/// for the formula's own).
///
/// - A reference has the format of its cell, a range that of its first cell.
/// - `+` and `-`: the format of the first operand that has one; a date less a
///   date is a number of days (General), a time less a time a time.
/// - `*` and `/`: likewise, but a percentage gives way to another format
///   (an amount times a rate is an amount) and what is divided by a cell of
///   the same kind of format is a ratio (General) - a date times or divided
///   by anything is no date.
/// - Functions that pass values on (SUM, AVERAGE, MIN, MAX, ROUND, IF,
///   INDEX, XLOOKUP ...) have the format of the argument the values come
///   from; date functions give a date (TODAY, DATE, EDATE, EOMONTH,
///   WORKDAY, DATEVALUE: `m/d/yyyy`; NOW: `m/d/yyyy h:mm`; TIME, TIMEVALUE:
///   `h:mm AM/PM`), payment functions an amount (PMT, PV, FV, NPV ...:
///   `$#,##0.00_);[Red]($#,##0.00)`), rate functions a percentage (IRR,
///   RATE ...: `0%`). Everything else (counts, text, comparisons, parts of
///   a date) is General.
let inheritedNumberFormat (formatAt: string option -> int -> int -> string option) (formula: string) : string option =
    let text = if formula.StartsWith "=" then formula else "=" + formula
    match Rbt.Spreadsheet.Formula.parseFormula text with
    | Result.Error _ -> None
    | Result.Ok ast ->
        let rec formatOf (expr: Rbt.Spreadsheet.Formula.Expr) : string option =
            match expr with
            | Rbt.Spreadsheet.Formula.CellRef cell -> inheritableFormat (formatAt cell.Sheet cell.Row cell.Col)
            | Rbt.Spreadsheet.Formula.RangeRef (first, last) -> inheritableFormat (formatAt first.Sheet (min first.Row last.Row) (min first.Col last.Col))
            | Rbt.Spreadsheet.Formula.SpillRef cell -> inheritableFormat (formatAt cell.Sheet cell.Row cell.Col)
            | Rbt.Spreadsheet.Formula.UnaryOp (op, operand) when op = "-" || op = "+" -> formatOf operand
            | Rbt.Spreadsheet.Formula.UnaryOp ("%", _) -> None
            | Rbt.Spreadsheet.Formula.BinOp (left, op, right) when op = "+" || op = "-" ->
                (match formatOf left, formatOf right with
                 | Some a, Some b when op = "-" && isDateCode a && isDateCode b -> None
                 // (A number of days added to a date - whichever side the date is on - is a date.)
                 | Some a, Some b when isDateCode b && not (CellText.isDateTimeNumberFormat a) -> Some b
                 | Some a, _ -> Some a
                 | None, other -> other)
            | Rbt.Spreadsheet.Formula.BinOp (left, op, right) when op = "*" || op = "/" ->
                let usable (code: string option) = code |> Option.filter (fun code -> not (CellText.isDateTimeNumberFormat code))
                (match usable (formatOf left), usable (formatOf right) with
                 | Some a, Some b when op = "/" && (a = b || (isPercentCode a && isPercentCode b)) -> None
                 | Some a, Some b when isPercentCode a && not (isPercentCode b) -> Some b
                 | Some a, Some b when op = "/" && isPercentCode b -> Some a
                 | Some a, _ -> Some a
                 // (One over an amount is no amount.)
                 | None, Some _ when op = "/" -> None
                 | None, other -> other)
            | Rbt.Spreadsheet.Formula.Function (name, args) ->
                let arg (index: int) = List.tryItem index args |> Option.bind formatOf
                let firstOf (indexes: int list) = indexes |> List.tryPick arg
                match name.ToUpperInvariant() with
                | "TODAY" | "DATE" | "EDATE" | "EOMONTH" | "WORKDAY" | "WORKDAY.INTL" | "DATEVALUE" -> Some "m/d/yyyy"
                | "NOW" -> Some "m/d/yyyy h:mm"
                | "TIME" | "TIMEVALUE" -> Some "h:mm AM/PM"
                | "PMT" | "PV" | "FV" | "NPV" | "XNPV" | "IPMT" | "PPMT" | "CUMIPMT" | "CUMPRINC" | "SLN" | "DB" | "DDB" | "SYD" -> Some "$#,##0.00_);[Red]($#,##0.00)"
                | "IRR" | "XIRR" | "MIRR" | "RATE" -> Some "0%"
                | "SUM" | "AVERAGE" | "AVERAGEA" | "MIN" | "MAX" | "MINA" | "MAXA" | "MEDIAN" | "MODE" | "MODE.SNGL" | "PRODUCT" | "SUMPRODUCT" | "LARGE" | "SMALL"
                | "TRIMMEAN" | "GEOMEAN" | "HARMEAN" | "PERCENTILE" | "PERCENTILE.INC" | "PERCENTILE.EXC" | "QUARTILE" | "QUARTILE.INC" | "QUARTILE.EXC" ->
                    args |> List.tryPick formatOf
                | "ROUND" | "ROUNDUP" | "ROUNDDOWN" | "MROUND" | "CEILING" | "CEILING.MATH" | "FLOOR" | "FLOOR.MATH" | "INT" | "TRUNC" | "ABS" | "N" | "EVEN" | "ODD"
                | "INDEX" | "OFFSET" | "SUMIFS" | "AVERAGEIFS" | "MAXIFS" | "MINIFS" | "SINGLE" | "ANCHORARRAY" | "SORT" | "UNIQUE" | "FILTER" | "TAKE" | "DROP" -> arg 0
                | "SUMIF" | "AVERAGEIF" -> firstOf [ 2; 0 ]
                | "SUBTOTAL" -> (match List.tryItem 0 args with
                                 // (2, 3, 102, 103: counts.)
                                 | Some (Rbt.Spreadsheet.Formula.Number n) when n = 2.0 || n = 3.0 || n = 102.0 || n = 103.0 -> None
                                 | _ -> args |> List.tail |> List.tryPick formatOf)
                | "AGGREGATE" -> (match List.tryItem 0 args with
                                  | Some (Rbt.Spreadsheet.Formula.Number n) when n = 2.0 || n = 3.0 -> None
                                  | _ -> (if args.Length > 2 then args |> List.skip 2 |> List.tryPick formatOf else None))
                | "IF" -> firstOf [ 1; 2 ]
                | "IFERROR" | "IFNA" -> firstOf [ 0; 1 ]
                | "IFS" -> args |> List.indexed |> List.tryPick (fun (index, value) -> if index % 2 = 1 then formatOf value else None)
                | "SWITCH" -> args |> List.indexed |> List.tryPick (fun (index, value) -> if index >= 2 && index % 2 = 0 then formatOf value else None)
                | "CHOOSE" -> (if args.IsEmpty then None else args |> List.tail |> List.tryPick formatOf)
                | "XLOOKUP" -> arg 2
                | "LOOKUP" -> firstOf [ 2; 1 ]
                | _ -> None
            | _ -> None
        formatOf ast
