module Rbt.Spreadsheet.XlsxImporter

open Fable.Core
open Fable.Core.JsInterop
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula

/// <summary>
/// XLSX import adapter for office-format library integration.
/// Converts office-format workbook data into Rbt.Grid SpreadsheetModel.
/// </summary>

    // Type definitions matching office-format API structure
    type ExcelBorder = {
        style: string option  // "thin", "medium", "thick", "dashed", "dotted", "double"
        color: string option
    }

    type ExcelCell = {
        value: obj option
        formula: string option
        bold: bool option
        italic: bool option
        underline: bool option
        underlineDouble: bool option
        strikethrough: bool option
        fontName: string option
        fontColor: string option
        fillColor: string option
        fontSize: float option
        numberFormat: string option
        horizontalAlign: string option  // "left", "center", "right", "justify"
        verticalAlign: string option    // "top", "middle", "bottom"
        wrapText: bool option
        borderTop: ExcelBorder option
        borderBottom: ExcelBorder option
        borderLeft: ExcelBorder option
        borderRight: ExcelBorder option
        borderDiagonal: ExcelBorder option
        diagonalUp: bool option
        diagonalDown: bool option
        hasComment: bool option
        hasListDropdown: bool option
    }

    type ExcelRichRun = {
        text: string
        bold: bool option
        italic: bool option
        underline: bool option
        underlineDouble: bool option
        strikethrough: bool option
        color: string option
        fontName: string option
        fontSize: float option
        vertAlign: string option
    }

    type ExcelWorkbook = obj
    type ExcelSheet = obj

    /// <summary>
    /// Import interface to office-format library (loaded at runtime).
    /// Expected API structure from office-format:
    ///   api().excel.load(buffer) -> workbook
    ///   api().excel.sheetNames(workbook) -> string[]
    ///   api().excel.sheet(name, workbook) -> sheet
    ///   api().excel.toArray(workbook, sheet) -> any[][]
    ///   api().excel.cellWithStyleAt(ref, sheet) -> CellWithStyle
    ///   api().excel.columnWidth(col, sheet) -> number
    ///   api().excel.rowHeight(row, sheet) -> number
    ///   api().excel.mergedCells(sheet) -> string[] (e.g., ["A1:B2", "C3:D5"])
    ///   api().excel.commentAt(ref, sheet) -> {text: string, author?: string} | null
    /// </summary>
    // Path expressed relative to the F# source dir, but chosen so Fable's
    // source→output rewrite produces "../office-formats/src/Main.js" in the
    // emitted JS — i.e., a sibling lookup inside dist/. No redbat reference.
    [<ImportAll("../../dist/office-formats/src/Main.js")>]
let private officeFormatApi: obj = jsNative

[<Emit("Array.from($0)")>]
let private arrayFrom (items: obj): obj[] = jsNative

[<Emit("typeof $0 === 'function'")>]
let private isFunction (value: obj): bool = jsNative

[<Emit("$0 == null")>]
let private isNullOrUndefined (value: obj): bool = jsNative

[<Emit("$0 == null ? null : $0[$1]")>]
let private getProp (value: obj) (name: string): obj = jsNative

let private getIndex (value: obj) (index: int): obj =
    getProp value (string index)

let private optStringProp source name =
    let value = getProp source name
    if isNullOrUndefined value then None else Some (string value)

let private optBoolProp source name =
    let value = getProp source name
    if isNullOrUndefined value then None else Some (unbox<bool> value)

let private optFloatProp source name =
    let value = getProp source name
    if isNullOrUndefined value then None else Some (unbox<float> value)

let private optIntProp source name =
    let value = getProp source name
    if isNullOrUndefined value then None else Some (unbox<int> value)

let private hexColor (value: string) =
    let trimmed = value.TrimStart('#')
    let cleaned = if trimmed.Length = 8 then trimmed.Substring(2) else trimmed
    if cleaned.Length = 6 then Some ("#" + cleaned) else None

/// Value of a non-formula cell, or the result a formula cell cached in the
/// file. `errorObj` is the error code ("DIV/0!") for error cells.
let private storedValue (rawValue: obj) (errorObj: obj): CellValue =
        if not (isNullOrUndefined errorObj) then
            CellValue.Error ((string errorObj).TrimStart('#'))
        elif isNullOrUndefined rawValue then
            CellValue.Empty
        else
            match rawValue with
            | :? float as n -> CellValue.Number n
            | :? bool as b -> CellValue.Boolean b
            | :? string as s -> CellValue.String s
            | other -> CellValue.String (string other)

    /// <summary>
    /// Convert office-format cell value to Rbt.Grid CellValue. A formula cell
    /// keeps the result the file cached for it: that is what Excel shows on
    /// open, and it stays correct where our engine lacks a function. Only a
    /// formula the file has no result for is left with an Empty value, for
    /// the workbook-level pass in `finishWorkbook` to compute.
    /// </summary>
let private convertCellValue (excelCell: ExcelCell) (errorObj: obj): CellValue =
        let stored =
            match excelCell.value with
            | Some v -> storedValue v errorObj
            | None -> storedValue null errorObj
        match excelCell.formula with
        | Some formula when not (System.String.IsNullOrWhiteSpace(formula)) ->
            CellValue.Formula (formula, stored)
        | _ ->
            match stored with
            | CellValue.String "" -> CellValue.Empty
            | other -> other

    /// <summary>
    /// Convert ExcelBorder to BorderStyle.
    /// </summary>
let private convertBorder (border: ExcelBorder option): BorderStyle option =
        border |> Option.bind (fun b ->
            match b.style with
            | Some style when not (System.String.IsNullOrWhiteSpace(style)) ->
                Some { Style = style; Color = b.color |> Option.bind hexColor }
            | _ -> None
        )

    /// <summary>
    /// Convert office-format cell style to Rbt.Spreadsheet CellFormat.
    /// </summary>
let private convertCellFormat (excelCell: ExcelCell): CellFormat =
        {
            Bold = excelCell.bold |> Option.defaultValue false
            Italic = excelCell.italic |> Option.defaultValue false
            Underline = excelCell.underline |> Option.defaultValue false
            UnderlineDouble = excelCell.underlineDouble |> Option.defaultValue false
            FontName = excelCell.fontName
            TextColor = excelCell.fontColor |> Option.bind hexColor
            BgColor = excelCell.fillColor |> Option.bind hexColor
            FontSize = excelCell.fontSize |> Option.map int  // Convert float to int
            HorizontalAlign = excelCell.horizontalAlign
            VerticalAlign = excelCell.verticalAlign
            WrapText = excelCell.wrapText |> Option.defaultValue false
            Indent = 0
            ShrinkToFit = false
            TextRotation = 0
            NumberFormat = excelCell.numberFormat
            BorderTop = convertBorder excelCell.borderTop
            BorderBottom = convertBorder excelCell.borderBottom
            BorderLeft = convertBorder excelCell.borderLeft
            BorderRight = convertBorder excelCell.borderRight
            BorderDiagonal = convertBorder excelCell.borderDiagonal
            DiagonalUp = excelCell.diagonalUp |> Option.defaultValue false
            DiagonalDown = excelCell.diagonalDown |> Option.defaultValue false
            FilterButton = false
            Strikethrough = excelCell.strikethrough |> Option.defaultValue false
            HasComment = excelCell.hasComment |> Option.defaultValue false
            HasListDropdown = excelCell.hasListDropdown |> Option.defaultValue false
            DataBar = None
            IconSet = None
        }

let private equalsIgnoreCase (left: string) (right: string) =
        if left.Length <> right.Length then
            false
        else
            let mutable index = 0
            let mutable equal = true
            while equal && index < left.Length do
                let l = left.[index]
                let r = right.[index]
                if l <> r && System.Char.ToUpperInvariant l <> System.Char.ToUpperInvariant r then
                    equal <- false
                index <- index + 1
            equal

let private resizeArrayToList (items: ResizeArray<'T>) =
        let mutable result = []
        for i = items.Count - 1 downto 0 do
            result <- items.[i] :: result
        result

let private convertRichRun (runObj: obj): string * RunOverride =
        let vertAlign =
            match optStringProp runObj "vertAlign" with
            | Some value when equalsIgnoreCase value "superscript" -> Some VertAlignSuperscript
            | Some value when equalsIgnoreCase value "subscript" -> Some VertAlignSubscript
            | _ -> None
        let text = optStringProp runObj "text" |> Option.defaultValue ""
        text, {
            Bold = optBoolProp runObj "bold"
            Italic = optBoolProp runObj "italic"
            Underline = optBoolProp runObj "underline"
            UnderlineDouble = optBoolProp runObj "underlineDouble"
            Strikethrough = optBoolProp runObj "strikethrough"
            Color = optStringProp runObj "color" |> Option.bind hexColor
            FontName = optStringProp runObj "fontName"
            FontSize = optFloatProp runObj "fontSize"
            VertAlign = vertAlign
        }

    /// <summary>
    /// Convert Excel column letter (A, B, ..., Z, AA, AB, ...) to zero-based index.
    /// </summary>
let private columnLetterToIndex (letter: string): int =
        let mutable result = 0
        for c in letter do
            let c = System.Char.ToUpperInvariant c
            result <- result * 26 + (int c - int 'A' + 1)
        result - 1

let private parseCellRefSpan (text: string) startIndex endExclusive =
        let mutable index = startIndex
        let mutable col = 0
        let mutable row = 0
        let mutable hasCol = false
        let mutable hasRow = false
        let mutable inRow = false
        let mutable valid = startIndex < endExclusive
        while valid && index < endExclusive do
            let c = text.[index]
            if System.Char.IsDigit c then
                inRow <- true
                hasRow <- true
                let digit = int c - int '0'
                if row > (System.Int32.MaxValue - digit) / 10 then
                    valid <- false
                else
                    row <- row * 10 + digit
            elif System.Char.IsLetter c && not inRow then
                hasCol <- true
                let upper = System.Char.ToUpperInvariant c
                let digit = int upper - int 'A' + 1
                if col > (System.Int32.MaxValue - digit) / 26 then
                    valid <- false
                else
                    col <- col * 26 + digit
            else
                valid <- false
            index <- index + 1
        if valid && hasCol && hasRow && row > 0 && col > 0 then
            Some (row - 1, col - 1)
        else
            None

let private parseMergedCellRange (refRange: string) =
        let colon = refRange.IndexOf(':')
        if colon >= 0 && colon = refRange.LastIndexOf(':') then
            match parseCellRefSpan refRange 0 colon, parseCellRefSpan refRange (colon + 1) refRange.Length with
            | Some (startRow, startCol), Some (endRow, endCol) ->
                Some {
                    StartRow = startRow
                    StartCol = startCol
                    EndRow = endRow
                    EndCol = endCol
                }
            | _ -> None
        else
            None

// ---------------------------------------------------------------------------
// Sheet metadata (api().excel.sheetMeta) -> model
// ---------------------------------------------------------------------------

module Wb = Rbt.Spreadsheet.Workbook

let private arrayProp (source: obj) (name: string): obj[] =
    let value = getProp source name
    if isNullOrUndefined value then [||] else arrayFrom value

let private boolProp (source: obj) (name: string) (fallback: bool) =
    optBoolProp source name |> Option.defaultValue fallback

let private intProp (source: obj) (name: string) (fallback: int) =
    optIntProp source name |> Option.defaultValue fallback

let private invariantNumber (text: string): float option =
    match System.Double.TryParse(text.Trim(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture) with
    | true, n when not (System.Double.IsNaN n) -> Some n
    | _ -> None

/// One corner of a range ("$B$7", "B", "7") as (row0, col0); a part that is
/// absent (whole column / whole row) is None.
let private parseCorner (text: string): (int option * int option) option =
    let mutable col = 0
    let mutable row = 0
    let mutable hasCol = false
    let mutable hasRow = false
    let mutable valid = text <> ""
    for c in text do
        if c = '$' then ()
        elif System.Char.IsLetter c && not hasRow && col < 20000 then
            hasCol <- true
            col <- col * 26 + (int (System.Char.ToUpperInvariant c) - int 'A' + 1)
        elif System.Char.IsDigit c && row < 10000000 then
            hasRow <- true
            row <- row * 10 + (int c - int '0')
        else
            valid <- false
    if not valid || (not hasCol && not hasRow) || (hasRow && row = 0) then None
    else Some ((if hasRow then Some (row - 1) else None), (if hasCol then Some (col - 1) else None))

/// "A1", "$A$1:$C$9", "A:C" or "2:5" as 0-based inclusive bounds, clipped
/// to `maxRow` / `maxCol`. None when malformed or wholly outside the clip.
let private parseRangeClipped (maxRow: int) (maxCol: int) (range: string): (int * int * int * int) option =
    let parts = range.Trim().Split(':')
    let corners =
        match parts.Length with
        | 1 -> parseCorner parts.[0] |> Option.map (fun c -> c, c)
        | 2 ->
            match parseCorner parts.[0], parseCorner parts.[1] with
            | Some a, Some b -> Some (a, b)
            | _ -> None
        | _ -> None
    corners
    |> Option.bind (fun ((r1, c1), (r2, c2)) ->
        let sr = r1 |> Option.defaultValue 0
        let sc = c1 |> Option.defaultValue 0
        let er = r2 |> Option.defaultValue maxRow
        let ec = c2 |> Option.defaultValue maxCol
        let sr, er = min sr er, max sr er
        let sc, ec = min sc ec, max sc ec
        if sr > maxRow || sc > maxCol then None
        else Some (sr, sc, min er maxRow, min ec maxCol))

/// The ranges of an sqref ("A1:A9 C3"), clipped.
let private parseSqref (maxRow: int) (maxCol: int) (sqref: string) =
    sqref.Split([| ' ' |], System.StringSplitOptions.RemoveEmptyEntries)
    |> Array.choose (parseRangeClipped maxRow maxCol)

/// A list validation whose source is a range or a defined name rather than
/// a literal; it can only be resolved once the sheets it may point at exist.
type private PendingList = {
    Bounds: (int * int * int * int) array
    Source: string
    /// The validation without its values (alert texts, style).
    Template: DataValidation
}

let private convertPageSetup (meta: obj): PageSetup =
    let setup =
        let p = getProp meta "pageSetup"
        if isNullOrUndefined p then defaultPageSetup
        else
            let w, h =
                optIntProp p "paperSize"
                |> Option.map paperSizePoints
                |> Option.defaultValue (defaultPageSetup.PageWidth, defaultPageSetup.PageHeight)
            { defaultPageSetup with
                PageWidth = w
                PageHeight = h
                Orientation = (if optStringProp p "orientation" = Some "landscape" then LandscapeOrientation else PortraitOrientation)
                FitToPage = boolProp p "fitToPage" false
                FitToWidth = intProp p "fitToWidth" 1
                FitToHeight = intProp p "fitToHeight" 1
                Scale = intProp p "scale" 100
                FirstPageNumber = intProp p "firstPageNumber" 1
                UseFirstPageNumber = boolProp p "useFirstPageNumber" false
                PageOrder = (if optStringProp p "pageOrder" = Some "overThenDown" then OverThenDown else DownThenOver)
                BlackAndWhite = boolProp p "blackAndWhite" false
                CellComments =
                    match optStringProp p "cellComments" with
                    | Some "asDisplayed" -> CellCommentsAsDisplayed
                    | Some "atEnd" -> CellCommentsAtEnd
                    | _ -> CellCommentsNone }
    let margins =
        let m = getProp meta "pageMargins"
        if isNullOrUndefined m then defaultPageMargins
        else
            let get name fallback = optFloatProp m name |> Option.defaultValue fallback
            ({ Top = get "top" 0.75; Bottom = get "bottom" 0.75; Left = get "left" 0.7; Right = get "right" 0.7
               Header = get "header" 0.3; Footer = get "footer" 0.3 } : PageMargins)
    let options =
        let o = getProp meta "printOptions"
        ({ GridLines = boolProp o "gridLines" false
           Headings = boolProp o "headings" false
           HorizontalCentered = boolProp o "horizontalCentered" false
           VerticalCentered = boolProp o "verticalCentered" false } : PrintOptions)
    let headerFooter =
        let h = getProp meta "headerFooter"
        ({ OddHeader = optStringProp h "oddHeader"
           OddFooter = optStringProp h "oddFooter"
           EvenHeader = optStringProp h "evenHeader"
           EvenFooter = optStringProp h "evenFooter"
           FirstHeader = optStringProp h "firstHeader"
           FirstFooter = optStringProp h "firstFooter"
           DifferentOddEven = boolProp h "differentOddEven" false
           DifferentFirst = boolProp h "differentFirst" false
           ScaleWithDoc = boolProp h "scaleWithDoc" true } : HeaderFooter)
    let wholeSheet = parseRangeClipped 1048575 16383
    { setup with
        Margins = margins
        Options = options
        HeaderFooter = headerFooter
        ManualRowBreaks = arrayProp meta "rowBreaks" |> Array.map unbox<int> |> Array.toList
        ManualColBreaks = arrayProp meta "colBreaks" |> Array.map unbox<int> |> Array.toList
        PrintArea =
            optStringProp meta "printArea"
            |> Option.bind wholeSheet
            |> Option.map (fun (sr, sc, er, ec) -> ({ From = { Row = sr; Col = sc }; To = { Row = er; Col = ec } } : CellRange))
        RepeatRows = optStringProp meta "printTitleRows" |> Option.bind wholeSheet |> Option.map (fun (sr, _, er, _) -> sr, er)
        RepeatCols = optStringProp meta "printTitleCols" |> Option.bind wholeSheet |> Option.map (fun (_, sc, _, ec) -> sc, ec) }

let private convertTables (meta: obj): Table list * Map<string, string list> =
    let mutable columns = Map.empty<string, string list>
    let tables =
        arrayProp meta "tables"
        |> Array.map (fun t ->
            let name = optStringProp t "name" |> Option.defaultValue ""
            columns <- Map.add name (arrayProp t "columns" |> Array.map string |> Array.toList) columns
            let style =
                let s = getProp t "style"
                if isNullOrUndefined s then None
                else
                    // HeaderFill / StripeFill make `applyTableStyles` paint the
                    // table at render time; nothing is baked into the cells.
                    Some ({ Name = optStringProp s "name" |> Option.defaultValue ""
                            ShowFirstColumn = boolProp s "showFirstColumn" false
                            ShowLastColumn = boolProp s "showLastColumn" false
                            ShowRowStripes = boolProp s "showRowStripes" false
                            ShowColumnStripes = boolProp s "showColumnStripes" false
                            HeaderFill = optStringProp s "headerFill"
                            StripeFill = optStringProp s "stripeFill" } : TableStyle)
            let autoFilter =
                match arrayProp t "autoFilter" with
                | [| c1; c2 |] -> Some (unbox<int> c1, unbox<int> c2)
                | _ -> None
            ({ Name = name
               StartRow = intProp t "startRow" 1
               StartCol = intProp t "startCol" 1
               EndRow = intProp t "endRow" 1
               EndCol = intProp t "endCol" 1
               HeaderRowCount = intProp t "headerRowCount" 1
               Style = style
               AutoFilter = autoFilter
               HasTotals = intProp t "totalsRowCount" 0 > 0 } : Table))
        |> Array.toList
    tables, columns

/// The sheet's conditional formatting as model rules, kept by range (a
/// whole-column rule stays one). The reader's rule records are turned into
/// the shared OOXML shape; `ConditionalFormatting.ofOoxml` does the rest.
let private convertConditionalFormats (meta: obj): ConditionalFormatRange list =
    let cfvos (items: obj array) =
        items |> Array.map (fun item -> optStringProp item "type" |> Option.defaultValue "min", optStringProp item "value") |> Array.toList
    [ for cf in arrayProp meta "conditionalFormats" ->
        ({ Sqref = optStringProp cf "range" |> Option.defaultValue ""
           Rules =
            [ for rule in arrayProp cf "rules" ->
                let style = getProp rule "style"
                ({ Type = optStringProp rule "type" |> Option.defaultValue "cellIs"
                   Operator = optStringProp rule "operator"
                   Formulas = arrayProp rule "formulas" |> Array.map string |> Array.toList
                   Text = optStringProp rule "text"
                   Rank = optIntProp rule "rank"
                   Percent = boolProp rule "percent" false
                   Bottom = boolProp rule "bottom" false
                   AboveAverage = boolProp rule "aboveAverage" true
                   EqualAverage = boolProp rule "equalAverage" false
                   StdDev = optIntProp rule "stdDev"
                   ColorScale =
                    arrayProp rule "colorScale"
                    |> Array.map (fun stop ->
                        optStringProp stop "type" |> Option.defaultValue "min",
                        optStringProp stop "value",
                        optStringProp stop "color" |> Option.defaultValue "FFFFFF")
                    |> Array.toList
                   DataBarColor = optStringProp rule "dataBarColor"
                   DataBarCfvo = cfvos (arrayProp rule "dataBarCfvo")
                   IconSet = optStringProp rule "iconSet"
                   IconSetCfvo = cfvos (arrayProp rule "iconSetCfvo")
                   IconSetReverse = boolProp rule "iconSetReverse" false
                   Priority = optIntProp rule "priority"
                   StopIfTrue = boolProp rule "stopIfTrue" false
                   Style =
                    if isNullOrUndefined style then None
                    else
                        Some ({ FontColor = optStringProp style "fontColor"
                                FillColor = optStringProp style "fillColor"
                                Bold = boolProp style "bold" false
                                Italic = boolProp style "italic" false } : ConditionalFormatting.OoxmlStyle) } : ConditionalFormatting.OoxmlRule) ] } : ConditionalFormatting.OoxmlConditionalFormat) ]
    |> ConditionalFormatting.ofOoxml

/// Validations with a literal definition become rules now; list validations
/// that point at a range or a name are returned as pending.
let private convertValidations (maxRow: int) (maxCol: int) (meta: obj): Map<CellAddress, DataValidation> * PendingList list =
    let mutable rules = Map.empty<CellAddress, DataValidation>
    let pending = ResizeArray<PendingList>()
    for dv in arrayProp meta "validations" do
        let bounds = optStringProp dv "range" |> Option.map (parseSqref maxRow maxCol) |> Option.defaultValue [||]
        let anchor =
            if bounds.Length > 0 then (let (sr, sc, _, _) = bounds.[0] in { Row = sr; Col = sc })
            else { Row = 0; Col = 0 }
        let ooxml : ConditionalFormatting.OoxmlValidation =
            { Type = optStringProp dv "type" |> Option.defaultValue "none"
              Operator = optStringProp dv "operator"
              Formula1 = optStringProp dv "formula1"
              Formula2 = optStringProp dv "formula2"
              AllowBlank = boolProp dv "allowBlank" false
              ShowErrorMessage = boolProp dv "showErrorMessage" false
              ShowInputMessage = boolProp dv "showInputMessage" false
              ErrorStyle = optStringProp dv "errorStyle"
              ErrorTitle = optStringProp dv "errorTitle"
              Error = optStringProp dv "error"
              PromptTitle = optStringProp dv "promptTitle"
              Prompt = optStringProp dv "prompt" }
        match ConditionalFormatting.validationRuleOfOoxml ooxml with
        | Some rule ->
            let validation = ConditionalFormatting.validationWithOoxml anchor ooxml rule
            for (sr, sc, er, ec) in bounds do
                for r in sr .. er do
                    for c in sc .. ec do
                        rules <- Map.add { Row = r; Col = c } validation rules
        | None ->
            match ooxml.Type, ooxml.Formula1 with
            | "list", Some source ->
                // The alert settings travel with the pending list; its
                // values are filled in once the source can be read.
                pending.Add { Bounds = bounds; Source = source.Trim()
                              Template = ConditionalFormatting.validationWithOoxml anchor ooxml (ListValues []) }
            | _ -> ()
    rules, List.ofSeq pending

/// Values of a list validation's source range. `lookupSheet None` is the
/// sheet that owns the validation.
let private resolveListSource (lookupSheet: string option -> SpreadsheetModel option) (lookupName: string -> string option) (source: string): string list option =
    let rec resolve (depth: int) (text: string) =
        let text = (if text.StartsWith "=" then text.Substring 1 else text).Trim()
        let sheetName, rangeText =
            match text.LastIndexOf '!' with
            | -1 -> None, text
            | idx ->
                let raw = text.Substring(0, idx)
                let name =
                    if raw.Length >= 2 && raw.StartsWith "'" && raw.EndsWith "'" then raw.Substring(1, raw.Length - 2).Replace("''", "'")
                    else raw
                Some name, text.Substring(idx + 1)
        match lookupSheet sheetName with
        | None -> None
        | Some model ->
            match parseRangeClipped (model.RowCount - 1) (model.ColCount - 1) rangeText with
            | Some (sr, sc, er, ec) ->
                [ for r in sr .. er do
                    for c in sc .. ec do
                        let text = displayValue (getCellValue { Row = r; Col = c } model)
                        if text <> "" then yield text ]
                |> Some
            | None ->
                // Not a range: a defined name whose refers-to is one.
                if depth > 0 || sheetName.IsSome then None
                else lookupName text |> Option.bind (resolve (depth + 1))
    resolve 0 source

let private applyPendingLists (lookupSheet: string option -> SpreadsheetModel option) (lookupName: string -> string option) (pending: PendingList list) (model: SpreadsheetModel): SpreadsheetModel =
    let mutable validations = model.Validations
    for item in pending do
        match resolveListSource lookupSheet lookupName item.Source with
        | Some values when not (List.isEmpty values) ->
            let validation = { item.Template with Rule = ListValues values }
            for (sr, sc, er, ec) in item.Bounds do
                for r in sr .. er do
                    for c in sc .. ec do
                        validations <- Map.add { Row = r; Col = c } validation validations
        | _ -> ()
    if obj.ReferenceEquals(validations, model.Validations) then model
    else { model with Validations = validations }

// ---------------------------------------------------------------------------
// Sheet import
// ---------------------------------------------------------------------------

type private SheetImport = {
    Model: SpreadsheetModel
    PendingLists: PendingList list
    /// `<tableColumn>` names per table, for structured references.
    TableColumns: Map<string, string list>
    State: string
}

let private excelApi (): obj =
    let api = officeFormatApi?``default``()
    api?excel

let private importSheetCore (excel: obj) (workbook: ExcelWorkbook) (sheet: ExcelSheet): SheetImport =
            let mutable rowCount = 0
            let mutable colCount = 0
            let setDimensionsFromSheet () =
                if isFunction excel?dimensions then
                    let dims = excel?dimensions(sheet)
                    rowCount <- optIntProp dims "rowCount" |> Option.defaultValue 0
                    colCount <- optIntProp dims "colCount" |> Option.defaultValue 0

            // Build cell map with styles
            let mutable cellMap = Map.empty<CellAddress, CellData>
            let mutable richRunsMap = Map.empty<CellAddress, (string * RunOverride) list>
            let mutable columnWidths = Map.empty<int, float>
            let mutable rowHeights = Map.empty<int, float>

            let extractBorder (borderObj: obj) =
                if isNullOrUndefined borderObj then None
                else
                    Some {
                        style = optStringProp borderObj "style"
                        color = optStringProp borderObj "color"
                    }

            let excelCellFromObj rawValue cellObj formulaOverride hasCommentOverride hasListDropdownOverride =
                {
                    value = if isNullOrUndefined rawValue then None else Some rawValue
                    formula =
                        match formulaOverride with
                        | Some value -> Some value
                        | None -> optStringProp cellObj "formula"
                    bold = optBoolProp cellObj "bold"
                    italic = optBoolProp cellObj "italic"
                    underline = optBoolProp cellObj "underline"
                    underlineDouble = optBoolProp cellObj "underlineDouble"
                    strikethrough = optBoolProp cellObj "strikethrough"
                    fontName = optStringProp cellObj "fontName"
                    fontColor = optStringProp cellObj "fontColor"
                    fillColor = optStringProp cellObj "fillColor"
                    fontSize = optFloatProp cellObj "fontSize"
                    numberFormat = optStringProp cellObj "numberFormat"
                    horizontalAlign = optStringProp cellObj "horizontalAlign"
                    verticalAlign = optStringProp cellObj "verticalAlign"
                    wrapText = optBoolProp cellObj "wrapText"
                    borderTop = extractBorder (getProp cellObj "borderTop")
                    borderBottom = extractBorder (getProp cellObj "borderBottom")
                    borderLeft = extractBorder (getProp cellObj "borderLeft")
                    borderRight = extractBorder (getProp cellObj "borderRight")
                    borderDiagonal = extractBorder (getProp cellObj "borderDiagonal")
                    diagonalUp = optBoolProp cellObj "diagonalUp"
                    diagonalDown = optBoolProp cellObj "diagonalDown"
                    hasComment =
                        match hasCommentOverride with
                        | Some value -> Some value
                        | None -> optBoolProp cellObj "hasComment"
                    hasListDropdown =
                        match hasListDropdownOverride with
                        | Some value -> Some value
                        | None -> optBoolProp cellObj "hasListDropdown"
                }

            let importCell row col excelRef rawValue cellObj directRichRuns knownHasComment =
                if not (isNullOrUndefined cellObj) || not (isNullOrUndefined rawValue) then
                    let excelCell: ExcelCell =
                        excelCellFromObj rawValue cellObj None None None

                    let cellValue = convertCellValue excelCell (getProp cellObj "error")
                    let cellFormat = convertCellFormat excelCell

                    // Read comment if present. Older office-format builds expose comments as a collection only.
                    let commentObj =
                        match knownHasComment with
                        | Some false -> null
                        | _ ->
                            if isFunction excel?commentAt then excel?commentAt(excelRef, sheet)
                            else null
                    let comment =
                        if isNullOrUndefined commentObj then None
                        else
                            let text = optStringProp commentObj "text" |> Option.defaultValue ""
                            let author = optStringProp commentObj "author"
                            if System.String.IsNullOrWhiteSpace(text) then None
                            else Some { Text = text; Author = author }

                    let cellFormat =
                        if comment.IsSome then { cellFormat with HasComment = true }
                        else cellFormat

                    // Keep style-only cells. Excel renders fills, borders,
                    // comment markers, and validation dropdowns even when
                    // the cell has no value.
                    if cellValue <> CellValue.Empty || cellFormat <> defaultFormat || comment.IsSome then
                        let addr = { Row = row; Col = col }
                        cellMap <- Map.add addr { Value = cellValue; Format = cellFormat; Comment = comment } cellMap

                        let richRunObjs =
                            match directRichRuns with
                            | Some runs -> runs
                            | None ->
                                if isFunction excel?richRunsAt then excel?richRunsAt(excelRef, workbook, sheet)
                                else null
                        if not (isNullOrUndefined richRunObjs) then
                            let richRunArray = arrayFrom richRunObjs
                            let runs = ResizeArray<string * RunOverride>(richRunArray.Length)
                            for runObj in richRunArray do
                                runs.Add(convertRichRun runObj)
                            if runs.Count > 0 then
                                let runs = resizeArrayToList runs
                                richRunsMap <- Map.add addr runs richRunsMap

            // Import actual cells with styles in one pass when the
            // office-format API supports it. The coordinate fallback is
            // kept for older browser bundles.
            let styleCache = System.Collections.Generic.Dictionary<int, CellFormat>()
            let styleFormat styleIndex =
                if styleIndex < 0 then
                    defaultFormat
                else
                    match styleCache.TryGetValue styleIndex with
                    | true, cached -> cached
                    | false, _ ->
                        let styleObj = excel?styleAt(styleIndex, workbook)
                        let excelCell = excelCellFromObj null styleObj None None None
                        let format = convertCellFormat excelCell
                        styleCache.[styleIndex] <- format
                        format

            let importCompactCell row col rawValue formulaObj (errorObj: obj) styleIndex hasComment hasListDropdown richRunObjs =
                if row >= 0 && col >= 0 then
                    if row + 1 > rowCount then rowCount <- row + 1
                    if col + 1 > colCount then colCount <- col + 1
                    let formula =
                        if isNullOrUndefined formulaObj then None else Some (string formulaObj)
                    let valueCell = excelCellFromObj rawValue null formula None None
                    let cellValue = convertCellValue valueCell errorObj
                    let baseFormat = styleFormat styleIndex
                    let cellFormat =
                        { baseFormat with
                            HasComment = baseFormat.HasComment || hasComment
                            HasListDropdown = baseFormat.HasListDropdown || hasListDropdown }
                    let comment =
                        if hasComment && isFunction excel?commentAt then
                            let excelRef = sprintf "%s%d" (columnToLetter col) (row + 1)
                            let commentObj = excel?commentAt(excelRef, sheet)
                            if isNullOrUndefined commentObj then None
                            else
                                let text = optStringProp commentObj "text" |> Option.defaultValue ""
                                let author = optStringProp commentObj "author"
                                if System.String.IsNullOrWhiteSpace(text) then None
                                else Some { Text = text; Author = author }
                        else
                            None
                    let cellFormat =
                        if comment.IsSome then { cellFormat with HasComment = true }
                        else cellFormat
                    if cellValue <> CellValue.Empty || cellFormat <> defaultFormat || comment.IsSome then
                        let addr = { Row = row; Col = col }
                        cellMap <- Map.add addr { Value = cellValue; Format = cellFormat; Comment = comment } cellMap

                        if not (isNullOrUndefined richRunObjs) then
                            let richRunArray = arrayFrom richRunObjs
                            let runs = ResizeArray<string * RunOverride>(richRunArray.Length)
                            for runObj in richRunArray do
                                runs.Add(convertRichRun runObj)
                            if runs.Count > 0 then
                                richRunsMap <- Map.add addr (resizeArrayToList runs) richRunsMap

            if isFunction excel?cellsColumnar && isFunction excel?styleAt then
                setDimensionsFromSheet ()
                let columnarCells = excel?cellsColumnar(workbook, sheet)
                let rows: int[] = unbox (getProp columnarCells "rows")
                let cols: int[] = unbox (getProp columnarCells "cols")
                let values: obj[] = unbox (getProp columnarCells "values")
                let formulas: obj[] = unbox (getProp columnarCells "formulas")
                let styleIndexes: int[] = unbox (getProp columnarCells "styleIndexes")
                let hasComments: bool[] = unbox (getProp columnarCells "hasComments")
                let hasListDropdowns: bool[] = unbox (getProp columnarCells "hasListDropdowns")
                let richRuns: obj[] = unbox (getProp columnarCells "richRuns")
                let errors: obj[] = arrayProp columnarCells "errors"
                let count =
                    [| rows.Length; cols.Length; values.Length; formulas.Length; styleIndexes.Length
                       hasComments.Length; hasListDropdowns.Length; richRuns.Length |]
                    |> Array.min
                for index in 0 .. count - 1 do
                    importCompactCell
                        rows.[index]
                        cols.[index]
                        values.[index]
                        formulas.[index]
                        (if index < errors.Length then errors.[index] else null)
                        styleIndexes.[index]
                        hasComments.[index]
                        hasListDropdowns.[index]
                        richRuns.[index]
            elif isFunction excel?cellsCompact && isFunction excel?styleAt then
                setDimensionsFromSheet ()
                let compactCells = excel?cellsCompact(workbook, sheet) |> arrayFrom
                for compactCell in compactCells do
                    let row = getIndex compactCell 0 |> unbox<int>
                    let col = getIndex compactCell 1 |> unbox<int>
                    importCompactCell
                        row
                        col
                        (getIndex compactCell 2)
                        (getIndex compactCell 3)
                        null
                        (getIndex compactCell 4 |> unbox<int>)
                        (getIndex compactCell 5 |> unbox<bool>)
                        (getIndex compactCell 6 |> unbox<bool>)
                        (getIndex compactCell 7)
            elif isFunction excel?cellsWithStyle then
                setDimensionsFromSheet ()
                let cellObjs = excel?cellsWithStyle(workbook, sheet) |> arrayFrom
                for cellObj in cellObjs do
                    let excelRef = optStringProp cellObj "ref" |> Option.defaultValue ""
                    match parseCellRef excelRef with
                    | Some (row, col) ->
                        if row + 1 > rowCount then rowCount <- row + 1
                        if col + 1 > colCount then colCount <- col + 1
                        let hasComment = optBoolProp cellObj "hasComment" |> Option.defaultValue false
                        importCell row col excelRef (getProp cellObj "value") cellObj (Some (getProp cellObj "richRuns")) (Some hasComment)
                    | None -> ()
            else
                // Older office-format bundles do not expose streaming cell
                // APIs. Import raw values from toArray without calling
                // cellWithStyleAt per coordinate; that lookup scans worksheet
                // rows and makes this legacy path quadratic.
                let dataArray: obj[][] = excel?toArray(workbook, sheet)
                rowCount <- dataArray.Length
                colCount <-
                    if rowCount > 0 then
                        dataArray |> Array.map (fun row -> row.Length) |> Array.max
                    else
                        0
                for row in 0 .. rowCount - 1 do
                    for col in 0 .. colCount - 1 do
                        let colLetter = columnToLetter col
                        let excelRef = sprintf "%s%d" colLetter (row + 1)
                        let rawValue =
                            if row < dataArray.Length && col < dataArray.[row].Length then dataArray.[row].[col]
                            else null
                        importCell row col excelRef rawValue null None None

            let addColumnWidth col width =
                if width > 0.0 then
                    // Convert Excel width units to pixels (approximate: 1 Excel unit ≈ 7 pixels)
                    let pixelWidth = width * 7.0
                    columnWidths <- Map.add col pixelWidth columnWidths

            let addRowHeight row height =
                if height > 0.0 then
                    // Convert points to pixels (1 point = 1.333 pixels at 96 DPI)
                    let pixelHeight = height * 1.333
                    rowHeights <- Map.add row pixelHeight rowHeights

            // Import column widths (office-format provides in Excel units)
            if isFunction excel?columnWidths then
                let widths = excel?columnWidths(colCount, sheet) |> arrayFrom
                let last = min colCount widths.Length
                if last > 0 then
                    for col in 0 .. last - 1 do
                        addColumnWidth col (unbox<float> widths.[col])
            else
                for col in 0 .. colCount - 1 do
                    let width: float =
                        if isFunction excel?columnWidth then excel?columnWidth(col, sheet)
                        else 0.0
                    addColumnWidth col width

            // Import row heights (office-format provides in points)
            if isFunction excel?rowHeights then
                let heights = excel?rowHeights(rowCount, sheet) |> arrayFrom
                let last = min rowCount heights.Length
                if last > 0 then
                    for row in 0 .. last - 1 do
                        addRowHeight row (unbox<float> heights.[row])
            else
                for row in 0 .. rowCount - 1 do
                    let height: float =
                        if isFunction excel?rowHeight then excel?rowHeight(row, sheet)
                        else 0.0
                    addRowHeight row height

            // Import merged cells (office-format provides as array of ref strings like "A1:B2")
            let mergedCells: string[] = excel?mergedCells(sheet)
            let mergedCellList =
                if isNull mergedCells then []
                else
                    let ranges = ResizeArray<MergedCell>()
                    for refRange in mergedCells do
                        try
                            match parseMergedCellRange refRange with
                            | Some range -> ranges.Add range
                            | None -> ()
                        with _ -> ()
                    resizeArrayToList ranges

            let modelRows = max rowCount 100  // Minimum 100 rows
            let modelCols = max colCount 26   // Minimum 26 columns (A-Z)

            // Everything that is not a cell: freeze panes, hidden rows and
            // columns, tables, validations, conditional formats, print setup.
            let meta: obj =
                if isFunction excel?sheetMeta then excel?sheetMeta(workbook, sheet) else null
            let hasMeta = not (isNullOrUndefined meta)

            // A size of 0 is how the model marks a hidden row / column.
            if hasMeta then
                for row in arrayProp meta "hiddenRows" do
                    rowHeights <- Map.add (unbox<int> row) 0.0 rowHeights
                for col in arrayProp meta "hiddenCols" do
                    columnWidths <- Map.add (unbox<int> col) 0.0 columnWidths

            // Rules over whole columns / rows are stored per cell, so they
            // are clipped to the used range plus room to grow.
            let ruleRows = modelRows + 200 - 1
            let ruleCols = modelCols + 50 - 1
            let validations, pendingLists =
                if hasMeta then convertValidations ruleRows ruleCols meta else Map.empty, []
            let tables, tableColumns =
                if hasMeta then convertTables meta else [], Map.empty

            // The model has no hyperlink store; show link cells the way Excel
            // styles them so they are at least recognisable.
            if hasMeta then
                for link in arrayProp meta "hyperlinks" do
                    match optStringProp link "range" |> Option.bind (parseRangeClipped (modelRows - 1) (modelCols - 1)) with
                    | Some (sr, sc, er, ec) when (er - sr + 1) * (ec - sc + 1) <= 4096 ->
                        for r in sr .. er do
                            for c in sc .. ec do
                                let addr = { Row = r; Col = c }
                                match Map.tryFind addr cellMap with
                                | Some data ->
                                    let format =
                                        { data.Format with
                                            TextColor = data.Format.TextColor |> Option.orElse (Some "#0563C1")
                                            Underline = true }
                                    cellMap <- Map.add addr { data with Format = format } cellMap
                                | None -> ()
                    | _ -> ()

            // Create model with imported data
            let importedModel = {
                Cells = cellMap
                RowCount = modelRows
                ColCount = modelCols
                Selection = SingleCell (0, 0)
                EditingCell = None
                EditBuffer = ""
                Clipboard = None
                History = { Past = []; Future = []; MaxSize = 50 }
                SortState = None
                Filters = Map.empty
                FilterVisibleRows = None
                FilterHiddenRows = [||]
                RowHeights = rowHeights
                ColumnWidths = columnWidths
                SearchQuery = None
                SearchCaseSensitive = None
                SearchResults = [||]
                SearchIndex = 0
                ConditionalFormats = if hasMeta then convertConditionalFormats meta else []
                RichRuns = richRunsMap
                MergedCells = mergedCellList
                FrozenRows = if hasMeta then intProp meta "frozenRows" 0 else 0
                FrozenCols = if hasMeta then intProp meta "frozenCols" 0 else 0
                PageSetup = if hasMeta then convertPageSetup meta else defaultPageSetup
                Tables = tables
                CfBoldCells = Set.empty
                ShowGridLines = if hasMeta then boolProp meta "showGridLines" true else true
                Spill = Map.empty
                Validations = validations
                Macros = ""
            }
            {
                Model = importedModel
                PendingLists = pendingLists
                TableColumns = tableColumns
                State = (if hasMeta then optStringProp meta "state" |> Option.defaultValue "visible" else "visible")
            }

// ---------------------------------------------------------------------------
// Workbook assembly
// ---------------------------------------------------------------------------

/// One sheet of an imported workbook. `State` is "visible", "hidden" or
/// "veryHidden"; hidden sheets are imported (formulas refer to them) but a
/// host should not offer a tab for them.
type ImportedSheet = {
    Name: string
    Model: SpreadsheetModel
    State: string
    Visible: bool
}

/// A whole imported workbook: every sheet in file order, and the workbook
/// context (sheets, structured-reference tables, defined names) the formula
/// engine needs to evaluate anything that looks outside its own sheet.
type ImportedWorkbook = {
    Sheets: ImportedSheet array
    Context: Rbt.Spreadsheet.Workbook.Workbook
    /// The file used the 1904 date system. Date serials have already been
    /// converted to the 1900 system the engine uses.
    Date1904: bool
}

/// Sheet id of the sheet at `index`. Must agree with the ids the Spreadsheet
/// component assigns (`sheetN`, 1-based).
let sheetIdOfIndex (index: int): string =
    sprintf "sheet%d" (index + 1)

/// Workbook context for a set of sheet models: every table on every sheet is
/// registered for structured references. `tableColumns` supplies declared
/// column names by table name; a table without them reads its header row.
let buildWorkbookContext (models: SpreadsheetModel array) (names: string array) (visible: bool array)
                         (tableColumns: Map<string, string list>)
                         (definedNames: Map<Rbt.Spreadsheet.Workbook.SheetId option * string, string>)
                         (activeIndex: int): Rbt.Spreadsheet.Workbook.Workbook =
    let sheetEntries =
        [ for i in 0 .. models.Length - 1 ->
            let id = sheetIdOfIndex i
            id, ({ Id = id
                   Name = (if i < names.Length then names.[i] else sprintf "Sheet%d" (i + 1))
                   Model = models.[i]
                   Visible = (if i < visible.Length then visible.[i] else true) } : Rbt.Spreadsheet.Workbook.Sheet) ]
    let tables =
        [ for i in 0 .. models.Length - 1 do
            for t in models.[i].Tables do
                if t.Name <> "" then
                    let columns = Map.tryFind t.Name tableColumns |> Option.defaultValue []
                    yield t.Name,
                          Wb.makeTableFromRegion (sheetIdOfIndex i) t.Name t.StartRow t.StartCol t.EndRow t.EndCol
                              t.HeaderRowCount (if t.HasTotals then 1 else 0) columns models.[i] ]
    { Sheets = Map.ofList sheetEntries
      SheetOrder = sheetEntries |> List.map fst
      ActiveSheet = sheetIdOfIndex (max 0 (min (models.Length - 1) activeIndex))
      Tables = Map.ofList tables
      Names = definedNames }

let private isUncached (cell: CellData) =
    match cell.Value with
    | CellValue.Formula (_, CellValue.Empty) -> true
    | _ -> false

/// Compute the formulas the file carries no result for (workbooks written by
/// libraries rather than by Excel). The calculation runs over the WHOLE
/// workbook — a sheet recalculated on its own turns every cross-sheet
/// reference into #REF! and every defined name into #NAME? — and afterwards
/// every formula that did have a cached result gets it back, so values Excel
/// computed are never replaced by ours.
let private computeUncachedFormulas (context: Rbt.Spreadsheet.Workbook.Workbook): Rbt.Spreadsheet.Workbook.Workbook =
    let needsWork =
        context.Sheets |> Map.exists (fun _ sheet -> sheet.Model.Cells |> Map.exists (fun _ cell -> isUncached cell))
    if not needsWork then context
    else
        let recalculated = recalculateWorkbook context
        let sheets =
            recalculated.Sheets
            |> Map.map (fun id sheet ->
                match Map.tryFind id context.Sheets with
                | None -> sheet
                | Some original ->
                    let mutable cells = sheet.Model.Cells
                    for KeyValue(addr, cell) in original.Model.Cells do
                        match cell.Value with
                        | CellValue.Formula (_, cached) when cached <> CellValue.Empty ->
                            cells <- Map.add addr cell cells
                        | _ -> ()
                    { sheet with Model = { sheet.Model with Cells = cells } })
        { recalculated with Sheets = sheets }

let private requireWorkbook (workbook: ExcelWorkbook): ExcelWorkbook =
    if isNullOrUndefined workbook then
        failwith "Not a readable XLSX workbook: the package has no workbook part"
    workbook

/// Convert a loaded office-format workbook: every sheet, plus the workbook
/// context. Throws when the workbook has no sheets.
let importWorkbook (workbook: ExcelWorkbook): ImportedWorkbook =
    let workbook = requireWorkbook workbook
    let excel = excelApi ()
    let names: string[] = excel?sheetNames(workbook) |> arrayFrom |> Array.map string
    if names.Length = 0 then failwith "XLSX workbook has no worksheets"
    let imports =
        names |> Array.map (fun name -> importSheetCore excel workbook (excel?sheet(name, workbook)))
    let info: obj = if isFunction excel?workbookInfo then excel?workbookInfo(workbook) else null

    // Defined names. `localSheetId` indexes the workbook's sheet list, which
    // is the list just imported (every sheet is loaded here).
    let definedNames =
        arrayProp info "definedNames"
        |> Array.choose (fun entry ->
            match optStringProp entry "name", optStringProp entry "refersTo" with
            | Some name, Some refersTo ->
                let scope = optIntProp entry "localSheetId" |> Option.map sheetIdOfIndex
                Some (scope, name, refersTo)
            | _ -> None)
        |> Wb.namesOfTriples

    // List validations that point at a range or a name, now that every sheet exists.
    let lookupName (index: int) (name: string) =
        let key = name.ToUpperInvariant()
        Map.tryFind (Some (sheetIdOfIndex index), key) definedNames
        |> Option.orElse (Map.tryFind (None, key) definedNames)
    let models =
        imports
        |> Array.mapi (fun index import ->
            let lookupSheet (name: string option) =
                match name with
                | None -> Some import.Model
                | Some name ->
                    names
                    |> Array.tryFindIndex (fun n -> System.String.Equals(n, name, System.StringComparison.OrdinalIgnoreCase))
                    |> Option.map (fun i -> imports.[i].Model)
            applyPendingLists lookupSheet (lookupName index) import.PendingLists import.Model)

    let visible = imports |> Array.map (fun import -> import.State = "visible")
    let tableColumns =
        imports |> Array.fold (fun acc import -> Map.fold (fun m k v -> Map.add k v m) acc import.TableColumns) Map.empty
    let firstVisible = visible |> Array.tryFindIndex id |> Option.defaultValue 0
    let context =
        buildWorkbookContext models names visible tableColumns definedNames firstVisible
        |> computeUncachedFormulas
    {
        Sheets =
            names
            |> Array.mapi (fun index name ->
                { Name = name
                  Model =
                    (match Map.tryFind (sheetIdOfIndex index) context.Sheets with
                     | Some sheet -> sheet.Model
                     | None -> models.[index])
                  State = imports.[index].State
                  Visible = visible.[index] })
        Context = context
        Date1904 = (not (isNullOrUndefined info)) && boolProp info "date1904" false
    }

    /// <summary>
    /// Import a sheet from office-format workbook into SpreadsheetModel.
    /// Formulas keep the values cached in the file. A formula the file has no
    /// value for is computed with this sheet as the only context — use
    /// `importWorkbook` when other sheets or defined names matter.
    /// Throws when the workbook is unreadable or the sheet does not exist.
    /// </summary>
    /// <param name="workbook">office-format workbook object</param>
    /// <param name="sheetName">Name of sheet to import (or first sheet if not specified)</param>
    /// <returns>SpreadsheetModel populated with sheet data</returns>
let importSheet (workbook: ExcelWorkbook) (sheetName: string option): SpreadsheetModel =
    let workbook = requireWorkbook workbook
    let excel = excelApi ()
    let actualSheetName =
        match sheetName with
        | Some name -> name
        | None ->
            let sheetNames = excel?sheetNames(workbook) |> arrayFrom |> Array.map string
            if sheetNames.Length > 0 then sheetNames.[0]
            else failwith "XLSX workbook has no worksheets"
    let sheet = excel?sheet(actualSheetName, workbook)
    if isNullOrUndefined sheet then
        failwith (sprintf "XLSX workbook has no worksheet named '%s'" actualSheetName)
    let import = importSheetCore excel workbook sheet
    let model =
        applyPendingLists (fun name -> if name.IsNone then Some import.Model else None) (fun _ -> None) import.PendingLists import.Model
    let context =
        buildWorkbookContext [| model |] [| actualSheetName |] [| true |] import.TableColumns Map.empty 0
        |> computeUncachedFormulas
    match Map.tryFind (sheetIdOfIndex 0) context.Sheets with
    | Some sheet -> sheet.Model
    | None -> model

// ---------------------------------------------------------------------------
// Buffer entry points. A failure (not a zip, corrupt XML, no workbook, an
// archive over the size limits) rejects / throws: it must never come back as
// an empty sheet the user could save over the original.
//
// The `…Promise` functions are for JavaScript callers, which cannot await an
// F# Async; the Async ones remain for F# callers.
// ---------------------------------------------------------------------------

let private loadWorkbook (buffer: JS.ArrayBuffer) (selected: string option option): JS.Promise<ExcelWorkbook> =
    let excel = excelApi ()
    let loading: JS.Promise<ExcelWorkbook> =
        match selected with
        | Some sheetName when isFunction excel?loadSelected -> excel?loadSelected(buffer, sheetName)
        | _ -> excel?load(buffer)
    loading.``then``(requireWorkbook)

/// Import one sheet. Only that sheet is parsed, so formulas keep their
/// cached values but nothing outside the sheet is available to recompute.
let importFromBufferPromise (buffer: JS.ArrayBuffer) (sheetName: string option): JS.Promise<SpreadsheetModel> =
    (loadWorkbook buffer (Some sheetName)).``then``(fun workbook -> importSheet workbook sheetName)

/// Import the whole workbook: sheets with their visibility, and the context
/// (tables, defined names) cross-sheet formulas need.
let importWorkbookPromise (buffer: JS.ArrayBuffer): JS.Promise<ImportedWorkbook> =
    (loadWorkbook buffer None).``then``(importWorkbook)

/// Import every sheet in source order, paired with its name (hidden sheets
/// included — see `importWorkbookPromise` for visibility).
let importAllSheetsPromise (buffer: JS.ArrayBuffer): JS.Promise<(string * SpreadsheetModel) array> =
    (importWorkbookPromise buffer).``then``(fun imported ->
        imported.Sheets |> Array.map (fun sheet -> sheet.Name, sheet.Model))

/// Name and visibility of every sheet: `{ name, state, visible }` where
/// state is "visible" | "hidden" | "veryHidden".
let getSheetInfosPromise (buffer: JS.ArrayBuffer): JS.Promise<obj[]> =
    (loadWorkbook buffer None).``then``(fun workbook ->
        let excel = excelApi ()
        let names: string[] = excel?sheetNames(workbook) |> arrayFrom |> Array.map string
        let states: string[] =
            if isFunction excel?sheetStates then excel?sheetStates(workbook) |> arrayFrom |> Array.map string
            else names |> Array.map (fun _ -> "visible")
        names
        |> Array.mapi (fun index name ->
            let state = if index < states.Length then states.[index] else "visible"
            createObj [ "name" ==> name; "state" ==> state; "visible" ==> (state = "visible") ]))

/// Names of every sheet in the workbook, hidden ones included.
let getSheetNamesPromise (buffer: JS.ArrayBuffer): JS.Promise<string[]> =
    (getSheetInfosPromise buffer).``then``(fun infos -> infos |> Array.map (fun info -> string (getProp info "name")))

/// Names of the sheets Excel shows a tab for.
let getVisibleSheetNamesPromise (buffer: JS.ArrayBuffer): JS.Promise<string[]> =
    (getSheetInfosPromise buffer).``then``(fun infos ->
        infos
        |> Array.filter (fun info -> unbox<bool> (getProp info "visible"))
        |> Array.map (fun info -> string (getProp info "name")))

    /// <summary>
    /// Import XLSX file from ArrayBuffer.
    /// </summary>
    /// <param name="buffer">ArrayBuffer containing XLSX file data</param>
    /// <param name="sheetName">Optional sheet name (or first sheet if not specified)</param>
    /// <returns>SpreadsheetModel populated with sheet data</returns>
let importFromBuffer (buffer: JS.ArrayBuffer) (sheetName: string option): Async<SpreadsheetModel> =
    importFromBufferPromise buffer sheetName |> Async.AwaitPromise

    /// <summary>
    /// Import every sheet in the workbook in source order, paired with
    /// its sheet name. Used by the multi-sheet load path so the chrome
    /// can render a tab strip and the user can swap between sheets
    /// without re-parsing the XLSX each time.
    /// </summary>
let importAllSheets (buffer: JS.ArrayBuffer): Async<(string * SpreadsheetModel) array> =
    importAllSheetsPromise buffer |> Async.AwaitPromise

    /// <summary>
    /// Get list of sheet names from workbook buffer.
    /// Useful for showing sheet selector before import.
    /// </summary>
    /// <param name="buffer">ArrayBuffer containing XLSX file data</param>
    /// <returns>Array of sheet names</returns>
let getSheetNames (buffer: JS.ArrayBuffer): Async<string[]> =
    getSheetNamesPromise buffer |> Async.AwaitPromise
