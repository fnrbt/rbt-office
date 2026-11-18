/// Excel's object model for macros (VBA run by Rbt.Vb), over the whole workbook: every sheet, typed
/// values, formats, row and column edits, and the formula engine for WorksheetFunction and Evaluate.
/// A run works on a copy of the workbook (`MacroWorkbook`); its host commits the result as one step of
/// the session's history (one undo step, the workbook marked as changed). Shared by the browser and the
/// native hosts.
module Rbt.Spreadsheet.MacroHost

open System
open Rbt.Vb.Values
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Workbook

// ---- errors ------------------------------------------------------------------------------------------

/// A run-time error of the object model, as Excel raises it: `On Error` sees its number.
let private fail number (description: string) : 'a = raiseVbErrorWith number description
let private failApp description = fail 1004 description
let private unsupported (typeName: string) (memberName: string) =
    fail 438 (sprintf "Object doesn't support this property or method (%s.%s)" typeName memberName)

// ---- Excel's constants -------------------------------------------------------------------------------

/// The xl* constants macros use, as host globals.
let constants : (string * Value) list =
    [ "xlUp", -4162; "xlDown", -4121; "xlToLeft", -4159; "xlToRight", -4161
      "xlLeft", -4131; "xlRight", -4152; "xlCenter", -4108; "xlTop", -4160; "xlBottom", -4107
      "xlGeneral", 1; "xlJustify", -4130; "xlDistributed", -4117; "xlCenterAcrossSelection", 7; "xlFill", 5
      "xlNone", -4142; "xlAutomatic", -4105; "xlSolid", 1
      "xlCalculationAutomatic", -4105; "xlCalculationManual", -4135; "xlCalculationSemiautomatic", 2
      "xlUnderlineStyleNone", -4142; "xlUnderlineStyleSingle", 2; "xlUnderlineStyleDouble", -4119
      "xlShiftUp", -4162; "xlShiftToLeft", -4159; "xlShiftDown", -4121; "xlShiftToRight", -4161
      "xlPasteAll", -4104; "xlPasteValues", -4163; "xlPasteFormats", -4122; "xlPasteFormulas", -4123
      "xlWhole", 1; "xlPart", 2; "xlValues", -4163; "xlFormulas", -4123; "xlComments", -4144
      "xlByRows", 1; "xlByColumns", 2; "xlNext", 1; "xlPrevious", 2
      "xlAscending", 1; "xlDescending", 2; "xlYes", 1; "xlNo", 2; "xlGuess", 0
      "xlSheetVisible", -1; "xlSheetHidden", 0; "xlSheetVeryHidden", 2
      "xlWorksheet", -4167; "xlCellTypeLastCell", 11; "xlCellTypeBlanks", 4; "xlCellTypeConstants", 2; "xlCellTypeFormulas", -4123
      "xlContinuous", 1; "xlThin", 2; "xlMedium", -4138; "xlThick", 4; "xlDash", -4115; "xlDot", -4118; "xlDouble", -4119; "xlLineStyleNone", -4142
      "xlEdgeLeft", 7; "xlEdgeTop", 8; "xlEdgeBottom", 9; "xlEdgeRight", 10; "xlInsideVertical", 11; "xlInsideHorizontal", 12
      "xlA1", 1; "xlR1C1", -4150
      "xlXYScatter", -4169; "xlXYScatterLines", 74; "xlXYScatterLinesNoMarkers", 75; "xlXYScatterSmooth", 72; "xlXYScatterSmoothNoMarkers", 73
      "xlLine", 4; "xlLineMarkers", 65; "xlColumnClustered", 51; "xlColumnStacked", 52; "xlBarClustered", 57; "xlPie", 5; "xlArea", 1
      "xlColumns", 2; "xlRows", 1
      "xlCellValue", 1; "xlExpression", 2; "xlBetween", 1; "xlNotBetween", 2; "xlEqual", 3; "xlNotEqual", 4
      "xlGreater", 5; "xlLess", 6; "xlGreaterEqual", 7; "xlLessEqual", 8
      "xlConditionValueNumber", 0; "xlConditionValueLowestValue", 1; "xlConditionValueHighestValue", 2
      "xlConditionValuePercent", 3; "xlConditionValueFormula", 4; "xlConditionValuePercentile", 5
      "xlHorizontal", -4128; "xlVertical", -4166; "xlUpward", -4171; "xlDownward", -4170
      "xlMaximized", -4137; "xlMinimized", -4140; "xlNormal", -4143
      "xlHairline", 1; "xlDashDot", 4; "xlDashDotDot", 5; "xlSlantDashDot", 13
      "xlDiagonalDown", 5; "xlDiagonalUp", 6; "xlColorIndexAutomatic", -4105; "xlColorIndexNone", -4142
      "xlThemeColorDark1", 1; "xlThemeColorLight1", 2; "xlThemeColorDark2", 3; "xlThemeColorLight2", 4
      "xlThemeColorAccent1", 5; "xlThemeColorAccent2", 6; "xlThemeColorAccent3", 7; "xlThemeColorAccent4", 8
      "xlThemeColorAccent5", 9; "xlThemeColorAccent6", 10; "xlThemeColorHyperlink", 11; "xlThemeColorFollowedHyperlink", 12
      "xlDisabled", 0; "xlInterrupt", 1; "xlErrorHandler", 2
      "xlErrDiv0", 2007; "xlErrNA", 2042; "xlErrName", 2029; "xlErrNull", 2000; "xlErrNum", 2036; "xlErrRef", 2023; "xlErrValue", 2015 ]
    @ MacroShapes.constants
    |> List.map (fun (name, value) -> name, VLong value)

// ---- values --------------------------------------------------------------------------------------------

let private eq (left: string) (right: string) = String.Equals(left, right, StringComparison.OrdinalIgnoreCase)

let rec private scalarOf (value: Value) =
    match value with
    | VRef (values, index) when index >= 0 && index < values.Length -> scalarOf values.[index]
    | VFieldRef (instance, field) ->
        match instance.Fields.TryGetValue field with
        | true, inner -> scalarOf inner
        | _ -> VEmpty
    | VHostObject host ->
        // An object used as a value: its default property (a Range's Value).
        match host.GetMember "_Default" with
        | Some inner -> scalarOf inner
        | None -> value
    | other -> other

let private toInt (value: Value) =
    match scalarOf value with
    | VInteger i | VLong i -> i
    | VByte b -> int b
    | VSingle f -> int (Math.Round (float f))
    | VDouble d -> int (Math.Round d)
    | VDate serial -> int (Math.Round serial)
    | VCurrency units -> int (Math.Round (float units / 10000.0))
    | VBoolean b -> if b then -1 else 0
    | VString s ->
        match Double.TryParse(s, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture) with
        | true, d -> int (Math.Round d)
        | _ -> fail 13 "Type mismatch"
    | VEmpty | VNull -> 0
    | _ -> fail 13 "Type mismatch"

let private toFloat (value: Value) =
    match scalarOf value with
    | VInteger i | VLong i -> float i
    | VByte b -> float b
    | VSingle f -> float (Rbt.Vb.NumberText.singleAsDouble f)
    | VDouble d -> d
    | VDate serial -> serial
    | VCurrency units -> float units / 10000.0
    | VBoolean b -> if b then -1.0 else 0.0
    | VString s ->
        match Double.TryParse(s, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture) with
        | true, d -> d
        | _ -> fail 13 "Type mismatch"
    | VEmpty | VNull -> 0.0
    | _ -> fail 13 "Type mismatch"

let private toText (value: Value) =
    match scalarOf value with
    | VString s -> s
    | VInteger i | VLong i -> string i
    | VByte b -> string b
    | VSingle _ -> Rbt.Vb.Api.displayValue value
    | VDouble d -> d.ToString("R", Globalization.CultureInfo.InvariantCulture)
    | VDate _ | VCurrency _ -> Rbt.Vb.Api.displayValue value
    | VBoolean b -> if b then "True" else "False"
    | VEmpty | VNull | VNothing | VUndefined -> ""
    | _ -> fail 13 "Type mismatch"

let private toBool (value: Value) =
    match scalarOf value with
    | VBoolean b -> b
    | VInteger i | VLong i -> i <> 0
    | VByte b -> b <> 0uy
    | VSingle f -> f <> 0.0f
    | VDouble d -> d <> 0.0
    | VString s when eq s "true" -> true
    | VString s when eq s "false" -> false
    | VEmpty -> false
    | _ -> fail 13 "Type mismatch"

let private isMissing (value: Value) = match value with VEmpty | VUndefined -> true | _ -> false

/// Excel's error values as VBA's Error codes (CVErr(xlErrNA) is 2042), both ways.
let private errorCodes =
    [ "NULL!", 2000; "DIV/0!", 2007; "VALUE!", 2015; "REF!", 2023; "NAME?", 2029; "NUM!", 2036; "N/A", 2042
      "GETTING_DATA", 2043; "SPILL!", 2045; "CONNECT!", 2046; "BLOCKED!", 2047; "UNKNOWN!", 2048; "FIELD!", 2049; "CALC!", 2050 ]
let private errorCodeOf (text: string) = errorCodes |> List.tryPick (fun (name, code) -> if name = text then Some code else None) |> Option.defaultValue 2015
let private errorTextOf (code: int) = errorCodes |> List.tryPick (fun (name, c) -> if c = code then Some name else None) |> Option.defaultValue "VALUE!"

/// What a cell holds, as VBA's Value2 gives it: numbers as Double (dates too), text, booleans, Error values.
let rec cellValue (value: CellValue) : Value =
    match value with
    | CellValue.Empty -> VEmpty
    | Number n -> VDouble n
    | String s -> VString s
    | Boolean b -> VBoolean b
    | Error e -> VError (errorCodeOf e)
    | Formula (_, cached) -> cellValue cached

/// What a cell holds as VBA's Value gives it: a number formatted as a date or time is a Date.
let cellValueAs (date1904: bool) (format: CellFormat) (value: CellValue) : Value =
    match cellValue value, format.NumberFormat with
    | VDouble n, Some code when code = "date" || CellText.isDateTimeNumberFormat code -> VDate (if date1904 then n + 1462.0 else n)
    | other, _ -> other

/// A 2-D array of a range's values, 1-based in both dimensions, as Range.Value gives it.
let private array2D (rows: int) (cols: int) (item: int -> int -> Value) =
    VArray { Data = Array.init (rows * cols) (fun i -> item (i / cols) (i % cols))
             LowerBounds = [| 1; 1 |]
             Lengths = [| rows; cols |] }

/// The values of an array assigned to a range, by (row, column) offset; a 1-D array fills a row.
let private arrayItem (array: VBArray) (row: int) (col: int) : Value option =
    match array.Lengths with
    | [| n |] -> if row = 0 && col < n then Some array.Data.[col] else None
    | [| rows; cols |] -> if row < rows && col < cols then Some array.Data.[row * cols + col] else None
    | _ -> None

// ---- colours: VBA's are BGR longs (RGB(r, g, b) = r + g * 256 + b * 65536) --------------------------------

let private colorToLong (color: string option) =
    match color with
    | Some text when text.StartsWith "#" && text.Length >= 7 ->
        let r = Convert.ToInt32(text.Substring(1, 2), 16)
        let g = Convert.ToInt32(text.Substring(3, 2), 16)
        let b = Convert.ToInt32(text.Substring(5, 2), 16)
        Some (r + g * 256 + b * 65536)
    | _ -> None

let private longToColor (value: int) =
    sprintf "#%02X%02X%02X" (value &&& 0xFF) ((value >>> 8) &&& 0xFF) ((value >>> 16) &&& 0xFF)

/// Excel's 56-colour palette for ColorIndex (index 1-based).
let private palette =
    [| "#000000"; "#FFFFFF"; "#FF0000"; "#00FF00"; "#0000FF"; "#FFFF00"; "#FF00FF"; "#00FFFF"
       "#800000"; "#008000"; "#000080"; "#808000"; "#800080"; "#008080"; "#C0C0C0"; "#808080"
       "#9999FF"; "#993366"; "#FFFFCC"; "#CCFFFF"; "#660066"; "#FF8080"; "#0066CC"; "#CCCCFF"
       "#000080"; "#FF00FF"; "#FFFF00"; "#00FFFF"; "#800080"; "#800000"; "#008080"; "#0000FF"
       "#00CCFF"; "#CCFFFF"; "#CCFFCC"; "#FFFF99"; "#99CCFF"; "#FF99CC"; "#CC99FF"; "#FFCC99"
       "#3366FF"; "#33CCCC"; "#99CC00"; "#FFCC00"; "#FF9900"; "#FF6600"; "#666699"; "#969696"
       "#003366"; "#339966"; "#003300"; "#333300"; "#993300"; "#993366"; "#333399"; "#333333" |]

/// The colours of Office's default theme by Excel's XlThemeColor (1-based): xlThemeColorDark1 is
/// "Background 1" (white) and xlThemeColorLight1 "Text 1" (black), as the macro recorder names them.
let private themeColors =
    [| "#FFFFFF"; "#000000"; "#E7E6E6"; "#44546A"; "#4472C4"; "#ED7D31"; "#A5A5A5"; "#FFC000"; "#5B9BD5"; "#70AD47"; "#0563C1"; "#954F72" |]

let private themeColor (index: int) =
    if index >= 1 && index <= themeColors.Length then themeColors.[index - 1] else fail 5 "Invalid procedure call or argument"

/// A colour lightened (TintAndShade > 0) or darkened (< 0), as Excel applies TintAndShade.
let private tinted (amount: float) (color: string) =
    let channel (at: int) =
        let c = float (Convert.ToInt32(color.Substring(at, 2), 16))
        let v = if amount >= 0.0 then c + (255.0 - c) * amount else c * (1.0 + amount)
        int (Math.Round (max 0.0 (min 255.0 v)))
    if color.Length < 7 || amount = 0.0 then color else sprintf "#%02X%02X%02X" (channel 1) (channel 3) (channel 5)

// ---- borders: Excel's LineStyle and Weight as the styles of the file format ----

let private xlContinuous = 1
let private xlLineStyleNone = -4142

/// (LineStyle, Weight) of a border style ("thin", "mediumDashed" ...).
let private lineOfStyle (style: string) =
    match style.ToLowerInvariant() with
    | "hair" -> 1, 1
    | "medium" -> 1, -4138
    | "thick" -> 1, 4
    | "dashed" -> -4115, 2
    | "mediumdashed" -> -4115, -4138
    | "dotted" -> -4118, 2
    | "dashdot" -> 4, 2
    | "mediumdashdot" -> 4, -4138
    | "dashdotdot" -> 5, 2
    | "mediumdashdotdot" -> 5, -4138
    | "slantdashdot" -> 13, -4138
    | "double" -> -4119, 4
    | _ -> 1, 2

/// The border style of a LineStyle and a Weight (None for xlLineStyleNone).
let private styleOfLine (lineStyle: int) (weight: int) : string option =
    let heavy = weight = -4138 || weight = 4
    match lineStyle with
    | -4142 -> None
    | -4115 -> Some (if heavy then "mediumDashed" else "dashed")
    | -4118 -> Some (if weight = 1 then "hair" else "dotted")
    | 4 -> Some (if heavy then "mediumDashDot" else "dashDot")
    | 5 -> Some (if heavy then "mediumDashDotDot" else "dashDotDot")
    | 13 -> Some "slantDashDot"
    | -4119 -> Some "double"
    | _ -> Some (match weight with 1 -> "hair" | -4138 -> "medium" | 4 -> "thick" | _ -> "thin")

/// A side of a cell: 0 left, 1 top, 2 bottom, 3 right, 4 diagonal down, 5 diagonal up.
let private sideOf (f: CellFormat) (side: int) =
    match side with
    | 0 -> f.BorderLeft
    | 1 -> f.BorderTop
    | 2 -> f.BorderBottom
    | 3 -> f.BorderRight
    | 4 -> if f.DiagonalDown then f.BorderDiagonal else None
    | _ -> if f.DiagonalUp then f.BorderDiagonal else None

let private withSide (side: int) (border: BorderStyle option) (f: CellFormat) =
    match side with
    | 0 -> { f with BorderLeft = border }
    | 1 -> { f with BorderTop = border }
    | 2 -> { f with BorderBottom = border }
    | 3 -> { f with BorderRight = border }
    | 4 -> { f with DiagonalDown = border.IsSome; BorderDiagonal = (if border.IsSome then border elif f.DiagonalUp then f.BorderDiagonal else None) }
    | _ -> { f with DiagonalUp = border.IsSome; BorderDiagonal = (if border.IsSome then border elif f.DiagonalDown then f.BorderDiagonal else None) }

// ---- addresses -------------------------------------------------------------------------------------------

let private columnName (col: int) =
    let mutable n = col + 1
    let mutable name = ""
    while n > 0 do
        let r = (n - 1) % 26
        name <- string (char (int 'A' + r)) + name
        n <- (n - 1) / 26
    name

let private cellAddress (row: int) (col: int) (rowAbsolute: bool) (colAbsolute: bool) =
    (if colAbsolute then "$" else "") + columnName col + (if rowAbsolute then "$" else "") + string (row + 1)

let maxRow = 1048575
let maxCol = 16383

/// A block's address as Excel's Range.Address gives it: whole rows as $2:$3, whole columns as $A:$B,
/// a single cell alone; R1C1 (absolute, as Address gives it without RelativeTo) with `r1c1`.
let private blockAddress (top: int) (left: int) (bottom: int) (right: int) (rowAbs: bool) (colAbs: bool) (r1c1: bool) =
    let row n = (if rowAbs then "$" else "") + string (n + 1)
    let col n = (if colAbs then "$" else "") + columnName n
    let wholeRows = left = 0 && right = maxCol
    let wholeCols = top = 0 && bottom = maxRow
    if r1c1 then
        let cell r c = sprintf "R%dC%d" (r + 1) (c + 1)
        if wholeRows then (if top = bottom then sprintf "R%d" (top + 1) else sprintf "R%d:R%d" (top + 1) (bottom + 1))
        elif wholeCols then (if left = right then sprintf "C%d" (left + 1) else sprintf "C%d:C%d" (left + 1) (right + 1))
        elif top = bottom && left = right then cell top left
        else cell top left + ":" + cell bottom right
    elif wholeRows then row top + ":" + row bottom
    elif wholeCols then col left + ":" + col right
    elif top = bottom && left = right then cellAddress top left rowAbs colAbs
    else cellAddress top left rowAbs colAbs + ":" + cellAddress bottom right rowAbs colAbs

/// A reference of a macro ("A1", "$B$2:C9", "A:A", "3:5", "Sheet2!A1", a defined name) as
/// (sheet name option, top, left, bottom, right), 0-based.
let private parseReference (text: string) : (string option * int * int * int * int) option =
    let text = text.Trim()
    let sheetName, body =
        match text.LastIndexOf '!' with
        | -1 -> None, text
        | at -> Some (text.Substring(0, at).Trim('\'').Replace("''", "'")), text.Substring(at + 1)
    let parsePart (part: string) =
        let part = part.Replace("$", "").Trim()
        let letters = System.String(part |> Seq.takeWhile Char.IsLetter |> Seq.toArray)
        let digits = part.Substring letters.Length
        let col =
            if letters = "" then None
            else Some ((letters.ToUpperInvariant() |> Seq.fold (fun acc ch -> acc * 26 + (int ch - int 'A' + 1)) 0) - 1)
        let row =
            match Int32.TryParse digits with
            | true, n when n >= 1 -> Some (n - 1)
            | _ -> None
        if digits <> "" && row.IsNone then None
        elif letters = "" && digits = "" then None
        else Some (row, col)
    let parts = body.Split ':'
    match parts |> Array.map parsePart with
    | [| Some (Some r, Some c) |] -> Some (sheetName, r, c, r, c)
    | [| Some (Some r0, Some c0); Some (Some r1, Some c1) |] -> Some (sheetName, min r0 r1, min c0 c1, max r0 r1, max c0 c1)
    | [| Some (None, Some c0); Some (None, Some c1) |] -> Some (sheetName, 0, min c0 c1, maxRow, max c0 c1)
    | [| Some (Some r0, None); Some (Some r1, None) |] -> Some (sheetName, min r0 r1, 0, max r0 r1, maxCol)
    | _ -> None

// ---- the workbook a run works on ---------------------------------------------------------------------------

/// What a host gives a run (MacroApi.MacroHooks): MsgBox / InputBox / Debug.Print, the file's name.
type MacroHooks = Rbt.Spreadsheet.MacroApi.MacroHooks
let defaultHooks = Rbt.Spreadsheet.MacroApi.defaultHooks
let private noMacroHooks = defaultHooks

/// The workbook a run reads and writes. A read after a write sees the formulas calculated again.
/// The cell a position on a sheet (points from its top left) falls in, and the offset into it, in pixels.
let cellAtPoints (model: SpreadsheetModel) (left: float) (top: float) =
    let walk (target: float) (size: int -> float) =
        let mutable index = 0
        let mutable edge = 0.0
        while edge + size index <= target && index < 16383 do
            edge <- edge + size index
            index <- index + 1
        index, target - edge
    let col, dx = walk (max 0.0 left * 4.0 / 3.0) (fun c -> model.ColumnWidths |> Map.tryFind c |> Option.defaultValue model.DefaultColumnWidth)
    let row, dy = walk (max 0.0 top * 4.0 / 3.0) (fun r -> model.RowHeights |> Map.tryFind r |> Option.defaultValue model.DefaultRowHeight)
    { Row = row; Col = col }, dx, dy

type MacroWorkbook(initial: Workbook, initialHooks: MacroHooks) =
    let mutable workbook = initial
    let mutable hooks = initialHooks
    let mutable engine : Recalc.Engine option = None
    let mutable dirty = false
    // The engine has not been told of writes since (a workbook without formulas): built again before it calculates.
    let mutable stale = false
    let hints = ResizeArray<SheetId * CellAddress>()
    let mutable selection : (SheetId * int * int * int * int) option = None
    // (Counts the changes, for a host that shows the workbook while the macro runs.)
    let mutable version = 0
    // (Keyed by lower-cased name: Fable has no case-insensitive comparer.)
    let settingValues = Collections.Generic.Dictionary<string, Value>()
    // The charts the run made, by their relationship id: their sheet and what they show.
    let charts = Collections.Generic.Dictionary<string, SheetId * Rbt.Spreadsheet.ChartAuthoring.MacroChart>()
    // The shapes the run drew (MacroShapes), by sheet and drawing id; those changed since their XML was written.
    let shapes = Collections.Generic.Dictionary<SheetId * int, MacroShapes.ShapeNode>()
    let staleShapes = Collections.Generic.HashSet<SheetId * int>()
    let mutable shapeIds = 100000

    let resolve (sheetId: SheetId) =
        match Map.tryFind sheetId workbook.Sheets with
        | Some sheet when sheet.Model.Deferred.IsSome ->
            let loaded = resolveSheet sheet.Model
            workbook <- { workbook with Sheets = Map.add sheetId { sheet with Model = loaded } workbook.Sheets }
            engine <- None
        | _ -> ()

    let calculate () =
        // A workbook without a formula has nothing to calculate after a write of values (most macros that
        // paint or fill cells): the engine is not told of them, and is built again when a formula is written.
        let formulaWritten () =
            hints |> Seq.exists (fun (sheetId, address) ->
                match Map.tryFind sheetId workbook.Sheets with
                | Some sheet -> (match getCell address sheet.Model with Some { Value = Formula _ } -> true | _ -> false)
                | None -> false)
        if dirty && (match engine with Some current -> Recalc.formulaCount current = 0 | None -> false) && not (formulaWritten ()) then
            hints.Clear()
            dirty <- false
            stale <- true
        if dirty then
            if stale then
                engine <- None
                stale <- false
            let next =
                match engine with
                | Some current ->
                    match Recalc.recalculate current workbook (Some (List.ofSeq hints)) with
                    | Some updated -> Some updated
                    | None -> None
                | None -> None
            match next with
            | Some updated -> workbook <- updated
            | None ->
                let updated, built = Recalc.recalculateAll workbook
                workbook <- updated
                engine <- Some built
            hints.Clear()
            dirty <- false

    member _.Hooks = hooks
    /// Start over on another workbook (a program's globals stay bound to this object across runs).
    member _.Reset(next: Workbook, nextHooks: MacroHooks) =
        workbook <- next
        hooks <- nextHooks
        engine <- None
        dirty <- false
        stale <- false
        hints.Clear()
        selection <- None
        settingValues.Clear()
        charts.Clear()
        shapes.Clear()
        staleShapes.Clear()
    member _.Setting(name: string) =
        match settingValues.TryGetValue(name.ToLowerInvariant()) with
        | true, value -> Some value
        | _ -> None
    member _.SetSetting(name: string, value: Value) = settingValues.[name.ToLowerInvariant()] <- value
    member _.Charts = charts
    member _.Shapes = shapes
    /// An id for a shape inside a group (the drawing's own ids number the sheet's objects).
    member _.NextShapeId() = shapeIds <- shapeIds + 1; shapeIds
    /// A drawn shape changed: its XML is written again before the workbook is next read.
    member _.ShapeChanged(sheetId: SheetId, id: int) =
        staleShapes.Add((sheetId, id)) |> ignore
        version <- version + 1
    member this.FlushShapes() =
        if staleShapes.Count > 0 then
            let pending = List.ofSeq staleShapes
            staleShapes.Clear()
            for (sheetId, id) in pending do
                match shapes.TryGetValue((sheetId, id)) with
                | true, node ->
                    let model = this.Model sheetId
                    let l, t, w, h = MacroShapes.frame node
                    let cell, dx, dy = cellAtPoints model l t
                    let xml = MacroShapes.anchorXml node cell.Row cell.Col dx dy
                    let anchor = SheetOps.pictureAnchor cell dx dy (max 1.0 (w * 4.0 / 3.0)) (max 1.0 (h * 4.0 / 3.0))
                    let drawn = { Id = id; Anchor = anchor; Name = Some node.Name; Description = None; Content = KeptObject ("shape", xml) }
                    let next =
                        if model.Drawings |> List.exists (fun d -> d.Id = id) then model.Drawings |> List.map (fun d -> if d.Id = id then drawn else d)
                        else model.Drawings @ [ drawn ]
                    workbook <- updateSheetModel sheetId { model with Drawings = next } workbook
                | _ -> ()
    member this.Workbook = calculate (); this.FlushShapes (); workbook
    member _.SheetOrder = workbook.SheetOrder
    member _.ActiveSheet = workbook.ActiveSheet
    member _.SetActiveSheet(sheetId: SheetId) =
        if workbook.ActiveSheet <> sheetId then version <- version + 1
        workbook <- setActiveSheet sheetId workbook
    member _.Version = version
    member _.Sheet(sheetId: SheetId) =
        resolve sheetId
        match Map.tryFind sheetId workbook.Sheets with
        | Some sheet -> sheet
        | None -> fail 9 "Subscript out of range"
    member this.Model(sheetId: SheetId) = (this.Sheet sheetId).Model
    member this.Cell(sheetId: SheetId, row: int, col: int) =
        resolve sheetId
        calculate ()
        getCell { Row = row; Col = col } (this.Model sheetId)
    member this.SetCell(sheetId: SheetId, row: int, col: int, cell: CellData option) =
        let model = this.Model sheetId
        let address = { Row = row; Col = col }
        let next =
            match cell with
            | Some data -> setCell address data model
            | None -> clearCell address model
        let next = if cell.IsSome && (row >= next.RowCount || col >= next.ColCount) then { next with RowCount = max next.RowCount (row + 1); ColCount = max next.ColCount (col + 1) } else next
        workbook <- updateSheetModel sheetId next workbook
        dirty <- true
        version <- version + 1
        hints.Add((sheetId, address))
    /// The sheet grown to show a cell (the grid shows RowCount x ColCount; a macro may write past them).
    member this.Grow(sheetId: SheetId, row: int, col: int) =
        let model = this.Model sheetId
        if row >= model.RowCount || col >= model.ColCount then
            workbook <- updateSheetModel sheetId { model with RowCount = max model.RowCount (row + 1); ColCount = max model.ColCount (col + 1) } workbook
    /// A change of formats, sizes or merges: no formula reads them.
    member this.ChangeModel(sheetId: SheetId, change: SpreadsheetModel -> SpreadsheetModel) =
        workbook <- updateSheetModel sheetId (change (this.Model sheetId)) workbook
        version <- version + 1
    member this.Structure(sheetId: SheetId, edit: StructureEdit) =
        resolve sheetId
        let next, _ = applyStructureEdit sheetId edit workbook
        workbook <- next
        version <- version + 1
        engine <- None
        dirty <- true
    member _.ChangeWorkbook(change: Workbook -> Workbook) =
        version <- version + 1
        workbook <- change workbook
        engine <- None
        dirty <- true
    member _.Selection
        with get () = selection
        and set (value) = selection <- value
    member _.Date1904 = date1904 workbook

// ---- the object model ----------------------------------------------------------------------------------------
//
// Each type of object (Range, Font, Worksheet, Shape ...) is a table of its members (Rbt.Vb.HostMembers),
// built once at the end of this section; an object is its type's table with the state of one instance (a
// Range: the workbook and a block of cells). A member is found by name in any case by one lookup; a name
// the table does not have is error 438 naming the type and the member, and the parameters of a method
// give it named arguments.

open Rbt.Vb.HostMembers

/// A block of cells of a sheet, 0-based.
type private RangeRef = { Sheet: SheetId; Top: int; Left: int; Bottom: int; Right: int }

let private rangeAddress (r: RangeRef) rowAbs colAbs r1c1 = blockAddress r.Top r.Left r.Bottom r.Right rowAbs colAbs r1c1

/// The parts of a formula outside its string literals, rewritten by `change` (literals kept as they are).
let private outsideStrings (change: string -> string) (formula: string) =
    let parts = formula.Split('"')
    String.Join("\"", parts |> Array.mapi (fun i part -> if i % 2 = 0 then change part else part))

/// A formula's A1 references as R1C1 ones, relative to the cell it is in (Range.FormulaR1C1).
let private toR1C1 (row: int) (col: int) (formula: string) =
    let pattern = Text.RegularExpressions.Regex(@"(?<![A-Za-z0-9_.])(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})(?![0-9A-Za-z_(!])")
    formula |> outsideStrings (fun text ->
        pattern.Replace(text, Text.RegularExpressions.MatchEvaluator(fun m ->
            match parseReference (m.Groups.[2].Value + m.Groups.[4].Value) with
            | Some (_, r, c, _, _) when r <= maxRow && c <= maxCol ->
                let part (absolute: bool) (letter: string) (target: int) (here: int) =
                    if absolute then letter + string (target + 1)
                    elif target = here then letter
                    else sprintf "%s[%d]" letter (target - here)
                part (m.Groups.[3].Value = "$") "R" r row + part (m.Groups.[1].Value = "$") "C" c col
            | _ -> m.Value)))

/// A formula's R1C1 references as A1 ones, for the cell it goes into (setting Range.FormulaR1C1).
let private fromR1C1 (row: int) (col: int) (formula: string) =
    let pattern = Text.RegularExpressions.Regex(@"(?<![A-Za-z0-9_.])R(\[-?[0-9]+\]|[0-9]+)?C(\[-?[0-9]+\]|[0-9]+)?(?![0-9A-Za-z_(])", Text.RegularExpressions.RegexOptions.IgnoreCase)
    formula |> outsideStrings (fun text ->
        pattern.Replace(text, Text.RegularExpressions.MatchEvaluator(fun m ->
            let resolve (part: string) (here: int) =
                if part = "" then false, here
                elif part.StartsWith "[" then false, here + int (part.Trim('[', ']'))
                else true, int part - 1
            let rowAbs, r = resolve m.Groups.[1].Value row
            let colAbs, c = resolve m.Groups.[2].Value col
            cellAddress r c rowAbs colAbs)))

let private sheetIdNamed (book: MacroWorkbook) (name: string) =
    book.SheetOrder |> List.tryFind (fun id -> eq (book.Sheet id).Name name)

let private quotedSheet (name: string) =
    if name |> Seq.forall (fun ch -> Char.IsLetterOrDigit ch || ch = '_') && not (Char.IsDigit name.[0]) then name
    else "'" + name.Replace("'", "''") + "'"

/// A block as a formula anywhere refers to it: 'Sheet 2'!$A$1:$B$3.
let private absoluteReference (book: MacroWorkbook) (r: RangeRef) =
    quotedSheet (book.Sheet r.Sheet).Name + "!" + cellAddress r.Top r.Left true true
    + (if r.Top = r.Bottom && r.Left = r.Right then "" else ":" + cellAddress r.Bottom r.Right true true)

/// A cell's value as text entered into it would make it ("123" a number, "=A1*2" a formula).
let private entered (book: MacroWorkbook) (sheetId: SheetId) (row: int) (col: int) (value: Value) : CellData option =
    let previous = book.Cell(sheetId, row, col)
    let format =
        match previous with
        | Some cell -> cell.Format
        | None -> SheetOps.effectiveFormat (book.Model sheetId) { Row = row; Col = col }
    let comment = previous |> Option.bind (fun cell -> cell.Comment)
    let withValue v = Some { Value = v; Format = format; Comment = comment }
    match scalarOf value with
    | VEmpty | VNull | VNothing | VUndefined -> if comment.IsSome || format <> defaultFormat then withValue CellValue.Empty else None
    | VInteger i | VLong i -> withValue (Number (float i))
    | VByte b -> withValue (Number (float b))
    // (A Single goes into a cell as its exact binary value, as Excel stores it - not its 7 printed digits.)
    | VSingle f -> withValue (Number (float f))
    | VDouble d -> withValue (Number d)
    | VBoolean b -> withValue (Boolean b)
    | VDate serial ->
        // A Date gets a date format when the cell has none (as Excel gives it one): a time alone a time's.
        let format =
            match format.NumberFormat with
            | None ->
                let code = if serial >= 0.0 && serial < 1.0 then "h:mm:ss AM/PM" elif serial <> Math.Floor serial then "m/d/yyyy h:mm" else "m/d/yyyy"
                { format with NumberFormat = Some code }
            | Some _ -> format
        Some { Value = Number (if book.Date1904 then serial - 1462.0 else serial); Format = format; Comment = comment }
    | VCurrency units ->
        // (A Currency gets Excel's currency format when the cell has none.)
        let format = if format.NumberFormat.IsNone then { format with NumberFormat = Some "$#,##0.00_);($#,##0.00)" } else format
        Some { Value = Number (float units / 10000.0); Format = format; Comment = comment }
    | VError code -> withValue (Error (errorTextOf code))
    | VString "" -> if comment.IsSome || format <> defaultFormat then withValue CellValue.Empty else None
    | VString text ->
        let cell = Rbt.Spreadsheet.Input.cellFromInputFor book.Date1904 DateTime.Now.Year previous text
        Some { cell with Comment = comment }
    | VArray _ -> fail 13 "Type mismatch"
    | _ -> fail 13 "Type mismatch"

/// A cell's format changed; a cell the sheet does not have yet starts from its column's default style.
let private formatCell (model: SpreadsheetModel) (address: CellAddress) (change: CellFormat -> CellFormat) =
    match getCell address model with
    | Some _ -> setCellFormat address change model
    | None ->
        let start = SheetOps.effectiveFormat model address
        let next = change start
        if next = defaultFormat then model else setCell address { Value = CellValue.Empty; Format = next; Comment = None } model

/// The default styles of columns with those of `first` .. `last` changed (None: removed).
let private restyledColumns (styles: ColumnStyle list) (first: int) (last: int) (change: (CellFormat -> CellFormat) option) =
    let outside =
        styles |> List.collect (fun style ->
            if style.Last < first || style.First > last then [ style ]
            else
                [ if style.First < first then yield { style with Last = first - 1 }
                  if style.Last > last then yield { style with First = last + 1 } ])
    match change with
    | None -> outside
    | Some change ->
        let formatOf col =
            styles |> List.tryFind (fun style -> col >= style.First && col <= style.Last) |> Option.map (fun style -> style.Format) |> Option.defaultValue defaultFormat
        // Runs of columns that end up with the same style.
        let runs = ResizeArray<ColumnStyle>()
        for col in first .. last do
            let format = change (formatOf col)
            if runs.Count > 0 && runs.[runs.Count - 1].Last = col - 1 && runs.[runs.Count - 1].Format = format then
                runs.[runs.Count - 1] <- { runs.[runs.Count - 1] with Last = col }
            else runs.Add { First = col; Last = col; Format = format; Source = -1 }
        outside @ (runs |> Seq.filter (fun style -> style.Format <> defaultFormat) |> List.ofSeq) |> List.sortBy (fun style -> style.First)

/// Excel's 3-colour scale as AddColorScale(3) makes it (red, yellow, green), and the 2-colour one.
let private defaultColorScale (stops: int) =
    if stops = 2 then [ { Kind = CfvoMin; Value = None }, "#F8696B"; { Kind = CfvoMax; Value = None }, "#63BE7B" ]
    else [ { Kind = CfvoMin; Value = None }, "#F8696B"; { Kind = CfvoPercentile; Value = Some "50" }, "#FFEB84"; { Kind = CfvoMax; Value = None }, "#63BE7B" ]

/// An argument of a call by its position; past the arguments given, VEmpty.
let private at (args: Value array) (index: int) = if index < args.Length then args.[index] else VEmpty

/// A setting of Excel's that changes nothing here: a write is accepted, and it is not read.
let private ignored (name: string) : Member<'T> = set name (fun _ _ -> ())

/// Excel's ColorIndex of a colour (`none` without one; xlColorIndexAutomatic off the palette), and back.
let private paletteIndex (color: string option) (none: int) =
    match color with
    | None -> none
    | Some c -> match Array.IndexOf(palette, c.ToUpperInvariant()) with -1 -> -4105 | i -> i + 1
let private paletteColor (index: int) = if index >= 1 && index <= 56 then Some palette.[index - 1] else None

let private alignmentName (value: Value) (horizontal: bool) =
    match toInt value with
    | -4131 -> Some "left"
    | -4152 -> Some "right"
    | -4108 -> if horizontal then Some "center" else Some "middle"
    | 7 -> Some "centerContinuous"
    | 5 -> Some "fill"
    | -4117 -> Some "distributed"
    | -4130 -> Some "justify"
    | -4160 -> Some "top"
    | -4107 -> Some "bottom"
    | _ -> None

let private alignmentValue (name: string option) (horizontal: bool) =
    match name with
    | Some "left" -> -4131 | Some "right" -> -4152 | Some "center" | Some "middle" -> -4108
    | Some "justify" -> -4130 | Some "top" -> -4160 | Some "bottom" -> -4107
    | Some "centerContinuous" -> 7 | Some "fill" -> 5 | Some "distributed" -> -4117
    | _ -> if horizontal then 1 else -4107

/// Range.Orientation as the format keeps it (TextRotation): xlHorizontal, xlVertical (stacked), xlUpward,
/// xlDownward, or degrees (-90 .. 90).
let private orientationOf (rotation: int) = match rotation with 0 -> -4128 | 255 -> -4166 | d when d > 90 -> -(d - 90) | d -> d
let private rotationOf (value: Value) =
    match toInt value with
    | -4128 -> 0
    | -4166 -> 255
    | -4171 -> 90
    | -4170 -> 180
    | d when d >= 0 && d <= 90 -> d
    | d when d < 0 && d >= -90 -> 90 - d
    | _ -> failApp "Unable to set the Orientation property of the Range class"

// ---- the state of each type of object ----

/// A Range: a block of cells (`Ref`) of the workbook a run works on, and what it stands for - the block,
/// its rows or columns (`Lines` 1, 2: Count, Item and For Each go over them; 3, 4: a sheet's, which take
/// "3:5" and "A" too), or several blocks (a union: `Areas`, the first of them `Ref`).
type private RangeState(book: MacroWorkbook, r: RangeRef, lines: int, areas: RangeRef list) =
    let rows = r.Bottom - r.Top + 1
    let cols = r.Right - r.Left + 1
    let wholeRows = r.Left = 0 && r.Right = maxCol
    let wholeCols = r.Top = 0 && r.Bottom = maxRow
    let vertical = lines = 2 || lines = 4
    let inside (a: CellAddress) = a.Row >= r.Top && a.Row <= r.Bottom && a.Col >= r.Left && a.Col <= r.Right
    let block (ref: RangeRef) = RangeState(book, ref, 0, [])
    let sub top left bottom right = block { r with Top = top; Left = left; Bottom = bottom; Right = right }
    let cellAt row col = book.Cell(r.Sheet, row, col)
    let valueAt row col = match cellAt row col with Some c -> cellValueAs book.Date1904 c.Format c.Value | None -> VEmpty
    let formulaAt row col =
        match cellAt row col with
        | Some { Value = Formula (text, _) } -> VString (if text.StartsWith "=" then text else "=" + text)
        | Some cell -> match cellValue cell.Value with VEmpty -> VString "" | other -> other
        | None -> VString ""
    let gather (item: int -> int -> Value) =
        if rows = 1 && cols = 1 then item r.Top r.Left
        else array2D rows cols (fun i j -> item (r.Top + i) (r.Left + j))
    let assign (value: Value) (cellOf: int -> int -> Value -> CellData option) =
        match scalarOf value with
        | VArray array ->
            for i in 0 .. rows - 1 do
                for j in 0 .. cols - 1 do
                    let item = arrayItem array i j |> Option.defaultValue (VString "#N/A")
                    book.SetCell(r.Sheet, r.Top + i, r.Left + j, cellOf (r.Top + i) (r.Left + j) item)
        | scalar ->
            for row in r.Top .. r.Bottom do
                for col in r.Left .. r.Right do
                    book.SetCell(r.Sheet, row, col, cellOf row col scalar)
    let format () = match cellAt r.Top r.Left with Some c -> c.Format | None -> defaultFormat
    let used () =
        let model = book.Model r.Sheet
        if Map.isEmpty model.Cells then None
        else
            let addresses = model.Cells |> Map.toSeq |> Seq.map fst
            Some (addresses |> Seq.map (fun a -> a.Row) |> Seq.min, addresses |> Seq.map (fun a -> a.Col) |> Seq.min,
                  addresses |> Seq.map (fun a -> a.Row) |> Seq.max, addresses |> Seq.map (fun a -> a.Col) |> Seq.max)
    let filled row col = match cellAt row col with Some { Value = v } -> unwrapFormulaValue v <> CellValue.Empty | None -> false
    // ---- borders: a line between two cells is one line (setting the bottom of A1 is setting the top of A2) ----
    let formatIn (model: SpreadsheetModel) row col = getCell { Row = row; Col = col } model |> Option.map (fun c -> c.Format)
    /// The sides Borders(index) stands for, (row, col, side) of the range's cells, each with the side of the
    /// cell beyond it that is the same line, when there is one.
    let bordersSides (index: int option) =
        // (Whole rows or columns: as far as the grid goes.)
        let model = book.Model r.Sheet
        let bottom = if wholeCols then min r.Bottom (max r.Top (model.RowCount - 1)) else min r.Bottom (r.Top + 100000)
        let right = if wholeRows then min r.Right (max r.Left (model.ColCount - 1)) else min r.Right (r.Left + 16384)
        let opposite side = match side with 0 -> 3 | 1 -> 2 | 2 -> 1 | 3 -> 0 | other -> other
        let beyond row col side = match side with 0 -> row, col - 1 | 1 -> row - 1, col | 2 -> row + 1, col | _ -> row, col + 1
        let sides =
            match index with
            | None -> [ for row in r.Top .. bottom do for col in r.Left .. right do for side in 0 .. 3 -> row, col, side ]
            | Some 7 | Some -4131 -> [ for row in r.Top .. bottom -> row, r.Left, 0 ]
            | Some 8 | Some -4160 -> [ for col in r.Left .. right -> r.Top, col, 1 ]
            | Some 9 | Some -4107 -> [ for col in r.Left .. right -> bottom, col, 2 ]
            | Some 10 | Some -4152 -> [ for row in r.Top .. bottom -> row, right, 3 ]
            | Some 11 -> [ for row in r.Top .. bottom do for col in r.Left .. right - 1 -> row, col, 3 ]
            | Some 12 -> [ for row in r.Top .. bottom - 1 do for col in r.Left .. right -> row, col, 2 ]
            | Some 5 -> [ for row in r.Top .. bottom do for col in r.Left .. right -> row, col, 4 ]
            | Some 6 -> [ for row in r.Top .. bottom do for col in r.Left .. right -> row, col, 5 ]
            | Some _ -> failApp "Unable to get the Item property of the Borders class"
        sides |> List.map (fun (row, col, side) ->
            let nr, nc = beyond row col side
            (row, col, side), (if side < 4 && nr >= 0 && nc >= 0 && nr <= maxRow && nc <= maxCol then Some (nr, nc, opposite side) else None))
    let borderIn (model: SpreadsheetModel) ((row, col, side), shared) =
        match formatIn model row col |> Option.bind (fun f -> sideOf f side) with
        | Some border -> Some border
        | None -> shared |> Option.bind (fun (nr, nc, ns) -> formatIn model nr nc |> Option.bind (fun f -> sideOf f ns))
    let changeBorders (index: int option) (update: BorderStyle option -> BorderStyle option) =
        let sides = bordersSides index
        book.ChangeModel(r.Sheet, fun model ->
            let mutable m = model
            for ((row, col, side), shared) as entry in sides do
                let next = update (borderIn m entry)
                m <- formatCell m { Row = row; Col = col } (withSide side next)
                // (The cell beyond has the same line, as Excel keeps both cells' borders: A1's bottom stays
                // when row 2 is deleted, A2's top reads what A1's bottom was set to.)
                match shared with
                | Some (nr, nc, ns) when next.IsSome || (formatIn m nr nc |> Option.bind (fun f -> sideOf f ns)).IsSome ->
                    m <- formatCell m { Row = nr; Col = nc } (withSide ns next)
                | _ -> ()
            m)
        book.Grow(r.Sheet, (if wholeCols then 0 else min r.Bottom (r.Top + 100000)), (if wholeRows then 0 else min r.Right (r.Left + 16384)))
    let note () = getCell { Row = r.Top; Col = r.Left } (book.Model r.Sheet) |> Option.bind (fun c -> c.Comment)
    let setNote (note: CellComment option) =
        book.ChangeModel(r.Sheet, fun model ->
            let address = { Row = r.Top; Col = r.Left }
            match Map.tryFind address model.Cells, note with
            | Some data, _ -> { model with Cells = Map.add address { data with Comment = note; Format = { data.Format with HasComment = note.IsSome } } model.Cells }
            | None, Some _ -> { model with Cells = Map.add address { Value = CellValue.Empty; Format = { defaultFormat with HasComment = true }; Comment = note } model.Cells }
            | None, None -> model)
    let lineCount = if vertical then cols else rows
    let lineItem (i: int) = if vertical then sub r.Top (r.Left + i - 1) r.Bottom (r.Left + i - 1) else sub (r.Top + i - 1) r.Left (r.Top + i - 1) r.Right
    /// A sheet's Rows(index) / Columns(index): a number, "3:5", or a column's letters.
    let sheetLine (args: Value array) =
        let whole = { r with Top = 0; Left = 0; Bottom = maxRow; Right = maxCol }
        match args with
        | [||] -> block whole
        | [| index |] ->
            match scalarOf index with
            | VString text when text.Contains ":" ->
                match parseReference text with
                | Some (_, a, b, c, d) -> block { Sheet = r.Sheet; Top = a; Left = b; Bottom = c; Right = d }
                | None -> failApp "Application-defined or object-defined error"
            | VString letters when vertical && letters <> "" && Char.IsLetter letters.[0] ->
                match parseReference (letters + "1") with
                | Some (_, _, c, _, _) -> block { whole with Left = c; Right = c }
                | None -> failApp "Application-defined or object-defined error"
            | v ->
                let i = toInt v - 1
                if vertical then block { whole with Left = i; Right = i } else block { whole with Top = i; Bottom = i }
        | _ -> fail 450 "Wrong number of arguments or invalid property assignment"

    member _.Book = book
    member _.Ref = r
    member _.Sheet = r.Sheet
    member _.Lines = lines
    member _.WholeRows = wholeRows
    member _.WholeCols = wholeCols
    member _.IsUnion = not areas.IsEmpty
    member _.Areas = if areas.IsEmpty then [ r ] else areas
    member _.Sub(top, left, bottom, right) = sub top left bottom right
    /// The block itself (not its rows or columns; a union's first area).
    member this.Plain() = if lines = 0 && areas.IsEmpty then this else block r
    member _.LinesOf(vertical: bool) = RangeState(book, r, (if vertical then 2 else 1), [])
    member _.CellAt(row, col) = cellAt row col
    member _.Format() = format ()
    member _.Used() = used ()

    // ---- values ----
    member _.Values() = gather valueAt
    member _.Values2() = gather (fun row col -> match cellAt row col with Some c -> cellValue c.Value | None -> VEmpty)
    member _.Formulas() = gather formulaAt
    member _.FormulasR1C1() =
        gather (fun row col -> match formulaAt row col with VString text when text.StartsWith "=" -> VString (toR1C1 row col text) | other -> other)
    member _.Enter(value: Value) = assign value (fun row col item -> entered book r.Sheet row col item)
    member _.EnterR1C1(value: Value) =
        assign value (fun row col item ->
            match scalarOf item with
            | VString text when text.StartsWith "=" -> entered book r.Sheet row col (VString (fromR1C1 row col text))
            | other -> entered book r.Sheet row col other)
    member _.Text() =
        match cellAt r.Top r.Left with
        | Some cell ->
            let width = book.Model(r.Sheet).ColumnWidths |> Map.tryFind r.Left |> Option.defaultValue (book.Model(r.Sheet).DefaultColumnWidth)
            VString (CellText.cellDisplayTextIn book.Date1904 None width cell.Format cell.Value)
        | None -> VString ""

    // ---- where it is ----
    member _.Address(args: Value array) =
        let rowAbs = if isMissing (at args 0) then true else toBool (at args 0)
        let colAbs = if isMissing (at args 1) then true else toBool (at args 1)
        let r1c1 = args.Length >= 3 && not (isMissing (at args 2)) && toInt (at args 2) = -4150
        let external = args.Length >= 4 && not (isMissing (at args 3)) && toBool (at args 3)
        let text = rangeAddress r rowAbs colAbs r1c1
        if external then "[" + book.Hooks.FileName + "]" + quotedSheet (book.Sheet r.Sheet).Name + "!" + text else text
    // (Count is a Long, as Excel's - which overflows past 2^31 cells; CountLarge is the Double.)
    member _.Count() =
        if lines > 0 then VLong lineCount
        elif not areas.IsEmpty then VDouble (areas |> List.sumBy (fun a -> float (a.Bottom - a.Top + 1) * float (a.Right - a.Left + 1)))
        else
            let cells = float rows * float cols
            if cells > 2147483647.0 then fail 6 "Overflow" else VLong (int cells)
    member _.CountLarge() =
        if areas.IsEmpty then VDouble (float rows * float cols)
        else VDouble (areas |> List.sumBy (fun a -> float (a.Bottom - a.Top + 1) * float (a.Right - a.Left + 1)))
    /// Cells(row, col) and Item(index): relative to the range, by rows then columns.
    member _.CellItem(args: Value array) =
        if args.Length >= 2 && not (isMissing args.[1]) then
            let row = r.Top + toInt args.[0] - 1
            let col =
                match scalarOf args.[1] with
                | VString letters when letters <> "" && Char.IsLetter letters.[0] ->
                    match parseReference (letters + "1") with
                    | Some (_, _, c, _, _) -> c
                    | None -> fail 13 "Type mismatch"
                | v -> r.Left + toInt v - 1
            sub row col row col
        else
            let index = toInt args.[0] - 1
            let row = r.Top + index / cols
            let col = r.Left + index % cols
            sub row col row col
    /// Item(...) and the range called (`r(2)`): a cell; of its rows or columns, a row or column (past its
    /// end too, as Excel).
    member this.ItemAt(args: Value array) =
        match lines with
        | 1 | 2 -> lineItem (toInt args.[0])
        | 3 | 4 -> sheetLine args
        | _ -> this.CellItem args
    member _.LineItem(i: int) = lineItem i
    /// How many of its rows or columns For Each goes over (whole rows / columns: as far as the sheet is used).
    member _.UsedLineCount() =
        if not (wholeRows || wholeCols) then lineCount
        else
            match used () with
            | Some (_, _, usedBottom, usedRight) -> if vertical then min lineCount (usedRight - r.Left + 1) else min lineCount (usedBottom - r.Top + 1)
            | None -> 0
    /// For Each over a range: its cells, row by row (whole rows / columns only as far as the sheet is used).
    member _.Cells() =
        let bottom, right =
            if wholeRows || wholeCols then
                match used () with
                | Some (_, _, usedBottom, usedRight) -> min r.Bottom usedBottom, min r.Right usedRight
                | None -> r.Top - 1, r.Left - 1
            else r.Bottom, r.Right
        seq { for row in r.Top .. bottom do
                for col in r.Left .. right -> sub row col row col }
    member _.Range(text: string) =
        match parseReference text with
        | Some (_, a, b, c, d) -> sub (r.Top + a) (r.Left + b) (r.Top + c) (r.Left + d)
        | None -> failApp "Method 'Range' of object 'Range' failed"
    member _.Offset(args: Value array) =
        let dr = if isMissing (at args 0) then 0 else toInt (at args 0)
        let dc = if isMissing (at args 1) then 0 else toInt (at args 1)
        if r.Top + dr < 0 || r.Left + dc < 0 || r.Bottom + dr > maxRow || r.Right + dc > maxCol then failApp "Application-defined or object-defined error"
        sub (r.Top + dr) (r.Left + dc) (r.Bottom + dr) (r.Right + dc)
    member _.Resize(args: Value array) =
        let rs = if isMissing (at args 0) then rows else toInt (at args 0)
        let cs = if isMissing (at args 1) then cols else toInt (at args 1)
        if rs < 1 || cs < 1 then failApp "Application-defined or object-defined error"
        sub r.Top r.Left (min maxRow (r.Top + rs - 1)) (min maxCol (r.Left + cs - 1))
    /// Ctrl+arrow from the range's first cell: to the edge of the block it is in, or of the next one.
    member _.EndOf(direction: int) =
        let dr, dc = match direction with -4162 -> -1, 0 | -4121 -> 1, 0 | -4159 -> 0, -1 | -4161 -> 0, 1 | _ -> fail 5 "Invalid procedure call or argument"
        let inBounds row col = row >= 0 && row <= maxRow && col >= 0 && col <= maxCol
        let mutable row, col = r.Top, r.Left
        let next () = row + dr, col + dc
        let lastFilled =
            match used () with
            | Some (_, _, bottom, right) -> bottom, right
            | None -> 0, 0
        let beyondData row col = (dr > 0 && row > fst lastFilled) || (dc > 0 && col > snd lastFilled)
        let nr, nc = next ()
        if not (inBounds nr nc) then ()
        elif filled row col && filled nr nc then
            // Along the block to its last filled cell.
            while (let a, b = next () in inBounds a b && filled a b) do
                let a, b = next ()
                row <- a; col <- b
        else
            // Over the gap to the next filled cell, or to the sheet's edge.
            let mutable a, b = nr, nc
            while inBounds a b && not (filled a b) && not (beyondData a b) do
                a <- a + dr; b <- b + dc
            if inBounds a b && filled a b then row <- a; col <- b
            else
                row <- (if dr > 0 then maxRow elif dr < 0 then 0 else row)
                col <- (if dc > 0 then maxCol elif dc < 0 then 0 else col)
        sub row col row col
    member _.CurrentRegion() =
        let mutable top, left, bottom, right = r.Top, r.Left, r.Bottom, r.Right
        let mutable grown = true
        let anyFilled r0 c0 r1 c1 =
            seq { for row in max 0 r0 .. min maxRow r1 do for col in max 0 c0 .. min maxCol c1 -> row, col } |> Seq.exists (fun (a, b) -> filled a b)
        while grown do
            grown <- false
            if top > 0 && anyFilled (top - 1) (left - 1) (top - 1) (right + 1) then top <- top - 1; grown <- true
            if bottom < maxRow && anyFilled (bottom + 1) (left - 1) (bottom + 1) (right + 1) then bottom <- bottom + 1; grown <- true
            if left > 0 && anyFilled (top - 1) (left - 1) (bottom + 1) (left - 1) then left <- left - 1; grown <- true
            if right < maxCol && anyFilled (top - 1) (right + 1) (bottom + 1) (right + 1) then right <- right + 1; grown <- true
        sub top left bottom right
    /// Find(What, After, LookIn, LookAt, SearchOrder, SearchDirection, MatchCase): the first cell, by rows.
    member _.Find(args: Value array) =
        let what = toText (at args 0)
        let whole = args.Length > 4 && not (isMissing (at args 4)) && toInt (at args 4) = 1
        let matchCase = args.Length > 6 && not (isMissing (at args 6)) && toBool (at args 6)
        let comparison = if matchCase then StringComparison.Ordinal else StringComparison.OrdinalIgnoreCase
        book.Model(r.Sheet).Cells |> Map.toSeq
        |> Seq.filter (fun (a, _) -> inside a)
        |> Seq.sortBy (fun (a, _) -> a.Row, a.Col)
        |> Seq.tryFind (fun (a, _) ->
            let text = toText (valueAt a.Row a.Col)
            if whole then String.Equals(text, what, comparison)
            elif matchCase then text.Contains what
            else text.ToLowerInvariant().Contains(what.ToLowerInvariant()))
        |> Option.map (fun (a, _) -> sub a.Row a.Col a.Row a.Col)
    // Left, Top, Width, Height: in points from the sheet's top left (where a shape over these cells goes).
    member _.Span(edge: int) =
        let model = book.Model(r.Sheet)
        // (The sum of sizes of a run of columns or rows: the default for each, with the ones set apart.)
        let span (sizes: Map<int, float>) (fallback: float) (first: int) (last: int) =
            if last < first then 0.0
            else
                let mutable total = float (last - first + 1) * fallback
                for KeyValue (index, size) in sizes do
                    if index >= first && index <= last then total <- total + size - fallback
                total
        let px =
            match edge with
            | 0 -> span model.ColumnWidths model.DefaultColumnWidth 0 (r.Left - 1)
            | 1 -> span model.RowHeights model.DefaultRowHeight 0 (r.Top - 1)
            | 2 -> span model.ColumnWidths model.DefaultColumnWidth r.Left r.Right
            | _ -> span model.RowHeights model.DefaultRowHeight r.Top r.Bottom
        VDouble (px * 0.75)

    // ---- formats ----
    member _.ChangeFormats(change: CellFormat -> CellFormat) =
        if wholeRows || wholeCols then
            // Whole columns, rows or the sheet (Cells.Interior.Color = vbWhite): the cells it has and those the
            // grid shows, and for whole columns their default style (a cell made later starts from it).
            book.ChangeModel(r.Sheet, fun model ->
                let mutable m = model
                for KeyValue(address, _) in model.Cells do
                    if inside address then m <- formatCell m address change
                for row in r.Top .. min r.Bottom (model.RowCount - 1) do
                    for col in r.Left .. min r.Right (model.ColCount - 1) do
                        let address = { Row = row; Col = col }
                        if not (Map.containsKey address model.Cells) then m <- formatCell m address change
                if wholeCols then { m with ColumnStyles = restyledColumns m.ColumnStyles r.Left r.Right (Some change) } else m)
        else
            let bottom, right = min r.Bottom (r.Top + 100000), min r.Right (r.Left + 16384)
            book.ChangeModel(r.Sheet, fun model ->
                let mutable m = model
                for row in r.Top .. bottom do
                    for col in r.Left .. right do
                        m <- formatCell m { Row = row; Col = col } change
                m)
            book.Grow(r.Sheet, bottom, right)
    member _.MergeCells() =
        book.Model(r.Sheet).MergedCells |> List.exists (fun m -> m.StartRow <= r.Top && m.EndRow >= r.Top && m.StartCol <= r.Left && m.EndCol >= r.Left)
    member _.Merge() =
        book.ChangeModel(r.Sheet, fun m ->
            let others = m.MergedCells |> List.filter (fun x -> x.EndRow < r.Top || x.StartRow > r.Bottom || x.EndCol < r.Left || x.StartCol > r.Right)
            { m with MergedCells = { StartRow = r.Top; StartCol = r.Left; EndRow = r.Bottom; EndCol = r.Right } :: others })
    member _.UnMerge() =
        book.ChangeModel(r.Sheet, fun m -> { m with MergedCells = m.MergedCells |> List.filter (fun x -> x.EndRow < r.Top || x.StartRow > r.Bottom || x.EndCol < r.Left || x.StartCol > r.Right) })
    member _.ColumnWidth() =
        let px = book.Model(r.Sheet).ColumnWidths |> Map.tryFind r.Left |> Option.defaultValue (book.Model(r.Sheet).DefaultColumnWidth)
        VDouble (if px < 12.0 then Math.Round(px / 12.0, 2) else Math.Round((px - 5.0) / 7.0, 2))
    member _.SetColumnWidth(width: float) =
        // (Excel: a width under one character is 12 px a character; from one on, 7 px each and 5 of padding.)
        let px = if width <= 0.0 then 0.0 elif width < 1.0 then Math.Round(width * 12.0) else Math.Floor((256.0 * width + 18.0) / 256.0 * 7.0) + 5.0
        if wholeRows then
            // Every column: the sheet's default width.
            book.ChangeModel(r.Sheet, fun m -> { m with DefaultColumnWidth = px; ColumnWidths = Map.empty })
        else
            book.ChangeModel(r.Sheet, fun m -> { m with ColumnWidths = [ r.Left .. r.Right ] |> List.fold (fun acc c -> Map.add c px acc) m.ColumnWidths })
            book.Grow(r.Sheet, 0, r.Right)
    member _.RowHeight() =
        let px = book.Model(r.Sheet).RowHeights |> Map.tryFind r.Top |> Option.defaultValue (book.Model(r.Sheet).DefaultRowHeight)
        VDouble (Math.Round(px * 0.75, 2))
    member _.SetRowHeight(points: float) =
        let px = points / 0.75
        if wholeCols then
            book.ChangeModel(r.Sheet, fun m -> { m with DefaultRowHeight = px; RowHeights = Map.empty })
        else
            book.ChangeModel(r.Sheet, fun m -> { m with RowHeights = [ r.Top .. min r.Bottom (r.Top + 100000) ] |> List.fold (fun acc row -> Map.add row px acc) m.RowHeights })
            book.Grow(r.Sheet, r.Bottom, 0)
    member _.Hidden() =
        if wholeCols then book.Model(r.Sheet).ColumnWidths |> Map.tryFind r.Left = Some 0.0 else book.Model(r.Sheet).RowHeights |> Map.tryFind r.Top = Some 0.0
    member _.SetHidden(hidden: bool) =
        if wholeCols then
            book.ChangeModel(r.Sheet, fun m -> { m with ColumnWidths = [ r.Left .. r.Right ] |> List.fold (fun acc c -> if hidden then Map.add c 0.0 acc else Map.remove c acc) m.ColumnWidths })
        else
            book.ChangeModel(r.Sheet, fun m -> { m with RowHeights = [ r.Top .. r.Bottom ] |> List.fold (fun acc row -> if hidden then Map.add row 0.0 acc else Map.remove row acc) m.RowHeights })
    /// Range.Name = "x": a workbook name for these cells (Names.Add).
    member _.SetName(text: string) =
        match definedNameProblem text with
        | Some problem -> failApp problem
        | None -> book.ChangeWorkbook(setDefinedName None text (absoluteReference book r) None)
    member _.BordersSides(index: int option) = bordersSides index
    member _.BorderIn(model: SpreadsheetModel, entry) = borderIn model entry
    member _.ChangeBorders(index: int option, update: BorderStyle option -> BorderStyle option) = changeBorders index update
    /// BorderAround(LineStyle, Weight, ColorIndex, Color, ThemeColor): the four edges.
    member _.BorderAround(args: Value array) =
        let lineStyle = if isMissing (at args 0) then xlContinuous else toInt (at args 0)
        let weight = if isMissing (at args 1) then 2 else toInt (at args 1)
        let color =
            if args.Length > 3 && not (isMissing (at args 3)) then Some (longToColor (toInt (at args 3)))
            elif args.Length > 4 && not (isMissing (at args 4)) then Some (themeColor (toInt (at args 4)))
            elif args.Length > 2 && not (isMissing (at args 2)) then paletteColor (toInt (at args 2))
            else None
        let border = styleOfLine lineStyle weight |> Option.map (fun style -> { Style = style; Color = color })
        for edge in [ 7; 8; 9; 10 ] do changeBorders (Some edge) (fun _ -> border)

    // ---- contents ----
    member _.ClearContents() =
        for (address, cell) in book.Model(r.Sheet).Cells |> Map.toList do
            if inside address then
                book.SetCell(r.Sheet, address.Row, address.Col, (if cell.Format = defaultFormat && cell.Comment.IsNone then None else Some { cell with Value = CellValue.Empty }))
    member _.ClearFormats() =
        if wholeCols then book.ChangeModel(r.Sheet, fun m -> { m with ColumnStyles = restyledColumns m.ColumnStyles r.Left r.Right None })
        for (address, cell) in book.Model(r.Sheet).Cells |> Map.toList do
            if inside address then
                book.SetCell(r.Sheet, address.Row, address.Col, (if cell.Value = CellValue.Empty && cell.Comment.IsNone then None else Some { cell with Format = defaultFormat }))
    member _.ClearComments() =
        let clear (a: CellAddress) (data: CellData) =
            if inside a && data.Comment.IsSome then { data with Comment = None; Format = { data.Format with HasComment = false } } else data
        book.ChangeModel(r.Sheet, fun model -> { model with Cells = model.Cells |> Map.map clear })
    /// The cells copied to the block at `target` (xlPasteAll, xlPasteValues, xlPasteFormats).
    member _.CopyTo(target: RangeRef, what: int) =
        let source = [ for row in r.Top .. r.Bottom do for col in r.Left .. r.Right -> row, col, cellAt row col ]
        let dr, dc = target.Top - r.Top, target.Left - r.Left
        for row, col, cell in source do
            let existing = book.Cell(target.Sheet, row + dr, col + dc)
            let next =
                match cell, what with
                | None, -4122 -> existing |> Option.map (fun c -> { c with Format = defaultFormat })
                | None, _ -> None
                | Some c, -4163 -> Some { Value = (match c.Value with Formula (_, cached) -> cached | v -> v); Format = (existing |> Option.map (fun e -> e.Format) |> Option.defaultValue defaultFormat); Comment = existing |> Option.bind (fun e -> e.Comment) }
                | Some c, -4122 -> Some { Value = (existing |> Option.map (fun e -> e.Value) |> Option.defaultValue CellValue.Empty); Format = c.Format; Comment = existing |> Option.bind (fun e -> e.Comment) }
                | Some c, _ ->
                    let value = match c.Value with Formula (text, _) -> Formula (Rbt.Spreadsheet.Formula.shiftFormulaText dr dc text, CellValue.Empty) | v -> v
                    Some { c with Value = value }
            book.SetCell(target.Sheet, row + dr, col + dc, next)
    /// Delete(Shift) / Insert(Shift): whole rows or columns, else the cells shifted up / left (Delete) or
    /// down / right (Insert).
    member _.Shift(inserting: bool, args: Value array) =
        if wholeRows then book.Structure(r.Sheet, (if inserting then InsertRows (r.Top, rows) else DeleteRows (r.Top, rows)))
        elif wholeCols then book.Structure(r.Sheet, (if inserting then InsertCols (r.Left, cols) else DeleteCols (r.Left, cols)))
        else
            let shift = if isMissing (at args 0) then (if inserting then -4121 else -4162) else toInt (at args 0)
            let model = book.Model r.Sheet
            let horizontal = shift = -4159 || shift = -4161
            let moved =
                model.Cells |> Map.toList
                |> List.filter (fun (a, _) ->
                    if horizontal then a.Row >= r.Top && a.Row <= r.Bottom && a.Col >= r.Left
                    else a.Col >= r.Left && a.Col <= r.Right && a.Row >= r.Top)
            for (a, _) in moved do book.SetCell(r.Sheet, a.Row, a.Col, None)
            for (a, cell) in moved do
                let row, col =
                    if inserting then (if horizontal then a.Row, a.Col + cols else a.Row + rows, a.Col)
                    else (if horizontal then a.Row, a.Col - cols else a.Row - rows, a.Col)
                if inserting || not (inside a) then
                    if row >= 0 && col >= 0 && row <= maxRow && col <= maxCol then book.SetCell(r.Sheet, row, col, Some cell)
    /// Replace(What, Replacement, LookAt, SearchOrder, MatchCase): in the values of the range's cells, each
    /// entered again as typing would (a cell left "" is empty).
    member _.Replace(args: Value array) =
        let what = toText (at args 0)
        let replacement = toText (at args 1)
        let whole = not (isMissing (at args 2)) && toInt (at args 2) = 1
        let matchCase = not (isMissing (at args 4)) && toBool (at args 4)
        if what <> "" then
            let cells = book.Model(r.Sheet).Cells |> Map.toList |> List.filter (fun (a, _) -> inside a)
            for (a, cell) in cells do
                let text = match cell.Value with Formula (f, _) -> (if f.StartsWith "=" then f else "=" + f) | v -> toText (cellValue v)
                let next =
                    if whole then (if String.Equals(text, what, (if matchCase then StringComparison.Ordinal else StringComparison.OrdinalIgnoreCase)) then replacement else text)
                    elif matchCase then text.Replace(what, replacement)
                    else Text.RegularExpressions.Regex.Replace(text, Text.RegularExpressions.Regex.Escape what, replacement.Replace("$", "$$"), Text.RegularExpressions.RegexOptions.IgnoreCase)
                if next <> text then book.SetCell(r.Sheet, a.Row, a.Col, entered book r.Sheet a.Row a.Col (VString next))
        VBoolean true
    member _.Select() =
        book.Selection <- Some (r.Sheet, r.Top, r.Left, r.Bottom, r.Right)
        book.SetActiveSheet r.Sheet

    // ---- notes (Range.Comment) and Range.ID ----
    member _.Note() = note ()
    member _.SetNote(next: CellComment option) = setNote next
    member _.NoteText() = note () |> Option.map (fun n -> n.Text) |> Option.defaultValue ""
    member _.IdKey = sprintf "__id\t%s\t%d\t%d" r.Sheet r.Top r.Left

    // ---- conditional formats of the range (FormatConditions) ----
    member _.Conditions() =
        let overlaps (cr: CellRange) = not (cr.To.Row < r.Top || cr.From.Row > r.Bottom || cr.To.Col < r.Left || cr.From.Col > r.Right)
        book.Model(r.Sheet).ConditionalFormats |> List.filter (fun e -> e.Ranges |> List.exists overlaps) |> List.sortBy (fun e -> e.Priority)
    /// A conditional format of the range with `rule` (and no style yet); its id.
    member _.AddCondition(rule: ConditionalFormatRule) =
        let highest = book.SheetOrder |> List.fold (fun value id -> max value (SheetOps.highestConditionalFormatId (book.Model id))) 0
        let id = highest + 1
        let thisRange : CellRange = { From = { Row = r.Top; Col = r.Left }; To = { Row = r.Bottom; Col = r.Right } }
        book.ChangeModel(r.Sheet, SheetOps.addConditionalFormatWithId id [ thisRange ] false { Rule = rule; Style = defaultFormat })
        id
    /// What the first cell shows with (DisplayFormat): its format, with a conditional format that holds on it.
    member _.Displayed() =
        let model = book.Model r.Sheet
        let address = { Row = r.Top; Col = r.Left }
        let own = SheetOps.effectiveFormat model address
        if model.ConditionalFormats.IsEmpty then own
        else
            cellAt r.Top r.Left |> ignore  // (calculated first)
            let context = ConditionalFormatting.createContext evaluateFormulaAt (book.Model r.Sheet)
            match ConditionalFormatting.effectAt context address with
            | Some effect -> ConditionalFormatting.applyEffect effect own
            | None -> own

let private rangeState (book: MacroWorkbook) (r: RangeRef) = RangeState(book, r, 0, [])

/// Font or Interior of a format kept elsewhere (a conditional format's style, DisplayFormat): its colours,
/// bold and italic, read from `read ()` and written with `change` (None: read only).
type private StyleState(typeName: string, read: unit -> CellFormat, change: ((CellFormat -> CellFormat) -> unit) option) =
    member _.TypeName = typeName
    member _.Read() = read ()
    member _.Write(f: CellFormat -> CellFormat) =
        match change with
        | Some apply -> apply f
        | None -> failApp (sprintf "Unable to set a property of the %s class" typeName)

/// Borders (index None: all the sides of the range's cells) or one Border of a range.
type private BorderState(range: RangeState, index: int option) =
    member _.Range = range
    /// What every side has of `pick`, else Null.
    member _.One(pick: BorderStyle option -> Value) =
        let model = range.Book.Model range.Sheet
        match range.BordersSides index |> List.map (fun entry -> pick (range.BorderIn(model, entry))) |> List.distinct with
        | [ v ] -> v
        | _ -> VNull
    member _.Change(update: BorderStyle option -> BorderStyle option) = range.ChangeBorders(index, update)

/// A conditional format of a sheet, by its id (FormatCondition), as a range's FormatConditions has it.
type private ConditionState(range: RangeState, id: int) =
    member _.Range = range
    member _.Id = id
    member _.Entry() =
        match (range.Book.Model range.Sheet).ConditionalFormats |> List.tryFind (fun e -> e.Id = id) with
        | Some e -> e
        | None -> failApp "Application-defined or object-defined error"
    member _.Update(f: ConditionalFormatRange -> ConditionalFormatRange) =
        range.Book.ChangeModel(range.Sheet, fun m -> { m with ConditionalFormats = m.ConditionalFormats |> List.map (fun e -> if e.Id = id then f e else e) })
    member this.ChangeStyle(f: CellFormat -> CellFormat) = this.Update(fun e -> { e with Format = { e.Format with Style = f e.Format.Style } })
    member this.Stops() = match (this.Entry ()).Format.Rule with ColorScale stops -> stops | _ -> failApp "This condition is not a colour scale"
    member this.ChangeStop(index: int, f: CfThreshold * string -> CfThreshold * string) =
        this.Update(fun e ->
            match e.Format.Rule with
            | ColorScale stops -> { e with Format = { e.Format with Rule = ColorScale (stops |> List.mapi (fun i stop -> if i = index - 1 then f stop else stop)) } }
            | _ -> e)

/// A stop of a colour scale (ColorScaleCriterion), 1-based.
type private CriterionState(condition: ConditionState, index: int) =
    member _.Condition = condition
    member _.Index = index
    member _.Stop() = match List.tryItem (index - 1) (condition.Stops ()) with Some s -> s | None -> fail 9 "Subscript out of range"
    member _.Change(f: CfThreshold * string -> CfThreshold * string) = condition.ChangeStop(index, f)

/// The areas of a range (one) or of a union.
type private AreasState(book: MacroWorkbook, areas: RangeRef list) =
    member _.Book = book
    member _.Areas = areas

type private WorksheetState(book: MacroWorkbook, id: SheetId) =
    member _.Book = book
    member _.Id = id
    member _.Whole = { Sheet = id; Top = 0; Left = 0; Bottom = maxRow; Right = maxCol }
    member _.Order = book.SheetOrder |> List.findIndex ((=) id)

/// The objects of a sheet (Shapes; ChartObjects: its charts only).
type private ShapesState(book: MacroWorkbook, sheetId: SheetId, chartsOnly: bool) =
    member _.Book = book
    member _.Sheet = sheetId
    member _.TypeName = if chartsOnly then "ChartObjects" else "Shapes"
    member _.Drawings() =
        (book.Model sheetId).Drawings
        |> List.filter (fun d -> not chartsOnly || (match d.Content with KeptObject ("chart", _) -> true | _ -> false))

/// A shape of a sheet the macro did not draw (a chart's frame, a picture, a kept object), by its id.
type private ShapeState(book: MacroWorkbook, sheetId: SheetId, id: int) =
    member _.Book = book
    member _.Id = id
    member _.Drawing() =
        match (book.Model sheetId).Drawings |> List.tryFind (fun d -> d.Id = id) with
        | Some d -> d
        | None -> failApp "The object has been deleted."
    member _.Rel(d: SheetDrawing) =
        match d.Content with
        | KeptObject ("chart", xml) ->
            let m = Text.RegularExpressions.Regex.Match(xml, "r:id=\"([^\"]+)\"")
            if m.Success then Some m.Groups.[1].Value else None
        | _ -> None
    member _.Change(f: SheetDrawing -> SheetDrawing) =
        book.ChangeModel(sheetId, fun m -> { m with Drawings = m.Drawings |> List.map (fun d -> if d.Id = id then f d else d) })
    member this.Delete() =
        let d = this.Drawing ()
        book.ChangeModel(sheetId, fun m ->
            { m with Drawings = m.Drawings |> List.filter (fun x -> x.Id <> id)
                     AuthoredCharts = match this.Rel d with Some rel -> Map.remove rel m.AuthoredCharts | None -> m.AuthoredCharts })

/// A colour of a shape's fill, line or text (ColorFormat): RGB, or a theme colour (ObjectThemeColor).
type private ColorState(read: unit -> int, write: int -> unit) =
    member _.Read() = read ()
    member _.Write(color: int) = write color

/// A shape the macro drew (Shapes.AddShape, AddLine, AddTextbox, Duplicate, Group): what Excel's Shape has
/// of it - place, size, rotation, fill, line, text, its group's members. A change to a group changes them.
type private DrawnState(book: MacroWorkbook, sheetId: SheetId, node: MacroShapes.ShapeNode) =
    let rec rootOf (n: MacroShapes.ShapeNode) = match n.Parent with Some p -> rootOf p | None -> n
    let changed () =
        let root = rootOf node
        if not (book.Shapes.ContainsKey((sheetId, root.Id))) then failApp "The object has been deleted."
        book.ShapeChanged(sheetId, root.Id)
    let leaves () = MacroShapes.all node |> Seq.filter (fun n -> n.Kind <> "group") |> List.ofSeq
    member _.Book = book
    member _.Sheet = sheetId
    member _.Node = node
    member _.Changed() = changed ()
    /// The first shape that is not a group (what a group's text and colours read as).
    member _.First() = match leaves () with n :: _ -> n | [] -> node
    member _.EachLeaf(f: MacroShapes.ShapeNode -> unit) = (for n in leaves () do f n); changed ()
    member _.ZOrderPosition() =
        let root = rootOf node
        1 + ((book.Model sheetId).Drawings |> List.findIndex (fun d -> d.Id = root.Id))
    member _.Reorder(command: int) =
        if node.Parent.IsNone then
            book.ChangeModel(sheetId, fun m ->
                let mine, others = m.Drawings |> List.partition (fun d -> d.Id = node.Id)
                let index = m.Drawings |> List.findIndex (fun d -> d.Id = node.Id)
                let next =
                    match command with
                    | 0 -> others @ mine
                    | 1 -> mine @ others
                    | 2 -> let at = min others.Length (index + 1) in List.take at others @ mine @ List.skip at others
                    | _ -> let at = max 0 (index - 1) in List.take at others @ mine @ List.skip at others
                { m with Drawings = next })
    member _.Delete() =
        match node.Parent with
        | Some parent ->
            parent.Members.Remove node |> ignore
            changed ()
        | None ->
            book.Shapes.Remove((sheetId, node.Id)) |> ignore
            book.ChangeModel(sheetId, fun m -> { m with Drawings = m.Drawings |> List.filter (fun d -> d.Id <> node.Id) })
    /// A copy at the front of the sheet, named as Excel names it ("Oval 7").
    member _.Duplicate() =
        let copy = MacroShapes.copy (fun () -> book.NextShapeId()) node
        let id = SheetOps.nextDrawingId (book.Model sheetId)
        copy.Id <- id
        let stem = Text.RegularExpressions.Regex.Replace(node.Name, " [0-9]+$", "")
        copy.Name <- sprintf "%s %d" stem id
        copy.Parent <- None
        book.Shapes.[(sheetId, id)] <- copy
        book.ShapeChanged(sheetId, id)
        book.FlushShapes()
        copy
    /// The group taken apart: its members the sheet's shapes again, by their new ids.
    member _.Ungroup() =
        if node.Kind <> "group" || node.Parent.IsSome then failApp "Ungroup method of Shape class failed"
        book.Shapes.Remove((sheetId, node.Id)) |> ignore
        book.ChangeModel(sheetId, fun m -> { m with Drawings = m.Drawings |> List.filter (fun d -> d.Id <> node.Id) })
        [ for m in List.ofSeq node.Members do
            let id = SheetOps.nextDrawingId (book.Model sheetId)
            m.Id <- id
            m.Parent <- None
            book.Shapes.[(sheetId, id)] <- m
            book.ShapeChanged(sheetId, id)
            book.FlushShapes()
            yield id ]
    member _.Flip(command: int) =
        let l, t, w, h = MacroShapes.frame node
        if command = 0 then
            for n in MacroShapes.all node do if n.Kind <> "group" then n.FlipH <- not n.FlipH
            MacroShapes.scaleAbout node (l + w / 2.0) (t + h / 2.0) -1.0 1.0
        else
            for n in MacroShapes.all node do if n.Kind <> "group" then n.FlipV <- not n.FlipV
            MacroShapes.scaleAbout node (l + w / 2.0) (t + h / 2.0) 1.0 -1.0
        // (scaleAbout by -1 leaves a negative size: the frame is put back the right way round.)
        for n in MacroShapes.all node do
            if n.Width < 0.0 then (n.Left <- n.Left + n.Width; n.Width <- -n.Width)
            if n.Height < 0.0 then (n.Top <- n.Top + n.Height; n.Height <- -n.Height)
        changed ()

/// Shapes.Range(...): several of a sheet's shapes, by id.
type private ShapeRangeState(book: MacroWorkbook, sheetId: SheetId, ids: int list) =
    member _.Book = book
    member _.Sheet = sheetId
    member _.Ids = ids

/// A chart a macro made, by its relationship id: its type, title, legend and series (each change draws it again).
type private ChartState(book: MacroWorkbook, rel: string) =
    member _.Book = book
    member _.Rel = rel
    member _.Current() = snd book.Charts.[rel]
    member _.Update(f: Rbt.Spreadsheet.ChartAuthoring.MacroChart -> Rbt.Spreadsheet.ChartAuthoring.MacroChart) =
        let sheetId, chart = book.Charts.[rel]
        let next = f chart
        book.Charts.[rel] <- (sheetId, next)
        book.ChangeModel(sheetId, fun m -> { m with AuthoredCharts = Map.add rel (Rbt.Spreadsheet.ChartAuthoring.macroChartXml next) m.AuthoredCharts })

/// A series of a chart, 1-based.
type private SeriesState(chart: ChartState, index: int) =
    member _.Chart = chart
    member _.Index = index
    member _.Pick() =
        match List.tryItem (index - 1) (chart.Current ()).Series with
        | Some s -> s
        | None -> fail 9 "Subscript out of range"
    member _.Change(f: Rbt.Spreadsheet.ChartAuthoring.MacroSeries -> Rbt.Spreadsheet.ChartAuthoring.MacroSeries) =
        chart.Update(fun c -> { c with Series = c.Series |> List.mapi (fun i s -> if i = index - 1 then f s else s) })

type private NameState(book: MacroWorkbook, name: DefinedName) =
    member _.Book = book
    member _.Name = name

// ---- the types that objects refer to before they are defined (their tables are assigned below) ----

let mutable private rangeType : Table<RangeState> = null
/// A range that stands for what the workbook has now (Selection, ActiveCell, Cells, Rows, Columns).
let mutable private liveRangeType : Table<unit -> RangeState> = null
let mutable private worksheetType : Table<WorksheetState> = null
/// A sheet that stands for what the workbook has now (ActiveSheet, Sheet1).
let mutable private liveSheetType : Table<unit -> WorksheetState> = null
let mutable private workbookType : Table<MacroWorkbook> = null
let mutable private drawnType : Table<DrawnState> = null
let mutable private chartShapeType : Table<ShapeState> = null
let mutable private shapeRangeType : Table<ShapeRangeState> = null

let private wrapRange (s: RangeState) = rangeType.Wrap s
let private rangeValue (book: MacroWorkbook) (r: RangeRef) = rangeType.Wrap (rangeState book r)
let private worksheetValue (book: MacroWorkbook) (sheetId: SheetId) = worksheetType.Wrap (WorksheetState(book, sheetId))
let private workbookValue (book: MacroWorkbook) = workbookType.Wrap book

/// The range a value stands for (a Range, Selection, ActiveCell, a union: its first area).
let private rangeStateOf (value: Value) : RangeState option =
    match rangeType.TryState value with
    | Some s -> Some s
    | None -> liveRangeType.TryState value |> Option.map (fun resolve -> resolve ())

let private rangeRefOf (value: Value) : RangeRef option = rangeStateOf value |> Option.map (fun s -> s.Ref)

/// The areas a value stands for: those of a union, else its one range.
let private areasOf (value: Value) : RangeRef list = match rangeStateOf value with Some s -> s.Areas | None -> []

let private worksheetStateOf (value: Value) : WorksheetState option =
    match worksheetType.TryState value with
    | Some s -> Some s
    | None -> liveSheetType.TryState value |> Option.map (fun resolve -> resolve ())

/// Each area of a range (a union's), as a range.
let private eachArea (s: RangeState) (f: RangeState -> unit) =
    if s.IsUnion then for a in s.Areas do f (rangeState s.Book a) else f s

/// Several areas (Union of separate blocks): its Address joins theirs ($A$1,$C$3), Areas lists them, Count
/// and For Each go over all their cells, a value or format given to it goes to each; anything else is its
/// first area's, as Excel reads a range.
let private unionValue (book: MacroWorkbook) (areas: RangeRef list) =
    match areas with
    | [ single ] -> rangeValue book single
    | [] -> VNothing
    | first :: _ -> rangeType.Wrap (RangeState(book, first, 0, areas))

let private addressOf (s: RangeState) (args: Value array) =
    if s.IsUnion then
        let rowAbs = if isMissing (at args 0) then true else toBool (at args 0)
        let colAbs = if isMissing (at args 1) then true else toBool (at args 1)
        String.Join(",", s.Areas |> List.map (fun a -> rangeAddress a rowAbs colAbs false))
    else s.Address args

/// For Each over a range: its cells; over its rows or columns, them; over a union, the cells of each area.
let private rangeItems (s: RangeState) : Value seq =
    if s.IsUnion then s.Areas |> Seq.collect (fun a -> (rangeState s.Book a).Cells () |> Seq.map wrapRange)
    elif s.Lines > 0 then seq { for i in 1 .. s.UsedLineCount () -> wrapRange (s.LineItem i) }
    else s.Cells () |> Seq.map wrapRange

/// Worksheet.Range(...): a reference ("A1:B2", a defined name), a range, or the block two corners span.
let private sheetRange (book: MacroWorkbook) (sheetId: SheetId) (args: Value array) : RangeState =
    match args with
    | [| first |] ->
        match rangeRefOf first with
        | Some r -> rangeState book { r with Sheet = sheetId }
        | None ->
            let text = toText first
            match parseReference text with
            | Some (_, a, b, c, d) -> rangeState book { Sheet = sheetId; Top = a; Left = b; Bottom = c; Right = d }
            | None ->
                // A defined name.
                match tryFindDefinedName (Some sheetId) text book.Workbook |> Option.orElse (tryFindDefinedName None text book.Workbook) with
                | Some name ->
                    match parseReference (name.RefersTo.TrimStart '=') with
                    | Some (sheetName, a, b, c, d) ->
                        let target = sheetName |> Option.bind (sheetIdNamed book) |> Option.defaultValue sheetId
                        rangeState book { Sheet = target; Top = a; Left = b; Bottom = c; Right = d }
                    | None -> failApp (sprintf "Method 'Range' of object '_Worksheet' failed (%s)" text)
                | None -> failApp (sprintf "Method 'Range' of object '_Worksheet' failed (%s)" text)
    | [| first; second |] ->
        let corner (v: Value) =
            match rangeRefOf v with
            | Some r -> r.Top, r.Left, r.Bottom, r.Right
            | None ->
                match parseReference (toText v) with
                | Some (_, a, b, c, d) -> a, b, c, d
                | None -> failApp "Method 'Range' of object '_Worksheet' failed"
        let a0, b0, c0, d0 = corner first
        let a1, b1, c1, d1 = corner second
        rangeState book { Sheet = sheetId; Top = min a0 a1; Left = min b0 b1; Bottom = max c0 c1; Right = max d0 d1 }
    | _ -> fail 450 "Wrong number of arguments or invalid property assignment"

let private usedRange (s: WorksheetState) =
    let model = s.Book.Model s.Id
    if Map.isEmpty model.Cells then rangeState s.Book { s.Whole with Bottom = 0; Right = 0 }
    else
        let addresses = model.Cells |> Map.toSeq |> Seq.map fst |> Seq.toList
        rangeState s.Book { Sheet = s.Id
                            Top = addresses |> List.map (fun a -> a.Row) |> List.min; Left = addresses |> List.map (fun a -> a.Col) |> List.min
                            Bottom = addresses |> List.map (fun a -> a.Row) |> List.max; Right = addresses |> List.map (fun a -> a.Col) |> List.max }

// ---- sheets and the workbook ----

let private sheetsItem (book: MacroWorkbook) (index: Value) =
    match scalarOf index with
    | VString name ->
        match sheetIdNamed book name with
        | Some id -> worksheetValue book id
        | None -> fail 9 "Subscript out of range"
    | v ->
        let i = toInt v
        if i < 1 || i > book.SheetOrder.Length then fail 9 "Subscript out of range"
        worksheetValue book book.SheetOrder.[i - 1]

/// Sheets.Add(Before, After): a new sheet, before or after the one given (else the active one), active.
let private addWorksheet (book: MacroWorkbook) (args: Value array) =
    let given (v: Value) = if isMissing v then None else worksheetStateOf v |> Option.map (fun s -> s.Id)
    let before = args |> Array.tryItem 0 |> Option.bind given
    let after = args |> Array.tryItem 1 |> Option.bind given
    let name = nextSheetName book.Workbook
    let mutable added = ""
    book.ChangeWorkbook(fun wb ->
        let grown = addSheet name wb
        added <- grown.SheetOrder |> List.find (fun id -> not (List.contains id wb.SheetOrder))
        let without = grown.SheetOrder |> List.filter ((<>) added)
        let at =
            match before, after with
            | Some b, _ -> List.findIndex ((=) b) without
            | None, Some a -> List.findIndex ((=) a) without + 1
            | None, None -> List.findIndex ((=) wb.ActiveSheet) without
        let order = (List.take at without) @ [ added ] @ (List.skip at without)
        reorderSheets order grown |> setActiveSheet added)
    worksheetValue book added

/// The cells the workbook has selected (the macro's last Select, else the sheet's selection), and its first.
let private activeRange (book: MacroWorkbook) =
    match book.Selection with
    | Some (sheet, top, left, bottom, right) when sheet = book.ActiveSheet -> { Sheet = sheet; Top = top; Left = left; Bottom = bottom; Right = right }
    | _ ->
        let model = book.Model book.ActiveSheet
        match model.Selection with
        | SingleCell (row, col) -> { Sheet = book.ActiveSheet; Top = row; Left = col; Bottom = row; Right = col }
        | Range (r0, c0, r1, c1) -> { Sheet = book.ActiveSheet; Top = min r0 r1; Left = min c0 c1; Bottom = max r0 r1; Right = max c0 c1 }
let private activeCell (book: MacroWorkbook) = let r = activeRange book in { r with Bottom = r.Top; Right = r.Left }
let private wholeSheet (sheetId: SheetId) = { Sheet = sheetId; Top = 0; Left = 0; Bottom = maxRow; Right = maxCol }

// ---- shapes and charts: the charts a macro makes (Shapes.AddChart, ChartObjects.Add), the shapes it draws ----

let private drawnValue (book: MacroWorkbook) (sheetId: SheetId) (node: MacroShapes.ShapeNode) = drawnType.Wrap (DrawnState(book, sheetId, node))

/// A shape of a sheet by its id: one the macro drew, else a chart's frame or a picture.
let private shapeValue (book: MacroWorkbook) (sheetId: SheetId) (id: int) =
    match book.Shapes.TryGetValue((sheetId, id)) with
    | true, node -> drawnValue book sheetId node
    | _ -> chartShapeType.Wrap (ShapeState(book, sheetId, id))

let private shapeRangeValue (book: MacroWorkbook) (sheetId: SheetId) (ids: int list) = shapeRangeType.Wrap (ShapeRangeState(book, sheetId, ids))

/// A drawing of the sheet's Shapes / ChartObjects by index or name.
let private findDrawing (s: ShapesState) (index: Value) =
    let list = s.Drawings ()
    let nameOf (d: SheetDrawing) =
        match s.Book.Shapes.TryGetValue((s.Sheet, d.Id)) with
        | true, node -> Some node.Name
        | _ -> d.Name
    match scalarOf index with
    | VString wanted ->
        match list |> List.tryFind (fun d -> nameOf d |> Option.exists (eq wanted)) with
        | Some d -> d
        | None -> failApp "The item with the specified name wasn't found."
    | v ->
        let i = toInt v
        if i < 1 || i > list.Length then failApp "The index into the specified collection is out of bounds."
        list.[i - 1]

let private shapesItem (s: ShapesState) (args: Value array) =
    if args.Length = 1 then shapeValue s.Book s.Sheet (findDrawing s args.[0]).Id
    else unsupported s.TypeName "Item"

/// Shapes.Range(index or names): the shapes it names, as a ShapeRange.
let private shapesRange (s: ShapesState) (index: Value) =
    let picked =
        match scalarOf index with
        | VArray a -> a.Data |> Array.toList |> List.map (findDrawing s)
        | _ -> [ findDrawing s index ]
    shapeRangeValue s.Book s.Sheet (picked |> List.map (fun d -> d.Id))

/// A chart, drawn by the grid's chart renderer from the series the macro gives it.
let private addChart (s: ShapesState) (chartType: int) (left: Value) (top: Value) (width: Value) (height: Value) =
    let book, sheetId = s.Book, s.Sheet
    let number v fallback = if isMissing v then fallback else toFloat v
    // (Excel's default chart: 360 x 216 points; placed here at C2 when no position is given.)
    let w, h = number width 360.0 * 4.0 / 3.0, number height 216.0 * 4.0 / 3.0
    let model = book.Model sheetId
    let cell, dx, dy =
        if isMissing left && isMissing top then { Row = 1; Col = 2 }, 0.0, 0.0
        else cellAtPoints model (number left 0.0) (number top 0.0)
    let id = SheetOps.nextDrawingId model
    let used = Collections.Generic.HashSet<string>(model.AuthoredCharts |> Map.toSeq |> Seq.map fst)
    for drawing in model.Drawings do
        match drawing.Content with
        | KeptObject (_, xml) -> for found in Text.RegularExpressions.Regex.Matches(xml, "rIdChart[0-9]+") do used.Add(found.Value) |> ignore
        | _ -> ()
    let mutable n = 1
    while used.Contains (sprintf "rIdChart%d" n) do n <- n + 1
    let rel = sprintf "rIdChart%d" n
    let name = sprintf "Chart %d" id
    let chart : Rbt.Spreadsheet.ChartAuthoring.MacroChart = { ChartType = chartType; Title = None; HasLegend = true; Series = [] }
    let anchorXml = Rbt.Spreadsheet.ChartAuthoring.standardChartAnchorXml id rel name cell.Row cell.Col w h
    let drawing = { Id = id; Anchor = SheetOps.pictureAnchor cell dx dy w h; Name = Some name; Description = None; Content = KeptObject ("chart", anchorXml) }
    book.Charts.[rel] <- (sheetId, chart)
    book.ChangeModel(sheetId, fun m -> { m with Drawings = m.Drawings @ [ drawing ]; AuthoredCharts = Map.add rel (Rbt.Spreadsheet.ChartAuthoring.macroChartXml chart) m.AuthoredCharts })
    chartShapeType.Wrap (ShapeState(book, sheetId, id))

/// A new shape at the front: numbered as the sheet's next object, named as Excel names it ("Oval 3").
let private drawShape (s: ShapesState) (node: MacroShapes.ShapeNode) (baseName: string) =
    let id = SheetOps.nextDrawingId (s.Book.Model s.Sheet)
    node.Id <- id
    if node.Name = "" then node.Name <- sprintf "%s %d" baseName id
    s.Book.Shapes.[(s.Sheet, id)] <- node
    s.Book.ShapeChanged(s.Sheet, id)
    s.Book.FlushShapes()
    drawnValue s.Book s.Sheet node

let private points (v: Value) = if isMissing v then 0.0 else toFloat v

/// AddShape(Type, Left, Top, Width, Height).
let private addShape (s: ShapesState) (args: Value array) =
    let autoShapeType = toInt (at args 0)
    let preset, baseName = MacroShapes.presetOf autoShapeType
    let node = MacroShapes.ShapeNode(0, "", "autoshape")
    node.Preset <- preset
    node.AutoShapeType <- autoShapeType
    node.Left <- points (at args 1); node.Top <- points (at args 2); node.Width <- points (at args 3); node.Height <- points (at args 4)
    drawShape s node baseName

/// AddTextbox(Orientation, Left, Top, Width, Height).
let private addTextbox (s: ShapesState) (args: Value array) =
    let node = MacroShapes.ShapeNode(0, "", "textbox")
    node.Left <- points (at args 1); node.Top <- points (at args 2); node.Width <- points (at args 3); node.Height <- points (at args 4)
    node.FillColor <- 0xFFFFFF
    node.LineColor <- 0xBFBFBF
    node.LineWeight <- 0.75
    node.TextColor <- 0
    node.TextAlign <- "l"
    node.TextAnchor <- "t"
    drawShape s node "TextBox"

/// AddLine(BeginX, BeginY, EndX, EndY).
let private addLine (s: ShapesState) (beginX: Value) (beginY: Value) (endX: Value) (endY: Value) =
    let x1, y1, x2, y2 = points beginX, points beginY, points endX, points endY
    let node = MacroShapes.ShapeNode(0, "", "line")
    node.Left <- min x1 x2; node.Top <- min y1 y2; node.Width <- abs (x2 - x1); node.Height <- abs (y2 - y1)
    node.FlipH <- x2 < x1
    node.FlipV <- y2 < y1
    node.LineColor <- 0xC47244
    node.LineWeight <- 0.75
    drawShape s node "Straight Connector"

/// The shapes of a ShapeRange, as objects.
let private shapeRangeObjects (s: ShapeRangeState) = s.Ids |> List.map (shapeValue s.Book s.Sheet)

let private firstShape (s: ShapeRangeState) =
    match shapeRangeObjects s with
    | VHostObject h :: _ -> h
    | _ -> failApp "The index into the specified collection is out of bounds."

/// ShapeRange.Group: one shape of them all (they must be shapes the macro drew).
let private groupShapes (s: ShapeRangeState) =
    let book, sheetId = s.Book, s.Sheet
    let members =
        s.Ids |> List.map (fun id ->
            match book.Shapes.TryGetValue((sheetId, id)) with
            | true, node -> Some node
            | _ -> None)
    if members |> List.exists Option.isNone || members.Length < 2 then failApp "Group method of ShapeRange class failed"
    let members = members |> List.choose id
    for m in members do book.Shapes.Remove((sheetId, m.Id)) |> ignore
    book.ChangeModel(sheetId, fun model -> { model with Drawings = model.Drawings |> List.filter (fun d -> not (List.contains d.Id s.Ids)) })
    let groupNode = MacroShapes.ShapeNode(0, "", "group")
    for m in members do
        // (Inside a group a shape is numbered apart from the sheet's objects.)
        for n in MacroShapes.all m do n.Id <- book.NextShapeId()
        m.Parent <- Some groupNode
        groupNode.Members.Add m
    let id = SheetOps.nextDrawingId (book.Model sheetId)
    groupNode.Id <- id
    groupNode.Name <- sprintf "Group %d" id
    book.Shapes.[(sheetId, id)] <- groupNode
    book.ShapeChanged(sheetId, id)
    book.FlushShapes()
    drawnValue book sheetId groupNode

/// What a chart's Values / XValues are given: a range (its formula and values) or an array (literal points).
let private chartData (book: MacroWorkbook) (value: Value) : Rbt.Spreadsheet.ChartAuthoring.MacroSeriesData =
    let point (v: Value) =
        match scalarOf v with
        | VString text ->
            match Double.TryParse(text, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
            | true, n -> text, Some n
            | _ -> text, None
        | VEmpty | VNull -> "", None
        | VError _ -> "", None
        | VBoolean _ as b -> toText b, None
        | other -> let n = toFloat other in n.ToString("R", Globalization.CultureInfo.InvariantCulture), Some n
    let ofRange (r: RangeRef) : Rbt.Spreadsheet.ChartAuthoring.MacroSeriesData =
        let points =
            [| for row in r.Top .. r.Bottom do
                 for col in r.Left .. r.Right ->
                    match book.Cell(r.Sheet, row, col) with
                    | Some c -> point (cellValue c.Value)
                    | None -> "", None |]
        { Formula = Some (absoluteReference book r); Points = points }
    match rangeRefOf value with
    | Some r -> ofRange r
    | None ->
        match scalarOf value with
        | VArray array -> { Formula = None; Points = array.Data |> Array.map point }
        | VString text when text.StartsWith "=" ->
            // A reference as text ("=Sheet1!$B$2:$B$9").
            match parseReference (text.TrimStart '=') with
            | Some (sheetName, a, b, c, d) ->
                let target = sheetName |> Option.bind (sheetIdNamed book) |> Option.defaultValue book.ActiveSheet
                { Formula = Some (text.TrimStart '='); Points = [| for row in a .. c do for col in b .. d -> match book.Cell(target, row, col) with Some x -> point (cellValue x.Value) | None -> "", None |] }
            | None -> { Formula = None; Points = [| point value |] }
        | other -> { Formula = None; Points = [| point other |] }

/// SetSourceData(range, PlotBy): a series of each column (xlColumns 2, the default for a tall range) or
/// row, the first column / row the categories when it holds no numbers.
let private setSourceData (chart: ChartState) (source: Value) (plotBy: Value) =
    let book = chart.Book
    match rangeRefOf source with
    | None -> failApp "Method 'SetSourceData' of object '_Chart' failed"
    | Some r ->
        let rows, cols = r.Bottom - r.Top + 1, r.Right - r.Left + 1
        let byColumns = if isMissing plotBy then rows >= cols else toInt plotBy = 2
        let numeric row col = match book.Cell(r.Sheet, row, col) with Some { Value = v } -> (match unwrapFormulaValue v with Number _ -> true | _ -> false) | None -> false
        let text row col = match book.Cell(r.Sheet, row, col) with Some c -> toText (cellValue c.Value) | None -> ""
        let data top left bottom right = chartData book (rangeValue book { r with Top = top; Left = left; Bottom = bottom; Right = right })
        let series =
            if byColumns then
                let header = rows > 1 && not (numeric r.Top (r.Right))
                let firstRow = if header then r.Top + 1 else r.Top
                let labels = cols > 1 && not (seq { firstRow .. r.Bottom } |> Seq.exists (fun row -> numeric row r.Left))
                let firstCol = if labels then r.Left + 1 else r.Left
                [ for col in firstCol .. r.Right ->
                    Rbt.Spreadsheet.ChartAuthoring.macroSeries (if header then Some (text r.Top col) else None)
                        (if labels then Some (data firstRow r.Left r.Bottom r.Left) else None)
                        (data firstRow col r.Bottom col) ]
            else
                let header = cols > 1 && not (numeric r.Bottom r.Left)
                let firstCol = if header then r.Left + 1 else r.Left
                let labels = rows > 1 && not (seq { firstCol .. r.Right } |> Seq.exists (fun col -> numeric r.Top col))
                let firstRow = if labels then r.Top + 1 else r.Top
                [ for row in firstRow .. r.Bottom ->
                    Rbt.Spreadsheet.ChartAuthoring.macroSeries (if header then Some (text row r.Left) else None)
                        (if labels then Some (data r.Top firstCol r.Top r.Right) else None)
                        (data row firstCol row r.Right) ]
        chart.Update(fun c -> { c with Series = series })

// ---- the formula engine, for WorksheetFunction and Evaluate ---------------------------------------------------------

/// A macro's argument as formula text: a range by its address, a value as a literal, an array as a constant.
let private formulaArgument (book: MacroWorkbook) (value: Value) =
    let literal (v: Value) =
        match scalarOf v with
        | VInteger i | VLong i -> string i
        | VByte b -> string b
        | VSingle f -> (Rbt.Vb.NumberText.singleAsDouble f).ToString("R", Globalization.CultureInfo.InvariantCulture)
        | VDouble d -> d.ToString("R", Globalization.CultureInfo.InvariantCulture)
        | VBoolean b -> if b then "TRUE" else "FALSE"
        | VString s -> "\"" + s.Replace("\"", "\"\"") + "\""
        | VDate serial -> (if book.Date1904 then serial - 1462.0 else serial).ToString("R", Globalization.CultureInfo.InvariantCulture)
        | VCurrency units -> (float units / 10000.0).ToString("R", Globalization.CultureInfo.InvariantCulture)
        | VError code -> "#" + errorTextOf code
        | VEmpty | VNull | VUndefined -> ""
        | _ -> fail 13 "Type mismatch"
    match rangeRefOf value with
    | Some r ->
        quotedSheet (book.Sheet r.Sheet).Name + "!" + cellAddress r.Top r.Left false false
        + (if r.Top = r.Bottom && r.Left = r.Right then "" else ":" + cellAddress r.Bottom r.Right false false)
    | None ->
        match scalarOf value with
        | VArray array ->
            match array.Lengths with
            | [| n |] -> "{" + String.Join(",", array.Data |> Array.map literal) + "}"
            | [| rows; cols |] ->
                "{" + String.Join(";", [| for i in 0 .. rows - 1 -> String.Join(",", [| for j in 0 .. cols - 1 -> literal array.Data.[i * cols + j] |]) |]) + "}"
            | _ -> fail 13 "Type mismatch"
        | _ -> literal value

/// The value of a formula (without its "="), evaluated on the active sheet.
let evaluate (book: MacroWorkbook) (formula: string) : Value =
    let workbook = book.Workbook
    let sheet = book.ActiveSheet
    let model = book.Model sheet
    // (An array or range result keeps its shape: a 2-D array, 1-based, as Excel hands it to VBA.)
    match Rbt.Spreadsheet.Formula.evaluateFormulaRows (Some workbook) (Some sheet) formula model { Row = 0; Col = 0 } with
    | Some [| [| single |] |] -> cellValue single
    | Some rows when rows.Length > 0 && rows.[0].Length > 0 ->
        let cols = rows.[0].Length
        array2D rows.Length cols (fun i j -> if j < rows.[i].Length then cellValue rows.[i].[j] else VError 2042)
    | _ -> cellValue (Rbt.Spreadsheet.Formula.evaluateFormulaShifted (Some workbook) (Some sheet) formula 0 0 model { Row = 0; Col = 0 })

/// A worksheet function called from a macro: FUNCTION(arguments) evaluated.
let private worksheetFunction (book: MacroWorkbook) (name: string) (args: Value array) =
    evaluate book (name.ToUpperInvariant() + "(" + String.Join(",", args |> Array.map (formulaArgument book)) + ")")

/// Evaluate("A1:B2") is that range; anything else is a formula.
let private evaluateArgument (book: MacroWorkbook) (text: string) =
    let text = text.Trim().TrimStart '='
    match parseReference text with
    | Some (sheetName, a, b, c, d) ->
        let target = sheetName |> Option.bind (sheetIdNamed book) |> Option.defaultValue book.ActiveSheet
        rangeValue book { Sheet = target; Top = a; Left = b; Bottom = c; Right = d }
    | None -> evaluate book text

let private intersect (a: RangeRef) (b: RangeRef) =
    if a.Sheet <> b.Sheet then None
    else
        let top, left, bottom, right = max a.Top b.Top, max a.Left b.Left, min a.Bottom b.Bottom, min a.Right b.Right
        if top > bottom || left > right then None else Some { a with Top = top; Left = left; Bottom = bottom; Right = right }

/// Intersect(range, range, ...): the cells they all have, or Nothing.
let private intersectOf (book: MacroWorkbook) (args: Value array) =
    let refs = args |> Array.choose rangeRefOf
    if refs.Length < 2 then failApp "Method 'Intersect' of object '_Application' failed"
    match refs |> Array.skip 1 |> Array.fold (fun acc next -> acc |> Option.bind (fun a -> intersect a next)) (Some refs.[0]) with
    | Some r -> rangeValue book r
    | None -> VNothing

/// Union(range, range, ...): separate blocks stay separate areas; blocks one contains are dropped.
let private unionOf (book: MacroWorkbook) (args: Value array) =
    let refs = args |> Array.filter (isMissing >> not) |> Array.toList |> List.collect areasOf
    if refs.IsEmpty then failApp "Method 'Union' of object '_Application' failed"
    let contains (a: RangeRef) (b: RangeRef) = a.Sheet = b.Sheet && a.Top <= b.Top && a.Left <= b.Left && a.Bottom >= b.Bottom && a.Right >= b.Right
    let kept =
        refs |> List.distinct
        |> fun all -> all |> List.filter (fun r -> not (all |> List.exists (fun other -> other <> r && contains other r)))
    unionValue book kept

/// The hooks of the VBA engine's run-time library for a run on the workbook.
let private runtimeHooksOf (book: MacroWorkbook) : Rbt.Vb.Values.RuntimeHooks =
    { Print = (fun text -> book.Hooks.Output text)
      MsgBox = (fun prompt buttons title -> book.Hooks.MessageBox prompt buttons title)
      InputBox = (fun prompt title defaultText -> book.Hooks.InputBox prompt title defaultText)
      CreateObject = (fun _ -> None)
      Environ = (fun _ -> "") }

// ---- the tables: each type's members ----------------------------------------------------------------------------

/// A table's members over another state that resolves to the table's (a global that stands for what the
/// workbook has now: ActiveSheet, Selection, Cells).
let private lift (resolve: 'S -> 'T) (m: Member<'T>) : Member<'S> =
    { Name = m.Name
      Parameters = m.Parameters
      Get = m.Get |> Option.map (fun read -> fun (s: 'S) -> read (resolve s))
      Set = m.Set |> Option.map (fun write -> fun (s: 'S) value -> write (resolve s) value)
      Call = m.Call |> Option.map (fun run -> fun (s: 'S) args -> run (resolve s) args) }

// (The Office theme's colours, as MsoThemeColorIndex numbers them.)
let private shapeThemeColor (index: int) =
    match index with
    | 1 | 13 -> Some 0x000000 | 2 | 14 -> Some 0xFFFFFF | 3 | 15 -> Some 0x546A44 | 4 | 16 -> Some 0xE1E7E7
    | 5 -> Some 0xC47244 | 6 -> Some 0x317DED | 7 -> Some 0xA5A5A5 | 8 -> Some 0x00C0FF | 9 -> Some 0xD59B5B | 10 -> Some 0x47AD70
    | _ -> None

let private colorFormatType =
    Table<ColorState>("ColorFormat", [
        prop "RGB" (fun c -> VLong (c.Read ())) (fun c v -> c.Write (toInt v))
        prop "ObjectThemeColor" (fun _ -> VLong 0) (fun c v -> match shapeThemeColor (toInt v) with Some color -> c.Write color | None -> ())
        prop "SchemeColor" (fun _ -> VLong 0) (fun c v -> match shapeThemeColor (toInt v) with Some color -> c.Write color | None -> ())
        accepted "TintAndShade" (VLong 0)
        accepted "Brightness" (VLong 0)
        get "Type" (fun _ -> VLong 1) ], defaultMember = "RGB")

let private colorValue (read: unit -> int) (write: int -> unit) = colorFormatType.Wrap (ColorState(read, write))

// ---- a range's Font and Interior ----

let private fontType =
    Table<RangeState>("Font", [
        prop "Bold" (fun s -> VBoolean (s.Format().Bold)) (fun s v -> let b = toBool v in s.ChangeFormats(fun f -> { f with Bold = b }))
        prop "Italic" (fun s -> VBoolean (s.Format().Italic)) (fun s v -> let b = toBool v in s.ChangeFormats(fun f -> { f with Italic = b }))
        prop "Underline"
            (fun s -> let f = s.Format() in VLong (if f.UnderlineDouble then -4119 elif f.Underline then 2 else -4142))
            (fun s v ->
                let u = match scalarOf v with VBoolean b -> (if b then 2 else -4142) | other -> toInt other
                s.ChangeFormats(fun f -> { f with Underline = (u = 2 || u = -4119 || u = 4 || u = 5); UnderlineDouble = (u = -4119) }))
        prop "Strikethrough" (fun s -> VBoolean (s.Format().Strikethrough)) (fun s v -> let b = toBool v in s.ChangeFormats(fun f -> { f with Strikethrough = b }))
        prop "Size" (fun s -> VDouble (s.Format().FontSize |> Option.map float |> Option.defaultValue 11.0)) (fun s v -> let size = int (Math.Round (toFloat v)) in s.ChangeFormats(fun f -> { f with FontSize = Some size }))
        prop "Name" (fun s -> VString (s.Format().FontName |> Option.defaultValue "Calibri")) (fun s v -> let name = toText v in s.ChangeFormats(fun f -> { f with FontName = Some name }))
        prop "Color" (fun s -> VDouble (float (colorToLong (s.Format().TextColor) |> Option.defaultValue 0))) (fun s v -> let color = longToColor (toInt v) in s.ChangeFormats(fun f -> { f with TextColor = Some color }))
        prop "ColorIndex" (fun s -> VLong (paletteIndex (s.Format().TextColor) -4105)) (fun s v -> let color = paletteColor (toInt v) in s.ChangeFormats(fun f -> { f with TextColor = color }))
        prop "ThemeColor" (fun _ -> VLong 0) (fun s v -> let color = themeColor (toInt v) in s.ChangeFormats(fun f -> { f with TextColor = Some color }))
        prop "TintAndShade" (fun _ -> VLong 0) (fun s v -> let amount = toFloat v in s.ChangeFormats(fun f -> { f with TextColor = f.TextColor |> Option.map (tinted amount) }))
        prop "FontStyle"
            (fun s -> let f = s.Format() in VString (match f.Bold, f.Italic with true, true -> "Bold Italic" | true, false -> "Bold" | false, true -> "Italic" | _ -> "Regular"))
            (fun s v -> let style = (toText v).ToLowerInvariant() in s.ChangeFormats(fun f -> { f with Bold = style.Contains "bold"; Italic = style.Contains "italic" })) ])

let private interiorType =
    Table<RangeState>("Interior", [
        prop "Color" (fun s -> VDouble (float (colorToLong (s.Format().BgColor) |> Option.defaultValue 16777215))) (fun s v -> let color = longToColor (toInt v) in s.ChangeFormats(fun f -> { f with BgColor = Some color }))
        prop "ColorIndex" (fun s -> VLong (paletteIndex (s.Format().BgColor) -4142)) (fun s v -> let color = paletteColor (toInt v) in s.ChangeFormats(fun f -> { f with BgColor = color }))
        prop "Pattern" (fun s -> VLong (if s.Format().BgColor.IsSome then 1 else -4142)) (fun s v -> if toInt v = -4142 then s.ChangeFormats(fun f -> { f with BgColor = None }))
        prop "ThemeColor" (fun _ -> VLong 0) (fun s v -> let color = themeColor (toInt v) in s.ChangeFormats(fun f -> { f with BgColor = Some color }))
        prop "TintAndShade" (fun _ -> VLong 0) (fun s v -> let amount = toFloat v in s.ChangeFormats(fun f -> { f with BgColor = f.BgColor |> Option.map (tinted amount) }))
        accepted "PatternTintAndShade" (VLong 0)
        accepted "PatternColorIndex" (VLong -4105)
        ignored "PatternColor"
        ignored "PatternThemeColor" ])

// ---- Font and Interior of a conditional format, of DisplayFormat ----

let private styleMembers (interior: bool) : Member<StyleState> list =
    let colorOf (f: CellFormat) = if interior then f.BgColor else f.TextColor
    let withColor (color: string option) (f: CellFormat) = if interior then { f with BgColor = color } else { f with TextColor = color }
    [ prop "Color" (fun s -> VLong (colorToLong (colorOf (s.Read ())) |> Option.defaultValue (if interior then 16777215 else 0))) (fun s v -> s.Write (withColor (Some (longToColor (toInt v)))))
      prop "ColorIndex" (fun s -> VLong (paletteIndex (colorOf (s.Read ())) (if interior then -4142 else -4105))) (fun s v -> let color = paletteColor (toInt v) in s.Write (withColor color))
      prop "ThemeColor" (fun _ -> VLong 0) (fun s v -> let color = themeColor (toInt v) in s.Write (withColor (Some color)))
      prop "TintAndShade" (fun _ -> VLong 0) (fun s v -> let amount = toFloat v in s.Write (fun f -> withColor (colorOf f |> Option.map (tinted amount)) f))
      ignored "PatternColorIndex"
      ignored "Strikethrough"
      ignored "Underline" ]

let private styleFontType =
    Table<StyleState>("Font", styleMembers false @ [
        prop "Bold" (fun s -> VBoolean (s.Read ()).Bold) (fun s v -> let b = toBool v in s.Write(fun f -> { f with Bold = b }))
        prop "Italic" (fun s -> VBoolean (s.Read ()).Italic) (fun s v -> let b = toBool v in s.Write(fun f -> { f with Italic = b })) ])

let private styleInteriorType =
    Table<StyleState>("Interior", styleMembers true @ [
        prop "Pattern" (fun s -> VLong (if (s.Read ()).BgColor.IsSome then 1 else -4142)) (fun s v -> if toInt v = -4142 then s.Write(fun f -> { f with BgColor = None })) ],
        typeNameOf = (fun s -> s.TypeName))

// ---- Borders and Border ----

let private borderMembers : Member<BorderState> list =
    let styled (current: BorderStyle option) = current |> Option.defaultValue { Style = "thin"; Color = None }
    let lineStyle =
        prop "LineStyle"
            (fun (b: BorderState) -> b.One(function Some border -> VLong (fst (lineOfStyle border.Style)) | None -> VLong xlLineStyleNone))
            (fun b v ->
                let lineStyle = match scalarOf v with VNull | VEmpty -> xlLineStyleNone | other -> toInt other
                b.Change(fun current ->
                    let weight = current |> Option.map (fun border -> snd (lineOfStyle border.Style)) |> Option.defaultValue 2
                    styleOfLine lineStyle weight |> Option.map (fun style -> { Style = style; Color = current |> Option.bind (fun border -> border.Color) })))
    [ lineStyle
      alias "Value" lineStyle
      prop "Weight"
          (fun b -> b.One(function Some border -> VLong (snd (lineOfStyle border.Style)) | None -> VLong 2))
          (fun b v ->
              let weight = toInt v
              b.Change(fun current ->
                  let lineStyle = current |> Option.map (fun border -> fst (lineOfStyle border.Style)) |> Option.defaultValue xlContinuous
                  styleOfLine lineStyle weight |> Option.map (fun style -> { Style = style; Color = current |> Option.bind (fun border -> border.Color) })))
      prop "Color"
          (fun b -> b.One(function Some { Color = Some c } -> VLong (colorToLong (Some c) |> Option.defaultValue 0) | _ -> VLong 0))
          (fun b v -> let color = longToColor (toInt v) in b.Change(fun current -> Some { styled current with Color = Some color }))
      prop "ColorIndex"
          (fun b -> b.One(function Some { Color = Some c } -> VLong (paletteIndex (Some c) -4105) | Some _ -> VLong -4105 | None -> VLong -4142))
          (fun b v -> let color = paletteColor (toInt v) in b.Change(fun current -> Some { styled current with Color = color }))
      prop "ThemeColor" (fun _ -> VLong 0) (fun b v -> let color = themeColor (toInt v) in b.Change(fun current -> Some { styled current with Color = Some color }))
      prop "TintAndShade" (fun _ -> VLong 0) (fun b v -> let amount = toFloat v in b.Change(Option.map (fun border -> { border with Color = border.Color |> Option.map (tinted amount) }))) ]

let private borderType = Table<BorderState>("Border", borderMembers)

let private bordersType =
    Table<BorderState>("Borders", borderMembers @ [
        get "Count" (fun _ -> VLong 6)
        method "Item" [ "Index" ] (fun b args -> if args.Length = 1 then borderType.Wrap (BorderState(b.Range, Some (toInt args.[0]))) else unsupported "Borders" "Item") ],
        defaultMember = "Item")

// ---- a range's conditional formats (FormatConditions), its DisplayFormat ----

let private colorScaleCriterionType =
    Table<CriterionState>("ColorScaleCriterion", [
        prop "Type"
            (fun c -> VLong (match (fst (c.Stop ())).Kind with CfvoNumber -> 0 | CfvoMin -> 1 | CfvoMax -> 2 | CfvoPercent -> 3 | CfvoFormula -> 4 | CfvoPercentile -> 5 | _ -> -1))
            (fun c v ->
                let kind = match toInt v with 0 -> CfvoNumber | 1 -> CfvoMin | 2 -> CfvoMax | 3 -> CfvoPercent | 4 -> CfvoFormula | 5 -> CfvoPercentile | _ -> CfvoMin
                c.Change(fun (t, color) -> { t with Kind = kind; Value = (if kind = CfvoMin || kind = CfvoMax then None else t.Value) }, color))
        prop "Value"
            (fun c ->
                match (fst (c.Stop ())).Value with
                | Some v -> (match Double.TryParse(v, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with | true, n -> VDouble n | _ -> VString v)
                | None -> VEmpty)
            (fun c v -> let text = (toText v).TrimStart '=' in c.Change(fun (t, color) -> { t with Value = Some text }, color))
        get "FormatColor" (fun c ->
            styleInteriorType.Wrap (StyleState("FormatColor", (fun () -> { defaultFormat with BgColor = Some (snd (c.Stop ())) }),
                                               Some (fun f -> c.Change(fun (t, color) -> t, (f { defaultFormat with BgColor = Some color }).BgColor |> Option.defaultValue color))))) ])

let private colorScaleCriteriaType =
    Table<ConditionState>("ColorScaleCriteria", [
        get "Count" (fun c -> VLong (c.Stops ()).Length)
        method "Item" [ "Index" ] (fun c args -> if args.Length = 1 then colorScaleCriterionType.Wrap (CriterionState(c, toInt args.[0])) else unsupported "ColorScaleCriteria" "Item") ],
        defaultMember = "Item",
        items = (fun c -> c.Stops () |> List.mapi (fun i _ -> colorScaleCriterionType.Wrap (CriterionState(c, i + 1))) |> Seq.ofList))

let private formatConditionType =
    Table<ConditionState>("FormatCondition", [
        get "Font" (fun c -> styleFontType.Wrap (StyleState("Font", (fun () -> (c.Entry ()).Format.Style), Some (fun f -> c.ChangeStyle f))))
        get "Interior" (fun c -> styleInteriorType.Wrap (StyleState("Interior", (fun () -> (c.Entry ()).Format.Style), Some (fun f -> c.ChangeStyle f))))
        get "Type" (fun c -> VLong (match (c.Entry ()).Format.Rule with CellIs _ -> 1 | Expression _ -> 2 | ColorScale _ -> 3 | DataBar _ | DataBarEx _ -> 4 | _ -> 1))
        get "Priority" (fun c -> VLong (c.Entry ()).Priority)
        prop "StopIfTrue" (fun c -> VBoolean (c.Entry ()).StopIfTrue) (fun c v -> let stop = toBool v in c.Update(fun e -> { e with StopIfTrue = stop }))
        get "Formula1" (fun c -> match (c.Entry ()).Format.Rule with CellIs (_, f, _) | Expression f -> VString ("=" + f) | _ -> VString "")
        method "SetFirstPriority" [] (fun c _ -> c.Range.Book.ChangeModel(c.Range.Sheet, SheetOps.moveConditionalFormat "SetFirstPriority" c.Id 1); VEmpty)
        method "Delete" [] (fun c _ -> c.Range.Book.ChangeModel(c.Range.Sheet, SheetOps.removeConditionalFormat "Delete" c.Id); VEmpty)
        get "ColorScaleCriteria" (fun c -> colorScaleCriteriaType.Wrap c) ])

let private conditionValue (s: RangeState) (id: int) = formatConditionType.Wrap (ConditionState(s, id))

/// FormatConditions.Add(Type, Operator, Formula1, Formula2): xlCellValue or xlExpression.
let private addFormatCondition (s: RangeState) (args: Value array) =
    let formula (v: Value) = (toText v).TrimStart '='
    match (if isMissing (at args 0) then 1 else toInt (at args 0)) with
    | 2 -> conditionValue s (s.AddCondition (Expression (formula (at args 2))))
    | 1 ->
        let op =
            match (if isMissing (at args 1) then 1 else toInt (at args 1)) with
            | 1 -> OpBetween | 2 -> OpNotBetween | 3 -> OpEqual | 4 -> OpNotEqual
            | 5 -> OpGreaterThan | 6 -> OpLessThan | 7 -> OpGreaterOrEqual | _ -> OpLessOrEqual
        conditionValue s (s.AddCondition (CellIs (op, formula (at args 2), (if isMissing (at args 3) then None else Some (formula (at args 3))))))
    | other -> failApp (sprintf "FormatConditions.Add of type %d is not supported" other)

let private formatConditionsType =
    Table<RangeState>("FormatConditions", [
        get "Count" (fun s -> VLong (s.Conditions ()).Length)
        method "Item" [ "Index" ] (fun s args ->
            if args.Length <> 1 then unsupported "FormatConditions" "Item"
            else
                match List.tryItem (toInt args.[0] - 1) (s.Conditions ()) with
                | Some e -> conditionValue s e.Id
                | None -> fail 9 "Subscript out of range")
        method "Add" [ "Type"; "Operator"; "Formula1"; "Formula2" ] addFormatCondition
        method "AddColorScale" [ "ColorScaleType" ] (fun s args -> conditionValue s (s.AddCondition (ColorScale (defaultColorScale (if isMissing (at args 0) then 3 else toInt args.[0])))))
        method "Delete" [] (fun s _ ->
            for e in s.Conditions () do s.Book.ChangeModel(s.Sheet, SheetOps.removeConditionalFormat "Delete" e.Id)
            VEmpty) ],
        defaultMember = "Item",
        items = (fun s -> s.Conditions () |> List.map (fun e -> conditionValue s e.Id) |> Seq.ofList))

let private displayFormatType =
    Table<RangeState>("DisplayFormat", [
        get "Interior" (fun s -> styleInteriorType.Wrap (StyleState("Interior", (fun () -> s.Displayed ()), None)))
        get "Font" (fun s -> styleFontType.Wrap (StyleState("Font", (fun () -> s.Displayed ()), None)))
        get "NumberFormat" (fun s -> VString ((s.Displayed ()).NumberFormat |> Option.defaultValue "General"))
        get "HorizontalAlignment" (fun _ -> VLong 1) ])

// ---- a cell's note (Range.Comment), a range's Areas ----

let private commentType =
    Table<RangeState>("Comment", [
        // Text(Text, Start, Overwrite): the note's text, replaced (or inserted at Start).
        indexed "Text" [ "Text"; "Start"; "Overwrite" ] (fun s -> VString (s.NoteText ())) (fun s args ->
            if not (isMissing (at args 0)) then
                let start = if isMissing (at args 1) then 0 else toInt args.[1]
                let overwrite = if isMissing (at args 2) then true else toBool args.[2]
                let current = s.NoteText ()
                let next =
                    if start <= 0 then toText args.[0]
                    else
                        let at = min current.Length (start - 1)
                        if overwrite then current.Substring(0, at) + toText args.[0] else current.Insert(at, toText args.[0])
                s.SetNote (Some { (s.Note () |> Option.defaultValue (plainNote "" None)) with Text = next; Runs = [] })
            VString (s.NoteText ()))
        get "Author" (fun s -> VString (s.Note () |> Option.bind (fun n -> n.Author) |> Option.defaultValue ""))
        accepted "Visible" (VBoolean false)
        method "Delete" [] (fun s _ -> s.SetNote None; VEmpty)
        get "Parent" (fun s -> wrapRange (s.Plain ())) ])

let private areasType =
    Table<AreasState>("Areas", [
        get "Count" (fun a -> VLong a.Areas.Length)
        // (A range's one area is the range, whatever the index.)
        method "Item" [ "Index" ] (fun a args ->
            if args.Length <> 1 then unsupported "Areas" "Item"
            else
                match a.Areas with
                | [ single ] -> rangeValue a.Book single
                | areas -> rangeValue a.Book areas.[toInt args.[0] - 1]) ],
        defaultMember = "Item",
        items = (fun a -> a.Areas |> List.map (rangeValue a.Book) |> Seq.ofList))

// ---- Range ----

let private rangeMembers : Member<RangeState> list =
    let value = prop "Value" (fun (s: RangeState) -> s.Values ()) (fun s v -> s.Enter v)
    let formula = prop "Formula" (fun (s: RangeState) -> s.Formulas ()) (fun s v -> s.Enter v)
    let numberFormat =
        prop "NumberFormat" (fun (s: RangeState) -> VString (s.Format().NumberFormat |> Option.defaultValue "General"))
            (fun s v -> let code = toText v in s.ChangeFormats(fun f -> { f with NumberFormat = (if eq code "General" then None else Some code) }))
    let worksheet = get "Worksheet" (fun (s: RangeState) -> worksheetValue s.Book s.Sheet)
    let select = method "Select" [] (fun (s: RangeState) _ -> s.Select (); VEmpty)
    let locked = "__locked"
    // A union's value or format goes to each of its areas.
    let onEveryArea (m: Member<RangeState>) =
        match m.Set with
        | Some write -> { m with Set = Some (fun s v -> if s.IsUnion then eachArea s (fun area -> write area v) else write s v) }
        | None -> m
    [ // The range read, written, called: `r(2)`, `Cells(1, 1) = 5`.
      { indexed "_Default" [ "RowIndex"; "ColumnIndex" ] (fun (s: RangeState) -> s.Values ())
            (fun s args -> if args.Length = 0 && s.Lines < 3 then s.Values () else wrapRange (s.ItemAt args))
        with Set = Some (fun s v -> s.Enter v) }
      method "Item" [ "RowIndex"; "ColumnIndex" ] (fun s args -> if args.Length = 0 then unsupported "Range" "Item" else wrapRange (s.ItemAt args))
      value
      prop "Value2" (fun s -> s.Values2 ()) (fun s v -> s.Enter v)
      formula
      alias "FormulaLocal" formula
      prop "FormulaR1C1" (fun s -> s.FormulasR1C1 ()) (fun s v -> s.EnterR1C1 v)
      get "Text" (fun s -> s.Text ())
      get "HasFormula" (fun s -> VBoolean (match s.CellAt(s.Ref.Top, s.Ref.Left) with Some { Value = Formula _ } -> true | _ -> false))
      indexed "Address" [ "RowAbsolute"; "ColumnAbsolute"; "ReferenceStyle"; "External"; "RelativeTo" ] (fun s -> VString (addressOf s [||])) (fun s args -> VString (addressOf s args))
      get "Row" (fun s -> VLong (s.Ref.Top + 1))
      get "Column" (fun s -> VLong (s.Ref.Left + 1))
      get "Count" (fun s -> s.Count ())
      get "CountLarge" (fun s -> s.CountLarge ())
      // Rows(i) / Columns(i): what Rows / Columns is called with (a row or column of the range).
      get "Rows" (fun s -> wrapRange (s.LinesOf false))
      get "Columns" (fun s -> wrapRange (s.LinesOf true))
      indexed "Cells" [ "RowIndex"; "ColumnIndex" ] (fun s -> wrapRange (s.Plain ())) (fun s args -> wrapRange (if args.Length = 0 then s.Plain () else s.CellItem args))
      method "Range" [ "Cell1"; "Cell2" ] (fun s args -> if args.Length = 0 then unsupported "Range" "Range" else wrapRange (s.Range (toText args.[0])))
      method "Offset" [ "RowOffset"; "ColumnOffset" ] (fun s args -> wrapRange (s.Offset args))
      method "Resize" [ "RowSize"; "ColumnSize" ] (fun s args -> wrapRange (s.Resize args))
      method "End" [ "Direction" ] (fun s args -> wrapRange (s.EndOf (toInt (at args 0))))
      worksheet
      alias "Parent" worksheet
      get "EntireRow" (fun s -> wrapRange (s.Sub(s.Ref.Top, 0, s.Ref.Bottom, maxCol)))
      get "EntireColumn" (fun s -> wrapRange (s.Sub(0, s.Ref.Left, maxRow, s.Ref.Right)))
      get "CurrentRegion" (fun s -> wrapRange (s.CurrentRegion ()))
      get "Areas" (fun s -> areasType.Wrap (AreasState(s.Book, s.Areas)))
      get "Next" (fun s -> wrapRange (s.Sub(s.Ref.Top, s.Ref.Left + 1, s.Ref.Top, s.Ref.Left + 1)))
      get "Previous" (fun s -> let col = max 0 (s.Ref.Left - 1) in wrapRange (s.Sub(s.Ref.Top, col, s.Ref.Top, col)))
      get "Application" (fun _ -> VString "Microsoft Excel")
      // Formats.
      numberFormat
      alias "NumberFormatLocal" numberFormat
      get "Font" (fun s -> fontType.Wrap s)
      get "Interior" (fun s -> interiorType.Wrap s)
      get "Borders" (fun s -> bordersType.Wrap (BorderState(s, None)))
      method "BorderAround" [ "LineStyle"; "Weight"; "ColorIndex"; "Color"; "ThemeColor" ] (fun s args -> s.BorderAround args; VBoolean true)
      get "FormatConditions" (fun s -> formatConditionsType.Wrap s)
      get "DisplayFormat" (fun s -> displayFormatType.Wrap s)
      prop "HorizontalAlignment" (fun s -> VLong (alignmentValue (s.Format().HorizontalAlign) true)) (fun s v -> let a = alignmentName v true in s.ChangeFormats(fun f -> { f with HorizontalAlign = a }))
      prop "VerticalAlignment" (fun s -> VLong (alignmentValue (s.Format().VerticalAlign) false)) (fun s v -> let a = alignmentName v false in s.ChangeFormats(fun f -> { f with VerticalAlign = a }))
      prop "WrapText" (fun s -> VBoolean (s.Format().WrapText)) (fun s v -> let b = toBool v in s.ChangeFormats(fun f -> { f with WrapText = b }))
      prop "IndentLevel" (fun s -> VLong (s.Format().Indent)) (fun s v -> let indent = toInt v in s.ChangeFormats(fun f -> { f with Indent = indent }))
      prop "Orientation" (fun s -> VLong (orientationOf (s.Format().TextRotation))) (fun s v -> let rotation = rotationOf v in s.ChangeFormats(fun f -> { f with TextRotation = rotation }))
      prop "MergeCells" (fun s -> VBoolean (s.MergeCells ())) (fun s v -> if toBool v then s.Merge () else s.UnMerge ())
      method "Merge" [] (fun s _ -> s.Merge (); VEmpty)
      method "UnMerge" [] (fun s _ -> s.UnMerge (); VEmpty)
      prop "ColumnWidth" (fun s -> s.ColumnWidth ()) (fun s v -> s.SetColumnWidth (toFloat v))
      prop "RowHeight" (fun s -> s.RowHeight ()) (fun s v -> s.SetRowHeight (toFloat v))
      method "AutoFit" [] (fun _ _ -> VEmpty)
      prop "Hidden" (fun s -> VBoolean (s.Hidden ())) (fun s v -> s.SetHidden (toBool v))
      get "Left" (fun s -> s.Span 0)
      get "Top" (fun s -> s.Span 1)
      get "Width" (fun s -> s.Span 2)
      get "Height" (fun s -> s.Span 3)
      // (Protection and the ID are kept for the macro to read back; the grid does not lock cells.)
      prop "ID" (fun s -> s.Book.Setting s.IdKey |> Option.defaultValue (VString "")) (fun s v -> s.Book.SetSetting(s.IdKey, VString (toText v)))
      prop "Locked" (fun s -> s.Book.Setting (locked + s.IdKey) |> Option.defaultValue (VBoolean true)) (fun s v -> s.Book.SetSetting(locked + s.IdKey, VBoolean (toBool v)))
      set "FormulaHidden" (fun s v -> s.Book.SetSetting(locked + s.IdKey, VBoolean (toBool v)))
      set "Name" (fun s v -> s.SetName (toText v))
      // Notes.
      get "Comment" (fun s -> if (s.Note ()).IsSome then commentType.Wrap s else VNothing)
      method "AddComment" [ "Text" ] (fun s args ->
          if (s.Note ()).IsSome then failApp "Application-defined or object-defined error"
          s.SetNote (Some (plainNote (if isMissing (at args 0) then "" else toText args.[0]) None))
          commentType.Wrap s)
      method "ClearComments" [] (fun s _ -> s.ClearComments (); VEmpty)
      // Contents.
      method "ClearContents" [] (fun s _ -> eachArea s (fun area -> area.ClearContents ()); VEmpty)
      method "Clear" [] (fun s _ -> eachArea s (fun area -> area.ClearContents (); area.ClearFormats ()); VEmpty)
      method "ClearFormats" [] (fun s _ -> eachArea s (fun area -> area.ClearFormats ()); VEmpty)
      method "Delete" [ "Shift" ] (fun s args -> s.Shift(false, args); VEmpty)
      method "Insert" [ "Shift"; "CopyOrigin" ] (fun s args -> s.Shift(true, args); VEmpty)
      method "Copy" [ "Destination" ] (fun s args ->
          match (if args.Length > 0 then rangeRefOf args.[0] else None) with
          | Some target -> s.CopyTo(target, -4104)
          | None -> s.Book.SetSetting("__Clipboard", wrapRange (s.Plain ()))
          VEmpty)
      method "PasteSpecial" [ "Paste"; "Operation"; "SkipBlanks"; "Transpose" ] (fun s args ->
          match s.Book.Setting "__Clipboard" |> Option.bind rangeRefOf with
          | Some source ->
              (rangeState s.Book source).CopyTo(s.Ref, (if isMissing (at args 0) then -4104 else toInt args.[0]))
              VEmpty
          | None -> failApp "PasteSpecial method of Range class failed")
      method "Find" [ "What"; "After"; "LookIn"; "LookAt"; "SearchOrder"; "SearchDirection"; "MatchCase"; "MatchByte"; "SearchFormat" ]
          (fun s args -> match s.Find args with Some found -> wrapRange found | None -> VNothing)
      method "Replace" [ "What"; "Replacement"; "LookAt"; "SearchOrder"; "MatchCase"; "MatchByte"; "SearchFormat"; "ReplaceFormat" ] (fun s args -> s.Replace args)
      select
      alias "Activate" select
      method "Calculate" [] (fun _ _ -> VEmpty) ]
    |> List.map onEveryArea

rangeType <- Table<RangeState>("Range", rangeMembers, defaultMember = "_Default", items = rangeItems)
liveRangeType <- Table<unit -> RangeState>("Range", rangeMembers |> List.map (lift (fun resolve -> resolve ())), defaultMember = "_Default", items = (fun resolve -> rangeItems (resolve ())))

// ---- a drawn shape's text, fill and line ----

/// The fill of a drawn shape's text: its colour (its other settings are accepted and change nothing).
let private textFillType =
    Table<DrawnState>("FillFormat", [ get "ForeColor" (fun s -> colorValue (fun () -> (s.First ()).TextColor) (fun c -> s.EachLeaf(fun n -> n.TextColor <- c))) ],
                      dynamic = { Read = (fun _ _ _ -> None); Write = (fun _ _ _ -> true) })

let private textFontType =
    Table<DrawnState>("Font", [
        prop "Bold" (fun s -> VBoolean (s.First ()).TextBold) (fun s v -> s.EachLeaf(fun n -> n.TextBold <- toBool v))
        prop "Italic" (fun s -> VBoolean (s.First ()).TextItalic) (fun s v -> s.EachLeaf(fun n -> n.TextItalic <- toBool v))
        prop "Size" (fun s -> VDouble (s.First ()).TextSize) (fun s v -> s.EachLeaf(fun n -> n.TextSize <- toFloat v))
        prop "Color" (fun s -> VLong (s.First ()).TextColor) (fun s v -> s.EachLeaf(fun n -> n.TextColor <- toInt v))
        accepted "Name" (VString "Calibri")
        ignored "Underline"
        ignored "ColorIndex"
        get "Fill" (fun s -> textFillType.Wrap s) ])

let private charactersType =
    let text = prop "Text" (fun (s: DrawnState) -> VString (s.First ()).Text) (fun s v -> s.EachLeaf(fun n -> n.Text <- toText v))
    Table<DrawnState>("Characters", [
        text
        alias "Caption" text
        get "Font" (fun s -> textFontType.Wrap s)
        get "Count" (fun s -> VLong (s.First ()).Text.Length) ], defaultMember = "Text")

let private textRangeType =
    Table<DrawnState>("TextRange", [
        prop "Text" (fun s -> VString (s.First ()).Text) (fun s v -> s.EachLeaf(fun n -> n.Text <- toText v))
        get "Font" (fun s -> textFontType.Wrap s)
        indexed "Characters" [ "Start"; "Length" ] (fun s -> charactersType.Wrap s) (fun s _ -> charactersType.Wrap s) ], defaultMember = "Text")

let private textFrameMembers : Member<DrawnState> list =
    let align (value: int) = match value with -4131 | 1 -> "l" | -4152 | 3 -> "r" | _ -> "ctr"
    let anchor (value: int) = match value with -4160 | 1 -> "t" | -4107 | 4 -> "b" | _ -> "ctr"
    [ indexed "Characters" [ "Start"; "Length" ] (fun s -> charactersType.Wrap s) (fun s _ -> charactersType.Wrap s)
      get "TextRange" (fun s -> textRangeType.Wrap s)
      get "HasText" (fun s -> VLong (if (s.First ()).Text <> "" then -1 else 0))
      prop "HorizontalAlignment" (fun s -> VLong (match (s.First ()).TextAlign with "l" -> -4131 | "r" -> -4152 | _ -> -4108)) (fun s v -> let a = align (toInt v) in s.EachLeaf(fun n -> n.TextAlign <- a))
      prop "VerticalAlignment" (fun s -> VLong (match (s.First ()).TextAnchor with "t" -> -4160 | "b" -> -4107 | _ -> -4108)) (fun s v -> let a = anchor (toInt v) in s.EachLeaf(fun n -> n.TextAnchor <- a))
      set "VerticalAnchor" (fun s v -> let a = anchor (toInt v) in s.EachLeaf(fun n -> n.TextAnchor <- a))
      ignored "AutoSize"
      ignored "WordWrap"
      ignored "Orientation"
      ignored "MarginLeft"
      ignored "MarginRight"
      ignored "MarginTop"
      ignored "MarginBottom" ]

let private textFrameType = Table<DrawnState>("TextFrame", textFrameMembers)
let private textFrame2Type = Table<DrawnState>("TextFrame2", textFrameMembers)

/// A drawn shape's FillFormat or LineFormat.
let private fillOrLineMembers (isLine: bool) : Member<DrawnState> list =
    let foreColor =
        get "ForeColor" (fun (s: DrawnState) ->
            colorValue (fun () -> let n = s.First () in if isLine then n.LineColor else n.FillColor)
                       (fun c -> s.EachLeaf(fun n -> if isLine then (n.LineColor <- c; n.LineVisible <- true) else (n.FillColor <- c; n.FillVisible <- true))))
    [ foreColor
      alias "BackColor" foreColor
      prop "Visible" (fun s -> let n = s.First () in VLong (if (if isLine then n.LineVisible else n.FillVisible) then -1 else 0))
          (fun s v -> s.EachLeaf(fun n -> if isLine then n.LineVisible <- toBool v else n.FillVisible <- toBool v))
      prop "Transparency" (fun s -> VDouble (if isLine then 0.0 else (s.First ()).FillTransparency)) (fun s v -> s.EachLeaf(fun n -> if not isLine then n.FillTransparency <- toFloat v))
      method "Solid" [] (fun s _ -> s.EachLeaf(fun n -> if not isLine then n.FillVisible <- true); VEmpty)
      accepted "DashStyle" (VLong 1)
      accepted "Style" (VLong 1)
      ignored "BeginArrowheadStyle"
      ignored "EndArrowheadStyle" ]

let private fillFormatType = Table<DrawnState>("FillFormat", fillOrLineMembers false)
let private lineFormatType =
    Table<DrawnState>("LineFormat", fillOrLineMembers true @ [
        prop "Weight" (fun s -> VDouble (s.First ()).LineWeight) (fun s v -> s.EachLeaf(fun n -> n.LineWeight <- toFloat v)) ])

let private groupShapesType =
    Table<DrawnState>("GroupShapes", [
        get "Count" (fun s -> VLong s.Node.Members.Count)
        method "Item" [ "Index" ] (fun s args ->
            if args.Length <> 1 then unsupported "GroupShapes" "Item"
            else
                match scalarOf args.[0] with
                | VString wanted ->
                    match s.Node.Members |> Seq.tryFind (fun m -> eq m.Name wanted) with
                    | Some m -> drawnValue s.Book s.Sheet m
                    | None -> failApp "The item with the specified name wasn't found."
                | v ->
                    let i = toInt v
                    if i < 1 || i > s.Node.Members.Count then failApp "The index into the specified collection is out of bounds."
                    drawnValue s.Book s.Sheet s.Node.Members.[i - 1]) ],
        defaultMember = "Item",
        items = (fun s -> s.Node.Members |> Seq.map (drawnValue s.Book s.Sheet) |> List.ofSeq |> Seq.ofList))

// ---- shapes ----

let private frameOf (pick: float * float * float * float -> float) (s: DrawnState) = VDouble (pick (MacroShapes.frame s.Node))

drawnType <-
    let select = method "Select" [] (fun (_: DrawnState) _ -> VEmpty)
    let changing (change: DrawnState -> Value -> unit) (s: DrawnState) (v: Value) = change s v; s.Changed ()
    Table<DrawnState>("Shape", [
        prop "Name" (fun s -> VString s.Node.Name) (changing (fun s v -> s.Node.Name <- toText v))
        prop "Left" (frameOf (fun (l, _, _, _) -> l)) (changing (fun s v -> MacroShapes.setLeft s.Node (toFloat v)))
        prop "Top" (frameOf (fun (_, t, _, _) -> t)) (changing (fun s v -> MacroShapes.setTop s.Node (toFloat v)))
        prop "Width" (frameOf (fun (_, _, w, _) -> w)) (changing (fun s v -> MacroShapes.setSize s.Node (Some (toFloat v)) None))
        prop "Height" (frameOf (fun (_, _, _, h) -> h)) (changing (fun s v -> MacroShapes.setSize s.Node None (Some (toFloat v))))
        prop "Rotation" (fun s -> VDouble s.Node.Rotation) (changing (fun s v -> MacroShapes.setRotation s.Node (toFloat v)))
        prop "Visible" (fun s -> VLong (if s.Node.Visible then -1 else 0)) (changing (fun s v -> s.Node.Visible <- toBool v))
        prop "OnAction" (fun s -> VString s.Node.OnAction) (changing (fun s v -> s.Node.OnAction <- toText v))
        get "ID" (fun s -> VLong s.Node.Id)
        get "Type" (fun s -> VLong (match s.Node.Kind with "group" -> 6 | "line" -> 9 | "textbox" -> 17 | _ -> 1))
        prop "AutoShapeType" (fun s -> VLong (if s.Node.Kind = "autoshape" then s.Node.AutoShapeType else -2))
            (fun s v -> let kind = toInt v in let preset, _ = MacroShapes.presetOf kind in s.EachLeaf(fun n -> n.AutoShapeType <- kind; n.Preset <- preset))
        get "HorizontalFlip" (fun s -> VLong (if s.Node.FlipH then -1 else 0))
        get "VerticalFlip" (fun s -> VLong (if s.Node.FlipV then -1 else 0))
        get "HasChart" (fun _ -> VBoolean false)
        get "Fill" (fun s -> fillFormatType.Wrap s)
        get "Line" (fun s -> lineFormatType.Wrap s)
        get "TextFrame" (fun s -> textFrameType.Wrap s)
        get "TextFrame2" (fun s -> textFrame2Type.Wrap s)
        get "ZOrderPosition" (fun s -> VLong (s.ZOrderPosition ()))
        get "GroupItems" (fun s -> if s.Node.Kind <> "group" then failApp "This member can only be accessed for a group." else groupShapesType.Wrap s)
        get "ParentGroup" (fun s ->
            match s.Node.Parent with
            | Some p -> drawnValue s.Book s.Sheet p
            | None -> failApp "This member can only be accessed for a group.")
        method "Duplicate" [] (fun s _ -> drawnValue s.Book s.Sheet (s.Duplicate ()))
        method "Delete" [] (fun s _ -> s.Delete (); VEmpty)
        method "Ungroup" [] (fun s _ -> shapeRangeValue s.Book s.Sheet (s.Ungroup ()))
        select
        alias "Activate" select
        accepted "LockAspectRatio" (VLong 0)
        accepted "Shadow" (VLong 0)
        ignored "Placement"
        ignored "Locked"
        method "IncrementLeft" [ "Increment" ] (fun s args -> MacroShapes.translate s.Node (toFloat (at args 0)) 0.0; s.Changed (); VEmpty)
        method "IncrementTop" [ "Increment" ] (fun s args -> MacroShapes.translate s.Node 0.0 (toFloat (at args 0)); s.Changed (); VEmpty)
        method "IncrementRotation" [ "Increment" ] (fun s args -> MacroShapes.setRotation s.Node (s.Node.Rotation + toFloat (at args 0)); s.Changed (); VEmpty)
        method "ScaleWidth" [ "Factor"; "RelativeToOriginalSize"; "Scale" ] (fun s args ->
            MacroShapes.scale s.Node (toFloat (at args 0)) (if isMissing (at args 2) then 0 else toInt args.[2]) true
            s.Changed ()
            VEmpty)
        method "ScaleHeight" [ "Factor"; "RelativeToOriginalSize"; "Scale" ] (fun s args ->
            MacroShapes.scale s.Node (toFloat (at args 0)) (if isMissing (at args 2) then 0 else toInt args.[2]) false
            s.Changed ()
            VEmpty)
        method "Flip" [ "FlipCmd" ] (fun s args -> s.Flip (toInt (at args 0)); VEmpty)
        method "ZOrder" [ "ZOrderCmd" ] (fun s args -> s.Reorder (toInt (at args 0)); VEmpty) ])

// ---- charts ----

let private seriesType =
    let points (d: Rbt.Spreadsheet.ChartAuthoring.MacroSeriesData) =
        VArray { Data = d.Points |> Array.map (fun (text, n) -> match n with Some x -> VDouble x | None -> VString text); LowerBounds = [| 1 |]; Lengths = [| d.Points.Length |] }
    Table<SeriesState>("Series", [
        prop "Values" (fun (s: SeriesState) -> points (s.Pick ()).YData) (fun s v -> let data = chartData s.Chart.Book v in s.Change(fun series -> { series with YData = data }))
        prop "XValues" (fun s -> (s.Pick ()).XData |> Option.map points |> Option.defaultValue (VArray { Data = [||]; LowerBounds = [| 1 |]; Lengths = [| 0 |] }))
            (fun s v -> let data = chartData s.Chart.Book v in s.Change(fun series -> { series with XData = Some data }))
        prop "Name" (fun s -> VString ((s.Pick ()).SeriesName |> Option.defaultValue (sprintf "Series%d" s.Index))) (fun s v -> let name = toText v in s.Change(fun series -> { series with SeriesName = Some name }))
        method "Delete" [] (fun s _ ->
            s.Chart.Update(fun chart -> { chart with Series = chart.Series |> List.indexed |> List.filter (fun (i, _) -> i <> s.Index - 1) |> List.map snd })
            VEmpty)
        ignored "ChartType"
        ignored "MarkerStyle"
        ignored "MarkerSize"
        ignored "Smooth" ])

let private seriesCollectionType =
    Table<ChartState>("SeriesCollection", [
        get "Count" (fun c -> VLong (c.Current ()).Series.Length)
        method "NewSeries" [] (fun c _ ->
            c.Update(fun chart -> { chart with Series = chart.Series @ [ Rbt.Spreadsheet.ChartAuthoring.macroSeries None None { Formula = None; Points = [||] } ] })
            seriesType.Wrap (SeriesState(c, (c.Current ()).Series.Length)))
        method "Item" [ "Index" ] (fun c args ->
            if args.Length <> 1 then unsupported "SeriesCollection" "Item"
            else
                match scalarOf args.[0] with
                | VString text ->
                    match (c.Current ()).Series |> List.tryFindIndex (fun s -> s.SeriesName |> Option.exists (eq text)) with
                    | Some i -> seriesType.Wrap (SeriesState(c, i + 1))
                    | None -> fail 1004 "Invalid parameter"
                | v -> seriesType.Wrap (SeriesState(c, toInt v))) ],
        defaultMember = "Item",
        items = (fun c -> (c.Current ()).Series |> List.mapi (fun i _ -> seriesType.Wrap (SeriesState(c, i + 1))) |> Seq.ofList))

let private chartTitleType =
    let text = prop "Text" (fun (c: ChartState) -> VString ((c.Current ()).Title |> Option.defaultValue "")) (fun c v -> let title = toText v in c.Update(fun chart -> { chart with Title = Some title }))
    Table<ChartState>("ChartTitle", [
        text
        alias "Caption" text ],
        // (Its other settings - Font, Format - are accepted and change nothing.)
        dynamic = { Read = (fun _ _ _ -> None); Write = (fun _ _ _ -> true) })

let private chartType =
    let refresh = method "Refresh" [] (fun (_: ChartState) _ -> VEmpty)
    Table<ChartState>("Chart", [
        prop "ChartType" (fun c -> VLong (c.Current ()).ChartType) (fun c v -> let kind = toInt v in c.Update(fun chart -> { chart with ChartType = kind }))
        prop "HasLegend" (fun c -> VBoolean (c.Current ()).HasLegend) (fun c v -> let legend = toBool v in c.Update(fun chart -> { chart with HasLegend = legend }))
        prop "HasTitle" (fun c -> VBoolean (c.Current ()).Title.IsSome)
            (fun c v -> let title = toBool v in c.Update(fun chart -> { chart with Title = (if title then Some (chart.Title |> Option.defaultValue "Chart Title") else None) }))
        get "ChartTitle" (fun c -> chartTitleType.Wrap c)
        get "SeriesCollection" (fun c -> seriesCollectionType.Wrap c)
        get "Name" (fun c -> VString c.Rel)
        method "SetSourceData" [ "Source"; "PlotBy" ] (fun c args -> setSourceData c (at args 0) (at args 1); VEmpty)
        refresh
        alias "Select" refresh
        alias "Activate" refresh
        ignored "ChartStyle"
        ignored "PlotBy" ])

chartShapeType <-
    let select = method "Select" [] (fun (s: ShapeState) _ -> s.Drawing () |> ignore; VEmpty)
    Table<ShapeState>("Shape", [
        prop "Name" (fun s -> VString ((s.Drawing ()).Name |> Option.defaultValue (sprintf "Object %d" s.Id))) (fun s v -> let name = toText v in s.Change(fun d -> { d with Name = Some name }))
        prop "Width" (fun s -> VDouble (((s.Drawing ()).Anchor.Width |> Option.defaultValue 0.0) * 0.75))
            (fun s v -> let width = toFloat v * 4.0 / 3.0 in s.Change(fun d -> { d with Anchor = { d.Anchor with Width = Some width } }))
        prop "Height" (fun s -> VDouble (((s.Drawing ()).Anchor.Height |> Option.defaultValue 0.0) * 0.75))
            (fun s v -> let height = toFloat v * 4.0 / 3.0 in s.Change(fun d -> { d with Anchor = { d.Anchor with Height = Some height } }))
        get "Chart" (fun s ->
            match s.Rel (s.Drawing ()) with
            | Some rel when s.Book.Charts.ContainsKey rel -> chartType.Wrap (ChartState(s.Book, rel))
            | _ -> failApp "This object has no chart the macro can change")
        get "HasChart" (fun s -> VBoolean (s.Rel (s.Drawing ())).IsSome)
        prop "Visible" (fun s -> s.Drawing () |> ignore; VLong -1) (fun _ _ -> ())
        ignored "Placement"
        get "ID" (fun s -> s.Drawing () |> ignore; VLong s.Id)
        method "Delete" [] (fun s _ -> s.Delete (); VEmpty)
        select
        alias "Activate" select ])

shapeRangeType <-
    Table<ShapeRangeState>("ShapeRange", [
        get "Count" (fun s -> VLong s.Ids.Length)
        method "Group" [] (fun s _ -> groupShapes s)
        method "Delete" [] (fun s _ ->
            for o in shapeRangeObjects s do
                match o with VHostObject h -> h.CallMethod "Delete" [||] |> ignore | _ -> ()
            VEmpty)
        method "Select" [] (fun _ _ -> VEmpty)
        method "Item" [ "Index" ] (fun s args ->
            if args.Length <> 1 then unsupported "ShapeRange" "Item"
            else
                let i = toInt args.[0]
                let all = shapeRangeObjects s
                if i < 1 || i > all.Length then failApp "The index into the specified collection is out of bounds."
                all.[i - 1]) ],
        defaultMember = "Item",
        items = (fun s -> shapeRangeObjects s |> Seq.ofList),
        // Anything else is its shapes': read of the first, set on each, a method called on each.
        dynamic =
            { Read = (fun s name args ->
                match args with
                | Some args when args.Length > 0 ->
                    let mutable result = None
                    for o in shapeRangeObjects s do
                        match o with VHostObject h -> result <- h.CallMethod name args | _ -> ()
                    result
                | _ -> (firstShape s).GetMember name)
              Write = (fun s name value ->
                for o in shapeRangeObjects s do
                    match o with VHostObject h -> h.SetMember name value |> ignore | _ -> ()
                true) })

/// Shapes and ChartObjects: by index or name, Count; AddChart / AddChart2 / Add (a chart), AddShape,
/// AddTextbox, AddLine (a shape the macro draws).
let private shapesMembers : Member<ShapesState> list =
    [ get "Count" (fun s -> VLong (s.Drawings ()).Length)
      method "Item" [ "Index" ] shapesItem
      method "Range" [ "Index" ] (fun s args -> if args.Length = 1 then shapesRange s args.[0] else unsupported s.TypeName "Range")
      method "SelectAll" [] (fun _ _ -> VEmpty) ]

let private shapesType =
    let textbox = method "AddTextbox" [ "Orientation"; "Left"; "Top"; "Width"; "Height" ] addTextbox
    Table<ShapesState>("Shapes", shapesMembers @ [
        method "AddChart" [ "Type"; "Left"; "Top"; "Width"; "Height" ] (fun s args -> addChart s (if isMissing (at args 0) then 51 else toInt args.[0]) (at args 1) (at args 2) (at args 3) (at args 4))
        method "AddChart2" [ "Style"; "XlChartType"; "Left"; "Top"; "Width"; "Height"; "NewLayout" ] (fun s args -> addChart s (if isMissing (at args 1) then 51 else toInt args.[1]) (at args 2) (at args 3) (at args 4) (at args 5))
        method "AddShape" [ "Type"; "Left"; "Top"; "Width"; "Height" ] addShape
        textbox
        alias "AddLabel" textbox
        method "AddLine" [ "BeginX"; "BeginY"; "EndX"; "EndY" ] (fun s args -> addLine s (at args 0) (at args 1) (at args 2) (at args 3))
        method "AddConnector" [ "Type"; "BeginX"; "BeginY"; "EndX"; "EndY" ] (fun s args -> addLine s (at args 1) (at args 2) (at args 3) (at args 4)) ],
        defaultMember = "Item",
        items = (fun s -> s.Drawings () |> List.map (fun d -> shapeValue s.Book s.Sheet d.Id) |> Seq.ofList))

let private chartObjectsType =
    Table<ShapesState>("ChartObjects", shapesMembers @ [
        // ChartObjects.Add(Left, Top, Width, Height).
        method "Add" [ "Left"; "Top"; "Width"; "Height" ] (fun s args -> addChart s 51 (at args 0) (at args 1) (at args 2) (at args 3)) ],
        defaultMember = "Item",
        items = (fun s -> s.Drawings () |> List.map (fun d -> shapeValue s.Book s.Sheet d.Id) |> Seq.ofList))

// ---- Worksheet ----

let private worksheetMembers : Member<WorksheetState> list =
    let activate = method "Activate" [] (fun (s: WorksheetState) _ -> s.Book.SetActiveSheet s.Id; VEmpty)
    let protectedKey (s: WorksheetState) = "__protected\t" + s.Id
    let protection = get "ProtectContents" (fun (s: WorksheetState) -> s.Book.Setting (protectedKey s) |> Option.defaultValue (VBoolean false))
    [ prop "Name" (fun s -> VString (s.Book.Sheet s.Id).Name) (fun s v ->
          let newName = toText v
          match sheetNameProblem newName with
          | Some problem -> failApp problem
          | None ->
              if sheetNameTaken (Some s.Id) newName s.Book.Workbook then failApp "That name is already taken. Try a different one."
              s.Book.ChangeWorkbook(renameSheetEverywhere s.Id newName))
      get "CodeName" (fun s -> VString (s.Book.Sheet s.Id).Name)
      get "Index" (fun s -> VLong (s.Order + 1))
      indexed "Cells" [ "RowIndex"; "ColumnIndex" ] (fun s -> rangeValue s.Book s.Whole) (fun s args ->
          let whole = rangeState s.Book s.Whole
          wrapRange (if args.Length = 0 then whole else whole.CellItem args))
      // Rows(i) / Columns(i): what Rows / Columns is called with ("3:5", "A:C" and a column's letters too).
      get "Rows" (fun s -> wrapRange (RangeState(s.Book, s.Whole, 3, [])))
      get "Columns" (fun s -> wrapRange (RangeState(s.Book, s.Whole, 4, [])))
      method "Range" [ "Cell1"; "Cell2" ] (fun s args -> wrapRange (sheetRange s.Book s.Id args))
      get "UsedRange" (fun s -> wrapRange (usedRange s))
      get "Shapes" (fun s -> shapesType.Wrap (ShapesState(s.Book, s.Id, false)))
      get "ChartObjects" (fun s -> chartObjectsType.Wrap (ShapesState(s.Book, s.Id, true)))
      prop "Visible" (fun s -> VLong (if (s.Book.Sheet s.Id).Visible then -1 elif (s.Book.Model s.Id).VeryHidden then 2 else 0)) (fun s v ->
          let visible = match scalarOf v with VBoolean b -> (if b then -1 else 0) | other -> toInt other
          s.Book.ChangeWorkbook(fun wb ->
              match Map.tryFind s.Id wb.Sheets with
              | Some sheet -> { wb with Sheets = Map.add s.Id { sheet with Visible = (visible = -1); Model = { sheet.Model with VeryHidden = (visible = 2) } } wb.Sheets }
              | None -> wb))
      accepted "DisplayPageBreaks" (VBoolean false)
      accepted "EnableCalculation" (VBoolean true)
      ignored "EnableSelection"
      ignored "ScrollArea"
      get "Parent" (fun s -> workbookValue s.Book)
      get "Type" (fun _ -> VLong -4167)
      get "Next" (fun s -> let order = s.Book.SheetOrder in let at = s.Order in if at + 1 < order.Length then worksheetValue s.Book order.[at + 1] else VNothing)
      get "Previous" (fun s -> let order = s.Book.SheetOrder in let at = s.Order in if at > 0 then worksheetValue s.Book order.[at - 1] else VNothing)
      activate
      alias "Select" activate
      method "Calculate" [] (fun _ _ -> VEmpty)
      // (Protection is kept for the macro to read back; the grid does not lock the sheet.)
      method "Protect" [] (fun s _ -> s.Book.SetSetting(protectedKey s, VBoolean true); VEmpty)
      method "Unprotect" [] (fun s _ -> s.Book.SetSetting(protectedKey s, VBoolean false); VEmpty)
      protection
      alias "ProtectDrawingObjects" protection
      alias "ProtectScenarios" protection
      method "Delete" [] (fun s _ ->
          match deleteSheet s.Id s.Book.Workbook with
          | Some _ -> s.Book.ChangeWorkbook(fun wb -> deleteSheet s.Id wb |> Option.defaultValue wb); VEmpty
          | None -> failApp "Delete method of Worksheet class failed") ]

worksheetType <- Table<WorksheetState>("Worksheet", worksheetMembers)
liveSheetType <- Table<unit -> WorksheetState>("Worksheet", worksheetMembers |> List.map (lift (fun resolve -> resolve ())))

// ---- the workbook: Sheets, Names ----

let private sheetsType =
    Table<MacroWorkbook>("Sheets", [
        get "Count" (fun b -> VLong b.SheetOrder.Length)
        method "Item" [ "Index" ] (fun b args -> if args.Length = 1 then sheetsItem b args.[0] else unsupported "Sheets" "Item")
        method "Add" [ "Before"; "After"; "Count"; "Type" ] addWorksheet ],
        defaultMember = "Item",
        items = (fun b -> b.SheetOrder |> List.map (worksheetValue b) |> Seq.ofList))

let private nameType =
    let refersTo = prop "RefersTo" (fun (n: NameState) -> VString ("=" + n.Name.RefersTo.TrimStart '=')) (fun n v -> n.Book.ChangeWorkbook(setDefinedName n.Name.Scope n.Name.Name ((toText v).TrimStart '=') n.Name.Comment))
    Table<NameState>("Name", [
        get "Name" (fun n -> VString n.Name.Name)
        refersTo
        { refersTo with Name = "Value"; Set = None }
        get "RefersToRange" (fun n ->
            match parseReference (n.Name.RefersTo.TrimStart '=') with
            | Some (sheetName, a, b, c, d) ->
                let target = sheetName |> Option.bind (sheetIdNamed n.Book) |> Option.defaultValue n.Book.ActiveSheet
                rangeValue n.Book { Sheet = target; Top = a; Left = b; Bottom = c; Right = d }
            | None -> failApp "Application-defined or object-defined error")
        method "Delete" [] (fun n _ -> n.Book.ChangeWorkbook(fun wb -> deleteDefinedName n.Name.Scope n.Name.Name wb |> Option.defaultValue wb); VEmpty) ],
        defaultMember = "Value")

let private nameValue (book: MacroWorkbook) (name: DefinedName) = nameType.Wrap (NameState(book, name))

let private namesType =
    Table<MacroWorkbook>("Names", [
        get "Count" (fun b -> VLong (definedNames b.Workbook).Length)
        method "Item" [ "Index" ] (fun b args ->
            if args.Length = 0 then unsupported "Names" "Item"
            else
                match scalarOf args.[0] with
                | VString text ->
                    match definedNames b.Workbook |> List.tryFind (fun x -> eq x.Name text) with
                    | Some found -> nameValue b found
                    | None -> fail 1004 "Application-defined or object-defined error"
                | v ->
                    match definedNames b.Workbook |> List.tryItem (toInt v - 1) with
                    | Some found -> nameValue b found
                    | None -> fail 9 "Subscript out of range")
        method "Add" [ "Name"; "RefersTo"; "Visible"; "MacroType"; "ShortcutKey"; "Category"; "NameLocal"; "RefersToLocal" ] (fun b args ->
            if args.Length < 2 then unsupported "Names" "Add"
            else
                let name = toText args.[0]
                let refersTo =
                    match rangeRefOf args.[1] with
                    | Some r -> absoluteReference b r
                    | None -> (toText args.[1]).TrimStart '='
                match definedNameProblem name with
                | Some problem -> failApp problem
                | None ->
                    b.ChangeWorkbook(setDefinedName None name refersTo None)
                    nameValue b (tryFindDefinedName None name b.Workbook |> Option.get)) ],
        defaultMember = "Item",
        items = (fun b -> definedNames b.Workbook |> List.map (nameValue b) |> Seq.ofList))

workbookType <-
    let name = get "Name" (fun (b: MacroWorkbook) -> VString b.Hooks.FileName)
    let sheets = get "Worksheets" (fun (b: MacroWorkbook) -> sheetsType.Wrap b)
    let save = method "Save" [] (fun (_: MacroWorkbook) _ -> VEmpty)
    Table<MacroWorkbook>("Workbook", [
        name
        alias "FullName" name
        get "Path" (fun _ -> VString "")
        sheets
        alias "Sheets" sheets
        get "ActiveSheet" (fun b -> worksheetValue b b.ActiveSheet)
        get "Names" (fun b -> namesType.Wrap b)
        accepted "Saved" (VBoolean false)
        save
        alias "Activate" save ])

// ---- Application and the globals ----

let private workbooksType =
    Table<MacroWorkbook>("Workbooks", [
        get "Count" (fun _ -> VLong 1)
        method "Item" [ "Index" ] (fun b args -> if args.Length = 1 then workbookValue b else unsupported "Workbooks" "Item") ],
        defaultMember = "Item",
        items = (fun b -> Seq.singleton (workbookValue b)))

let private worksheetFunctionType =
    // Any worksheet function: WorksheetFunction.Sum(...); one without arguments reads as a member (Pi).
    Table<MacroWorkbook>("WorksheetFunction", [],
        dynamic =
            { Read = (fun b name args ->
                match args with
                | None ->
                    if name = "" then unsupported "WorksheetFunction" name
                    else
                        match evaluate b (name.ToUpperInvariant() + "()") with
                        | VError 2029 -> None
                        | VError _ -> failApp (sprintf "Unable to get the %s property of the WorksheetFunction class" name)
                        | value -> Some value
                | Some args ->
                    match worksheetFunction b name args with
                    | VError _ -> failApp (sprintf "Unable to get the %s property of the WorksheetFunction class" name)
                    | value -> Some value)
              Write = (fun _ _ _ -> false) })

/// The window of the workbook (ActiveWindow): gridlines, headings, zoom, frozen panes and scrolling of the active sheet.
let private windowType =
    let model (b: MacroWorkbook) = b.Model b.ActiveSheet
    let change (b: MacroWorkbook) f = b.ChangeModel(b.ActiveSheet, f)
    let display (name: string) (read: SpreadsheetModel -> bool) (write: bool -> SpreadsheetModel -> SpreadsheetModel) =
        prop name (fun b -> VBoolean (read (model b))) (fun b v -> let on = toBool v in change b (write on))
    // (Settings kept for the macro to read back.)
    let kept (name: string) (key: string) (fallback: Value) = prop name (fun (b: MacroWorkbook) -> b.Setting key |> Option.defaultValue fallback) (fun b v -> b.SetSetting(key, v))
    let remembered (name: string) = set name (fun (b: MacroWorkbook) v -> b.SetSetting(name, v))
    let scroll = method "SmallScroll" [] (fun (_: MacroWorkbook) _ -> VEmpty)
    Table<MacroWorkbook>("Window", [
        display "DisplayGridlines" (fun m -> m.ShowGridLines) (fun on m -> { m with ShowGridLines = on })
        display "DisplayHeadings" (fun m -> m.ShowHeaders) (fun on m -> { m with ShowHeaders = on })
        display "DisplayFormulas" (fun m -> m.ShowFormulas) (fun on m -> { m with ShowFormulas = on })
        display "DisplayZeros" (fun m -> m.ShowZeros) (fun on m -> { m with ShowZeros = on })
        prop "Zoom" (fun b -> VDouble (Math.Round((model b).Zoom * 100.0))) (fun b v ->
            match scalarOf v with
            | VBoolean _ -> ()
            | other -> let z = toFloat other in change b (fun m -> { m with Zoom = (max 10.0 (min 400.0 z)) / 100.0 }))
        prop "FreezePanes" (fun b -> VBoolean ((model b).FrozenRows > 0 || (model b).FrozenCols > 0)) (fun b v ->
            if toBool v then
                // Above and to the left of the active cell, as Excel freezes them.
                let row, col =
                    match b.Selection with
                    | Some (sheet, top, left, _, _) when sheet = b.ActiveSheet -> top, left
                    | _ -> match (model b).Selection with SingleCell (r, c) -> r, c | Range (r0, c0, _, _) -> r0, c0
                change b (fun m -> { m with FrozenRows = row; FrozenCols = col })
            else change b (fun m -> { m with FrozenRows = 0; FrozenCols = 0 }))
        prop "SplitRow" (fun b -> VLong (model b).FrozenRows) (fun b v -> let rows = toInt v in change b (fun m -> { m with FrozenRows = rows }))
        prop "SplitColumn" (fun b -> VLong (model b).FrozenCols) (fun b v -> let cols = toInt v in change b (fun m -> { m with FrozenCols = cols }))
        kept "ScrollRow" "WindowScrollRow" (VLong 1)
        kept "ScrollColumn" "WindowScrollColumn" (VLong 1)
        kept "WindowState" "WindowState" (VLong -4137)
        prop "Caption" (fun b -> VString b.Hooks.FileName) (fun b v -> b.SetSetting("Caption", v))
        remembered "Split"
        remembered "DisplayWorkbookTabs"
        remembered "DisplayHorizontalScrollBar"
        remembered "DisplayVerticalScrollBar"
        get "ActiveSheet" (fun b -> worksheetValue b b.ActiveSheet)
        scroll
        alias "LargeScroll" scroll
        alias "Activate" scroll ])

/// Application's settings kept for the macro to read back (ScreenUpdating is also whether a run shows its frames).
let private applicationSettings =
    [ "ScreenUpdating"; "EnableEvents"; "DisplayAlerts"; "StatusBar"; "CutCopyMode"; "DisplayStatusBar"; "Interactive"; "EnableCancelKey"; "AskToUpdateLinks" ]
    |> List.map (fun name -> prop name (fun (b: MacroWorkbook) -> b.Setting name |> Option.defaultValue (VBoolean (name <> "StatusBar"))) (fun b v -> b.SetSetting(name, v)))

let private rangeArguments = [ for i in 1 .. 30 -> sprintf "Arg%d" i ]

let private applicationType =
    let workbook = get "ActiveWorkbook" (fun (b: MacroWorkbook) -> workbookValue b)
    let sheets = get "Worksheets" (fun (b: MacroWorkbook) -> sheetsType.Wrap b)
    let calculate = method "Calculate" [] (fun (_: MacroWorkbook) _ -> VEmpty)
    let excel = get "Name" (fun (_: MacroWorkbook) -> VString "Microsoft Excel")
    Table<MacroWorkbook>("Application", [
        workbook
        alias "ThisWorkbook" workbook
        get "ActiveSheet" (fun b -> worksheetValue b b.ActiveSheet)
        sheets
        alias "Sheets" sheets
        get "Workbooks" (fun b -> workbooksType.Wrap b)
        get "Selection" (fun b -> rangeValue b (activeRange b))
        get "ActiveCell" (fun b -> rangeValue b (activeCell b))
        get "ActiveWindow" (fun b -> windowType.Wrap b)
        get "WorksheetFunction" (fun b -> worksheetFunctionType.Wrap b)
        get "Names" (fun b -> namesType.Wrap b)
        get "Version" (fun _ -> VString "16.0")
        excel
        alias "Application" excel
        get "UserName" (fun _ -> VString "")
        prop "Calculation" (fun b -> b.Setting "Calculation" |> Option.defaultValue (VLong -4105)) (fun b v -> b.SetSetting("Calculation", v))
        calculate
        alias "CalculateFull" calculate
        method "Range" [ "Cell1"; "Cell2" ] (fun b args -> wrapRange (sheetRange b b.ActiveSheet args))
        indexed "Cells" [ "RowIndex"; "ColumnIndex" ] (fun b -> rangeValue b (wholeSheet b.ActiveSheet)) (fun b args ->
            let whole = rangeState b (wholeSheet b.ActiveSheet)
            wrapRange (if args.Length = 0 then whole else whole.CellItem args))
        method "Evaluate" [ "Name" ] (fun b args -> if args.Length = 1 then evaluateArgument b (toText args.[0]) else unsupported "Application" "Evaluate")
        method "Intersect" rangeArguments intersectOf
        method "Union" rangeArguments unionOf
        method "Goto" [ "Reference"; "Scroll" ] (fun b args ->
            if args.Length = 0 then unsupported "Application" "Goto"
            else
                match rangeRefOf args.[0] with
                | Some r -> b.SetActiveSheet r.Sheet; b.Selection <- Some (r.Sheet, r.Top, r.Left, r.Bottom, r.Right); VEmpty
                | None -> VEmpty)
        method "Wait" [ "Time" ] (fun b args ->
            // (A run a slice at a time waits until then: MacroProgram.WaitMs; a run to the end does not.)
            let target = if args.Length > 0 then toFloat args.[0] else 0.0
            let now = (DateTime.Now - DateTime(1899, 12, 30)).TotalDays
            let until = if target < 1.0 then Math.Floor now + target else target
            b.SetSetting("__waitUntil", VDouble until)
            VBoolean true)
        method "OnTime" [] (fun _ _ -> VEmpty)
        method "InputBox" [ "Prompt"; "Title"; "Default"; "Left"; "Top"; "HelpFile"; "HelpContextID"; "Type" ] (fun b args ->
            let prompt = if args.Length > 0 then toText args.[0] else ""
            let title = if isMissing (at args 1) then "Microsoft Excel" else toText args.[1]
            let defaultText = if isMissing (at args 2) then "" else toText args.[2]
            match b.Hooks.InputBox prompt title defaultText with Some text -> VString text | None -> VBoolean false) ]
        @ applicationSettings,
        // Application.Max(...), Application.Text(...): the worksheet function, its error as an Error value.
        dynamic =
            { Read = (fun b name args ->
                match args with
                | Some args when args.Length > 0 ->
                    match worksheetFunction b name args with
                    | VError 2029 -> None
                    | value -> Some value
                | _ -> None)
              Write = (fun _ _ _ -> false) })

/// `VBA.RGB(...)`, `VBA.Len(...)`, `VBA.vbRed`: the VBA library by its name, as Excel's references give
/// it to macros (the engine's built-in functions and constants).
let private vbaType =
    Table<MacroWorkbook>("VBA", [],
        dynamic =
            { Read = (fun b name args ->
                match args with
                | None ->
                    match Rbt.Vb.Builtins.tryConstant name with
                    | Some value -> Some value
                    | None -> Some (Rbt.Vb.Builtins.callBuiltin (runtimeHooksOf b) false (name.ToLowerInvariant()) [||])
                | Some args -> Some (Rbt.Vb.Builtins.callBuiltin (runtimeHooksOf b) false (name.ToLowerInvariant()) (args |> Array.map scalarOf)))
              Write = (fun _ _ _ -> false) })

/// A global that is called (Evaluate(...), Intersect(...)): its one member, the default.
let private callable (typeName: string) (parameters: string list) (run: MacroWorkbook -> Value array -> Value) =
    Table<MacroWorkbook>(typeName, [ method "_Default" parameters run ], defaultMember = "_Default")

let private evaluateGlobal =
    callable "Evaluate" [ "Name" ] (fun b args -> if args.Length = 1 then evaluateArgument b (toText args.[0]) else unsupported "Evaluate" "")
let private intersectGlobal = callable "Intersect" rangeArguments intersectOf
let private unionGlobal = callable "Union" rangeArguments unionOf
/// Range(...): of the active sheet.
let private rangeGlobal =
    callable "Range" [ "Cell1"; "Cell2" ] (fun b args -> if args.Length = 0 then unsupported "Application" "Range" else wrapRange (sheetRange b b.ActiveSheet args))
// (DoEvents: the engine has no event loop to yield to.)
let private doEventsGlobal = callable "DoEvents" [] (fun _ _ -> VLong 0)

/// The objects a macro starts from, by name: Application and the globals Excel puts in scope (ActiveSheet,
/// Range, Cells, Worksheets, Selection, ...), and the xl* constants.
let hostGlobals (book: MacroWorkbook) : (string * Value) list =
    // (ActiveSheet, Selection, Cells ... stand for what the workbook has when they are used.)
    let live (resolve: unit -> RangeState) = liveRangeType.Wrap resolve
    let activeSheetLines (lines: int) () = RangeState(book, wholeSheet book.ActiveSheet, lines, [])
    [ "Application", applicationType.Wrap book
      "ActiveWorkbook", workbookValue book
      "ThisWorkbook", workbookValue book
      "ActiveSheet", liveSheetType.Wrap (fun () -> WorksheetState(book, book.ActiveSheet))
      "ActiveCell", live (fun () -> rangeState book (activeCell book))
      "ActiveWindow", windowType.Wrap book
      "Selection", live (fun () -> rangeState book (activeRange book))
      "Worksheets", sheetsType.Wrap book
      "Sheets", sheetsType.Wrap book
      "Range", rangeGlobal.Wrap book
      "Cells", live (activeSheetLines 0)
      "Rows", live (activeSheetLines 3)
      "Columns", live (activeSheetLines 4)
      "Names", namesType.Wrap book
      "VBA", vbaType.Wrap book
      "WorksheetFunction", worksheetFunctionType.Wrap book
      "Evaluate", evaluateGlobal.Wrap book
      "Intersect", intersectGlobal.Wrap book
      "Union", unionGlobal.Wrap book
      "DoEvents", doEventsGlobal.Wrap book ]
    @ constants

/// Sheet1, Sheet2 ...: a sheet by its code name, as Excel gives every module (`Sheet1.Range("A1")`). A
/// sheet's own module of that name is the project's; the others are Excel's default code names, the sheets
/// in order.
let private codeNameGlobals (book: MacroWorkbook) (modules: Rbt.Vb.Api.ModuleSource list) : (string * Value) list =
    let taken = modules |> List.map (fun m -> m.Name.ToLowerInvariant()) |> Set.ofList
    [ for index in 1 .. 64 do
        let name = sprintf "Sheet%d" index
        if not (taken.Contains (name.ToLowerInvariant())) then
            let sheet () =
                match List.tryItem (index - 1) book.SheetOrder with
                | Some id -> WorksheetState(book, id)
                | None -> fail 424 "Object required"
            yield name, liveSheetType.Wrap sheet ]

// ---- running macros ---------------------------------------------------------------------------------------------

open Rbt.Vb.Api

/// What a run gives back: the workbook as the macro left it (calculated) and the cells it selected
/// (sheet, top, left, bottom, right), or the error it stopped on (number, description, module, line).
type MacroResult =
    | MacroDone of workbook: Workbook * selection: (SheetId * int * int * int * int) option
    | MacroFailed of VbFailure

/// How many VBA instructions a run may take before it is stopped as a runaway (about ten seconds).
let defaultInstructionLimit = 300_000_000

/// The kind of a module of a VBA project, from its Attribute lines: ThisWorkbook, sheet modules and
/// forms are document modules, class modules are classes, the rest standard modules.
let moduleKind (source: string) : ModuleKind =
    let header = ModuleSource.parseHeader source
    let attribute name = header.Attributes |> List.tryPick (fun (key, value) -> if eq key name then Some (value.Trim('"')) else None)
    match attribute "VB_Base" with
    | Some _ -> Document
    | None ->
        match attribute "VB_Creatable", attribute "VB_Exposed", attribute "VB_PredeclaredId" with
        | None, None, None -> Standard
        | _ -> Class

/// A module of the project, its name from its VB_Name (else the stream's name).
let macroModule (streamName: string) (source: string) : ModuleSource =
    // (A module exported from the VBA editor - a .cls - starts with VERSION 1.0 CLASS / BEGIN ... END.)
    let source =
        let lines = source.Replace("\r\n", "\n").Split('\n')
        let first = lines |> Array.tryFindIndex (fun line -> line.Trim() <> "" && not (line.TrimStart().StartsWith "'"))
        match first with
        | Some at when lines.[at].Trim().StartsWith("VERSION ", StringComparison.OrdinalIgnoreCase) ->
            let mutable index = at + 1
            if index < lines.Length && lines.[index].Trim().Equals("BEGIN", StringComparison.OrdinalIgnoreCase) then
                while index < lines.Length && not (lines.[index].Trim().Equals("END", StringComparison.OrdinalIgnoreCase)) do index <- index + 1
                index <- index + 1
            String.concat "\n" (Array.concat [ lines.[.. at - 1]; Array.create (index - at) ""; lines.[index ..] ])
        | _ -> source
    let header = ModuleSource.parseHeader source
    { Name = header.Name |> Option.defaultValue streamName; Kind = moduleKind source; Source = source }

/// A workbook's VBA project compiled against Excel's object model, with its module-level state kept
/// between runs (as Excel keeps it while the workbook is open). One run at a time.
type MacroProgram private (project: Project, book: MacroWorkbook, session: Session) =
    let mutable lastFrame = -1
    let hooksOf (book: MacroWorkbook) : RuntimeHooks =
        { Print = (fun text -> book.Hooks.Output text)
          MsgBox = (fun prompt buttons title -> book.Hooks.MessageBox prompt buttons title)
          InputBox = (fun prompt title defaultText -> book.Hooks.InputBox prompt title defaultText)
          CreateObject = (fun _ -> None)
          Environ = (fun _ -> "") }
    let rec finish (status: ExecutionStatus) =
        match status with
        | Paused _ ->
            // (Stop: no debugger is attached, so the run carries on.)
            finish (session.RunToEnd defaultInstructionLimit)
        | Finished _ -> MacroDone (book.Workbook, book.Selection)
        | Failed failure -> MacroFailed failure
        | Running ->
            session.Cancel()
            match session.Step 1 with
            | Failed failure -> MacroFailed { failure with Description = "The macro ran too long and was stopped" }
            | _ -> MacroFailed { Kind = Runtime; Number = 18; Description = "The macro ran too long and was stopped"; Source = ""; Line = 0; Column = 0; Procedure = ""; Module = "" }

    /// Compile the modules (`macroModule` for those of a file); the compile errors when they do not.
    static member Compile(modules: ModuleSource list) : Result<MacroProgram, VbFailure list> =
        let book = MacroWorkbook(createWorkbook (), noMacroHooks)
        let runtime : RuntimeHooks =
            { Print = (fun text -> book.Hooks.Output text)
              MsgBox = (fun prompt buttons title -> book.Hooks.MessageBox prompt buttons title)
              InputBox = (fun prompt title defaultText -> book.Hooks.InputBox prompt title defaultText)
              CreateObject = (fun _ -> None)
              Environ = (fun _ -> "") }
        let options =
            ProjectOptions.vba
            |> ProjectOptions.withHostGlobals (Map.ofList (hostGlobals book @ codeNameGlobals book modules))
            |> ProjectOptions.withHooks runtime
        match (try Project.compile options modules
               with e -> Result.Error [ { Kind = Compile; Number = 0; Description = "The VBA engine could not compile the project: " + e.Message; Source = ""; Line = 0; Column = 0; Procedure = ""; Module = "" } ]) with
        | Result.Error failures -> Result.Error failures
        | Result.Ok project ->
            // (Module-level code runs once, on a blank workbook: Excel's modules have none to speak of.)
            // (An engine that throws on a project - bytecode it does not accept - fails it as a compile error.)
            match (try Session.Create(project, defaultInstructionLimit)
                   with e -> Result.Error { Kind = Compile; Number = 0; Description = "The VBA engine could not load the project: " + e.Message; Source = ""; Line = 0; Column = 0; Procedure = ""; Module = "" }) with
            | Result.Ok session -> Result.Ok (MacroProgram(project, book, session))
            | Result.Error failure -> Result.Error [ failure ]

    member _.Project = project
    member _.Modules = project.Modules
    /// The macros a user can run (Excel's Macros dialog): public Subs without required arguments; those of
    /// sheet and workbook modules as `Module.Name`.
    member _.Macros : string list =
        Project.macros project
        |> List.map (fun p -> if p.ModuleKind = Standard then p.Name else p.QualifiedName)
    member _.Procedures = Project.procedures project
    /// Whether the project has that procedure (`Name` or `Module.Name`).
    member _.Has(name: string) = (Project.tryFindProcedure name project).IsSome
    member _.Session = session

    /// Begin running `name` with arguments on the workbook; drive it with `Step`, end with `Result`.
    member _.Begin(workbook: Workbook, hooks: MacroHooks, name: string, args: Value list) : Result<unit, VbFailure> =
        book.Reset(workbook, hooks)
        lastFrame <- book.Version
        session.SetHooks(hooksOf book)
        session.Call(name, args)
    /// Begin an Immediate-window line (`?expr`, `Print expr`, statements) on the workbook.
    member _.BeginImmediate(workbook: Workbook, hooks: MacroHooks, line: string, ?moduleName: string) : Result<unit, VbFailure> =
        book.Reset(workbook, hooks)
        session.SetHooks(hooksOf book)
        session.Immediate(line, ?moduleName = moduleName)
    member _.Step(maxInstructions: int) = session.Step maxInstructions
    /// The workbook so far, when it changed since the last frame and the screen is on (ScreenUpdating).
    member _.Frame() : Workbook option =
        match book.Setting "ScreenUpdating" with
        | Some (VBoolean false) -> None
        | _ ->
            if book.Version = lastFrame then None
            else
                lastFrame <- book.Version
                Some book.Workbook
    /// How much longer Application.Wait asked to wait, in milliseconds (0: not waiting).
    member _.WaitMs() : float =
        match book.Setting "__waitUntil" with
        | Some (VDouble until) ->
            let left = (until - (DateTime.Now - DateTime(1899, 12, 30)).TotalDays) * 86400000.0
            if left <= 0.0 then book.SetSetting("__waitUntil", VEmpty); 0.0 else left
        | _ -> 0.0
    member _.Cancel() = session.Cancel()
    member _.Location = session.Location
    member _.IsRunning = session.IsRunning
    /// The outcome of the run that ended with `status`.
    member _.Result(status: ExecutionStatus) = finish status

    /// Run `name` to its end on the workbook (stopped after `maxInstructions`).
    member this.Run(workbook: Workbook, hooks: MacroHooks, name: string, ?args: Value list, ?maxInstructions: int) : MacroResult =
        match this.Begin(workbook, hooks, name, defaultArg args []) with
        | Result.Error failure -> MacroFailed failure
        | Result.Ok () -> finish (session.RunToEnd(defaultArg maxInstructions defaultInstructionLimit))
    /// Run an Immediate-window line to its end on the workbook; `?expr` / `Print expr` prints its value.
    member this.Immediate(workbook: Workbook, hooks: MacroHooks, line: string, ?moduleName: string, ?maxInstructions: int) : MacroResult =
        match this.BeginImmediate(workbook, hooks, line, ?moduleName = moduleName) with
        | Result.Error failure -> MacroFailed failure
        | Result.Ok () ->
            let status = session.RunToEnd(defaultArg maxInstructions defaultInstructionLimit)
            match status with
            | Finished value when MacroProgram.Prints line -> hooks.Output (displayValue value)
            | _ -> ()
            finish status
    /// Whether an Immediate-window line asks for a value (`?expr`, `Print expr`).
    static member Prints(line: string) =
        let text = line.Trim()
        text.StartsWith "?" || text.StartsWith("print ", StringComparison.OrdinalIgnoreCase)
    /// Run an event handler of a module (Workbook_Open of ThisWorkbook, Worksheet_Change of a sheet's
    /// module) when the project has it; its Target a range of the sheet (sheet, top, left, bottom, right).
    member this.RunEvent(workbook: Workbook, hooks: MacroHooks, handler: string, target: (SheetId * int * int * int * int) option) : MacroResult option =
        if not (this.Has handler) then None
        else
            let args =
                match target with
                | Some (sheet, top, left, bottom, right) -> [ rangeValue book { Sheet = sheet; Top = top; Left = left; Bottom = bottom; Right = right } ]
                | None -> []
            book.Reset(workbook, hooks)
            session.SetHooks(hooksOf book)
            match session.Call(handler, args) with
            | Result.Error failure -> Some (MacroFailed failure)
            | Result.Ok () -> Some (finish (session.RunToEnd defaultInstructionLimit))

// ---- the macros of an opened workbook -----------------------------------------------------------------------

let private keptProject (preserved: XlsxReader.WorkbookPreservation) =
    preserved.WorkbookRelationships
    |> List.tryPick (fun (_, relType, target, isExternal) ->
        if not isExternal && OpcXml.relationshipKind relType = "vbaProject" then Some (ZipInterop.normalizePath target) else None)
    |> Option.bind (fun path -> preserved.Parts |> List.tryPick (fun (part, data) -> if part = path then Some data else None))

/// The modules of the VBA project an opened workbook keeps, with their kinds (Attribute lines and all).
let macroModulesOf (preserved: XlsxReader.WorkbookPreservation) : ModuleSource list =
    match keptProject preserved with
    | Some bytes ->
        (try VbaProject.readModules bytes with _ -> [])
        |> List.filter (fun (_, source) -> not (String.IsNullOrWhiteSpace source))
        |> List.map (fun (stream, source) -> macroModule stream source)
    | None -> []

/// Macro source typed or edited by hand (the native host's macro dialog), as one standard module.
let editedModules (source: string) : ModuleSource list =
    if String.IsNullOrWhiteSpace source then [] else [ ModuleSource.standard "Module1" source ]

/// A module's source as (header, code): the export header and Attribute lines Excel's editor does not
/// show, and the code after them.
let splitHeader (source: string) : string * string =
    let lines = source.Replace("\r\n", "\n").Split('\n')
    let mutable index = 0
    let mutable inBegin = false
    let isHeader (line: string) =
        let text = line.Trim()
        if inBegin then
            (if text.Equals("END", StringComparison.OrdinalIgnoreCase) then inBegin <- false)
            true
        elif text.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase) then true
        elif text.StartsWith("VERSION ", StringComparison.OrdinalIgnoreCase) then true
        elif text.Equals("BEGIN", StringComparison.OrdinalIgnoreCase) then inBegin <- true; true
        else false
    while index < lines.Length && isHeader lines.[index] do index <- index + 1
    String.concat "\n" lines.[.. index - 1], String.concat "\n" lines.[index ..]

/// The header of a module again in front of its (edited) code.
let joinHeader (header: string) (code: string) =
    if header = "" then code else header + "\n" + code.Replace("\r\n", "\n")

/// A project of the modules, as vbaProject.bin (source only: Office compiles it when it opens the file).
let projectBytes (modules: ModuleSource list) : byte[] =
    modules
    |> List.map (fun m ->
        let kind =
            match m.Kind with
            | Standard -> VbaProject.StandardModule
            | Class -> VbaProject.ClassModule
            | Document -> VbaProject.DocumentModule
        m.Name, kind, m.Source)
    |> VbaProject.writeProjectModules
