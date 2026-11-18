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
      "xlErrDiv0", 2007; "xlErrNA", 2042; "xlErrName", 2029; "xlErrNull", 2000; "xlErrNum", 2036; "xlErrRef", 2023; "xlErrValue", 2015 ]
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
    member _.Setting(name: string) =
        match settingValues.TryGetValue(name.ToLowerInvariant()) with
        | true, value -> Some value
        | _ -> None
    member _.SetSetting(name: string, value: Value) = settingValues.[name.ToLowerInvariant()] <- value
    member _.Charts = charts
    member _.Workbook = calculate (); workbook
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

/// The parameters of Excel's methods by name, for calls with named arguments (`Find What:="x"`,
/// `Worksheets.Add After:=ws`): by (object type, method), lower case.
let private parameterNames : Map<string * string, string list> =
    [ ("Range", "find"), [ "What"; "After"; "LookIn"; "LookAt"; "SearchOrder"; "SearchDirection"; "MatchCase"; "MatchByte"; "SearchFormat" ]
      ("Range", "address"), [ "RowAbsolute"; "ColumnAbsolute"; "ReferenceStyle"; "External"; "RelativeTo" ]
      ("Range", "offset"), [ "RowOffset"; "ColumnOffset" ]
      ("Range", "resize"), [ "RowSize"; "ColumnSize" ]
      ("Range", "copy"), [ "Destination" ]
      ("Range", "pastespecial"), [ "Paste"; "Operation"; "SkipBlanks"; "Transpose" ]
      ("Range", "delete"), [ "Shift" ]
      ("Range", "insert"), [ "Shift"; "CopyOrigin" ]
      ("Range", "end"), [ "Direction" ]
      ("Range", "borderaround"), [ "LineStyle"; "Weight"; "ColorIndex"; "Color"; "ThemeColor" ]
      ("Range", "addcomment"), [ "Text" ]
      ("Range", "replace"), [ "What"; "Replacement"; "LookAt"; "SearchOrder"; "MatchCase"; "MatchByte"; "SearchFormat"; "ReplaceFormat" ]
      ("FormatConditions", "add"), [ "Type"; "Operator"; "Formula1"; "Formula2" ]
      ("FormatConditions", "addcolorscale"), [ "ColorScaleType" ]
      ("Range", "cells"), [ "RowIndex"; "ColumnIndex" ]
      ("Range", "item"), [ "RowIndex"; "ColumnIndex" ]
      ("Sheets", "add"), [ "Before"; "After"; "Count"; "Type" ]
      ("Names", "add"), [ "Name"; "RefersTo"; "Visible"; "MacroType"; "ShortcutKey"; "Category"; "NameLocal"; "RefersToLocal" ]
      ("Application", "inputbox"), [ "Prompt"; "Title"; "Default"; "Left"; "Top"; "HelpFile"; "HelpContextID"; "Type" ]
      ("Application", "intersect"), [ "Arg1"; "Arg2"; "Arg3"; "Arg4" ]
      ("Application", "union"), [ "Arg1"; "Arg2"; "Arg3"; "Arg4" ] ]
    |> List.map (fun ((typeName, methodName), names) -> (typeName, methodName), names |> List.map (fun n -> n.ToLowerInvariant()))
    |> Map.ofList

type private Host(typeName: string, get: string -> Value option, set: string -> Value -> bool,
                  call: string -> Value array -> Value option, enumerate: (unit -> Value seq) option) =
    interface IHostObject with
        member _.TypeName = typeName
        member _.GetMember name = get name
        member _.SetMember name value = set name value
        member _.CallMethod name args = call name args
        member _.CallNamed name args names =
            // Named arguments into their places; the holes are omitted arguments.
            // (Globals that stand for a collection or range name their parameters as it does.)
            let kind =
                match typeName with
                | "Worksheets" -> "Sheets"
                | "Selection" | "ActiveCell" | "Cells" | "Rows" | "Columns" -> "Range"
                | "ActiveSheet" -> "Worksheet"
                | other -> other
            match Map.tryFind (kind, name.ToLowerInvariant()) parameterNames with
            | None -> call name args
            | Some parameters ->
                let slots = ResizeArray<Value>()
                let put index value =
                    while slots.Count <= index do slots.Add VUndefined
                    slots.[index] <- value
                args |> Array.iteri (fun i arg ->
                    let argName = if i < names.Length then names.[i] else ""
                    if argName = "" then put i arg
                    else
                        match parameters |> List.tryFindIndex ((=) (argName.ToLowerInvariant())) with
                        | Some index -> put index arg
                        | None -> raiseVbErrorWith 448 ("Named argument not found: " + argName))
                call name (slots.ToArray())
        member _.Enumerate () = enumerate |> Option.map (fun items -> items ())

let private host typeName get set call = VHostObject (Host(typeName, get, set, call, None) :> IHostObject)
let private collection typeName get set call items = VHostObject (Host(typeName, get, set, call, Some items) :> IHostObject)

/// A range, recognised among host objects (for WorksheetFunction arguments, Union, Intersect, Copy).
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

/// The range a host object stands for (a Range, a global that forwards to one - Selection, ActiveCell -,
/// the first area of a union).
let private rangeRefOf (value: Value) : RangeRef option =
    match value with
    | VHostObject h ->
        match (try h.GetMember "__RbtRange" with _ -> None) with
        | Some (VString text) ->
            let parts = text.Split '\t'
            Some { Sheet = parts.[0]; Top = int parts.[1]; Left = int parts.[2]; Bottom = int parts.[3]; Right = int parts.[4] }
        | _ -> None
    | _ -> None

/// The areas a host object stands for: those of a union, else its one range.
let private areasOf (value: Value) : RangeRef list =
    match value with
    | VHostObject h ->
        match (try h.GetMember "__RbtAreas" with _ -> None) with
        | Some (VString text) when text <> "" ->
            [ for area in text.Split '\n' ->
                let parts = area.Split '\t'
                { Sheet = parts.[0]; Top = int parts.[1]; Left = int parts.[2]; Bottom = int parts.[3]; Right = int parts.[4] } ]
        | _ -> Option.toList (rangeRefOf value)
    | _ -> []

let private sheetIdNamed (book: MacroWorkbook) (name: string) =
    book.SheetOrder |> List.tryFind (fun id -> eq (book.Sheet id).Name name)

let private quotedSheet (name: string) =
    if name |> Seq.forall (fun ch -> Char.IsLetterOrDigit ch || ch = '_') && not (Char.IsDigit name.[0]) then name
    else "'" + name.Replace("'", "''") + "'"

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

/// Font or Interior of a format kept elsewhere (a conditional format's style, DisplayFormat): its colours,
/// bold and italic, read from `format ()` and written with `change` (None: read only).
let private styleObject (typeName: string) (interior: bool) (format: unit -> CellFormat) (change: ((CellFormat -> CellFormat) -> unit) option) =
    let colorOf (f: CellFormat) = if interior then f.BgColor else f.TextColor
    let withColor (color: string option) (f: CellFormat) = if interior then { f with BgColor = color } else { f with TextColor = color }
    let write (f: CellFormat -> CellFormat) =
        match change with
        | Some apply -> apply f
        | None -> failApp (sprintf "Unable to set a property of the %s class" typeName)
    host typeName
        (fun name ->
            let f = format ()
            if eq name "Color" then Some (VLong (colorToLong (colorOf f) |> Option.defaultValue (if interior then 16777215 else 0)))
            elif eq name "ColorIndex" then
                Some (VLong (match colorOf f with
                                | None -> if interior then -4142 else -4105
                                | Some c -> match Array.IndexOf(palette, c.ToUpperInvariant()) with -1 -> -4105 | i -> i + 1))
            elif eq name "Pattern" && interior then Some (VLong (if f.BgColor.IsSome then 1 else -4142))
            elif eq name "Bold" && not interior then Some (VBoolean f.Bold)
            elif eq name "Italic" && not interior then Some (VBoolean f.Italic)
            elif eq name "ThemeColor" || eq name "TintAndShade" then Some (VLong 0)
            else unsupported typeName name)
        (fun name value ->
            if eq name "Color" then write (withColor (Some (longToColor (toInt value)))); true
            elif eq name "ColorIndex" then let i = toInt value in write (withColor (if i >= 1 && i <= 56 then Some palette.[i - 1] else None)); true
            elif eq name "ThemeColor" then let c = themeColor (toInt value) in write (withColor (Some c)); true
            elif eq name "TintAndShade" then let t = toFloat value in write (fun f -> withColor (colorOf f |> Option.map (tinted t)) f); true
            elif eq name "Pattern" && interior then (if toInt value = -4142 then write (withColor None)); true
            elif eq name "Bold" && not interior then let b = toBool value in write (fun f -> { f with Bold = b }); true
            elif eq name "Italic" && not interior then let b = toBool value in write (fun f -> { f with Italic = b }); true
            elif eq name "PatternColorIndex" || eq name "Strikethrough" || eq name "Underline" then true
            else unsupported typeName name)
        (fun name _ -> unsupported typeName name)

/// Excel's 3-colour scale as AddColorScale(3) makes it (red, yellow, green), and the 2-colour one.
let private defaultColorScale (stops: int) =
    if stops = 2 then [ { Kind = CfvoMin; Value = None }, "#F8696B"; { Kind = CfvoMax; Value = None }, "#63BE7B" ]
    else [ { Kind = CfvoMin; Value = None }, "#F8696B"; { Kind = CfvoPercentile; Value = Some "50" }, "#FFEB84"; { Kind = CfvoMax; Value = None }, "#63BE7B" ]

let rec private rangeObject (book: MacroWorkbook) (r: RangeRef) : Value =
    let rows = r.Bottom - r.Top + 1
    let cols = r.Right - r.Left + 1
    let sheetName () = (book.Sheet r.Sheet).Name
    let sub top left bottom right = rangeObject book { r with Top = top; Left = left; Bottom = bottom; Right = right }
    let cellAt row col = book.Cell(r.Sheet, row, col)
    let valueAt row col = cellAt row col |> Option.map (fun c -> cellValueAs book.Date1904 c.Format c.Value) |> Option.defaultValue VEmpty
    let value2At row col = cellAt row col |> Option.map (fun c -> cellValue c.Value) |> Option.defaultValue VEmpty
    let formulaAt row col =
        match cellAt row col with
        | Some { Value = Formula (text, _) } -> VString (if text.StartsWith "=" then text else "=" + text)
        | Some cell -> match cellValue cell.Value with VEmpty -> VString "" | other -> other
        | None -> VString ""
    let textAt row col =
        match cellAt row col with
        | Some cell ->
            let width = book.Model(r.Sheet).ColumnWidths |> Map.tryFind col |> Option.defaultValue (book.Model(r.Sheet).DefaultColumnWidth)
            VString (CellText.cellDisplayTextIn book.Date1904 None width cell.Format cell.Value)
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
    let format () = cellAt r.Top r.Left |> Option.map (fun c -> c.Format) |> Option.defaultValue defaultFormat
    let wholeRows = r.Left = 0 && r.Right = maxCol
    let wholeCols = r.Top = 0 && r.Bottom = maxRow
    let changeFormats (change: CellFormat -> CellFormat) =
        if wholeRows || wholeCols then
            // Whole columns, rows or the sheet (Cells.Interior.Color = vbWhite): the cells it has and those the
            // grid shows, and for whole columns their default style (a cell made later starts from it).
            book.ChangeModel(r.Sheet, fun model ->
                let mutable m = model
                let inside (a: CellAddress) = a.Row >= r.Top && a.Row <= r.Bottom && a.Col >= r.Left && a.Col <= r.Right
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
    let alignmentName (value: Value) (horizontal: bool) =
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
    let alignmentValue (name: string option) (horizontal: bool) =
        match name with
        | Some "left" -> -4131 | Some "right" -> -4152 | Some "center" | Some "middle" -> -4108
        | Some "justify" -> -4130 | Some "top" -> -4160 | Some "bottom" -> -4107
        | Some "centerContinuous" -> 7 | Some "fill" -> 5 | Some "distributed" -> -4117
        | _ -> if horizontal then 1 else -4107
    let used () =
        let model = book.Model r.Sheet
        if Map.isEmpty model.Cells then None
        else
            let addresses = model.Cells |> Map.toSeq |> Seq.map fst
            Some (addresses |> Seq.map (fun a -> a.Row) |> Seq.min, addresses |> Seq.map (fun a -> a.Col) |> Seq.min,
                  addresses |> Seq.map (fun a -> a.Row) |> Seq.max, addresses |> Seq.map (fun a -> a.Col) |> Seq.max)
    let filled row col = match cellAt row col with Some { Value = v } -> unwrapFormulaValue v <> CellValue.Empty | None -> false
    /// Ctrl+arrow from the range's first cell: to the edge of the block it is in, or of the next one.
    let endOf (direction: int) =
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
    let currentRegion () =
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
    let clearContents () =
        for (address, cell) in book.Model(r.Sheet).Cells |> Map.toList do
            if address.Row >= r.Top && address.Row <= r.Bottom && address.Col >= r.Left && address.Col <= r.Right then
                book.SetCell(r.Sheet, address.Row, address.Col, (if cell.Format = defaultFormat && cell.Comment.IsNone then None else Some { cell with Value = CellValue.Empty }))
    let clearFormats () =
        if wholeCols then book.ChangeModel(r.Sheet, fun m -> { m with ColumnStyles = restyledColumns m.ColumnStyles r.Left r.Right None })
        for (address, cell) in book.Model(r.Sheet).Cells |> Map.toList do
            if address.Row >= r.Top && address.Row <= r.Bottom && address.Col >= r.Left && address.Col <= r.Right then
                book.SetCell(r.Sheet, address.Row, address.Col, (if cell.Value = CellValue.Empty && cell.Comment.IsNone then None else Some { cell with Format = defaultFormat }))
    let copyTo (target: RangeRef) (what: int) =
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
        ignore sides
    let rec borderObject (index: int option) =
        let read (name: string) : Value option =
            let one (pick: BorderStyle option -> Value) =
                let model = book.Model r.Sheet
                match bordersSides index |> List.map (borderIn model >> pick) |> List.distinct with
                | [ v ] -> v
                | _ -> VNull
            if eq name "LineStyle" || eq name "Value" then Some (one (function Some b -> VLong (fst (lineOfStyle b.Style)) | None -> VLong xlLineStyleNone))
            elif eq name "Weight" then Some (one (function Some b -> VLong (snd (lineOfStyle b.Style)) | None -> VLong 2))
            elif eq name "Color" then Some (one (function Some { Color = Some c } -> VLong (colorToLong (Some c) |> Option.defaultValue 0) | _ -> VLong 0))
            elif eq name "ColorIndex" then
                Some (one (function
                    | Some { Color = Some c } -> VLong (match Array.IndexOf(palette, c.ToUpperInvariant()) with -1 -> -4105 | i -> i + 1)
                    | Some _ -> VLong -4105
                    | None -> VLong -4142))
            elif eq name "ThemeColor" || eq name "TintAndShade" then Some (VLong 0)
            elif eq name "Count" && index.IsNone then Some (VLong 6)
            else None
        let write (name: string) (value: Value) =
            let styled (current: BorderStyle option) = current |> Option.defaultValue { Style = "thin"; Color = None }
            if eq name "LineStyle" || eq name "Value" then
                let lineStyle = match scalarOf value with VNull | VEmpty -> xlLineStyleNone | v -> toInt v
                changeBorders index (fun current ->
                    let weight = current |> Option.map (fun b -> snd (lineOfStyle b.Style)) |> Option.defaultValue 2
                    styleOfLine lineStyle weight |> Option.map (fun style -> { Style = style; Color = current |> Option.bind (fun b -> b.Color) }))
            elif eq name "Weight" then
                let weight = toInt value
                changeBorders index (fun current ->
                    let lineStyle = current |> Option.map (fun b -> fst (lineOfStyle b.Style)) |> Option.defaultValue xlContinuous
                    styleOfLine lineStyle weight |> Option.map (fun style -> { Style = style; Color = current |> Option.bind (fun b -> b.Color) }))
            elif eq name "Color" then
                let color = longToColor (toInt value)
                changeBorders index (fun current -> Some { styled current with Color = Some color })
            elif eq name "ColorIndex" then
                let i = toInt value
                changeBorders index (fun current -> Some { styled current with Color = (if i >= 1 && i <= 56 then Some palette.[i - 1] else None) })
            elif eq name "ThemeColor" then
                let color = themeColor (toInt value)
                changeBorders index (fun current -> Some { styled current with Color = Some color })
            elif eq name "TintAndShade" then
                let amount = toFloat value
                changeBorders index (Option.map (fun b -> { b with Color = b.Color |> Option.map (tinted amount) }))
            else unsupported "Border" name
            true
        host (if index.IsNone then "Borders" else "Border")
            (fun name -> match read name with Some v -> Some v | None -> unsupported (if index.IsNone then "Borders" else "Border") name)
            write
            (fun name args ->
                if index.IsNone && (name = "" || eq name "_Default" || eq name "Item") && args.Length = 1 then Some (borderObjectAt (toInt args.[0]))
                else match read name with Some v when args.Length = 0 -> Some v | _ -> unsupported "Borders" name)
    and borderObjectAt (index: int) = borderObject (Some index)
    // ---- notes (Range.Comment) and Range.ID ----
    let noteOf () = getCell { Row = r.Top; Col = r.Left } (book.Model r.Sheet) |> Option.bind (fun c -> c.Comment)
    let setNote (note: CellComment option) =
        book.ChangeModel(r.Sheet, fun model ->
            let address = { Row = r.Top; Col = r.Left }
            match Map.tryFind address model.Cells, note with
            | Some data, _ -> { model with Cells = Map.add address { data with Comment = note; Format = { data.Format with HasComment = note.IsSome } } model.Cells }
            | None, Some _ -> { model with Cells = Map.add address { Value = CellValue.Empty; Format = { defaultFormat with HasComment = true }; Comment = note } model.Cells }
            | None, None -> model)
    let commentObject () =
        let text () = noteOf () |> Option.map (fun n -> n.Text) |> Option.defaultValue ""
        host "Comment"
            (fun name ->
                if eq name "Text" then Some (VString (text ()))
                elif eq name "Author" then Some (VString (noteOf () |> Option.bind (fun n -> n.Author) |> Option.defaultValue ""))
                elif eq name "Visible" then Some (VBoolean false)
                elif eq name "Delete" then setNote None; Some VEmpty
                elif eq name "Parent" then Some (rangeObject book r)
                else unsupported "Comment" name)
            (fun name _ -> if eq name "Visible" then true else unsupported "Comment" name)
            (fun name args ->
                if eq name "Text" then
                    // Text(Text, Start, Overwrite): the note's text, replaced (or inserted at Start).
                    if args.Length >= 1 && not (isMissing args.[0]) then
                        let start = if args.Length >= 2 && not (isMissing args.[1]) then toInt args.[1] else 0
                        let overwrite = if args.Length >= 3 && not (isMissing args.[2]) then toBool args.[2] else true
                        let current = text ()
                        let next =
                            if start <= 0 then toText args.[0]
                            else
                                let at = min current.Length (start - 1)
                                if overwrite then current.Substring(0, at) + toText args.[0] else current.Insert(at, toText args.[0])
                        setNote (Some { (noteOf () |> Option.defaultValue (plainNote "" None)) with Text = next; Runs = [] })
                    Some (VString (text ()))
                elif eq name "Delete" then setNote None; Some VEmpty
                else unsupported "Comment" name)
    let idKey () = sprintf "__id\t%s\t%d\t%d" r.Sheet r.Top r.Left
    // ---- conditional formats of the range (FormatConditions) ----
    let thisRange : CellRange = { From = { Row = r.Top; Col = r.Left }; To = { Row = r.Bottom; Col = r.Right } }
    let overlaps (cr: CellRange) = not (cr.To.Row < r.Top || cr.From.Row > r.Bottom || cr.To.Col < r.Left || cr.From.Col > r.Right)
    let conditions () =
        book.Model(r.Sheet).ConditionalFormats |> List.filter (fun e -> e.Ranges |> List.exists overlaps) |> List.sortBy (fun e -> e.Priority)
    let conditionObject (id: int) =
        let entry () =
            match book.Model(r.Sheet).ConditionalFormats |> List.tryFind (fun e -> e.Id = id) with
            | Some e -> e
            | None -> failApp "Application-defined or object-defined error"
        let update (f: ConditionalFormatRange -> ConditionalFormatRange) =
            book.ChangeModel(r.Sheet, fun m -> { m with ConditionalFormats = m.ConditionalFormats |> List.map (fun e -> if e.Id = id then f e else e) })
        let changeStyle (f: CellFormat -> CellFormat) = update (fun e -> { e with Format = { e.Format with Style = f e.Format.Style } })
        let stops () = match (entry ()).Format.Rule with ColorScale stops -> stops | _ -> failApp "This condition is not a colour scale"
        let changeStop (index: int) (f: CfThreshold * string -> CfThreshold * string) =
            update (fun e ->
                match e.Format.Rule with
                | ColorScale stops -> { e with Format = { e.Format with Rule = ColorScale (stops |> List.mapi (fun i stop -> if i = index - 1 then f stop else stop)) } }
                | _ -> e)
        let criterion (index: int) =
            let stop () = match List.tryItem (index - 1) (stops ()) with Some s -> s | None -> fail 9 "Subscript out of range"
            let kindValue (kind: CfThresholdKind) =
                match kind with CfvoNumber -> 0 | CfvoMin -> 1 | CfvoMax -> 2 | CfvoPercent -> 3 | CfvoFormula -> 4 | CfvoPercentile -> 5 | _ -> -1
            host "ColorScaleCriterion"
                (fun name ->
                    let threshold, color = stop ()
                    if eq name "Type" then Some (VLong (kindValue threshold.Kind))
                    elif eq name "Value" then Some (match threshold.Value with Some v -> (match Double.TryParse(v, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with | true, n -> VDouble n | _ -> VString v) | None -> VEmpty)
                    elif eq name "FormatColor" then
                        Some (styleObject "FormatColor" true (fun () -> { defaultFormat with BgColor = Some (snd (stop ())) })
                                (Some (fun f -> changeStop index (fun (t, c) -> t, (f { defaultFormat with BgColor = Some c }).BgColor |> Option.defaultValue c))))
                    else unsupported "ColorScaleCriterion" name)
                (fun name value ->
                    if eq name "Type" then
                        let kind = match toInt value with 0 -> CfvoNumber | 1 -> CfvoMin | 2 -> CfvoMax | 3 -> CfvoPercent | 4 -> CfvoFormula | 5 -> CfvoPercentile | _ -> CfvoMin
                        changeStop index (fun (t, c) -> { t with Kind = kind; Value = (if kind = CfvoMin || kind = CfvoMax then None else t.Value) }, c); true
                    elif eq name "Value" then changeStop index (fun (t, c) -> { t with Value = Some ((toText value).TrimStart '=') }, c); true
                    else unsupported "ColorScaleCriterion" name)
                (fun name _ -> unsupported "ColorScaleCriterion" name)
        host "FormatCondition"
            (fun name ->
                if eq name "Font" then Some (styleObject "Font" false (fun () -> (entry ()).Format.Style) (Some changeStyle))
                elif eq name "Interior" then Some (styleObject "Interior" true (fun () -> (entry ()).Format.Style) (Some changeStyle))
                elif eq name "Type" then Some (VLong (match (entry ()).Format.Rule with CellIs _ -> 1 | Expression _ -> 2 | ColorScale _ -> 3 | DataBar _ | DataBarEx _ -> 4 | _ -> 1))
                elif eq name "Priority" then Some (VLong (entry ()).Priority)
                elif eq name "StopIfTrue" then Some (VBoolean (entry ()).StopIfTrue)
                elif eq name "Formula1" then Some (match (entry ()).Format.Rule with CellIs (_, f, _) | Expression f -> VString ("=" + f) | _ -> VString "")
                elif eq name "SetFirstPriority" then book.ChangeModel(r.Sheet, SheetOps.moveConditionalFormat "SetFirstPriority" id 1); Some VEmpty
                elif eq name "Delete" then book.ChangeModel(r.Sheet, SheetOps.removeConditionalFormat "Delete" id); Some VEmpty
                elif eq name "ColorScaleCriteria" then
                    Some (collection "ColorScaleCriteria"
                            (fun n -> if eq n "Count" then Some (VLong (stops ()).Length) else unsupported "ColorScaleCriteria" n)
                            (fun n _ -> unsupported "ColorScaleCriteria" n)
                            (fun n args -> if (n = "" || eq n "Item") && args.Length = 1 then Some (criterion (toInt args.[0])) else unsupported "ColorScaleCriteria" n)
                            (fun () -> (stops ()) |> List.mapi (fun i _ -> criterion (i + 1)) |> Seq.ofList))
                else None
                |> function Some v -> Some v | None -> unsupported "FormatCondition" name)
            (fun name value ->
                if eq name "StopIfTrue" then update (fun e -> { e with StopIfTrue = toBool value }); true
                else unsupported "FormatCondition" name)
            (fun name args ->
                if eq name "ColorScaleCriteria" && args.Length = 1 then Some (criterion (toInt args.[0]))
                elif eq name "SetFirstPriority" then book.ChangeModel(r.Sheet, SheetOps.moveConditionalFormat "SetFirstPriority" id 1); Some VEmpty
                elif eq name "Delete" then book.ChangeModel(r.Sheet, SheetOps.removeConditionalFormat "Delete" id); Some VEmpty
                else unsupported "FormatCondition" name)
    let addCondition (rule: ConditionalFormatRule) =
        let highest = book.SheetOrder |> List.fold (fun value id -> max value (SheetOps.highestConditionalFormatId (book.Model id))) 0
        let id = highest + 1
        book.ChangeModel(r.Sheet, SheetOps.addConditionalFormatWithId id [ thisRange ] false { Rule = rule; Style = defaultFormat })
        conditionObject id
    let formatConditions () =
        let item (i: int) =
            match List.tryItem (i - 1) (conditions ()) with
            | Some e -> conditionObject e.Id
            | None -> fail 9 "Subscript out of range"
        let add (args: Value array) =
            let arg i = if i < args.Length then args.[i] else VEmpty
            let formula (v: Value) = (toText v).TrimStart '='
            match (if isMissing (arg 0) then 1 else toInt (arg 0)) with
            | 2 -> addCondition (Expression (formula (arg 2)))
            | 1 ->
                let op =
                    match (if isMissing (arg 1) then 1 else toInt (arg 1)) with
                    | 1 -> OpBetween | 2 -> OpNotBetween | 3 -> OpEqual | 4 -> OpNotEqual
                    | 5 -> OpGreaterThan | 6 -> OpLessThan | 7 -> OpGreaterOrEqual | _ -> OpLessOrEqual
                addCondition (CellIs (op, formula (arg 2), (if isMissing (arg 3) then None else Some (formula (arg 3)))))
            | other -> failApp (sprintf "FormatConditions.Add of type %d is not supported" other)
        collection "FormatConditions"
            (fun name ->
                if eq name "Count" then Some (VLong (conditions ()).Length)
                elif eq name "Delete" then
                    for e in conditions () do book.ChangeModel(r.Sheet, SheetOps.removeConditionalFormat "Delete" e.Id)
                    Some VEmpty
                else unsupported "FormatConditions" name)
            (fun name _ -> unsupported "FormatConditions" name)
            (fun name args ->
                if (name = "" || eq name "Item") && args.Length = 1 then Some (item (toInt args.[0]))
                elif eq name "Add" then Some (add args)
                elif eq name "AddColorScale" then Some (addCondition (ColorScale (defaultColorScale (if args.Length > 0 && not (isMissing args.[0]) then toInt args.[0] else 3))))
                elif eq name "Delete" then
                    for e in conditions () do book.ChangeModel(r.Sheet, SheetOps.removeConditionalFormat "Delete" e.Id)
                    Some VEmpty
                elif eq name "Count" then Some (VLong (conditions ()).Length)
                else unsupported "FormatConditions" name)
            (fun () -> conditions () |> List.map (fun e -> conditionObject e.Id) |> Seq.ofList)
    /// What the first cell shows with (DisplayFormat): its format, with a conditional format that holds on it.
    let displayed () =
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
    let displayFormatObject () =
        host "DisplayFormat"
            (fun name ->
                if eq name "Interior" then Some (styleObject "Interior" true displayed None)
                elif eq name "Font" then Some (styleObject "Font" false displayed None)
                elif eq name "NumberFormat" then Some (VString ((displayed ()).NumberFormat |> Option.defaultValue "General"))
                elif eq name "HorizontalAlignment" then Some (VLong 1)
                else unsupported "DisplayFormat" name)
            (fun name _ -> unsupported "DisplayFormat" name)
            (fun name _ -> unsupported "DisplayFormat" name)
    /// Replace(What, Replacement, LookAt, SearchOrder, MatchCase): in the values of the range's cells, each
    /// entered again as typing would (a cell left "" is empty).
    let replace (args: Value array) =
        let arg i = if i < args.Length then args.[i] else VEmpty
        let what = toText (arg 0)
        let replacement = toText (arg 1)
        let whole = not (isMissing (arg 2)) && toInt (arg 2) = 1
        let matchCase = not (isMissing (arg 4)) && toBool (arg 4)
        if what <> "" then
            let cells = book.Model(r.Sheet).Cells |> Map.toList |> List.filter (fun (a, _) -> a.Row >= r.Top && a.Row <= r.Bottom && a.Col >= r.Left && a.Col <= r.Right)
            for (a, cell) in cells do
                let text = match cell.Value with Formula (f, _) -> (if f.StartsWith "=" then f else "=" + f) | v -> toText (cellValue v)
                let next =
                    if whole then (if String.Equals(text, what, (if matchCase then StringComparison.Ordinal else StringComparison.OrdinalIgnoreCase)) then replacement else text)
                    elif matchCase then text.Replace(what, replacement)
                    else Text.RegularExpressions.Regex.Replace(text, Text.RegularExpressions.Regex.Escape what, replacement.Replace("$", "$$"), Text.RegularExpressions.RegexOptions.IgnoreCase)
                if next <> text then book.SetCell(r.Sheet, a.Row, a.Col, entered book r.Sheet a.Row a.Col (VString next))
        VBoolean true
    let fontObject () =
        let get (name: string) =
            let f = format ()
            if eq name "Bold" then Some (VBoolean f.Bold)
            elif eq name "Italic" then Some (VBoolean f.Italic)
            elif eq name "Underline" then Some (VLong (if f.UnderlineDouble then -4119 elif f.Underline then 2 else -4142))
            elif eq name "Strikethrough" then Some (VBoolean f.Strikethrough)
            elif eq name "Size" then Some (VDouble (f.FontSize |> Option.map float |> Option.defaultValue 11.0))
            elif eq name "Name" then Some (VString (f.FontName |> Option.defaultValue "Calibri"))
            elif eq name "Color" then Some (VDouble (float (colorToLong f.TextColor |> Option.defaultValue 0)))
            elif eq name "ColorIndex" then Some (VLong (match f.TextColor with None -> -4105 | Some c -> (Array.IndexOf(palette, c.ToUpperInvariant()) + 1) |> fun i -> if i = 0 then -4105 else i))
            elif eq name "ThemeColor" || eq name "TintAndShade" then Some (VLong 0)
            elif eq name "FontStyle" then Some (VString (match f.Bold, f.Italic with true, true -> "Bold Italic" | true, false -> "Bold" | false, true -> "Italic" | _ -> "Regular"))
            else None
        let set (name: string) (value: Value) =
            if eq name "Bold" then changeFormats (fun f -> { f with Bold = toBool value }); true
            elif eq name "Italic" then changeFormats (fun f -> { f with Italic = toBool value }); true
            elif eq name "Underline" then
                let u = match scalarOf value with VBoolean b -> (if b then 2 else -4142) | v -> toInt v
                changeFormats (fun f -> { f with Underline = (u = 2 || u = -4119 || u = 4 || u = 5); UnderlineDouble = (u = -4119) }); true
            elif eq name "Strikethrough" then changeFormats (fun f -> { f with Strikethrough = toBool value }); true
            elif eq name "Size" then changeFormats (fun f -> { f with FontSize = Some (int (Math.Round (toFloat value))) }); true
            elif eq name "Name" then changeFormats (fun f -> { f with FontName = Some (toText value) }); true
            elif eq name "Color" then changeFormats (fun f -> { f with TextColor = Some (longToColor (toInt value)) }); true
            elif eq name "ColorIndex" then
                let i = toInt value
                changeFormats (fun f -> { f with TextColor = (if i >= 1 && i <= 56 then Some palette.[i - 1] else None) }); true
            elif eq name "ThemeColor" then
                let color = themeColor (toInt value)
                changeFormats (fun f -> { f with TextColor = Some color }); true
            elif eq name "TintAndShade" then
                let amount = toFloat value
                changeFormats (fun f -> { f with TextColor = f.TextColor |> Option.map (tinted amount) }); true
            elif eq name "FontStyle" then
                let style = (toText value).ToLowerInvariant()
                changeFormats (fun f -> { f with Bold = style.Contains "bold"; Italic = style.Contains "italic" }); true
            else unsupported "Font" name
        host "Font" (fun name -> match get name with Some v -> Some v | None -> unsupported "Font" name) set (fun name _ -> unsupported "Font" name)
    let interiorObject () =
        let get (name: string) : Value option =
            let f = format ()
            if eq name "Color" then Some (VDouble (float (colorToLong f.BgColor |> Option.defaultValue 16777215)))
            elif eq name "ColorIndex" then Some (VLong (match f.BgColor with None -> -4142 | Some c -> (Array.IndexOf(palette, c.ToUpperInvariant()) + 1) |> fun i -> if i = 0 then -4105 else i))
            elif eq name "Pattern" then Some (VLong (if f.BgColor.IsSome then 1 else -4142))
            elif eq name "ThemeColor" || eq name "TintAndShade" || eq name "PatternTintAndShade" then Some (VLong 0)
            elif eq name "PatternColorIndex" then Some (VLong -4105)
            else unsupported "Interior" name
        let set (name: string) (value: Value) =
            if eq name "Color" then changeFormats (fun f -> { f with BgColor = Some (longToColor (toInt value)) }); true
            elif eq name "ColorIndex" then
                let i = toInt value
                changeFormats (fun f -> { f with BgColor = (if i >= 1 && i <= 56 then Some palette.[i - 1] else None) }); true
            elif eq name "Pattern" then
                (if toInt value = -4142 then changeFormats (fun f -> { f with BgColor = None })); true
            elif eq name "ThemeColor" then
                let color = themeColor (toInt value)
                changeFormats (fun f -> { f with BgColor = Some color }); true
            elif eq name "TintAndShade" then
                let amount = toFloat value
                changeFormats (fun f -> { f with BgColor = f.BgColor |> Option.map (tinted amount) }); true
            elif eq name "PatternColorIndex" || eq name "PatternTintAndShade" || eq name "PatternColor" || eq name "PatternThemeColor" then true
            else unsupported "Interior" name
        host "Interior" get set (fun name _ -> unsupported "Interior" name)
    let lines (vertical: bool) =
        // Rows / Columns of the range: a collection with Count and Item, and the range itself otherwise.
        let count = if vertical then cols else rows
        let item (i: int) = if vertical then sub r.Top (r.Left + i - 1) r.Bottom (r.Left + i - 1) else sub (r.Top + i - 1) r.Left (r.Top + i - 1) r.Right
        let self = rangeObject book r
        let selfHost = match self with VHostObject h -> h | _ -> failwith "range"
        let usedCount () =
            if not (wholeRows || wholeCols) then count
            else
                match used () with
                | Some (_, _, usedBottom, usedRight) -> if vertical then min count (usedRight - r.Left + 1) else min count (usedBottom - r.Top + 1)
                | None -> 0
        collection "Range"
            (fun name -> if eq name "Count" then Some (VLong count) else selfHost.GetMember name)
            (fun name value -> selfHost.SetMember name value)
            (fun name args ->
                if (name = "" || eq name "_Default" || eq name "Item") && args.Length >= 1 then Some (item (toInt args.[0]))
                else selfHost.CallMethod name args)
            (fun () -> seq { for i in 1 .. usedCount () -> item i })
    let get (name: string) : Value option =
        if name = "__RbtRange" then Some (VString (String.Join("\t", [| r.Sheet; string r.Top; string r.Left; string r.Bottom; string r.Right |])))
        elif eq name "Value2" then Some (gather value2At)
        elif eq name "Value" || eq name "_Default" || name = "" then Some (gather valueAt)
        elif eq name "Formula" || eq name "FormulaLocal" then Some (gather formulaAt)
        elif eq name "FormulaR1C1" then
            Some (gather (fun row col -> match formulaAt row col with VString text when text.StartsWith "=" -> VString (toR1C1 row col text) | other -> other))
        elif eq name "Text" then Some (textAt r.Top r.Left)
        elif eq name "HasFormula" then Some (VBoolean (match cellAt r.Top r.Left with Some { Value = Formula _ } -> true | _ -> false))
        elif eq name "Address" then Some (VString (rangeAddress r true true false))
        elif eq name "Row" then Some (VLong (r.Top + 1))
        elif eq name "Column" then Some (VLong (r.Left + 1))
        // (Count is a Long, as Excel's - which overflows past 2^31 cells; CountLarge is the Double.)
        elif eq name "Count" then
            let cells = float rows * float cols
            if cells > 2147483647.0 then fail 6 "Overflow" else Some (VLong (int cells))
        elif eq name "CountLarge" then Some (VDouble (float rows * float cols))
        elif eq name "Rows" then Some (lines false)
        elif eq name "Columns" then Some (lines true)
        elif eq name "Cells" then Some (rangeObject book r)
        elif eq name "Worksheet" || eq name "Parent" then Some (worksheetObject book r.Sheet)
        elif eq name "EntireRow" then Some (sub r.Top 0 r.Bottom maxCol)
        elif eq name "EntireColumn" then Some (sub 0 r.Left maxRow r.Right)
        elif eq name "CurrentRegion" then Some (currentRegion ())
        elif eq name "NumberFormat" || eq name "NumberFormatLocal" then Some (VString (format().NumberFormat |> Option.defaultValue "General"))
        elif eq name "Font" then Some (fontObject ())
        elif eq name "Interior" then Some (interiorObject ())
        elif eq name "Borders" then Some (borderObject None)
        elif eq name "FormatConditions" then Some (formatConditions ())
        elif eq name "DisplayFormat" then Some (displayFormatObject ())
        elif eq name "Comment" then Some (if (noteOf ()).IsSome then commentObject () else VNothing)
        elif eq name "ClearComments" then
            let inside (a: CellAddress) = a.Row >= r.Top && a.Row <= r.Bottom && a.Col >= r.Left && a.Col <= r.Right
            let clear (a: CellAddress) (data: CellData) =
                if inside a && data.Comment.IsSome then { data with Comment = None; Format = { data.Format with HasComment = false } } else data
            book.ChangeModel(r.Sheet, fun model -> { model with Cells = model.Cells |> Map.map clear })
            Some VEmpty
        elif eq name "ID" then Some (book.Setting (idKey ()) |> Option.defaultValue (VString ""))
        elif eq name "Locked" then Some (book.Setting ("__locked" + idKey ()) |> Option.defaultValue (VBoolean true))
        elif eq name "HorizontalAlignment" then Some (VLong (alignmentValue (format().HorizontalAlign) true))
        elif eq name "VerticalAlignment" then Some (VLong (alignmentValue (format().VerticalAlign) false))
        elif eq name "WrapText" then Some (VBoolean (format().WrapText))
        elif eq name "IndentLevel" then Some (VLong (format().Indent))
        elif eq name "Orientation" then
            Some (VLong (match format().TextRotation with 0 -> -4128 | 255 -> -4166 | d when d > 90 -> -(d - 90) | d -> d))
        elif eq name "MergeCells" then
            Some (VBoolean (book.Model(r.Sheet).MergedCells |> List.exists (fun m -> m.StartRow <= r.Top && m.EndRow >= r.Top && m.StartCol <= r.Left && m.EndCol >= r.Left)))
        elif eq name "ColumnWidth" then
            let px = book.Model(r.Sheet).ColumnWidths |> Map.tryFind r.Left |> Option.defaultValue (book.Model(r.Sheet).DefaultColumnWidth)
            Some (VDouble (if px < 12.0 then Math.Round(px / 12.0, 2) else Math.Round((px - 5.0) / 7.0, 2)))
        elif eq name "RowHeight" then
            let px = book.Model(r.Sheet).RowHeights |> Map.tryFind r.Top |> Option.defaultValue (book.Model(r.Sheet).DefaultRowHeight)
            Some (VDouble (Math.Round(px * 0.75, 2)))
        elif eq name "Hidden" then
            Some (VBoolean (if wholeCols then book.Model(r.Sheet).ColumnWidths |> Map.tryFind r.Left = Some 0.0 else book.Model(r.Sheet).RowHeights |> Map.tryFind r.Top = Some 0.0))
        elif eq name "Areas" then Some (collection "Areas" (fun n -> if eq n "Count" then Some (VLong 1) else unsupported "Areas" n) (fun n _ -> unsupported "Areas" n) (fun n args -> if (n = "" || eq n "Item") && args.Length = 1 then Some (rangeObject book r) else unsupported "Areas" n) (fun () -> Seq.singleton (rangeObject book r)))
        elif eq name "Next" then Some (sub r.Top (r.Left + 1) r.Top (r.Left + 1))
        elif eq name "Previous" then Some (sub r.Top (max 0 (r.Left - 1)) r.Top (max 0 (r.Left - 1)))
        elif eq name "Application" then Some (VString "Microsoft Excel")
        // Methods called without parentheses read as members.
        elif eq name "ClearContents" then clearContents (); Some VEmpty
        elif eq name "Clear" then clearContents (); clearFormats (); Some VEmpty
        elif eq name "ClearFormats" then clearFormats (); Some VEmpty
        elif eq name "Select" || eq name "Activate" then
            book.Selection <- Some (r.Sheet, r.Top, r.Left, r.Bottom, r.Right)
            book.SetActiveSheet r.Sheet
            Some VEmpty
        else None
    let set (name: string) (value: Value) : bool =
        if eq name "Value" || eq name "Value2" || eq name "_Default" || name = "" || eq name "Formula" || eq name "FormulaLocal" then
            assign value (fun row col item -> entered book r.Sheet row col item); true
        elif eq name "FormulaR1C1" then
            assign value (fun row col item ->
                match scalarOf item with
                | VString text when text.StartsWith "=" -> entered book r.Sheet row col (VString (fromR1C1 row col text))
                | other -> entered book r.Sheet row col other)
            true
        elif eq name "NumberFormat" || eq name "NumberFormatLocal" then
            let code = toText value
            changeFormats (fun f -> { f with NumberFormat = (if eq code "General" then None else Some code) }); true
        elif eq name "HorizontalAlignment" then changeFormats (fun f -> { f with HorizontalAlign = alignmentName value true }); true
        elif eq name "VerticalAlignment" then changeFormats (fun f -> { f with VerticalAlign = alignmentName value false }); true
        elif eq name "WrapText" then changeFormats (fun f -> { f with WrapText = toBool value }); true
        elif eq name "IndentLevel" then changeFormats (fun f -> { f with Indent = toInt value }); true
        elif eq name "Orientation" then
            // xlHorizontal, xlVertical (stacked), xlUpward, xlDownward, or degrees (-90 .. 90).
            let rotation =
                match toInt value with
                | -4128 -> 0
                | -4166 -> 255
                | -4171 -> 90
                | -4170 -> 180
                | d when d >= 0 && d <= 90 -> d
                | d when d < 0 && d >= -90 -> 90 - d
                | _ -> failApp "Unable to set the Orientation property of the Range class"
            changeFormats (fun f -> { f with TextRotation = rotation }); true
        elif eq name "ColumnWidth" then
            // (Excel: a width under one character is 12 px a character; from one on, 7 px each and 5 of padding.)
            let width = toFloat value
            let px = if width <= 0.0 then 0.0 elif width < 1.0 then Math.Round(width * 12.0) else Math.Floor((256.0 * width + 18.0) / 256.0 * 7.0) + 5.0
            if wholeRows then
                // Every column: the sheet's default width.
                book.ChangeModel(r.Sheet, fun m -> { m with DefaultColumnWidth = px; ColumnWidths = Map.empty })
            else
                book.ChangeModel(r.Sheet, fun m -> { m with ColumnWidths = [ r.Left .. r.Right ] |> List.fold (fun acc c -> Map.add c px acc) m.ColumnWidths })
            (if not wholeRows then book.Grow(r.Sheet, 0, r.Right))
            true
        elif eq name "RowHeight" then
            let px = toFloat value / 0.75
            if wholeCols then
                book.ChangeModel(r.Sheet, fun m -> { m with DefaultRowHeight = px; RowHeights = Map.empty })
            else
                book.ChangeModel(r.Sheet, fun m -> { m with RowHeights = [ r.Top .. min r.Bottom (r.Top + 100000) ] |> List.fold (fun acc row -> Map.add row px acc) m.RowHeights })
            (if not wholeCols then book.Grow(r.Sheet, r.Bottom, 0))
            true
        elif eq name "Hidden" then
            let hidden = toBool value
            if wholeCols then
                book.ChangeModel(r.Sheet, fun m -> { m with ColumnWidths = [ r.Left .. r.Right ] |> List.fold (fun acc c -> if hidden then Map.add c 0.0 acc else Map.remove c acc) m.ColumnWidths })
            else
                book.ChangeModel(r.Sheet, fun m -> { m with RowHeights = [ r.Top .. r.Bottom ] |> List.fold (fun acc row -> if hidden then Map.add row 0.0 acc else Map.remove row acc) m.RowHeights })
            true
        elif eq name "ID" then book.SetSetting(idKey (), VString (toText value)); true
        elif eq name "Locked" || eq name "FormulaHidden" then book.SetSetting("__locked" + idKey (), VBoolean (toBool value)); true
        elif eq name "MergeCells" then
            (if toBool value then (rangeObject book r |> function VHostObject h -> h.CallMethod "Merge" [||] |> ignore | _ -> ())
             else (rangeObject book r |> function VHostObject h -> h.CallMethod "UnMerge" [||] |> ignore | _ -> ()))
            true
        else unsupported "Range" name
    let call (name: string) (args: Value array) : Value option =
        let arg i = if i < args.Length then args.[i] else VEmpty
        if (name = "" || eq name "_Default" || eq name "Item" || eq name "Cells") && args.Length >= 1 then
            // Cells(row, col) and Item(index): relative to the range, by rows then columns.
            if args.Length >= 2 && not (isMissing args.[1]) then
                let row = r.Top + toInt args.[0] - 1
                let col =
                    match scalarOf args.[1] with
                    | VString letters when letters <> "" && Char.IsLetter letters.[0] ->
                        match parseReference (letters + "1") with
                        | Some (_, _, c, _, _) -> c
                        | None -> fail 13 "Type mismatch"
                    | v -> r.Left + toInt v - 1
                Some (sub row col row col)
            else
                let index = toInt args.[0] - 1
                let row = r.Top + index / cols
                let col = r.Left + index % cols
                Some (sub row col row col)
        elif name = "" && args.Length = 0 then Some (gather valueAt)
        elif eq name "Range" && args.Length >= 1 then
            match parseReference (toText args.[0]) with
            | Some (_, a, b, c, d) -> Some (sub (r.Top + a) (r.Left + b) (r.Top + c) (r.Left + d))
            | None -> failApp "Method 'Range' of object 'Range' failed"
        elif eq name "Offset" then
            let dr = if isMissing (arg 0) then 0 else toInt (arg 0)
            let dc = if isMissing (arg 1) then 0 else toInt (arg 1)
            if r.Top + dr < 0 || r.Left + dc < 0 || r.Bottom + dr > maxRow || r.Right + dc > maxCol then failApp "Application-defined or object-defined error"
            Some (sub (r.Top + dr) (r.Left + dc) (r.Bottom + dr) (r.Right + dc))
        elif eq name "Resize" then
            let rs = if isMissing (arg 0) then rows else toInt (arg 0)
            let cs = if isMissing (arg 1) then cols else toInt (arg 1)
            if rs < 1 || cs < 1 then failApp "Application-defined or object-defined error"
            Some (sub r.Top r.Left (min maxRow (r.Top + rs - 1)) (min maxCol (r.Left + cs - 1)))
        elif eq name "End" then Some (endOf (toInt (arg 0)))
        elif eq name "Address" then
            let rowAbs = if isMissing (arg 0) then true else toBool (arg 0)
            let colAbs = if isMissing (arg 1) then true else toBool (arg 1)
            let r1c1 = args.Length >= 3 && not (isMissing (arg 2)) && toInt (arg 2) = -4150
            let external = args.Length >= 4 && not (isMissing (arg 3)) && toBool (arg 3)
            let text = rangeAddress r rowAbs colAbs r1c1
            Some (VString (if external then "[" + book.Hooks.FileName + "]" + quotedSheet (sheetName ()) + "!" + text else text))
        elif eq name "ClearContents" then clearContents (); Some VEmpty
        elif eq name "Clear" then clearContents (); clearFormats (); Some VEmpty
        elif eq name "ClearFormats" then clearFormats (); Some VEmpty
        elif eq name "Select" || eq name "Activate" then
            book.Selection <- Some (r.Sheet, r.Top, r.Left, r.Bottom, r.Right)
            book.SetActiveSheet r.Sheet
            Some VEmpty
        elif eq name "Delete" || eq name "Insert" then
            let inserting = eq name "Insert"
            if wholeRows then book.Structure(r.Sheet, (if inserting then InsertRows (r.Top, rows) else DeleteRows (r.Top, rows)))
            elif wholeCols then book.Structure(r.Sheet, (if inserting then InsertCols (r.Left, cols) else DeleteCols (r.Left, cols)))
            else
                // Cells shifted within the sheet: up / left on Delete, down / right on Insert.
                let shift = if isMissing (arg 0) then (if inserting then -4121 else -4162) else toInt (arg 0)
                let model = book.Model r.Sheet
                let horizontal = shift = -4159 || shift = -4161
                let moved =
                    model.Cells |> Map.toList
                    |> List.filter (fun (a, _) ->
                        if horizontal then a.Row >= r.Top && a.Row <= r.Bottom && a.Col >= r.Left
                        else a.Col >= r.Left && a.Col <= r.Right && a.Row >= r.Top)
                for (a, _) in moved do book.SetCell(r.Sheet, a.Row, a.Col, None)
                for (a, cell) in moved do
                    let inside = a.Row >= r.Top && a.Row <= r.Bottom && a.Col >= r.Left && a.Col <= r.Right
                    let row, col =
                        if inserting then (if horizontal then a.Row, a.Col + cols else a.Row + rows, a.Col)
                        else (if horizontal then a.Row, a.Col - cols else a.Row - rows, a.Col)
                    if inserting || not inside then
                        if row >= 0 && col >= 0 && row <= maxRow && col <= maxCol then book.SetCell(r.Sheet, row, col, Some cell)
            Some VEmpty
        elif eq name "Copy" then
            match args |> Array.tryHead |> Option.bind (scalarOf >> fun _ -> rangeRefOf args.[0]) with
            | Some target -> copyTo target -4104; Some VEmpty
            | None ->
                book.SetSetting("__Clipboard", rangeObject book r)
                Some VEmpty
        elif eq name "PasteSpecial" then
            match book.Setting "__Clipboard" |> Option.bind rangeRefOf with
            | Some source ->
                let what = if isMissing (arg 0) then -4104 else toInt (arg 0)
                (match rangeObject book source with
                 | VHostObject h -> h.CallMethod "__CopyTo" [| rangeObject book r; VLong what |] |> ignore
                 | _ -> ())
                Some VEmpty
            | None -> failApp "PasteSpecial method of Range class failed"
        elif eq name "__CopyTo" then
            match rangeRefOf args.[0] with
            | Some target -> copyTo target (toInt args.[1]); Some VEmpty
            | None -> None
        elif eq name "Merge" then
            book.ChangeModel(r.Sheet, fun m ->
                let others = m.MergedCells |> List.filter (fun x -> x.EndRow < r.Top || x.StartRow > r.Bottom || x.EndCol < r.Left || x.StartCol > r.Right)
                { m with MergedCells = { StartRow = r.Top; StartCol = r.Left; EndRow = r.Bottom; EndCol = r.Right } :: others })
            Some VEmpty
        elif eq name "UnMerge" then
            book.ChangeModel(r.Sheet, fun m -> { m with MergedCells = m.MergedCells |> List.filter (fun x -> x.EndRow < r.Top || x.StartRow > r.Bottom || x.EndCol < r.Left || x.StartCol > r.Right) })
            Some VEmpty
        elif eq name "AutoFit" then Some VEmpty
        elif eq name "Borders" && args.Length = 1 then Some (borderObject (Some (toInt args.[0])))
        elif eq name "FormatConditions" && args.Length = 1 then
            match formatConditions () with VHostObject h -> h.CallMethod "Item" args | _ -> None
        elif eq name "Replace" then Some (replace args)
        elif eq name "BorderAround" then
            let lineStyle = if isMissing (arg 0) then xlContinuous else toInt (arg 0)
            let weight = if isMissing (arg 1) then 2 else toInt (arg 1)
            let color =
                if args.Length > 3 && not (isMissing (arg 3)) then Some (longToColor (toInt (arg 3)))
                elif args.Length > 4 && not (isMissing (arg 4)) then Some (themeColor (toInt (arg 4)))
                elif args.Length > 2 && not (isMissing (arg 2)) then (let i = toInt (arg 2) in if i >= 1 && i <= 56 then Some palette.[i - 1] else None)
                else None
            let border = styleOfLine lineStyle weight |> Option.map (fun style -> { Style = style; Color = color })
            for edge in [ 7; 8; 9; 10 ] do changeBorders (Some edge) (fun _ -> border)
            Some (VBoolean true)
        elif eq name "AddComment" then
            if (noteOf ()).IsSome then failApp "Application-defined or object-defined error"
            setNote (Some (plainNote (if args.Length > 0 && not (isMissing args.[0]) then toText args.[0] else "") None))
            Some (commentObject ())
        elif eq name "Find" then
            let what = toText (arg 0)
            let whole = args.Length > 4 && not (isMissing (arg 4)) && toInt (arg 4) = 1
            let matchCase = args.Length > 6 && not (isMissing (arg 6)) && toBool (arg 6)
            let comparison = if matchCase then StringComparison.Ordinal else StringComparison.OrdinalIgnoreCase
            let found =
                book.Model(r.Sheet).Cells |> Map.toSeq
                |> Seq.filter (fun (a, _) -> a.Row >= r.Top && a.Row <= r.Bottom && a.Col >= r.Left && a.Col <= r.Right)
                |> Seq.sortBy (fun (a, _) -> a.Row, a.Col)
                |> Seq.tryFind (fun (a, _) ->
                    let text = toText (valueAt a.Row a.Col)
                    if whole then String.Equals(text, what, comparison)
                    elif matchCase then text.Contains what
                    else text.ToLowerInvariant().Contains(what.ToLowerInvariant()))
            Some (match found with Some (a, _) -> sub a.Row a.Col a.Row a.Col | None -> VNothing)
        elif eq name "Calculate" then Some VEmpty
        else
            match get name with
            | Some value when args.Length = 0 -> Some value
            | _ -> unsupported "Range" name
    // For Each over a range: its cells, row by row (whole rows / columns only as far as the sheet is used).
    let items () =
        let bottom, right =
            if wholeRows || wholeCols then
                match used () with
                | Some (_, _, usedBottom, usedRight) -> min r.Bottom usedBottom, min r.Right usedRight
                | None -> r.Top - 1, r.Left - 1
            else r.Bottom, r.Right
        seq { for row in r.Top .. bottom do
                for col in r.Left .. right -> sub row col row col }
    collection "Range" (fun name -> match get name with Some v -> Some v | None -> unsupported "Range" name) set call items

/// Several areas (Union of separate blocks): its Address joins theirs ($A$1,$C$3), Areas lists them, Count
/// and For Each go over all their cells; anything else is its first area's, as Excel reads a range.
and private areasObject (book: MacroWorkbook) (areas: RangeRef list) : Value =
    match areas with
    | [ single ] -> rangeObject book single
    | [] -> VNothing
    | first :: _ ->
        let firstHost () = match rangeObject book first with VHostObject h -> h | _ -> failwith "range"
        let areaList () =
            collection "Areas"
                (fun n -> if eq n "Count" then Some (VLong areas.Length) else unsupported "Areas" n)
                (fun n _ -> unsupported "Areas" n)
                (fun n args -> if (n = "" || eq n "_Default" || eq n "Item") && args.Length = 1 then Some (rangeObject book areas.[toInt args.[0] - 1]) else unsupported "Areas" n)
                (fun () -> areas |> List.map (rangeObject book) |> Seq.ofList)
        let encoded () = String.Join("\n", [| for a in areas -> String.Join("\t", [| a.Sheet; string a.Top; string a.Left; string a.Bottom; string a.Right |]) |])
        collection "Range"
            (fun n ->
                if n = "__RbtAreas" then Some (VString (encoded ()))
                elif eq n "Address" then Some (VString (String.Join(",", areas |> List.map (fun a -> rangeAddress a true true false))))
                elif eq n "Count" || eq n "CountLarge" then Some (VDouble (areas |> List.sumBy (fun a -> float (a.Bottom - a.Top + 1) * float (a.Right - a.Left + 1))))
                elif eq n "Areas" then Some (areaList ())
                else (firstHost ()).GetMember n)
            (fun n value ->
                // (A value or format given to a union goes to every area.)
                areas |> List.iter (fun a -> match rangeObject book a with VHostObject h -> h.SetMember n value |> ignore | _ -> ())
                true)
            (fun n args ->
                if eq n "Address" then
                    let rowAbs = if args.Length > 0 && not (isMissing args.[0]) then toBool args.[0] else true
                    let colAbs = if args.Length > 1 && not (isMissing args.[1]) then toBool args.[1] else true
                    Some (VString (String.Join(",", areas |> List.map (fun a -> rangeAddress a rowAbs colAbs false))))
                elif eq n "ClearContents" || eq n "Clear" || eq n "ClearFormats" then
                    areas |> List.iter (fun a -> match rangeObject book a with VHostObject h -> h.CallMethod n args |> ignore | _ -> ())
                    Some VEmpty
                else (firstHost ()).CallMethod n args)
            (fun () -> areas |> Seq.collect (fun a -> match rangeObject book a with VHostObject h -> (h.Enumerate () |> Option.defaultValue Seq.empty) | _ -> Seq.empty))

and private worksheetObject (book: MacroWorkbook) (sheetId: SheetId) : Value =
    let whole = { Sheet = sheetId; Top = 0; Left = 0; Bottom = maxRow; Right = maxCol }
    let rangeOf (args: Value array) =
        match args with
        | [| first |] ->
            match rangeRefOf first with
            | Some r -> rangeObject book { r with Sheet = sheetId }
            | None ->
                let text = toText first
                match parseReference text with
                | Some (_, a, b, c, d) -> rangeObject book { Sheet = sheetId; Top = a; Left = b; Bottom = c; Right = d }
                | None ->
                    // A defined name.
                    match tryFindDefinedName (Some sheetId) text book.Workbook |> Option.orElse (tryFindDefinedName None text book.Workbook) with
                    | Some name ->
                        match parseReference (name.RefersTo.TrimStart '=') with
                        | Some (sheetName, a, b, c, d) ->
                            let target = sheetName |> Option.bind (sheetIdNamed book) |> Option.defaultValue sheetId
                            rangeObject book { Sheet = target; Top = a; Left = b; Bottom = c; Right = d }
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
            rangeObject book { Sheet = sheetId; Top = min a0 a1; Left = min b0 b1; Bottom = max c0 c1; Right = max d0 d1 }
        | _ -> fail 450 "Wrong number of arguments or invalid property assignment"
    let usedRange () =
        let model = book.Model sheetId
        if Map.isEmpty model.Cells then rangeObject book { whole with Bottom = 0; Right = 0 }
        else
            let addresses = model.Cells |> Map.toSeq |> Seq.map fst |> Seq.toList
            rangeObject book { Sheet = sheetId
                               Top = addresses |> List.map (fun a -> a.Row) |> List.min; Left = addresses |> List.map (fun a -> a.Col) |> List.min
                               Bottom = addresses |> List.map (fun a -> a.Row) |> List.max; Right = addresses |> List.map (fun a -> a.Col) |> List.max }
    let lineRange (vertical: bool) (args: Value array) =
        match args with
        | [||] -> rangeObject book whole
        | [| index |] ->
            match scalarOf index with
            | VString text when text.Contains ":" ->
                match parseReference text with
                | Some (_, a, b, c, d) -> rangeObject book { Sheet = sheetId; Top = a; Left = b; Bottom = c; Right = d }
                | None -> failApp "Application-defined or object-defined error"
            | VString letters when not vertical |> not && letters <> "" && Char.IsLetter letters.[0] ->
                match parseReference (letters + "1") with
                | Some (_, _, c, _, _) -> rangeObject book { whole with Left = c; Right = c }
                | None -> failApp "Application-defined or object-defined error"
            | v ->
                let i = toInt v - 1
                if vertical then rangeObject book { whole with Left = i; Right = i } else rangeObject book { whole with Top = i; Bottom = i }
        | _ -> fail 450 "Wrong number of arguments or invalid property assignment"
    let lines (vertical: bool) =
        collection "Range"
            (fun name ->
                if eq name "Count" then Some (VLong (if vertical then maxCol + 1 else maxRow + 1))
                else match rangeObject book whole with VHostObject h -> h.GetMember name | _ -> None)
            (fun name value -> match rangeObject book whole with VHostObject h -> h.SetMember name value | _ -> false)
            (fun name args ->
                if name = "" || eq name "_Default" || eq name "Item" then Some (lineRange vertical args)
                else match rangeObject book whole with VHostObject h -> h.CallMethod name args | _ -> None)
            (fun () ->
                // (As far as the sheet is used.)
                let model = book.Model sheetId
                let last = if Map.isEmpty model.Cells then -1 else model.Cells |> Map.toSeq |> Seq.map (fun (a, _) -> if vertical then a.Col else a.Row) |> Seq.max
                seq { for i in 0 .. last -> lineRange vertical [| VLong (i + 1) |] })
    let get (name: string) =
        if eq name "Name" then Some (VString (book.Sheet sheetId).Name)
        elif eq name "CodeName" then Some (VString (book.Sheet sheetId).Name)
        elif eq name "Index" then Some (VLong ((book.SheetOrder |> List.findIndex ((=) sheetId)) + 1))
        elif eq name "Cells" then Some (rangeObject book whole)
        elif eq name "Rows" then Some (lines false)
        elif eq name "Columns" then Some (lines true)
        elif eq name "UsedRange" then Some (usedRange ())
        elif eq name "Shapes" then Some (shapesObject book sheetId false)
        elif eq name "ChartObjects" then Some (shapesObject book sheetId true)
        elif eq name "Visible" then Some (VLong (if (book.Sheet sheetId).Visible then -1 elif (book.Model sheetId).VeryHidden then 2 else 0))
        elif eq name "Parent" then Some (workbookObject book)
        elif eq name "Type" then Some (VLong -4167)
        elif eq name "Next" then
            let order = book.SheetOrder
            let at = order |> List.findIndex ((=) sheetId)
            Some (if at + 1 < order.Length then worksheetObject book order.[at + 1] else VNothing)
        elif eq name "Previous" then
            let order = book.SheetOrder
            let at = order |> List.findIndex ((=) sheetId)
            Some (if at > 0 then worksheetObject book order.[at - 1] else VNothing)
        elif eq name "Activate" || eq name "Select" then book.SetActiveSheet sheetId; Some VEmpty
        elif eq name "Calculate" then Some VEmpty
        // (Protection is kept for the macro to read back; the grid does not lock the sheet.)
        elif eq name "Protect" then book.SetSetting("__protected\t" + sheetId, VBoolean true); Some VEmpty
        elif eq name "Unprotect" then book.SetSetting("__protected\t" + sheetId, VBoolean false); Some VEmpty
        elif eq name "ProtectContents" || eq name "ProtectDrawingObjects" || eq name "ProtectScenarios" then
            Some (book.Setting ("__protected\t" + sheetId) |> Option.defaultValue (VBoolean false))
        elif eq name "Delete" then
            (match deleteSheet sheetId book.Workbook with
             | Some _ -> book.ChangeWorkbook(fun wb -> deleteSheet sheetId wb |> Option.defaultValue wb)
             | None -> failApp "Delete method of Worksheet class failed")
            Some VEmpty
        else None
    let set (name: string) (value: Value) =
        if eq name "Name" then
            let newName = toText value
            match sheetNameProblem newName with
            | Some problem -> failApp problem
            | None ->
                if sheetNameTaken (Some sheetId) newName book.Workbook then failApp "That name is already taken. Try a different one."
                book.ChangeWorkbook(renameSheetEverywhere sheetId newName); true
        elif eq name "Visible" then
            let v = match scalarOf value with VBoolean b -> (if b then -1 else 0) | other -> toInt other
            book.ChangeWorkbook(fun wb ->
                match Map.tryFind sheetId wb.Sheets with
                | Some sheet -> { wb with Sheets = Map.add sheetId { sheet with Visible = (v = -1); Model = { sheet.Model with VeryHidden = (v = 2) } } wb.Sheets }
                | None -> wb)
            true
        else unsupported "Worksheet" name
    let call (name: string) (args: Value array) =
        if eq name "Range" then Some (rangeOf args)
        elif eq name "Cells" then
            match args with
            | [||] -> Some (rangeObject book whole)
            | _ -> match rangeObject book whole with VHostObject h -> h.CallMethod "Cells" args | _ -> None
        elif eq name "Rows" then Some (lineRange false args)
        elif eq name "Columns" then Some (lineRange true args)
        elif eq name "Protect" || eq name "Unprotect" then get name
        elif (eq name "Shapes" || eq name "ChartObjects") && args.Length = 1 then
            match shapesObject book sheetId (eq name "ChartObjects") with
            | VHostObject h -> h.CallMethod "Item" args
            | _ -> None
        else
            match get name with
            | Some value when args.Length = 0 -> Some value
            | _ -> unsupported "Worksheet" name
    host "Worksheet" (fun name -> match get name with Some v -> Some v | None -> unsupported "Worksheet" name) set call

// ---- shapes and charts: the charts a macro makes (Shapes.AddChart, ChartObjects.Add) ----

/// The objects of a sheet (Shapes; ChartObjects: its charts only): by index or name, Count, AddChart /
/// AddChart2 / Add (a chart, drawn by the grid's chart renderer from the series the macro gives it).
and private shapesObject (book: MacroWorkbook) (sheetId: SheetId) (chartsOnly: bool) : Value =
    let typeName = if chartsOnly then "ChartObjects" else "Shapes"
    let drawings () =
        (book.Model sheetId).Drawings
        |> List.filter (fun d -> not chartsOnly || (match d.Content with KeptObject ("chart", _) -> true | _ -> false))
    let pointsToPixels (points: float) = points * 4.0 / 3.0
    /// The cell a position (in points from the sheet's top left) falls in, and the offset into it, in pixels.
    let cellAt (left: float) (top: float) =
        let model = book.Model sheetId
        let walk (target: float) (size: int -> float) =
            let mutable index = 0
            let mutable edge = 0.0
            while edge + size index <= target && index < 16383 do
                edge <- edge + size index
                index <- index + 1
            index, target - edge
        let col, dx = walk (pointsToPixels left) (fun c -> model.ColumnWidths |> Map.tryFind c |> Option.defaultValue model.DefaultColumnWidth)
        let row, dy = walk (pointsToPixels top) (fun r -> model.RowHeights |> Map.tryFind r |> Option.defaultValue model.DefaultRowHeight)
        { Row = row; Col = col }, dx, dy
    let addChart (chartType: int) (left: Value) (top: Value) (width: Value) (height: Value) =
        let number v fallback = if isMissing v then fallback else toFloat v
        // (Excel's default chart: 360 x 216 points; placed here at C2 when no position is given.)
        let w, h = pointsToPixels (number width 360.0), pointsToPixels (number height 216.0)
        let cell, dx, dy =
            if isMissing left && isMissing top then { Row = 1; Col = 2 }, 0.0, 0.0
            else cellAt (number left 0.0) (number top 0.0)
        let model = book.Model sheetId
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
        shapeObject book sheetId id
    let item (index: Value) =
        let list = drawings ()
        match scalarOf index with
        | VString name ->
            match list |> List.tryFind (fun d -> d.Name |> Option.exists (eq name)) with
            | Some d -> shapeObject book sheetId d.Id
            | None -> failApp "The item with the specified name wasn't found."
        | v ->
            let i = toInt v
            if i < 1 || i > list.Length then failApp "The index into the specified collection is out of bounds."
            shapeObject book sheetId list.[i - 1].Id
    collection typeName
        (fun name ->
            if eq name "Count" then Some (VLong (drawings ()).Length)
            elif eq name "AddChart" || eq name "Add" && chartsOnly then Some (addChart 51 VEmpty VEmpty VEmpty VEmpty)
            elif eq name "AddChart2" then Some (addChart 51 VEmpty VEmpty VEmpty VEmpty)
            else unsupported typeName name)
        (fun name _ -> unsupported typeName name)
        (fun name args ->
            let arg i = if i < args.Length then args.[i] else VEmpty
            if (name = "" || eq name "_Default" || eq name "Item") && args.Length = 1 then Some (item args.[0])
            elif eq name "Count" then Some (VLong (drawings ()).Length)
            // AddChart(Type, Left, Top, Width, Height); AddChart2(Style, XlChartType, Left, Top, Width, Height); ChartObjects.Add(Left, Top, Width, Height).
            elif eq name "AddChart" && not chartsOnly then Some (addChart (if isMissing (arg 0) then 51 else toInt (arg 0)) (arg 1) (arg 2) (arg 3) (arg 4))
            elif eq name "AddChart2" && not chartsOnly then Some (addChart (if isMissing (arg 1) then 51 else toInt (arg 1)) (arg 2) (arg 3) (arg 4) (arg 5))
            elif eq name "Add" && chartsOnly then Some (addChart 51 (arg 0) (arg 1) (arg 2) (arg 3))
            else unsupported typeName name)
        (fun () -> drawings () |> List.map (fun d -> shapeObject book sheetId d.Id) |> Seq.ofList)

/// A shape of a sheet (a chart's frame, a picture, a kept object): its name, place and size, its chart.
and private shapeObject (book: MacroWorkbook) (sheetId: SheetId) (id: int) : Value =
    let drawing () =
        match (book.Model sheetId).Drawings |> List.tryFind (fun d -> d.Id = id) with
        | Some d -> d
        | None -> failApp "The object has been deleted."
    let relOf (d: SheetDrawing) =
        match d.Content with
        | KeptObject ("chart", xml) ->
            let m = Text.RegularExpressions.Regex.Match(xml, "r:id=\"([^\"]+)\"")
            if m.Success then Some m.Groups.[1].Value else None
        | _ -> None
    let change (f: SheetDrawing -> SheetDrawing) =
        book.ChangeModel(sheetId, fun m -> { m with Drawings = m.Drawings |> List.map (fun d -> if d.Id = id then f d else d) })
    let get (name: string) : Value option =
        let d = drawing ()
        if eq name "Name" then Some (VString (d.Name |> Option.defaultValue (sprintf "Object %d" id)))
        elif eq name "Width" then Some (VDouble ((d.Anchor.Width |> Option.defaultValue 0.0) * 0.75))
        elif eq name "Height" then Some (VDouble ((d.Anchor.Height |> Option.defaultValue 0.0) * 0.75))
        elif eq name "Chart" then
            match relOf d with
            | Some rel when book.Charts.ContainsKey rel -> Some (chartObject book rel)
            | _ -> failApp "This object has no chart the macro can change"
        elif eq name "HasChart" then Some (VBoolean (relOf d).IsSome)
        elif eq name "Visible" then Some (VLong -1)
        elif eq name "ID" then Some (VLong id)
        elif eq name "Delete" then
            book.ChangeModel(sheetId, fun m ->
                { m with Drawings = m.Drawings |> List.filter (fun x -> x.Id <> id)
                         AuthoredCharts = match relOf d with Some rel -> Map.remove rel m.AuthoredCharts | None -> m.AuthoredCharts })
            Some VEmpty
        elif eq name "Select" || eq name "Activate" then Some VEmpty
        else None
    host "Shape"
        (fun name -> match get name with Some v -> Some v | None -> unsupported "Shape" name)
        (fun name value ->
            if eq name "Name" then change (fun d -> { d with Name = Some (toText value) }); true
            elif eq name "Width" then change (fun d -> { d with Anchor = { d.Anchor with Width = Some (toFloat value * 4.0 / 3.0) } }); true
            elif eq name "Height" then change (fun d -> { d with Anchor = { d.Anchor with Height = Some (toFloat value * 4.0 / 3.0) } }); true
            elif eq name "Visible" || eq name "Placement" then true
            else unsupported "Shape" name)
        (fun name args -> match get name with Some v when args.Length = 0 -> Some v | _ -> unsupported "Shape" name)

/// A chart a macro made: its type, title, legend and series (each change draws it again).
and private chartObject (book: MacroWorkbook) (rel: string) : Value =
    let current () = snd book.Charts.[rel]
    let update (f: Rbt.Spreadsheet.ChartAuthoring.MacroChart -> Rbt.Spreadsheet.ChartAuthoring.MacroChart) =
        let sheetId, chart = book.Charts.[rel]
        let next = f chart
        book.Charts.[rel] <- (sheetId, next)
        book.ChangeModel(sheetId, fun m -> { m with AuthoredCharts = Map.add rel (Rbt.Spreadsheet.ChartAuthoring.macroChartXml next) m.AuthoredCharts })
    /// What Values / XValues are given: a range (its formula and values) or an array (literal points).
    let dataOf (value: Value) : Rbt.Spreadsheet.ChartAuthoring.MacroSeriesData =
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
        match rangeRefOf value with
        | Some r ->
            let formula =
                quotedSheet (book.Sheet r.Sheet).Name + "!" + cellAddress r.Top r.Left true true
                + (if r.Top = r.Bottom && r.Left = r.Right then "" else ":" + cellAddress r.Bottom r.Right true true)
            let points =
                [| for row in r.Top .. r.Bottom do
                     for col in r.Left .. r.Right ->
                        match book.Cell(r.Sheet, row, col) with
                        | Some c -> point (cellValue c.Value)
                        | None -> "", None |]
            { Formula = Some formula; Points = points }
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
    let rec seriesObject (index: int) =
        let pick () =
            match List.tryItem (index - 1) (current ()).Series with
            | Some s -> s
            | None -> fail 9 "Subscript out of range"
        let change (f: Rbt.Spreadsheet.ChartAuthoring.MacroSeries -> Rbt.Spreadsheet.ChartAuthoring.MacroSeries) =
            update (fun chart -> { chart with Series = chart.Series |> List.mapi (fun i s -> if i = index - 1 then f s else s) })
        let points (d: Rbt.Spreadsheet.ChartAuthoring.MacroSeriesData) =
            VArray { Data = d.Points |> Array.map (fun (text, n) -> match n with Some x -> VDouble x | None -> VString text); LowerBounds = [| 1 |]; Lengths = [| d.Points.Length |] }
        host "Series"
            (fun name ->
                if eq name "Values" then Some (points (pick ()).YData)
                elif eq name "XValues" then Some ((pick ()).XData |> Option.map points |> Option.defaultValue (VArray { Data = [||]; LowerBounds = [| 1 |]; Lengths = [| 0 |] }))
                elif eq name "Name" then Some (VString ((pick ()).SeriesName |> Option.defaultValue (sprintf "Series%d" index)))
                elif eq name "Delete" then update (fun chart -> { chart with Series = chart.Series |> List.indexed |> List.filter (fun (i, _) -> i <> index - 1) |> List.map snd }); Some VEmpty
                else unsupported "Series" name)
            (fun name value ->
                if eq name "Values" then change (fun s -> { s with YData = dataOf value }); true
                elif eq name "XValues" then change (fun s -> { s with XData = Some (dataOf value) }); true
                elif eq name "Name" then change (fun s -> { s with SeriesName = Some (toText value) }); true
                elif eq name "ChartType" || eq name "MarkerStyle" || eq name "MarkerSize" || eq name "Smooth" then true
                else unsupported "Series" name)
            (fun name _ -> unsupported "Series" name)
    let newSeries () =
        update (fun chart -> { chart with Series = chart.Series @ [ Rbt.Spreadsheet.ChartAuthoring.macroSeries None None { Formula = None; Points = [||] } ] })
        seriesObject (current ()).Series.Length
    let seriesCollection =
        collection "SeriesCollection"
            (fun name ->
                if eq name "Count" then Some (VLong (current ()).Series.Length)
                elif eq name "NewSeries" then Some (newSeries ())
                else unsupported "SeriesCollection" name)
            (fun name _ -> unsupported "SeriesCollection" name)
            (fun name args ->
                if (name = "" || eq name "_Default" || eq name "Item") && args.Length = 1 then
                    match scalarOf args.[0] with
                    | VString text ->
                        match (current ()).Series |> List.tryFindIndex (fun s -> s.SeriesName |> Option.exists (eq text)) with
                        | Some i -> Some (seriesObject (i + 1))
                        | None -> fail 1004 "Invalid parameter"
                    | v -> Some (seriesObject (toInt v))
                elif eq name "NewSeries" then Some (newSeries ())
                elif eq name "Count" then Some (VLong (current ()).Series.Length)
                else unsupported "SeriesCollection" name)
            (fun () -> (current ()).Series |> List.mapi (fun i _ -> seriesObject (i + 1)) |> Seq.ofList)
    /// SetSourceData(range, PlotBy): a series of each column (xlColumns 2, the default for a tall range) or
    /// row, the first column / row the categories when it holds no numbers.
    let setSourceData (source: Value) (plotBy: Value) =
        match rangeRefOf source with
        | None -> failApp "Method 'SetSourceData' of object '_Chart' failed"
        | Some r ->
            let rows, cols = r.Bottom - r.Top + 1, r.Right - r.Left + 1
            let byColumns = if isMissing plotBy then rows >= cols else toInt plotBy = 2
            let numeric row col = match book.Cell(r.Sheet, row, col) with Some { Value = v } -> (match unwrapFormulaValue v with Number _ -> true | _ -> false) | None -> false
            let text row col = match book.Cell(r.Sheet, row, col) with Some c -> toText (cellValue c.Value) | None -> ""
            let range top left bottom right = rangeObject book { r with Top = top; Left = left; Bottom = bottom; Right = right }
            let series =
                if byColumns then
                    let header = rows > 1 && not (numeric r.Top (r.Right))
                    let firstRow = if header then r.Top + 1 else r.Top
                    let labels = cols > 1 && not (seq { firstRow .. r.Bottom } |> Seq.exists (fun row -> numeric row r.Left))
                    let firstCol = if labels then r.Left + 1 else r.Left
                    [ for col in firstCol .. r.Right ->
                        Rbt.Spreadsheet.ChartAuthoring.macroSeries (if header then Some (text r.Top col) else None)
                            (if labels then Some (dataOf (range firstRow r.Left r.Bottom r.Left)) else None)
                            (dataOf (range firstRow col r.Bottom col)) ]
                else
                    let header = cols > 1 && not (numeric r.Bottom r.Left)
                    let firstCol = if header then r.Left + 1 else r.Left
                    let labels = rows > 1 && not (seq { firstCol .. r.Right } |> Seq.exists (fun col -> numeric r.Top col))
                    let firstRow = if labels then r.Top + 1 else r.Top
                    [ for row in firstRow .. r.Bottom ->
                        Rbt.Spreadsheet.ChartAuthoring.macroSeries (if header then Some (text row r.Left) else None)
                            (if labels then Some (dataOf (range r.Top firstCol r.Top r.Right)) else None)
                            (dataOf (range row firstCol row r.Right)) ]
            update (fun chart -> { chart with Series = series })
    let titleObject () =
        host "ChartTitle"
            (fun name -> if eq name "Text" || eq name "Caption" then Some (VString ((current ()).Title |> Option.defaultValue "")) else unsupported "ChartTitle" name)
            (fun name value -> if eq name "Text" || eq name "Caption" then update (fun c -> { c with Title = Some (toText value) }); true else true)
            (fun name _ -> unsupported "ChartTitle" name)
    let get (name: string) : Value option =
        if eq name "ChartType" then Some (VLong (current ()).ChartType)
        elif eq name "HasLegend" then Some (VBoolean (current ()).HasLegend)
        elif eq name "HasTitle" then Some (VBoolean (current ()).Title.IsSome)
        elif eq name "ChartTitle" then Some (titleObject ())
        elif eq name "SeriesCollection" then Some seriesCollection
        elif eq name "Name" then Some (VString rel)
        elif eq name "Refresh" || eq name "Select" || eq name "Activate" then Some VEmpty
        else None
    host "Chart"
        (fun name -> match get name with Some v -> Some v | None -> unsupported "Chart" name)
        (fun name value ->
            if eq name "ChartType" then update (fun c -> { c with ChartType = toInt value }); true
            elif eq name "HasLegend" then update (fun c -> { c with HasLegend = toBool value }); true
            elif eq name "HasTitle" then update (fun c -> { c with Title = (if toBool value then Some (c.Title |> Option.defaultValue "Chart Title") else None) }); true
            elif eq name "ChartStyle" || eq name "PlotBy" then true
            else unsupported "Chart" name)
        (fun name args ->
            let arg i = if i < args.Length then args.[i] else VEmpty
            if eq name "SeriesCollection" then
                if args.Length = 0 then Some seriesCollection
                else match seriesCollection with VHostObject h -> h.CallMethod "Item" args | _ -> None
            elif eq name "SetSourceData" then setSourceData (arg 0) (arg 1); Some VEmpty
            else match get name with Some v when args.Length = 0 -> Some v | _ -> unsupported "Chart" name)

and private sheetsObject (book: MacroWorkbook) : Value =
    let item (index: Value) =
        match scalarOf index with
        | VString name ->
            match sheetIdNamed book name with
            | Some id -> worksheetObject book id
            | None -> fail 9 "Subscript out of range"
        | v ->
            let i = toInt v
            if i < 1 || i > book.SheetOrder.Length then fail 9 "Subscript out of range"
            worksheetObject book book.SheetOrder.[i - 1]
    let add (args: Value array) =
        let before = args |> Array.tryItem 0 |> Option.bind (fun v -> if isMissing v then None else Some v)
        let after = args |> Array.tryItem 1 |> Option.bind (fun v -> if isMissing v then None else Some v)
        let sheetIdOf (v: Value) =
            match v with
            | VHostObject h ->
                match h.GetMember "Name" with
                | Some (VString n) -> sheetIdNamed book n
                | _ -> None
            | _ -> None
        let name = nextSheetName book.Workbook
        let mutable added = ""
        book.ChangeWorkbook(fun wb ->
            let grown = addSheet name wb
            added <- grown.SheetOrder |> List.find (fun id -> not (List.contains id wb.SheetOrder))
            let without = grown.SheetOrder |> List.filter ((<>) added)
            let at =
                match before |> Option.bind sheetIdOf, after |> Option.bind sheetIdOf with
                | Some b, _ -> List.findIndex ((=) b) without
                | None, Some a -> List.findIndex ((=) a) without + 1
                | None, None -> List.findIndex ((=) wb.ActiveSheet) without
            let order = (List.take at without) @ [ added ] @ (List.skip at without)
            reorderSheets order grown |> setActiveSheet added)
        worksheetObject book added
    collection "Sheets"
        (fun name ->
            if eq name "Count" then Some (VLong book.SheetOrder.Length)
            elif eq name "Add" then Some (add [||])
            else unsupported "Sheets" name)
        (fun name _ -> unsupported "Sheets" name)
        (fun name args ->
            if (name = "" || eq name "_Default" || eq name "Item") && args.Length = 1 then Some (item args.[0])
            elif eq name "Add" then Some (add args)
            elif eq name "Count" then Some (VLong book.SheetOrder.Length)
            else unsupported "Sheets" name)
        (fun () -> book.SheetOrder |> List.map (worksheetObject book) |> Seq.ofList)

and private namesObject (book: MacroWorkbook) : Value =
    let names () = definedNames book.Workbook
    let nameObject (name: DefinedName) =
        host "Name"
            (fun n ->
                if eq n "Name" then Some (VString name.Name)
                elif eq n "RefersTo" || eq n "Value" || eq n "_Default" then Some (VString ("=" + name.RefersTo.TrimStart '='))
                elif eq n "RefersToRange" then
                    match parseReference (name.RefersTo.TrimStart '=') with
                    | Some (sheetName, a, b, c, d) ->
                        let target = sheetName |> Option.bind (sheetIdNamed book) |> Option.defaultValue book.ActiveSheet
                        Some (rangeObject book { Sheet = target; Top = a; Left = b; Bottom = c; Right = d })
                    | None -> failApp "Application-defined or object-defined error"
                elif eq n "Delete" then book.ChangeWorkbook(fun wb -> deleteDefinedName name.Scope name.Name wb |> Option.defaultValue wb); Some VEmpty
                else unsupported "Name" n)
            (fun n value ->
                if eq n "RefersTo" then book.ChangeWorkbook(setDefinedName name.Scope name.Name ((toText value).TrimStart '=') name.Comment); true
                else unsupported "Name" n)
            (fun n _ -> if eq n "Delete" then book.ChangeWorkbook(fun wb -> deleteDefinedName name.Scope name.Name wb |> Option.defaultValue wb); Some VEmpty else unsupported "Name" n)
    collection "Names"
        (fun n -> if eq n "Count" then Some (VLong (names ()).Length) else unsupported "Names" n)
        (fun n _ -> unsupported "Names" n)
        (fun n args ->
            if (n = "" || eq n "_Default" || eq n "Item") && args.Length >= 1 then
                match scalarOf args.[0] with
                | VString text ->
                    match names () |> List.tryFind (fun x -> eq x.Name text) with
                    | Some found -> Some (nameObject found)
                    | None -> fail 1004 "Application-defined or object-defined error"
                | v ->
                    let i = toInt v
                    match names () |> List.tryItem (i - 1) with
                    | Some found -> Some (nameObject found)
                    | None -> fail 9 "Subscript out of range"
            elif eq n "Add" && args.Length >= 2 then
                let name = toText args.[0]
                let refersTo =
                    match rangeRefOf args.[1] with
                    | Some r -> quotedSheet (book.Sheet r.Sheet).Name + "!" + cellAddress r.Top r.Left true true + (if r.Top = r.Bottom && r.Left = r.Right then "" else ":" + cellAddress r.Bottom r.Right true true)
                    | None -> (toText args.[1]).TrimStart '='
                match definedNameProblem name with
                | Some problem -> failApp problem
                | None ->
                    book.ChangeWorkbook(setDefinedName None name refersTo None)
                    Some (nameObject (tryFindDefinedName None name book.Workbook |> Option.get))
            else unsupported "Names" n)
        (fun () -> names () |> List.map nameObject |> Seq.ofList)

and private workbookObject (book: MacroWorkbook) : Value =
    let get (name: string) =
        if eq name "Name" then Some (VString book.Hooks.FileName)
        elif eq name "FullName" then Some (VString book.Hooks.FileName)
        elif eq name "Path" then Some (VString "")
        elif eq name "Worksheets" || eq name "Sheets" then Some (sheetsObject book)
        elif eq name "ActiveSheet" then Some (worksheetObject book book.ActiveSheet)
        elif eq name "Names" then Some (namesObject book)
        elif eq name "Saved" then Some (VBoolean false)
        elif eq name "Save" || eq name "Activate" then Some VEmpty
        else None
    host "Workbook"
        (fun name -> match get name with Some v -> Some v | None -> unsupported "Workbook" name)
        (fun name _ -> if eq name "Saved" then true else unsupported "Workbook" name)
        (fun name args ->
            if eq name "Worksheets" || eq name "Sheets" then
                match args with
                | [||] -> Some (sheetsObject book)
                | _ -> match sheetsObject book with VHostObject h -> h.CallMethod "" args | _ -> None
            else
                match get name with
                | Some value when args.Length = 0 -> Some value
                | _ -> unsupported "Workbook" name)

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

let private worksheetFunctionObject (book: MacroWorkbook) =
    host "WorksheetFunction"
        // (A function without arguments reads as a member: WorksheetFunction.Pi.)
        (fun name ->
            if name = "" || name.StartsWith "__" then unsupported "WorksheetFunction" name
            else
                match evaluate book (name.ToUpperInvariant() + "()") with
                | VError 2029 -> unsupported "WorksheetFunction" name
                | VError _ -> failApp (sprintf "Unable to get the %s property of the WorksheetFunction class" name)
                | value -> Some value)
        (fun name _ -> unsupported "WorksheetFunction" name)
        (fun name args ->
            let formula = name.ToUpperInvariant() + "(" + String.Join(",", args |> Array.map (formulaArgument book)) + ")"
            match evaluate book formula with
            | VError _ -> failApp (sprintf "Unable to get the %s property of the WorksheetFunction class" name)
            | value -> Some value)

/// Evaluate("A1:B2") is that range; anything else is a formula.
let private evaluateArgument (book: MacroWorkbook) (text: string) =
    let text = text.Trim().TrimStart '='
    match parseReference text with
    | Some (sheetName, a, b, c, d) ->
        let target = sheetName |> Option.bind (sheetIdNamed book) |> Option.defaultValue book.ActiveSheet
        rangeObject book { Sheet = target; Top = a; Left = b; Bottom = c; Right = d }
    | None -> evaluate book text

let private intersect (a: RangeRef) (b: RangeRef) =
    if a.Sheet <> b.Sheet then None
    else
        let top, left, bottom, right = max a.Top b.Top, max a.Left b.Left, min a.Bottom b.Bottom, min a.Right b.Right
        if top > bottom || left > right then None else Some { a with Top = top; Left = left; Bottom = bottom; Right = right }

/// The window of the workbook (ActiveWindow): gridlines, headings, zoom, frozen panes and scrolling of the active sheet.
let private windowObject (book: MacroWorkbook) =
    let model () = book.Model book.ActiveSheet
    let change f = book.ChangeModel(book.ActiveSheet, f)
    let get (name: string) =
        if eq name "DisplayGridlines" then Some (VBoolean (model ()).ShowGridLines)
        elif eq name "DisplayHeadings" then Some (VBoolean (model ()).ShowHeaders)
        elif eq name "DisplayFormulas" then Some (VBoolean (model ()).ShowFormulas)
        elif eq name "DisplayZeros" then Some (VBoolean (model ()).ShowZeros)
        elif eq name "Zoom" then Some (VDouble (Math.Round((model ()).Zoom * 100.0)))
        elif eq name "FreezePanes" then Some (VBoolean ((model ()).FrozenRows > 0 || (model ()).FrozenCols > 0))
        elif eq name "SplitRow" then Some (VLong (model ()).FrozenRows)
        elif eq name "SplitColumn" then Some (VLong (model ()).FrozenCols)
        elif eq name "ScrollRow" || eq name "ScrollColumn" then Some (book.Setting ("Window" + name) |> Option.defaultValue (VLong 1))
        elif eq name "WindowState" then Some (book.Setting "WindowState" |> Option.defaultValue (VLong -4137))
        elif eq name "Caption" then Some (VString book.Hooks.FileName)
        elif eq name "ActiveSheet" then Some (worksheetObject book book.ActiveSheet)
        elif eq name "ActiveCell" || eq name "Selection" || eq name "VisibleRange" then None
        elif eq name "SmallScroll" || eq name "LargeScroll" || eq name "Activate" then Some VEmpty
        else None
    let set (name: string) (value: Value) =
        if eq name "DisplayGridlines" then change (fun m -> { m with ShowGridLines = toBool value }); true
        elif eq name "DisplayHeadings" then change (fun m -> { m with ShowHeaders = toBool value }); true
        elif eq name "DisplayFormulas" then change (fun m -> { m with ShowFormulas = toBool value }); true
        elif eq name "DisplayZeros" then change (fun m -> { m with ShowZeros = toBool value }); true
        elif eq name "Zoom" then
            (match scalarOf value with
             | VBoolean _ -> ()
             | v -> let z = toFloat v in change (fun m -> { m with Zoom = (max 10.0 (min 400.0 z)) / 100.0 }))
            true
        elif eq name "FreezePanes" then
            if toBool value then
                // Above and to the left of the active cell, as Excel freezes them.
                let row, col =
                    match book.Selection with
                    | Some (sheet, top, left, _, _) when sheet = book.ActiveSheet -> top, left
                    | _ -> match (model ()).Selection with SingleCell (r, c) -> r, c | Range (r0, c0, _, _) -> r0, c0
                change (fun m -> { m with FrozenRows = row; FrozenCols = col })
            else change (fun m -> { m with FrozenRows = 0; FrozenCols = 0 })
            true
        elif eq name "SplitRow" then change (fun m -> { m with FrozenRows = toInt value }); true
        elif eq name "SplitColumn" then change (fun m -> { m with FrozenCols = toInt value }); true
        elif eq name "ScrollRow" || eq name "ScrollColumn" || eq name "WindowState" || eq name "Split" || eq name "Caption"
             || eq name "DisplayWorkbookTabs" || eq name "DisplayHorizontalScrollBar" || eq name "DisplayVerticalScrollBar" then
            book.SetSetting((if eq name "ScrollRow" || eq name "ScrollColumn" then "Window" + name else name), value); true
        else unsupported "Window" name
    host "Window"
        (fun name -> match get name with Some v -> Some v | None -> unsupported "Window" name)
        set
        (fun name args ->
            match get name with
            | Some value -> Some value
            | None -> unsupported "Window" name)

let private applicationObject (book: MacroWorkbook) =
    let activeRange () =
        match book.Selection with
        | Some (sheet, top, left, bottom, right) when sheet = book.ActiveSheet -> Some { Sheet = sheet; Top = top; Left = left; Bottom = bottom; Right = right }
        | _ ->
            let model = book.Model book.ActiveSheet
            match model.Selection with
            | SingleCell (row, col) -> Some { Sheet = book.ActiveSheet; Top = row; Left = col; Bottom = row; Right = col }
            | Range (r0, c0, r1, c1) -> Some { Sheet = book.ActiveSheet; Top = min r0 r1; Left = min c0 c1; Bottom = max r0 r1; Right = max c0 c1 }
    let get (name: string) =
        if eq name "ActiveWorkbook" || eq name "ThisWorkbook" then Some (workbookObject book)
        elif eq name "ActiveSheet" then Some (worksheetObject book book.ActiveSheet)
        elif eq name "Worksheets" || eq name "Sheets" then Some (sheetsObject book)
        elif eq name "Workbooks" then
            Some (collection "Workbooks"
                    (fun n -> if eq n "Count" then Some (VLong 1) else unsupported "Workbooks" n)
                    (fun n _ -> unsupported "Workbooks" n)
                    (fun n args -> if (n = "" || eq n "Item") && args.Length = 1 then Some (workbookObject book) else unsupported "Workbooks" n)
                    (fun () -> Seq.singleton (workbookObject book)))
        elif eq name "Selection" then Some (activeRange () |> Option.map (rangeObject book) |> Option.defaultValue VNothing)
        elif eq name "ActiveCell" then
            Some (activeRange () |> Option.map (fun r -> rangeObject book { r with Bottom = r.Top; Right = r.Left }) |> Option.defaultValue VNothing)
        elif eq name "WorksheetFunction" then Some (worksheetFunctionObject book)
        elif eq name "Names" then Some (namesObject book)
        elif eq name "ActiveWindow" then Some (windowObject book)
        elif eq name "Version" then Some (VString "16.0")
        elif eq name "Name" then Some (VString "Microsoft Excel")
        elif eq name "Application" then Some (VString "Microsoft Excel")
        elif eq name "UserName" then Some (VString "")
        elif eq name "Calculation" then Some (book.Setting "Calculation" |> Option.defaultValue (VLong -4105))
        elif eq name "Calculate" || eq name "CalculateFull" then Some VEmpty
        elif ["ScreenUpdating"; "EnableEvents"; "DisplayAlerts"; "StatusBar"; "CutCopyMode"; "DisplayStatusBar"; "Interactive"; "EnableCancelKey"; "AskToUpdateLinks"]
             |> List.exists (eq name) then
            Some (book.Setting name |> Option.defaultValue (if eq name "StatusBar" then VBoolean false else VBoolean true))
        else None
    host "Application"
        (fun name -> match get name with Some v -> Some v | None -> unsupported "Application" name)
        (fun name value ->
            if ["ScreenUpdating"; "EnableEvents"; "DisplayAlerts"; "StatusBar"; "CutCopyMode"; "Calculation"; "DisplayStatusBar"; "Interactive"; "EnableCancelKey"; "AskToUpdateLinks"]
               |> List.exists (eq name) then
                book.SetSetting(name, value); true
            else unsupported "Application" name)
        (fun name args ->
            if eq name "Range" then
                match worksheetObject book book.ActiveSheet with
                | VHostObject h -> h.CallMethod "Range" args
                | _ -> None
            elif eq name "Cells" then
                match worksheetObject book book.ActiveSheet with
                | VHostObject h -> h.CallMethod "Cells" args
                | _ -> None
            elif eq name "Worksheets" || eq name "Sheets" then
                match args with
                | [||] -> Some (sheetsObject book)
                | _ -> match sheetsObject book with VHostObject h -> h.CallMethod "" args | _ -> None
            elif eq name "Evaluate" && args.Length = 1 then Some (evaluateArgument book (toText args.[0]))
            elif eq name "Intersect" then
                let refs = args |> Array.choose rangeRefOf
                if refs.Length < 2 then failApp "Method 'Intersect' of object '_Application' failed"
                match refs |> Array.skip 1 |> Array.fold (fun acc next -> acc |> Option.bind (fun a -> intersect a next)) (Some refs.[0]) with
                | Some r -> Some (rangeObject book r)
                | None -> Some VNothing
            elif eq name "Union" then
                // Separate blocks stay separate areas; blocks one contains are dropped.
                let refs = args |> Array.filter (isMissing >> not) |> Array.toList |> List.collect areasOf
                if refs.IsEmpty then failApp "Method 'Union' of object '_Application' failed"
                let contains (a: RangeRef) (b: RangeRef) = a.Sheet = b.Sheet && a.Top <= b.Top && a.Left <= b.Left && a.Bottom >= b.Bottom && a.Right >= b.Right
                let kept =
                    refs |> List.distinct
                    |> fun all -> all |> List.filter (fun r -> not (all |> List.exists (fun other -> other <> r && contains other r)))
                Some (areasObject book kept)
            elif eq name "Goto" && args.Length >= 1 then
                match rangeRefOf args.[0] with
                | Some r -> book.SetActiveSheet r.Sheet; book.Selection <- Some (r.Sheet, r.Top, r.Left, r.Bottom, r.Right); Some VEmpty
                | None -> Some VEmpty
            elif eq name "Wait" then
                // (A run a slice at a time waits until then: MacroProgram.WaitMs; a run to the end does not.)
                let target = if args.Length > 0 then toFloat args.[0] else 0.0
                let now = (DateTime.Now - DateTime(1899, 12, 30)).TotalDays
                let until = if target < 1.0 then Math.Floor now + target else target
                book.SetSetting("__waitUntil", VDouble until)
                Some (VBoolean true)
            elif eq name "OnTime" then Some VEmpty
            elif eq name "InputBox" then
                let prompt = if args.Length > 0 then toText args.[0] else ""
                let title = if args.Length > 1 && not (isMissing args.[1]) then toText args.[1] else "Microsoft Excel"
                let defaultText = if args.Length > 2 && not (isMissing args.[2]) then toText args.[2] else ""
                Some (match book.Hooks.InputBox prompt title defaultText with Some text -> VString text | None -> VBoolean false)
            elif eq name "WorksheetFunction" then Some (worksheetFunctionObject book)
            else
                match get name with
                | Some value when args.Length = 0 -> Some value
                | _ when args.Length > 0 ->
                    // Application.Max(...), Application.Text(...): the worksheet function, its error as an Error value.
                    let formula = name.ToUpperInvariant() + "(" + String.Join(",", args |> Array.map (formulaArgument book)) + ")"
                    match evaluate book formula with
                    | VError 2029 -> unsupported "Application" name
                    | value -> Some value
                | _ -> unsupported "Application" name)

/// A range object of the workbook (an event's Target).
let private rangeValue (book: MacroWorkbook) (r: RangeRef) = rangeObject book r

/// The hooks of the VBA engine's run-time library for a run on the workbook.
let private runtimeHooksOf (book: MacroWorkbook) : Rbt.Vb.Values.RuntimeHooks =
    { Print = (fun text -> book.Hooks.Output text)
      MsgBox = (fun prompt buttons title -> book.Hooks.MessageBox prompt buttons title)
      InputBox = (fun prompt title defaultText -> book.Hooks.InputBox prompt title defaultText)
      CreateObject = (fun _ -> None)
      Environ = (fun _ -> "") }

/// `VBA.RGB(...)`, `VBA.Len(...)`, `VBA.vbRed`: the VBA library by its name, as Excel's references give
/// it to macros (the engine's built-in functions and constants).
let private vbaLibrary (book: MacroWorkbook) =
    host "VBA"
        (fun name ->
            match Rbt.Vb.Builtins.tryConstant name with
            | Some value -> Some value
            | None -> Some (Rbt.Vb.Builtins.callBuiltin (runtimeHooksOf book) false (name.ToLowerInvariant()) [||]))
        (fun name _ -> unsupported "VBA" name)
        (fun name args -> Some (Rbt.Vb.Builtins.callBuiltin (runtimeHooksOf book) false (name.ToLowerInvariant()) (args |> Array.map scalarOf)))

/// The objects a macro starts from, by name: Application and the globals Excel puts in scope (ActiveSheet,
/// Range, Cells, Worksheets, Selection, ...), and the xl* constants.
let hostGlobals (book: MacroWorkbook) : (string * Value) list =
    let app = applicationObject book
    let appHost = match app with VHostObject h -> h | _ -> failwith "application"
    /// A global that stands for what the active sheet / Application has now (ActiveSheet, Range, Rows ...):
    /// resolved on each use, with its members, its default call and For Each.
    let forward (name: string) (target: unit -> Value option) (defaultCall: Value array -> Value option) =
        let current () = match target () with Some (VHostObject h) -> Some h | _ -> None
        // (It is what it stands for: TypeName(ActiveSheet) is "Worksheet".)
        VHostObject
            { new IHostObject with
                member _.TypeName = match current () with Some h -> h.TypeName | None -> name
                member _.GetMember member' = match current () with Some h -> h.GetMember member' | None -> unsupported name member'
                member _.SetMember member' value = match current () with Some h -> h.SetMember member' value | None -> false
                member _.CallMethod member' args =
                    if member' = "" || eq member' "_Default" then (if args.Length = 0 then (match current () with Some h -> h.CallMethod "" args | None -> None) else defaultCall args)
                    else match current () with Some h -> h.CallMethod member' args | None -> unsupported name member'
                member _.CallNamed member' args names =
                    match current () with
                    | Some h when member' <> "" && not (eq member' "_Default") -> h.CallNamed member' args names
                    | _ -> defaultCall args
                member _.Enumerate () = current () |> Option.bind (fun h -> h.Enumerate ()) }
    let viaApp (name: string) = forward name (fun () -> appHost.GetMember name) (fun args -> appHost.CallMethod name args)
    let viaSheet (name: string) =
        let sheet () = match worksheetObject book book.ActiveSheet with VHostObject h -> h | _ -> failwith "sheet"
        forward name (fun () -> (sheet ()).GetMember name) (fun args -> (sheet ()).CallMethod name args)
    [ "Application", app
      "ActiveWorkbook", workbookObject book
      "ThisWorkbook", workbookObject book
      "ActiveSheet", viaApp "ActiveSheet"
      "ActiveCell", viaApp "ActiveCell"
      "ActiveWindow", viaApp "ActiveWindow"
      "Selection", viaApp "Selection"
      "Worksheets", viaApp "Worksheets"
      "Sheets", viaApp "Sheets"
      "Range", viaApp "Range"
      "Cells", viaSheet "Cells"
      "Rows", viaSheet "Rows"
      "Columns", viaSheet "Columns"
      "Names", namesObject book
      "VBA", vbaLibrary book
      "WorksheetFunction", worksheetFunctionObject book
      "Evaluate", host "Evaluate" (fun n -> unsupported "Evaluate" n) (fun _ _ -> false) (fun n args -> if (n = "" || eq n "_Default") && args.Length = 1 then Some (evaluateArgument book (toText args.[0])) else unsupported "Evaluate" n)
      "Intersect", host "Intersect" (fun n -> unsupported "Intersect" n) (fun _ _ -> false) (fun n args -> if n = "" || eq n "_Default" then appHost.CallMethod "Intersect" args else unsupported "Intersect" n)
      "Union", host "Union" (fun n -> unsupported "Union" n) (fun _ _ -> false) (fun n args -> if n = "" || eq n "_Default" then appHost.CallMethod "Union" args else unsupported "Union" n)
      // (DoEvents: the engine has no event loop to yield to.)
      "DoEvents", host "DoEvents" (fun _ -> Some (VLong 0)) (fun _ _ -> false) (fun _ _ -> Some (VLong 0)) ]
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
                | Some id -> (match worksheetObject book id with VHostObject h -> h | _ -> failwith "sheet")
                | None -> fail 424 "Object required"
            yield name,
                  host "Worksheet"
                    (fun member' -> (sheet ()).GetMember member')
                    (fun member' value -> (sheet ()).SetMember member' value)
                    (fun member' args -> (sheet ()).CallMethod member' args) ]

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
