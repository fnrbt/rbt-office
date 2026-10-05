/// The plain JS objects the component's host API takes and returns, and
/// their conversion to and from model values. Keys are camelCase; a value
/// that is absent is `null`. (Browser only: the objects are JS objects.)
module Rbt.Spreadsheet.HostShapes

open Fable.Core
open Fable.Core.JsInterop
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.SheetOps

// ---------------------------------------------------------------------------
// JS values
// ---------------------------------------------------------------------------

/// Whether a JS value is `undefined` or `null` (an optional argument that
/// was left out arrives as either).
[<Emit("($0 == null)")>]
let isNullish (value: obj) : bool = jsNative

[<Emit("(typeof $0)")>]
let jsTypeOf (value: obj) : string = jsNative

[<Emit("Array.isArray($0)")>]
let isJsArray (value: obj) : bool = jsNative

/// Whether a JS value is an object with keys (not null, not an array).
[<Emit("($0 !== null && typeof $0 === 'object' && !Array.isArray($0))")>]
let isJsObject (value: obj) : bool = jsNative

[<Emit("Object.prototype.hasOwnProperty.call($0, $1)")>]
let hasKey (source: obj) (key: string) : bool = jsNative

[<Emit("$0[$1]")>]
let getKey (source: obj) (key: string) : obj = jsNative

[<Emit("Object.keys($0)")>]
let keysOf (source: obj) : string array = jsNative

[<Emit("Number.isInteger($0)")>]
let isJsInteger (value: obj) : bool = jsNative

[<Emit("(typeof $0 === 'number' && isFinite($0))")>]
let isJsNumber (value: obj) : bool = jsNative

/// An option as a JS value: `null` for None.
let orNull (value: 'T option) : obj =
    match value with
    | Some value -> box value
    | None -> null

/// Refuse a host call: throws an Error whose message starts with the member's name.
let hostFail (memberName: string) (message: string) : 'T =
    failwith (memberName + ": " + message)

/// What a value is, for an error message.
let private describe (value: obj) : string =
    if isNullish value then "null"
    elif isJsArray value then "an array"
    elif jsTypeOf value = "string" then "\"" + unbox<string> value + "\""
    elif jsTypeOf value = "object" then "an object"
    else string value

// --- typed reads of an argument or of a key of an argument ------------------
// `what` names the value in the error ("row", "format.fontSize").

let stringArg (memberName: string) (what: string) (value: obj) : string =
    if jsTypeOf value = "string" then unbox<string> value
    else hostFail memberName (sprintf "%s must be a string, got %s" what (describe value))

let boolArg (memberName: string) (what: string) (value: obj) : bool =
    if jsTypeOf value = "boolean" then unbox<bool> value
    else hostFail memberName (sprintf "%s must be true or false, got %s" what (describe value))

let numberArg (memberName: string) (what: string) (value: obj) : float =
    if isJsNumber value then unbox<float> value
    else hostFail memberName (sprintf "%s must be a number, got %s" what (describe value))

let intArg (memberName: string) (what: string) (value: obj) : int =
    if isJsInteger value then unbox<int> value
    else hostFail memberName (sprintf "%s must be a whole number, got %s" what (describe value))

/// An optional argument: None when it is undefined or null.
let optionalArg (read: obj -> 'T) (value: obj) : 'T option =
    if isNullish value then None else Some (read value)

/// A 0-based row / column argument within the sheet.
let rowArg (memberName: string) (model: SpreadsheetModel) (value: obj) : int =
    let row = intArg memberName "row" value
    if row < 0 || row >= model.RowCount then hostFail memberName (sprintf "row %d is outside the sheet (0 to %d)" row (model.RowCount - 1))
    else row

let colArg (memberName: string) (model: SpreadsheetModel) (value: obj) : int =
    let col = intArg memberName "col" value
    if col < 0 || col >= model.ColCount then hostFail memberName (sprintf "col %d is outside the sheet (0 to %d)" col (model.ColCount - 1))
    else col

/// An A1 range argument of the active sheet ("B2", "B2:D5", "A:C", "3:5").
let rangeArg (memberName: string) (value: obj) : CellRange =
    parseRange memberName (stringArg memberName "the range" value)

// ---------------------------------------------------------------------------
// Formats
// ---------------------------------------------------------------------------
// The `format` object, the same in both directions:
//   { bold, italic, underline, doubleUnderline, strikethrough: boolean,
//     fontFamily: string | null,           (null: the default font)
//     fontSize: number | null,             (points; null: the default, 11)
//     textColor: "#RRGGBB" | null,         (null: automatic)
//     fill: "#RRGGBB" | null,              (null: no fill)
//     horizontalAlign: "left" | "center" | "right" | "justify" | "fill" | "centerContinuous" | "distributed" | null,   (null: general)
//     verticalAlign: "top" | "middle" | "bottom" | "justify" | "distributed" | null,                                 (null: bottom)
//     wrap: boolean, indent: number (0..250), shrinkToFit: boolean,
//     textRotation: number,                (degrees: 0..90 counter-clockwise, -1..-90 clockwise, 255 stacked)
//     numberFormat: string | null,         (an Excel format code; null: General)
//     borders: { top, bottom, left, right, diagonal: { style, color: "#RRGGBB" } | null,
//                diagonalUp: boolean, diagonalDown: boolean } }
// Read from the host it is a patch: only the keys that are present change
// the format, and `null` puts a nullable property back to its default.

let private borderStyles =
    [ "thin"; "medium"; "thick"; "dashed"; "dotted"; "double"; "hair"; "mediumDashed"; "dashDot"; "mediumDashDot"
      "dashDotDot"; "mediumDashDotDot"; "slantDashDot" ]

let private horizontalAligns = [ "left"; "center"; "right"; "justify"; "fill"; "centerContinuous"; "distributed"; "general" ]
let private verticalAligns = [ "top"; "middle"; "bottom"; "justify"; "distributed" ]

/// The model's text rotation (0..90 up, 91..180 = 90 + degrees down, 255
/// stacked) as degrees (-90..90, or 255).
let rotationDegrees (rotation: int) : int =
    if rotation = 255 then 255 elif rotation > 90 then 90 - rotation else rotation

/// Degrees (-90..90, or 255 for stacked text) as the model's text rotation.
let rotationOfDegrees (degrees: int) : int =
    if degrees = 255 then 255 elif degrees < 0 then 90 - degrees else degrees

let borderToObj (border: BorderStyle option) : obj =
    match border with
    | Some border when border.Style <> "none" && border.Style <> "" ->
        createObj [ "style" ==> border.Style; "color" ==> (border.Color |> Option.defaultValue "#000000") ]
    | _ -> null

/// The model holds "middle" (set in the editor) or "center" (from a file) for the same alignment.
let private verticalOrNull (align: string option) : obj =
    match align with
    | Some "center" -> box "middle"
    | other -> orNull other

/// "General" (any case) and "" are no number format.
let private numberFormatOrNull (code: string option) : obj =
    match code with
    | Some code when code <> "" && not (code.Equals("General", System.StringComparison.OrdinalIgnoreCase)) -> box code
    | _ -> null

/// A cell format as the `format` object (every key present).
let formatToObj (format: CellFormat) : obj =
    createObj [
        "bold" ==> format.Bold
        "italic" ==> format.Italic
        "underline" ==> (format.Underline || format.UnderlineDouble)
        "doubleUnderline" ==> format.UnderlineDouble
        "strikethrough" ==> format.Strikethrough
        "fontFamily" ==> orNull format.FontName
        "fontSize" ==> orNull format.FontSize
        "textColor" ==> orNull format.TextColor
        "fill" ==> orNull format.BgColor
        "horizontalAlign" ==> orNull format.HorizontalAlign
        "verticalAlign" ==> verticalOrNull format.VerticalAlign
        "wrap" ==> format.WrapText
        "indent" ==> format.Indent
        "shrinkToFit" ==> format.ShrinkToFit
        "textRotation" ==> rotationDegrees format.TextRotation
        "numberFormat" ==> numberFormatOrNull format.NumberFormat
        "borders" ==> createObj [
            "top" ==> borderToObj format.BorderTop
            "bottom" ==> borderToObj format.BorderBottom
            "left" ==> borderToObj format.BorderLeft
            "right" ==> borderToObj format.BorderRight
            "diagonal" ==> borderToObj format.BorderDiagonal
            "diagonalUp" ==> format.DiagonalUp
            "diagonalDown" ==> format.DiagonalDown
        ]
    ]

/// What the cells of a selection have in common, as the `format` object
/// with "mixed" in place of every property the cells differ in. `borders`
/// has "mixed" per side (top, bottom, left, right, diagonal).
let formatSummaryToObj (summary: FormatSummary) : obj =
    let format = summary.Format
    let unlessMixed (field: int) (value: obj) : obj = if summary.Mixed &&& field <> 0 then box "mixed" else value
    createObj [
        "bold" ==> unlessMixed FormatField.Bold (box format.Bold)
        "italic" ==> unlessMixed FormatField.Italic (box format.Italic)
        "underline" ==> unlessMixed (FormatField.Underline ||| FormatField.UnderlineDouble) (box (format.Underline || format.UnderlineDouble))
        "doubleUnderline" ==> unlessMixed FormatField.UnderlineDouble (box format.UnderlineDouble)
        "strikethrough" ==> unlessMixed FormatField.Strikethrough (box format.Strikethrough)
        "fontFamily" ==> unlessMixed FormatField.FontName (orNull format.FontName)
        "fontSize" ==> unlessMixed FormatField.FontSize (orNull format.FontSize)
        "textColor" ==> unlessMixed FormatField.TextColor (orNull format.TextColor)
        "fill" ==> unlessMixed FormatField.Fill (orNull format.BgColor)
        "horizontalAlign" ==> unlessMixed FormatField.HorizontalAlign (orNull format.HorizontalAlign)
        "verticalAlign" ==> unlessMixed FormatField.VerticalAlign (verticalOrNull format.VerticalAlign)
        "wrap" ==> unlessMixed FormatField.WrapText (box format.WrapText)
        "indent" ==> unlessMixed FormatField.Indent (box format.Indent)
        "shrinkToFit" ==> unlessMixed FormatField.ShrinkToFit (box format.ShrinkToFit)
        "textRotation" ==> unlessMixed FormatField.TextRotation (box (rotationDegrees format.TextRotation))
        "numberFormat" ==> unlessMixed FormatField.NumberFormat (numberFormatOrNull format.NumberFormat)
        "borders" ==> createObj [
            "top" ==> unlessMixed FormatField.BorderTop (borderToObj format.BorderTop)
            "bottom" ==> unlessMixed FormatField.BorderBottom (borderToObj format.BorderBottom)
            "left" ==> unlessMixed FormatField.BorderLeft (borderToObj format.BorderLeft)
            "right" ==> unlessMixed FormatField.BorderRight (borderToObj format.BorderRight)
            "diagonal" ==> unlessMixed FormatField.BorderDiagonal (borderToObj format.BorderDiagonal)
        ]
    ]

/// A colour argument: "#RGB" or "#RRGGBB" (the `#` may be left out), given back as "#RRGGBB".
let colorArg (memberName: string) (what: string) (value: obj) : string =
    let text = (stringArg memberName what value).Trim()
    let hex = if text.StartsWith "#" then text.Substring 1 else text
    let isHex = hex |> Seq.forall (fun c -> (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F'))
    if isHex && hex.Length = 6 then "#" + hex.ToUpperInvariant()
    elif isHex && hex.Length = 3 then "#" + (hex |> Seq.map (fun c -> string c + string c) |> String.concat "").ToUpperInvariant()
    else hostFail memberName (sprintf "%s must be a colour like \"#RRGGBB\", got \"%s\"" what text)

let private oneOf (memberName: string) (what: string) (allowed: string list) (value: obj) : string =
    let text = stringArg memberName what value
    if List.contains text allowed then text
    else hostFail memberName (sprintf "%s must be one of %s, got \"%s\"" what (allowed |> List.map (fun name -> "\"" + name + "\"") |> String.concat ", ") text)

/// A border of the `format` object: null (no border) or { style, color }.
let borderOfObj (memberName: string) (what: string) (value: obj) : BorderStyle option =
    if isNullish value then None
    elif not (isJsObject value) then hostFail memberName (sprintf "%s must be null or { style, color }" what)
    else
        let style = oneOf memberName (what + ".style") ("none" :: borderStyles) (getKey value "style")
        if style = "none" then None
        else
            let color = if isNullish (getKey value "color") then "#000000" else colorArg memberName (what + ".color") (getKey value "color")
            Some { Style = style; Color = Some color }

let private formatKeys =
    [ "bold"; "italic"; "underline"; "doubleUnderline"; "strikethrough"; "fontFamily"; "fontSize"; "textColor"; "fill"
      "horizontalAlign"; "verticalAlign"; "wrap"; "indent"; "shrinkToFit"; "textRotation"; "numberFormat"; "borders" ]
let private borderKeys = [ "top"; "bottom"; "left"; "right"; "diagonal"; "diagonalUp"; "diagonalDown" ]

/// Read a `format` object from the host as a change to a format: the keys
/// that are present are set, the others stay as they are. Every key is
/// checked now (an unknown key or a value of the wrong kind throws), so the
/// function that comes back cannot fail. `what` names the argument.
let formatPatchOfObj (memberName: string) (what: string) (value: obj) : CellFormat -> CellFormat =
    if not (isJsObject value) then hostFail memberName (sprintf "%s must be a format object, got %s" what (describe value))
    let changes = ResizeArray<CellFormat -> CellFormat>()
    for key in keysOf value do
        if not (List.contains key formatKeys) then
            hostFail memberName (sprintf "%s has an unknown key \"%s\" (known: %s)" what key (String.concat ", " formatKeys))
    let present (key: string) = hasKey value key && jsTypeOf (getKey value key) <> "undefined"
    let name (key: string) = what + "." + key
    let flag (key: string) (set: bool -> CellFormat -> CellFormat) =
        if present key then
            let on = boolArg memberName (name key) (getKey value key)
            changes.Add (set on)
    flag "bold" (fun on format -> { format with Bold = on })
    flag "italic" (fun on format -> { format with Italic = on })
    // (A double underline is an underline: switching underline off clears both.)
    flag "underline" (fun on format -> if on then { format with Underline = true } else { format with Underline = false; UnderlineDouble = false })
    flag "doubleUnderline" (fun on format -> if on then { format with Underline = true; UnderlineDouble = true } else { format with UnderlineDouble = false })
    flag "strikethrough" (fun on format -> { format with Strikethrough = on })
    flag "wrap" (fun on format -> { format with WrapText = on })
    flag "shrinkToFit" (fun on format -> { format with ShrinkToFit = on })
    if present "fontFamily" then
        let family = optionalArg (stringArg memberName (name "fontFamily")) (getKey value "fontFamily") |> Option.map (fun text -> text.Trim())
        if family = Some "" then hostFail memberName (name "fontFamily" + " must not be empty (null for the default font)")
        changes.Add (fun format -> { format with FontName = family })
    if present "fontSize" then
        let size = optionalArg (numberArg memberName (name "fontSize")) (getKey value "fontSize")
        match size with
        | Some size when size < 1.0 || size > 409.0 -> hostFail memberName (sprintf "%s must be between 1 and 409 points, got %g" (name "fontSize") size)
        | _ -> ()
        // (The model holds whole points.)
        let size = size |> Option.map (fun size -> int (System.Math.Floor (size + 0.5)))
        changes.Add (fun format -> { format with FontSize = size })
    if present "textColor" then
        let color = optionalArg (colorArg memberName (name "textColor")) (getKey value "textColor")
        changes.Add (fun format -> { format with TextColor = color })
    if present "fill" then
        let color = optionalArg (colorArg memberName (name "fill")) (getKey value "fill")
        changes.Add (fun format -> { format with BgColor = color })
    if present "horizontalAlign" then
        let align = optionalArg (oneOf memberName (name "horizontalAlign") horizontalAligns) (getKey value "horizontalAlign") |> Option.filter (fun align -> align <> "general")
        changes.Add (fun format -> { format with HorizontalAlign = align })
    if present "verticalAlign" then
        let align = optionalArg (oneOf memberName (name "verticalAlign") verticalAligns) (getKey value "verticalAlign")
        changes.Add (fun format -> { format with VerticalAlign = align })
    if present "indent" then
        let indent = intArg memberName (name "indent") (getKey value "indent")
        if indent < 0 || indent > 250 then hostFail memberName (sprintf "%s must be between 0 and 250, got %d" (name "indent") indent)
        changes.Add (fun format -> { format with Indent = indent })
    if present "textRotation" then
        let degrees = intArg memberName (name "textRotation") (getKey value "textRotation")
        if degrees <> 255 && (degrees < -90 || degrees > 90) then
            hostFail memberName (sprintf "%s must be between -90 and 90 degrees, or 255 for stacked text, got %d" (name "textRotation") degrees)
        changes.Add (fun format -> { format with TextRotation = rotationOfDegrees degrees })
    if present "numberFormat" then
        let code = optionalArg (stringArg memberName (name "numberFormat")) (getKey value "numberFormat")
        let code = code |> Option.filter (fun code -> code <> "" && not (code.Equals("General", System.StringComparison.OrdinalIgnoreCase)))
        changes.Add (fun format -> { format with NumberFormat = code })
    if present "borders" then
        let borders = getKey value "borders"
        if not (isJsObject borders) then hostFail memberName (name "borders" + " must be an object { top, bottom, left, right, diagonal, diagonalUp, diagonalDown }")
        for key in keysOf borders do
            if not (List.contains key borderKeys) then
                hostFail memberName (sprintf "%s has an unknown key \"%s\" (known: %s)" (name "borders") key (String.concat ", " borderKeys))
        let side (key: string) (set: BorderStyle option -> CellFormat -> CellFormat) =
            if hasKey borders key && jsTypeOf (getKey borders key) <> "undefined" then
                let border = borderOfObj memberName (name "borders." + key) (getKey borders key)
                changes.Add (set border)
        side "top" (fun border format -> { format with BorderTop = border })
        side "bottom" (fun border format -> { format with BorderBottom = border })
        side "left" (fun border format -> { format with BorderLeft = border })
        side "right" (fun border format -> { format with BorderRight = border })
        side "diagonal" (fun border format -> { format with BorderDiagonal = border })
        let diagonalFlag (key: string) (set: bool -> CellFormat -> CellFormat) =
            if hasKey borders key && jsTypeOf (getKey borders key) <> "undefined" then
                let on = boolArg memberName (name "borders." + key) (getKey borders key)
                changes.Add (set on)
        diagonalFlag "diagonalUp" (fun on format -> { format with DiagonalUp = on })
        diagonalFlag "diagonalDown" (fun on format -> { format with DiagonalDown = on })
    let changes = changes.ToArray()
    fun format ->
        let mutable result = format
        for change in changes do result <- change result
        result

/// A complete format from a `format` object: the default format with the
/// object's keys set (what a format painter applies: keys left out are the
/// defaults, not the target cell's).
let formatOfObj (memberName: string) (what: string) (value: obj) : CellFormat =
    formatPatchOfObj memberName what value defaultFormat

// ---------------------------------------------------------------------------
// Cells
// ---------------------------------------------------------------------------

/// A cell value as a JS value: a number, a string, a boolean, the text of
/// an error ("#DIV/0!"), or null for an empty cell. A formula gives its result.
let valueToObj (value: CellValue) : obj =
    match unwrapFormulaValue value with
    | CellValue.Empty -> null
    | CellValue.Number number -> box number
    | CellValue.String text -> box text
    | CellValue.Boolean flag -> box flag
    | CellValue.Error code -> box (if code.StartsWith "#" then code else "#" + code)
    | CellValue.Formula _ -> null

/// What kind of value a cell holds (of a formula: its result).
let valueKind (value: CellValue) : string =
    match unwrapFormulaValue value with
    | CellValue.Empty -> "empty"
    | CellValue.Number _ -> "number"
    | CellValue.String _ -> "string"
    | CellValue.Boolean _ -> "boolean"
    | CellValue.Error _ -> "error"
    | CellValue.Formula _ -> "empty"

/// A hyperlink: { range, target, isInternal, tooltip, display, safe } -
/// `isInternal`: a place in the workbook ("Sheet2!A1", a defined name), else
/// an address; `safe`: an address a host may open (http, https, mailto, tel).
let hyperlinkToObj (link: SheetHyperlink) : obj =
    createObj [ "range" ==> rangeText link.Range; "target" ==> link.Target; "isInternal" ==> link.IsInternal
                "tooltip" ==> orNull link.Tooltip; "display" ==> orNull link.Display; "safe" ==> hyperlinkIsSafe link ]

// --- data validation ---------------------------------------------------------
// The `rule` object of a validation, the same in both directions (`type` says
// which keys count):
//   { type: "any" }                                          (every value; carries an input message only)
//   { type: "list", values: string[], dropdown }             (the allowed values written out)
//   { type: "list", source: "$A$1:$A$9" | "Lists!$B$2:$B$20" | "NameOfARange", dropdown }   (read from cells when needed)
//   { type: "whole" | "decimal" | "date" | "time" | "textLength",
//     operator: "between" | "notBetween" | "equal" | "notEqual" | "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual",
//     formula1, formula2 (between / notBetween only) }
//   { type: "custom", formula }
// and for every type:
//   allowBlank (true), showInput (true), inputTitle, inputMessage,
//   showError (true), errorStyle: "stop" (default) | "warning" | "information", errorTitle, errorMessage.
// formula1 / formula2 / formula are a number or a formula without the leading
// "=" (one is accepted and dropped), written for the top-left cell of the
// validation's first range; a relative reference moves with the cell. A date
// may be given as "2024-03-31" or a Date, a time as "14:30"; both are read
// back as serial numbers. `value1` / `value2` are accepted for formula1 / formula2.

[<Emit("($0 instanceof Date)")>]
let private isJsDate (value: obj) : bool = jsNative

[<Emit("((d) => { const p = (n) => String(n).padStart(2, '0'); const day = d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()); return (d.getHours() || d.getMinutes() || d.getSeconds()) ? day + ' ' + p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds()) : day; })($0)")>]
let private jsDateText (value: obj) : string = jsNative

/// A validation as its `rule` object (see above); the keys that do not
/// count for its type are null.
let validationToObj (validation: DataValidation) : obj =
    let d = describeValidation validation
    let custom = d.Type = "custom"
    createObj [
        "type" ==> d.Type
        "operator" ==> orNull d.Operator
        "formula1" ==> (if custom then null else orNull d.Formula1)
        "formula2" ==> orNull d.Formula2
        "formula" ==> (if custom then orNull d.Formula1 else null)
        "values" ==> (match d.Values with Some values -> box (List.toArray values) | None -> null)
        "source" ==> orNull d.Source
        "dropdown" ==> (if d.Type = "list" then box d.Dropdown else null)
        "allowBlank" ==> d.AllowBlank
        "showInput" ==> d.ShowInput
        "inputTitle" ==> orNull d.InputTitle
        "inputMessage" ==> orNull d.InputMessage
        "showError" ==> d.ShowError
        "errorStyle" ==> d.ErrorStyle
        "errorTitle" ==> orNull d.ErrorTitle
        "errorMessage" ==> orNull d.ErrorMessage
    ]

/// A validation of a sheet: { id, range: "B2:B9" (areas separated by spaces), rule }.
let validationEntryToObj (entry: ValidationRange) : obj =
    createObj [ "id" ==> entry.Id; "range" ==> rangesText entry.Ranges; "rule" ==> validationToObj entry.Validation ]

/// The validation on a cell: its `rule` object with two more keys, `id` and
/// `range` (which `validationOfObj` ignores, so that the object can be given
/// back to SetValidation for other cells).
let validationAtToObj (entry: ValidationRange) : obj =
    let rule = validationToObj entry.Validation
    rule?id <- entry.Id
    rule?range <- rangesText entry.Ranges
    rule

let private validationKeys =
    [ "type"; "operator"; "formula1"; "formula2"; "value1"; "value2"; "formula"; "values"; "source"; "dropdown"
      "allowBlank"; "showInput"; "inputTitle"; "inputMessage"; "showError"; "errorStyle"; "errorTitle"; "errorMessage"; "id"; "range" ]

/// Read a validation `rule` object. Types and keys are checked here; what
/// the rule needs is checked by `SheetOps.validationOfDescription`.
let validationOfObj (memberName: string) (value: obj) : ValidationDescription =
    if not (isJsObject value) then hostFail memberName (sprintf "rule must be an object { type, ... }, got %s" (describe value))
    for key in keysOf value do
        if not (List.contains key validationKeys) then
            hostFail memberName (sprintf "a validation rule has no key \"%s\" (it takes: %s)" key (String.concat ", " validationKeys))
    let get (key: string) = getKey value key
    let kind = stringArg memberName "rule.type" (get "type")
    let operand (key: string) (alias: string) : string option =
        let given = if isNullish (get key) then get alias else get key
        if isNullish given then None
        elif isJsNumber given then Some (string (unbox<float> given))
        elif jsTypeOf given = "string" then Some (unbox<string> given)
        elif isJsDate given then Some (jsDateText given)
        else hostFail memberName (sprintf "rule.%s must be a number, a string or a Date, got %s" key (describe given))
    let text (key: string) = optionalArg (stringArg memberName ("rule." + key)) (get key)
    let flag (key: string) (fallback: bool) = optionalArg (boolArg memberName ("rule." + key)) (get key) |> Option.defaultValue fallback
    let values =
        let items = get "values"
        if isNullish items then None
        elif isJsArray items then
            Some [ for item in unbox<obj array> items ->
                    if isJsNumber item then string (unbox<float> item)
                    else stringArg memberName "an item of rule.values" item ]
        else hostFail memberName (sprintf "rule.values must be an array of texts, got %s" (describe items))
    { emptyValidationDescription kind with
        Operator = text "operator"
        Formula1 = (if kind = "custom" then (match operand "formula" "formula1" with Some f -> Some f | None -> None) else operand "formula1" "value1")
        Formula2 = operand "formula2" "value2"
        Values = values
        Source = text "source"
        Dropdown = flag "dropdown" true
        AllowBlank = flag "allowBlank" true
        ShowInput = flag "showInput" true
        InputTitle = text "inputTitle"
        InputMessage = text "inputMessage"
        ShowError = flag "showError" true
        ErrorStyle = (text "errorStyle" |> Option.defaultValue "stop")
        ErrorTitle = text "errorTitle"
        ErrorMessage = text "errorMessage" }

/// One cell as the host reads it (`GetCell`, `GetRange`):
/// A comment of a cell: { text, author: string | null, thread } where
/// `thread` is null for a note and, for a comment thread,
/// { id, author, date, text, resolved: boolean, replies: [{ id, author, date, text }] }
/// (`text` and `author` are then those of the thread's first comment).
let commentToObj (info: CommentInfo) : obj =
    let entry (comment: CommentThread) =
        [ "id" ==> comment.Id; "author" ==> orNull comment.Author; "date" ==> orNull comment.Date; "text" ==> comment.Text ]
    createObj [
        "text" ==> info.Text
        "author" ==> orNull info.Author
        "thread" ==>
            (match info.Thread with
             | Some thread ->
                createObj (entry thread @ [ "resolved" ==> thread.Done
                                            "replies" ==> (thread.Replies |> List.map (fun reply -> createObj (entry reply)) |> List.toArray) ])
             | None -> null)
    ]

/// { row, col, address: "B2", value, valueType, text, formula, format,
///   isMerged, mergeRange, comment, hyperlink, validation } - see GetCell.
/// `text` is the display text the caller worked out for the cell.
let cellToObj (model: SpreadsheetModel) (addr: CellAddress) (text: string) : obj =
    let data = Map.tryFind addr model.Cells
    let value = match data with Some data -> data.Value | None -> CellValue.Empty
    let merge = mergeRangeAt model addr
    let comment : obj = match commentAt model addr with Some info -> commentToObj info | None -> null
    createObj [
        "row" ==> addr.Row
        "col" ==> addr.Col
        "address" ==> formatCellRef addr.Row addr.Col
        "value" ==> valueToObj value
        "valueType" ==> valueKind value
        "text" ==> text
        "formula" ==> (match value with CellValue.Formula (formula, _) -> box (formulaText formula) | _ -> null)
        "format" ==> formatToObj (effectiveFormat model addr)
        "isMerged" ==> merge.IsSome
        "mergeRange" ==> (match merge with Some range -> box (rangeText range) | None -> null)
        "comment" ==> comment
        "hyperlink" ==> (match hyperlinkAt model addr with Some link -> hyperlinkToObj link | None -> null)
        "validation" ==> (match validationEntryAt addr model with Some entry -> validationAtToObj entry | None -> null)
    ]

// ---------------------------------------------------------------------------
// Conditional formats
// ---------------------------------------------------------------------------
// The `rule` object, the same in both directions (`type` says which keys count):
//   { type: "cellIs", operator: "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual" | "equal" | "notEqual"
//                               | "between" | "notBetween",
//     formula1, formula2 (between / notBetween only), format }
//   { type: "expression", formula, format }
//   { type: "containsText" | "notContainsText" | "beginsWith" | "endsWith", text, format }
//   { type: "top", rank: number, percent: boolean, bottom: boolean, format }
//   { type: "aboveAverage", above: boolean, orEqual: boolean, stdDev: 0..3, format }
//   { type: "duplicateValues" | "uniqueValues" | "blanks" | "noBlanks" | "errors" | "noErrors", format }
//   { type: "dateOccurring", period: "today" | "yesterday" | "tomorrow" | "last7Days" | "thisWeek" | "lastWeek" | "nextWeek"
//                                    | "thisMonth" | "lastMonth" | "nextMonth", format }
//   { type: "colorScale", stops: [threshold + { color: "#RRGGBB" }] (2 or 3, ascending) }
//   { type: "dataBar", min: threshold, max: threshold, color: "#RRGGBB" }
//   { type: "iconSet", iconSet: "3TrafficLights1" | ..., thresholds: [threshold] (one per icon), reverse: boolean }
// threshold: { type: "min" | "max" | "number" | "percent" | "percentile" | "formula", value: number | string | null }
// Formulas are written without the leading "=" (one is accepted and dropped):
// a number, a quoted text ("\"north\""), a reference or any formula, written
// for the top-left cell of the rule's first range. A number may be given as one.
// `format` is a `format` object of which a conditional format applies
// textColor, fill, bold, italic, underline and strikethrough.

let private ruleFormatKeys = [ "textColor"; "fill"; "bold"; "italic"; "underline"; "strikethrough" ]

/// The format a rule applies, as a `format` object with the keys a rule has.
let ruleFormatToObj (format: CellFormat) : obj =
    createObj [
        "textColor" ==> orNull format.TextColor
        "fill" ==> orNull format.BgColor
        "bold" ==> format.Bold
        "italic" ==> format.Italic
        "underline" ==> format.Underline
        "strikethrough" ==> format.Strikethrough
    ]

/// Read the `format` of a rule as a change to a rule's format; other keys
/// of the `format` object are refused (a rule cannot apply them).
let ruleFormatPatchOfObj (memberName: string) (what: string) (value: obj) : CellFormat -> CellFormat =
    if not (isJsObject value) then hostFail memberName (sprintf "%s must be a format object, got %s" what (describe value))
    for key in keysOf value do
        if not (List.contains key ruleFormatKeys) then
            hostFail memberName (sprintf "%s.%s: a conditional format applies only %s" what key (String.concat ", " ruleFormatKeys))
    formatPatchOfObj memberName what value

let private thresholdToObj (threshold: ThresholdDescription) (color: string option) : obj =
    let value : obj =
        match threshold.Value with
        | Some text when threshold.Kind <> "formula" ->
            (match System.Double.TryParse(text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture) with
             | true, number -> box number
             | _ -> box text)
        | Some text -> box text
        | None -> null
    createObj [
        yield "type" ==> threshold.Kind
        yield "value" ==> value
        match color with
        | Some color -> yield "color" ==> color
        | None -> () ]

/// A rule description as the `rule` object.
let ruleToObj (rule: RuleDescription) : obj =
    let format : obj = match rule.Style with Some style -> ruleFormatToObj style | None -> null
    createObj [
        yield "type" ==> rule.Type
        match rule.Type with
        | "cellIs" ->
            yield "operator" ==> orNull rule.Operator
            yield "formula1" ==> orNull rule.Formula1
            yield "formula2" ==> orNull rule.Formula2
        | "expression" -> yield "formula" ==> orNull rule.Formula1
        | "containsText" | "notContainsText" | "beginsWith" | "endsWith" -> yield "text" ==> orNull rule.Text
        | "dateOccurring" -> yield "period" ==> orNull rule.Period
        | "top" ->
            yield "rank" ==> rule.Rank
            yield "percent" ==> rule.Percent
            yield "bottom" ==> rule.Bottom
        | "aboveAverage" ->
            yield "above" ==> rule.Above
            yield "orEqual" ==> rule.OrEqual
            yield "stdDev" ==> rule.StdDev
        | "colorScale" ->
            yield "stops" ==> (rule.Stops |> List.map (fun (threshold, color) -> thresholdToObj threshold (Some color)) |> List.toArray)
        | "dataBar" ->
            yield "min" ==> (match rule.Low with Some low -> thresholdToObj low None | None -> null)
            yield "max" ==> (match rule.High with Some high -> thresholdToObj high None | None -> null)
            yield "color" ==> orNull rule.Color
        | "iconSet" ->
            yield "iconSet" ==> orNull rule.IconSet
            yield "thresholds" ==> (rule.Thresholds |> List.map (fun threshold -> thresholdToObj threshold None) |> List.toArray)
            yield "reverse" ==> rule.Reverse
        | _ -> ()
        yield "format" ==> format ]

/// A conditional format of a sheet: { id, range, priority, stopIfTrue, rule }.
/// `priority` is its place in the order of precedence (1 = first).
let conditionalFormatToObj (place: int) (entry: ConditionalFormatRange) : obj =
    createObj [
        "id" ==> entry.Id
        "range" ==> rangesText entry.Ranges
        "priority" ==> place
        "stopIfTrue" ==> entry.StopIfTrue
        "rule" ==> ruleToObj (describeRule entry.Format)
    ]

/// A number or a string of a rule, as text.
let private operandText (memberName: string) (what: string) (value: obj) : string =
    if isJsNumber value then string (unbox<float> value)
    elif jsTypeOf value = "string" then unbox<string> value
    else hostFail memberName (sprintf "%s must be a number or a string, got %s" what (describe value))

let private thresholdOfObj (memberName: string) (what: string) (allowColor: bool) (value: obj) : ThresholdDescription * string option =
    if not (isJsObject value) then hostFail memberName (sprintf "%s must be { type, value%s }, got %s" what (if allowColor then ", color" else "") (describe value))
    for key in keysOf value do
        if not (key = "type" || key = "value" || (allowColor && key = "color")) then
            hostFail memberName (sprintf "%s has an unknown key \"%s\"" what key)
    let kind = stringArg memberName (what + ".type") (getKey value "type")
    let text = optionalArg (operandText memberName (what + ".value")) (getKey value "value")
    let color = if allowColor then Some (colorArg memberName (what + ".color") (getKey value "color")) else None
    { Kind = kind; Value = text }, color

/// Read a `rule` object. Types and keys are checked here; what the rule
/// needs is checked by `SheetOps.ruleOfDescription`. The description's
/// `Style` is the `format` key applied to `baseStyle`, or None without one.
let ruleOfObj (memberName: string) (baseStyle: CellFormat) (value: obj) : RuleDescription =
    if not (isJsObject value) then hostFail memberName (sprintf "rule must be an object { type, ... }, got %s" (describe value))
    let kind = stringArg memberName "rule.type" (getKey value "type")
    let keys =
        match kind with
        | "cellIs" -> [ "operator"; "formula1"; "formula2" ]
        | "expression" -> [ "formula" ]
        | "containsText" | "notContainsText" | "beginsWith" | "endsWith" -> [ "text" ]
        | "dateOccurring" -> [ "period" ]
        | "top" -> [ "rank"; "percent"; "bottom" ]
        | "aboveAverage" -> [ "above"; "orEqual"; "stdDev" ]
        | "colorScale" -> [ "stops" ]
        | "dataBar" -> [ "min"; "max"; "color" ]
        | "iconSet" -> [ "iconSet"; "thresholds"; "reverse" ]
        | other when List.contains other ruleTypes -> []
        | other -> hostFail memberName (sprintf "unknown rule type \"%s\" (known: %s)" other (String.concat ", " ruleTypes))
    for key in keysOf value do
        if key <> "type" && key <> "format" && not (List.contains key keys) then
            hostFail memberName (sprintf "a \"%s\" rule has no key \"%s\" (it takes: %s)" kind key (String.concat ", " ("type" :: keys @ [ "format" ])))
    let get (key: string) = getKey value key
    let text (key: string) = optionalArg (operandText memberName ("rule." + key)) (get key)
    let flag (key: string) (fallback: bool) = optionalArg (boolArg memberName ("rule." + key)) (get key) |> Option.defaultValue fallback
    let whole (key: string) (fallback: int) = optionalArg (intArg memberName ("rule." + key)) (get key) |> Option.defaultValue fallback
    let list (key: string) : obj array =
        let items = get key
        if isNullish items then [||]
        elif isJsArray items then unbox<obj array> items
        else hostFail memberName (sprintf "rule.%s must be an array, got %s" key (describe items))
    let empty = emptyRuleDescription kind
    { empty with
        Operator = optionalArg (stringArg memberName "rule.operator") (get "operator")
        Formula1 = (if kind = "expression" then text "formula" else text "formula1")
        Formula2 = text "formula2"
        Text = optionalArg (stringArg memberName "rule.text") (get "text")
        Rank = whole "rank" 10
        Percent = flag "percent" false
        Bottom = flag "bottom" false
        Above = flag "above" true
        OrEqual = flag "orEqual" false
        StdDev = whole "stdDev" 0
        Stops =
            list "stops"
            |> Array.mapi (fun index stop ->
                let threshold, color = thresholdOfObj memberName (sprintf "rule.stops[%d]" index) true stop
                threshold, defaultArg color "")
            |> Array.toList
        Low = optionalArg (fun low -> fst (thresholdOfObj memberName "rule.min" false low)) (get "min")
        High = optionalArg (fun high -> fst (thresholdOfObj memberName "rule.max" false high)) (get "max")
        Color = optionalArg (colorArg memberName "rule.color") (get "color")
        IconSet = optionalArg (stringArg memberName "rule.iconSet") (get "iconSet")
        Thresholds = list "thresholds" |> Array.mapi (fun index item -> fst (thresholdOfObj memberName (sprintf "rule.thresholds[%d]" index) false item)) |> Array.toList
        Reverse = flag "reverse" false
        Period = optionalArg (stringArg memberName "rule.period") (get "period")
        Style = optionalArg (fun format -> ruleFormatPatchOfObj memberName "rule.format" format baseStyle) (get "format") }

// ---------------------------------------------------------------------------
// Sort keys and AutoFilter criteria
// ---------------------------------------------------------------------------
// A column is a 0-based sheet column number or its letters ("C").
//
// A sort key:  { column, descending?: boolean, by?: "value" | "cellColor" | "fontColor",
//                color?: "#RRGGBB" | null }   (color: with a colour sort, the colour that goes first)
//
// Filter criteria, one of:
//   { values: string[], blanks?: boolean }         the cells showing one of the texts (and the blank cells)
//   { conditions: [{ operator, value }], and?: boolean }   one or two comparisons, joined by AND (default) or OR;
//        operator: "equals" | "notEquals" | "greaterThan" | "greaterOrEqual" | "lessThan" | "lessOrEqual"
//                  | "contains" | "notContains" | "beginsWith" | "endsWith";
//        value: a number, a text (with "equals" / "notEquals" it may hold the wildcards * and ?), a Date or
//               a date written "2024-03-31"
//   { top: { count, percent?: boolean, bottom?: boolean } }   the largest (bottom: smallest) numbers
//   { color: "#RRGGBB" | null, by?: "cellColor" | "fontColor" }   null: no fill / the automatic text colour
//   { dynamic: "aboveAverage" | "belowAverage" }
// and, read only, { kept: true } for a filter of an opened file that the editor does not evaluate.

/// A column argument: a 0-based number or the column's letters.
let columnArg (memberName: string) (what: string) (model: SpreadsheetModel) (value: obj) : int =
    let col =
        if jsTypeOf value = "string" then
            let text = (unbox<string> value).Trim().TrimStart('$').ToUpperInvariant()
            if text <> "" && text.Length <= 3 && text |> Seq.forall (fun c -> c >= 'A' && c <= 'Z') then letterToColumn text
            else hostFail memberName (sprintf "%s must be a 0-based column number or column letters, got \"%s\"" what (unbox<string> value))
        else intArg memberName what value
    if col < 0 || col >= model.ColCount then hostFail memberName (sprintf "%s %s is outside the sheet" what (if col >= 0 then columnToLetter col else string col))
    else col

/// The keys of a sort (an array of key objects, or one key object).
let sortKeysOfObj (memberName: string) (model: SpreadsheetModel) (value: obj) : SortKey list =
    let one (item: obj) : SortKey =
        if not (isJsObject item) then
            // (A bare column is a key too.)
            { Column = columnArg memberName "a sort key" model item; Descending = false; By = SortByValue; Color = "" }
        else
            for key in keysOf item do
                if not (List.contains key [ "column"; "descending"; "by"; "color" ]) then
                    hostFail memberName (sprintf "a sort key has no key \"%s\" (it takes: column, descending, by, color)" key)
            let by =
                match optionalArg (stringArg memberName "key.by") (getKey item "by") with
                | None | Some "value" -> SortByValue
                | Some "cellColor" -> SortByCellColor
                | Some "fontColor" -> SortByFontColor
                | Some other -> hostFail memberName (sprintf "key.by must be \"value\", \"cellColor\" or \"fontColor\", got \"%s\"" other)
            let color = getKey item "color"
            if by = SortByValue && not (isNullish color) then hostFail memberName "key.color goes with by: \"cellColor\" or \"fontColor\""
            { Column = columnArg memberName "key.column" model (getKey item "column")
              Descending = (optionalArg (boolArg memberName "key.descending") (getKey item "descending") |> Option.defaultValue false)
              By = by
              Color = (if isNullish color then "" else colorArg memberName "key.color" color) }
    if isJsArray value then
        let items = unbox<obj array> value
        if items.Length = 0 then hostFail memberName "keys must hold at least one sort key"
        [ for item in items -> one item ]
    elif isNullish value then hostFail memberName "keys is required: [{ column, descending?, by? }]"
    else [ one value ]

/// The options of a sort: { hasHeader?, caseSensitive? } (both false when left out).
let sortOptionsOfObj (memberName: string) (value: obj) : SortOptions =
    if isNullish value then { HasHeader = false; CaseSensitive = false }
    elif not (isJsObject value) then hostFail memberName "options must be an object { hasHeader?, caseSensitive? }"
    else
        for key in keysOf value do
            if key <> "hasHeader" && key <> "caseSensitive" then hostFail memberName (sprintf "the sort options have no key \"%s\" (they take: hasHeader, caseSensitive)" key)
        { HasHeader = (optionalArg (boolArg memberName "options.hasHeader") (getKey value "hasHeader") |> Option.defaultValue false)
          CaseSensitive = (optionalArg (boolArg memberName "options.caseSensitive") (getKey value "caseSensitive") |> Option.defaultValue false) }

/// Filter criteria as a host reads them.
let filterToObj (criteria: FilterCriteria) : obj =
    let description = describeFilter criteria
    match description.Kind with
    | "values" -> createObj [ "values" ==> Array.ofList description.Values; "blanks" ==> description.Blanks ]
    | "conditions" ->
        createObj [
            "conditions" ==> [| for condition in description.Conditions -> createObj [ "operator" ==> condition.Operator; "value" ==> condition.Value ] |]
            "and" ==> description.And
        ]
    | "top" -> createObj [ "top" ==> createObj [ "count" ==> description.Count; "percent" ==> description.Percent; "bottom" ==> description.Bottom ] ]
    | "color" -> createObj [ "color" ==> (if description.Color = "" then null else box description.Color); "by" ==> description.By ]
    | "dynamic" -> createObj [ "dynamic" ==> description.Dynamic ]
    | _ -> createObj [ "kept" ==> true ]

/// Filter criteria a host gives, checked for their shape (`SheetOps.filterOfDescription` checks the rest).
let filterOfObj (memberName: string) (value: obj) : FilterDescription =
    if not (isJsObject value) then hostFail memberName "criteria must be an object: { values }, { conditions }, { top }, { color } or { dynamic }"
    let get (key: string) = getKey value key
    let kinds = [ "values"; "conditions"; "top"; "color"; "dynamic" ] |> List.filter (fun key -> hasKey value key)
    for key in keysOf value do
        if not (List.contains key [ "values"; "blanks"; "conditions"; "and"; "top"; "color"; "by"; "dynamic" ]) then
            hostFail memberName (sprintf "filter criteria have no key \"%s\" (they take: values, blanks, conditions, and, top, color, by, dynamic)" key)
    let text (what: string) (item: obj) : string =
        if isJsNumber item then string (unbox<float> item)
        elif jsTypeOf item = "string" then unbox<string> item
        elif isJsDate item then jsDateText item
        else hostFail memberName (sprintf "%s must be a text, a number or a Date" what)
    let flag (key: string) (fallback: bool) = optionalArg (boolArg memberName ("criteria." + key)) (get key) |> Option.defaultValue fallback
    match kinds with
    | [ "values" ] ->
        let items = get "values"
        if not (isJsArray items) then hostFail memberName "criteria.values must be an array of texts"
        { emptyFilterDescription "values" with
            Values = [ for item in unbox<obj array> items -> if isNullish item then "" else text "an item of criteria.values" item ]
            Blanks = flag "blanks" false }
    | [ "conditions" ] ->
        let items = get "conditions"
        let items = if isJsArray items then unbox<obj array> items elif isJsObject items then [| items |] else hostFail memberName "criteria.conditions must be an array of { operator, value }"
        { emptyFilterDescription "conditions" with
            Conditions =
                [ for item in items ->
                    if not (isJsObject item) then hostFail memberName "a condition must be an object { operator, value }"
                    { Operator = stringArg memberName "condition.operator" (getKey item "operator")
                      Value = (let given = getKey item "value" in if isNullish given then "" else text "condition.value" given) } ]
            And = flag "and" true }
    | [ "top" ] ->
        let top = get "top"
        if isJsNumber top then { emptyFilterDescription "top" with Count = intArg memberName "criteria.top" top }
        elif not (isJsObject top) then hostFail memberName "criteria.top must be an object { count, percent?, bottom? }"
        else
            { emptyFilterDescription "top" with
                Count = intArg memberName "top.count" (getKey top "count")
                Percent = (optionalArg (boolArg memberName "top.percent") (getKey top "percent") |> Option.defaultValue false)
                Bottom = (optionalArg (boolArg memberName "top.bottom") (getKey top "bottom") |> Option.defaultValue false) }
    | [ "color" ] ->
        let color = get "color"
        { emptyFilterDescription "color" with
            Color = (if isNullish color then "" else colorArg memberName "criteria.color" color)
            By = (optionalArg (stringArg memberName "criteria.by") (get "by") |> Option.defaultValue "cellColor") }
    | [ "dynamic" ] -> { emptyFilterDescription "dynamic" with Dynamic = stringArg memberName "criteria.dynamic" (get "dynamic") }
    | [] -> hostFail memberName "criteria need one of the keys values, conditions, top, color, dynamic"
    | several -> hostFail memberName (sprintf "criteria take one kind of filter, got %s" (String.concat " and " several))

// ---------------------------------------------------------------------------
// Pictures
// ---------------------------------------------------------------------------

[<Emit("(function (v) { try { if (v instanceof Uint8Array) return v; if (v instanceof ArrayBuffer) return new Uint8Array(v); if (ArrayBuffer.isView(v)) return new Uint8Array(v.buffer, v.byteOffset, v.byteLength); if (typeof v === 'string') { var m = /^data:[^,;]*((?:;[^,;]*)*),/.exec(v); if (!m) return null; var body = v.slice(m[0].length); if (/;base64/i.test(m[1])) { var bin = atob(body.replace(/\\s+/g, '')); var out = new Uint8Array(bin.length); for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i); return out; } return new TextEncoder().encode(decodeURIComponent(body)); } } catch (e) { } return null; })($0)")>]
let private bytesOfSource (value: obj) : byte[] = jsNative

/// The bytes of a picture as a host hands it over: a Uint8Array, an
/// ArrayBuffer, or a data URL ("data:image/png;base64,...").
let pictureBytesArg (memberName: string) (value: obj) : byte[] =
    if isNullish value then hostFail memberName "the picture is missing (a Uint8Array, an ArrayBuffer or a data URL)"
    else
        let bytes = bytesOfSource value
        if isNull (box bytes) then hostFail memberName "the picture is not a Uint8Array, an ArrayBuffer or a data URL (\"data:image/png;base64,...\")"
        elif bytes.Length = 0 then hostFail memberName "the picture is empty"
        else bytes

/// A cell as a host names one: "B3", or { row, col } (0-based).
let cellArg (memberName: string) (what: string) (model: SpreadsheetModel) (value: obj) : CellAddress =
    if isNullish value then hostFail memberName (what + " is missing")
    elif jsTypeOf value = "string" then
        match tryParseCell (unbox<string> value) with
        | Some addr when addr.Row < model.RowCount && addr.Col < model.ColCount -> addr
        | Some _ -> hostFail memberName (sprintf "%s %s is outside the sheet" what (unbox<string> value))
        | None -> hostFail memberName (sprintf "%s \"%s\" is not a cell (\"B3\", or { row, col })" what (unbox<string> value))
    elif isJsObject value && hasKey value "row" && hasKey value "col" then
        { Row = rowArg memberName model (getKey value "row"); Col = colArg memberName model (getKey value "col") }
    else hostFail memberName (what + " is not a cell (\"B3\", or { row, col })")

/// A drawing object of a sheet: { id, kind: "picture" | "chart" | "shape" | "other",
/// name, description, contentType (a picture's, else null), row, col (the
/// cell it is anchored to), x, y, width, height (sheet pixels at 100%) }.
let drawingToObj (drawing: SheetDrawing) (x: float) (y: float) (width: float) (height: float) : obj =
    createObj [
        "id" ==> drawing.Id
        "kind" ==> (match drawing.Content with Picture _ -> "picture" | KeptObject (kind, _) -> kind)
        "name" ==> orNull drawing.Name
        "description" ==> orNull drawing.Description
        "contentType" ==> (match drawing.Content with Picture (contentType, _, _) -> box contentType | KeptObject _ -> null)
        "row" ==> drawing.Anchor.FromRow
        "col" ==> drawing.Anchor.FromCol
        "x" ==> x
        "y" ==> y
        "width" ==> width
        "height" ==> height
    ]
