/// Information functions of the function table that ask about references,
/// sheets and the host: SHEET, SHEETS, FORMULATEXT, ISOMITTED, CELL, INFO.
module Rbt.Spreadsheet.Functions.FnInformation

open System
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Functions.Kit

module Wb = Rbt.Spreadsheet.Workbook

let moduleName = "Information"

/// The first area of a reference argument (a 3-D reference arrives as the union of its sheets).
let rec private firstArea (e: Expr) =
    match e with
    | Function (",", first :: _) -> firstArea first
    | other -> other

/// The sheet (its id) a reference is on: the formula's own when it names none.
let private sheetOfRef (a: Args) (sheet: string option) : Wb.SheetId option =
    match sheet, a.Ctx.Workbook with
    | None, _ -> a.Ctx.Sheet
    | Some name, Some wb -> Wb.resolveSheetName name wb
    | Some _, None -> None

/// The 1-based position of a sheet in the tabs (hidden sheets count).
let private sheetNumber (a: Args) (id: Wb.SheetId option) : Result<CellValue, string> =
    match a.Ctx.Workbook, id with
    | Some wb, Some id ->
        match wb.SheetOrder |> List.tryFindIndex ((=) id) with
        | Some i -> Ok (CellValue.Number (float (i + 1)))
        | None -> Result.Error Err.na
    | None, _ -> Ok (CellValue.Number 1.0)
    | Some _, None -> Result.Error Err.na

/// The top-left cell of a reference argument: (sheet qualifier, row, col).
let private topLeftOf (a: Args) (i: int) : (string option * int * int) option =
    let arg = firstArea a.Exprs.[i]
    let reference =
        match arg with
        | CellRef _ | RangeRef _ -> Some arg
        | _ -> a.Ev.Reference arg a.Ctx
    match reference with
    | Some (CellRef r) -> Some (r.Sheet, r.Row, r.Col)
    | Some (RangeRef (p, q)) -> Some (p.Sheet, min p.Row q.Row, min p.Col q.Col)
    | _ -> None

/// The model of the sheet a reference is on.
let private modelOfRef (a: Args) (sheet: string option) : SpreadsheetModel option =
    match sheet, a.Ctx.Workbook with
    | None, _ -> Some a.Ctx.Model
    | Some name, Some wb ->
        Wb.resolveSheetName name wb
        |> Option.bind (fun id -> Map.tryFind id wb.Sheets)
        |> Option.map (fun s -> resolveSheet s.Model)
    | Some _, None -> None

let private columnLetters (col: int) =
    let mutable n = col + 1
    let mutable text = ""
    while n > 0 do
        let d = (n - 1) % 26
        text <- string (char (int 'A' + d)) + text
        n <- (n - 1) / 26
    text

let private quoteSheet (name: string) =
    if name |> Seq.forall (fun c -> Char.IsLetterOrDigit c || c = '_' || c = '.') && not (Char.IsDigit name.[0]) then name
    else "'" + name.Replace("'", "''") + "'"

/// CELL("format")'s code of a number format: G, F2, ,2, C2, P2, S2, D1..D9,
/// with "-" when negative numbers have a colour and "()" when positive ones
/// stand in parentheses.
let formatCodeOf (format: string option) : string =
    let format = defaultArg format ""
    let sections = format.Split(';')
    let first = sections.[0]
    let plain = Text.RegularExpressions.Regex.Replace(first, "\"[^\"]*\"|\\[[^\\]]*\\]|\\\\.|_.|\\*.", "")
    let decimals =
        let dot = plain.IndexOf '.'
        if dot < 0 then 0 else plain.Substring(dot + 1) |> Seq.takeWhile (fun c -> c = '0' || c = '#' || c = '?') |> Seq.filter ((=) '0') |> Seq.length
    let lower = plain.ToLowerInvariant()
    let code =
        if format = "" || lower.Trim() = "general" then "G"
        elif lower.Contains "e+" || lower.Contains "e-" then "S" + string decimals
        elif lower.Contains "%" then "P" + string decimals
        elif first.Contains "$" || first.Contains "€" || first.Contains "£" || first.Contains "[$" then "C" + string decimals
        elif lower.Contains "h" || lower.Contains "d" || lower.Contains "y" || lower.Contains "m" || lower.Contains "s" then
            let hasDay = lower.Contains "d"
            let hasYear = lower.Contains "y"
            let hasHour = lower.Contains "h"
            let hasSecond = lower.Contains "s"
            let ampm = lower.Contains "am/pm" || lower.Contains "a/p"
            if hasHour && not hasDay && not hasYear then
                if ampm then (if hasSecond then "D6" else "D7") else (if hasSecond then "D8" else "D9")
            elif lower.Contains "mmm" then
                if hasDay && hasYear then "D1" elif hasDay then "D2" else "D3"
            elif hasDay && hasYear then "D4"
            elif hasDay then "D5"
            else "D4"
        elif plain.Contains "," then "," + string decimals
        else "F" + string decimals
    let negativeColour = sections.Length > 1 && sections.[1].Contains "["
    let parentheses = first.Contains "("
    code + (if negativeColour then "-" else "") + (if parentheses then "()" else "")

let private cellInfo (a: Args) : Result<CellValue, string> =
    res {
        let! kind = a.Text 0
        let kind = kind.Trim().ToLowerInvariant()
        let target =
            if a.Supplied 1 then topLeftOf a 1
            else a.Ctx.CurrentCell |> Option.map (fun c -> None, c.Row, c.Col)
        do! check ((not (a.Supplied 1)) || target.IsSome) Err.value
        let sheet, row, col = defaultArg target (None, 0, 0)
        let model = modelOfRef a sheet
        let data = model |> Option.bind (getCellAt row col)
        let value = match data with Some d -> unwrapValue d.Value | None -> CellValue.Empty
        let format = data |> Option.bind (fun d -> d.Format.NumberFormat)
        let isText = match value with CellValue.String _ -> true | _ -> false
        let isConstant = match data with Some { Value = CellValue.Formula _ } -> false | _ -> true
        let sheetName =
            match sheetOfRef a sheet, a.Ctx.Workbook with
            | Some id, Some wb -> Map.tryFind id wb.Sheets |> Option.map (fun s -> s.Name)
            | _ -> None
        match kind with
        | "address" ->
            let local = "$" + columnLetters col + "$" + string (row + 1)
            match sheet with
            | Some _ when sheetOfRef a sheet <> a.Ctx.Sheet -> return CellValue.String (quoteSheet (defaultArg sheetName "") + "!" + local)
            | _ -> return CellValue.String local
        | "col" -> return CellValue.Number (float (col + 1))
        | "row" -> return CellValue.Number (float (row + 1))
        | "contents" -> return value
        | "type" ->
            return CellValue.String (match value with CellValue.Empty -> "b" | CellValue.String _ when isConstant -> "l" | _ -> "v")
        | "format" -> return CellValue.String (formatCodeOf format)
        | "color" ->
            let sections = (defaultArg format "").Split(';')
            return CellValue.Number (if sections.Length > 1 && sections.[1].Contains "[" then 1.0 else 0.0)
        | "parentheses" -> return CellValue.Number (if (defaultArg format "").Split(';').[0].Contains "(" then 1.0 else 0.0)
        | "prefix" ->
            let align = data |> Option.bind (fun d -> d.Format.HorizontalAlign)
            return CellValue.String (
                if not isText || not isConstant then ""
                else
                    match align with
                    | Some "left" -> "'"
                    | Some "right" -> "\""
                    | Some "center" -> "^"
                    | Some "fill" -> "\\"
                    | _ -> "")
        | "protect" -> return CellValue.Number 1.0
        | "width" ->
            match model with
            | Some m ->
                let px = match Map.tryFind col m.ColumnWidths with Some w -> w | None -> m.DefaultColumnWidth
                return CellValue.Number (Math.Max (0.0, Math.Floor ((px - 5.0) / 7.0)))
            | None -> return! fail Err.ref
        | "filename" ->
            match hostInfo "filename" with
            | Some file when file <> "" ->
                return CellValue.String (defaultArg (hostInfo "directory") "" + "[" + file + "]" + defaultArg sheetName "")
            | _ -> return CellValue.String ""
        | _ -> return! fail Err.value }

let private infoValue (a: Args) : Result<CellValue, string> =
    res {
        let! field = a.Text 0
        let field = field.Trim().ToLowerInvariant()
        let given = hostInfo field
        match field with
        | "directory" | "osversion" -> return CellValue.String (defaultArg given "")
        | "origin" -> return CellValue.String (defaultArg given "$A:$A$1")
        | "system" -> return CellValue.String (defaultArg given "pcdos")
        | "release" -> return CellValue.String (defaultArg given "16.0")
        | "recalc" -> return CellValue.String (defaultArg given "Automatic")
        | "numfile" ->
            match given with
            | Some text ->
                match Double.TryParse(text, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
                | true, n -> return CellValue.Number n
                | _ -> return CellValue.String text
            | None ->
                return CellValue.Number (match a.Ctx.Workbook with Some wb -> float wb.SheetOrder.Length | None -> 1.0)
        | "memavail" | "memused" | "totmem" -> return! fail Err.na
        | _ -> return! fail Err.value }

let functions : FunctionDef list =
    [ scalar "SHEET" 0 1 "" (fun a ->
        if not (a.Supplied 0) then sheetNumber a a.Ctx.Sheet
        else
            match topLeftOf a 0 with
            | Some (sheet, _, _) -> sheetNumber a (sheetOfRef a sheet)
            | None ->
                match a.Value 0 with
                | CellValue.String name ->
                    match a.Ctx.Workbook with
                    | Some wb -> sheetNumber a (Wb.resolveSheetName name wb)
                    | None -> Result.Error Err.na
                | CellValue.Error e -> Result.Error e
                | _ -> Result.Error Err.na)
      scalar "SHEETS" 0 1 "" (fun a ->
        match a.Ctx.Workbook with
        | None -> Ok (CellValue.Number 1.0)
        | Some wb ->
            if not (a.Supplied 0) then Ok (CellValue.Number (float wb.SheetOrder.Length))
            else
                let sheetQualifier (e: Expr) =
                    match e with
                    | CellRef r -> Some r.Sheet
                    | RangeRef (p, _) -> Some p.Sheet
                    | other -> match a.Ev.Reference other a.Ctx with
                               | Some (CellRef r) -> Some r.Sheet
                               | Some (RangeRef (p, _)) -> Some p.Sheet
                               | _ -> None
                let parts = match a.Expr 0 with Function (",", parts) -> parts | e -> [ e ]
                let sheets =
                    parts |> List.map (fun part ->
                        match sheetQualifier part with
                        | Some (Some q) when q.IndexOf ':' >= 0 -> Some (threeDSheetIds wb q)
                        | Some q -> Some (Option.toList (sheetOfRef a q))
                        | None -> None)
                if sheets |> List.exists Option.isNone then Result.Error Err.ref
                else Ok (CellValue.Number (float (sheets |> List.collect Option.get |> List.distinct |> List.length))))
      scalar "FORMULATEXT" 1 1 "" (fun a ->
        match topLeftOf a 0 with
        | None -> (match a.Value 0 with CellValue.Error e -> Result.Error e | _ -> Result.Error Err.na)
        | Some (sheet, row, col) ->
            match modelOfRef a sheet with
            | None -> Result.Error Err.ref
            | Some m ->
                match getCellAt row col m with
                | Some { Value = CellValue.Formula (formula, _) } ->
                    Ok (CellValue.String (if formula.StartsWith "=" then formula else "=" + formula))
                | _ -> Result.Error Err.na)
      scalar "AREAS" 1 1 "|R" (fun a ->
        // A union (`(A1:B2,C3)`, a 3-D reference) is its areas; a reference is one.
        let rec areas (e: Expr) =
            match e with
            | Function (",", parts) -> parts |> List.map areas |> List.fold (fun acc n -> match acc, n with Some x, Some y -> Some (x + y) | _ -> None) (Some 0)
            | CellRef _ | RangeRef _ -> Some 1
            | NameRef _ ->
                match a.Ctx.Workbook |> Option.bind (fun wb -> resolveNameExpr (match e with NameRef n -> n | _ -> "") a.Ctx.Sheet wb) with
                | Some resolved -> areas resolved
                | None -> None
            | other -> if (a.Ev.Reference other a.Ctx).IsSome then Some 1 else None
        match areas (a.Expr 0) with
        | Some n -> Ok (CellValue.Number (float n))
        | None -> (match a.Value 0 with CellValue.Error e -> Result.Error e | _ -> Result.Error Err.value))
      scalar "ISOMITTED" 1 1 "" (fun a -> Ok (CellValue.Boolean (not (a.Supplied 0))))
      volatileFn (scalar "CELL" 1 2 "" cellInfo)
      volatileFn (scalar "INFO" 1 1 "|S" infoValue) ]

let install () = registerFunctions moduleName functions
