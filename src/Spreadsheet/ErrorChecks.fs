/// The background error checks Excel 2019 runs by default and marks with a green triangle at a cell's top left,
/// as far as they are known here. Each checked in Excel 2019 on generated cells (synthetic-error-checks.xlsx,
/// tools/excel-parity/error-probe.vbs: `Range.Errors(n).Value`) and on the real-estate models:
/// - a number stored as text: text that would be a number typed in ("123", "1,000", " 12", "12%", "1e3", "$5",
///   "-4.5", "(7)", "0012"; a date is not: "abc" and "1/1/99" are not);
/// - a text date of a two-digit year ("1/1/99", "Jan-99");
/// - a formula whose value is an error, but #N/A (=1/0 is marked, =NA() is not, a typed #REF! is not);
/// - a formula unlike the formulas on both sides of it in its row, or above and below it in its column, which are
///   alike (as R1C1 formulas): =F9*5 between =E9+1 and =G9+1;
/// - a formula whose range of a row or a column stops just short of a number (=SUM(B1:B3) over B4; not when that
///   cell holds a formula).
/// The cells of `<ignoredErrors>` are not marked for the checks it names.
module Rbt.Spreadsheet.ErrorChecks

open System
open System.Text.RegularExpressions
open Rbt.Spreadsheet.Model

let private columnOf (letters: string) =
    letters.ToUpperInvariant() |> Seq.fold (fun total ch -> total * 26 + (int ch - int 'A' + 1)) 0

let private referencePattern = Regex(@"(?<![A-Za-z0-9_.$])(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})(?![A-Za-z0-9_(])", RegexOptions.Compiled)
let private rangePattern = Regex(@"(?<![A-Za-z0-9_.$!])\$?([A-Za-z]{1,3})\$?([0-9]{1,7}):\$?([A-Za-z]{1,3})\$?([0-9]{1,7})(?![A-Za-z0-9_(])", RegexOptions.Compiled)

/// The parts of a formula outside its string literals (odd items are the literals).
let private outsideStrings (formula: string) : string list =
    formula.Split('"') |> Array.toList

/// A formula with its references relative to the cell it is in (R1C1-like), to tell formulas that are alike.
let relativeForm (formula: string) (at: CellAddress) : string =
    outsideStrings formula
    |> List.mapi (fun index part ->
        if index % 2 = 1 then "\"" + part + "\""
        else
            referencePattern.Replace(part, (fun (m: Match) ->
                let col = columnOf m.Groups.[2].Value - 1
                let row = int m.Groups.[4].Value - 1
                if col > 16383 || row > 1048575 then m.Value
                else
                    let r = if m.Groups.[3].Value = "$" then sprintf "R%d" row else sprintf "R[%d]" (row - at.Row)
                    let c = if m.Groups.[1].Value = "$" then sprintf "C%d" col else sprintf "C[%d]" (col - at.Col)
                    r + c)))
    |> String.concat ""

/// Areas of an `sqref` ("A1:A11 L1:L2").
let private areasOf (sqref: string) : (int * int * int * int) list =
    sqref.Split([| ' ' |], StringSplitOptions.RemoveEmptyEntries)
    |> Array.toList
    |> List.choose (fun part ->
        let ends = part.Replace("$", "").Split(':')
        let cell (text: string) =
            let m = Regex.Match(text, "^([A-Za-z]{1,3})([0-9]+)$")
            if m.Success then Some (int m.Groups.[2].Value - 1, columnOf m.Groups.[1].Value - 1) else None
        match cell ends.[0], cell ends.[ends.Length - 1] with
        | Some (r0, c0), Some (r1, c1) -> Some (min r0 r1, min c0 c1, max r0 r1, max c0 c1)
        | _ -> None)

let private isNumberConstant (cell: CellData option) =
    match cell with
    | Some { Value = Number _ } -> true
    | _ -> false

/// What a text cell is to the checks: a number stored as text, a text date of a two-digit year, or neither.
let private textCheck (text: string) : string option =
    if String.IsNullOrWhiteSpace text then None
    else
        let parsed = Input.parseInputIn 2026 text
        match parsed.Value with
        | Number _ ->
            match parsed.NumberFormat with
            | Some code when CellText.isDateTimeNumberFormat code ->
                // (A date whose year is written with two digits.)
                if Regex.IsMatch(text, "[0-9]{4}") then None
                elif Regex.IsMatch(text, "[0-9]{1,2}[/-][0-9]{1,2}[/-][0-9]{2}$|[A-Za-z]+[- ][0-9]{2}$") then Some "twoDigitTextYear"
                else None
            | _ -> Some "numberStoredAsText"
        | _ -> None

/// The cells of a sheet Excel marks, and which check marks each.
let markedCells (cells: Map<CellAddress, CellData>) (ignored: (string * string list) list) : Map<CellAddress, string> =
    let ignoredAreas = ignored |> List.map (fun (sqref, checks) -> areasOf sqref, Set.ofList checks)
    let isIgnored (at: CellAddress) (check: string) =
        ignoredAreas |> List.exists (fun (areas, checks) ->
            Set.contains check checks && areas |> List.exists (fun (r0, c0, r1, c1) -> at.Row >= r0 && at.Row <= r1 && at.Col >= c0 && at.Col <= c1))
    let forms = Collections.Generic.Dictionary<CellAddress, string>()
    let formOf (at: CellAddress) : string option =
        match forms.TryGetValue at with
        | true, form -> Some form
        | _ ->
            match Map.tryFind at cells with
            | Some { Value = Formula (text, _) } ->
                let form = relativeForm text at
                forms.[at] <- form
                Some form
            | _ -> None
    let mutable marked = Map.empty
    let mark (at: CellAddress) (check: string) = if not (isIgnored at check) then marked <- Map.add at check marked
    for KeyValue (at, cell) in cells do
        match cell.Value with
        | String text -> (match textCheck text with Some check -> mark at check | None -> ())
        | Formula (text, value) ->
            (match unwrapFormulaValue value with
             | Error code when code.Trim('#') <> "N/A" -> mark at "evalError"
             | _ -> ())
            if not (Map.containsKey at marked) then
                let mine = formOf at
                let alike (a: CellAddress) (b: CellAddress) =
                    match formOf a, formOf b with
                    | Some x, Some y -> x = y && Some x <> mine
                    | _ -> false
                let left, right = { at with Col = at.Col - 1 }, { at with Col = at.Col + 1 }
                let above, below = { at with Row = at.Row - 1 }, { at with Row = at.Row + 1 }
                if (at.Col > 0 && alike left right) || (at.Row > 0 && alike above below) then mark at "formula"
            if not (Map.containsKey at marked) then
                let omits =
                    outsideStrings text
                    |> List.indexed
                    |> List.filter (fun (index, _) -> index % 2 = 0)
                    |> List.exists (fun (_, part) ->
                        rangePattern.Matches(part)
                        |> Seq.cast<Match>
                        |> Seq.exists (fun m ->
                            let c0, r0 = columnOf m.Groups.[1].Value - 1, int m.Groups.[2].Value - 1
                            let c1, r1 = columnOf m.Groups.[3].Value - 1, int m.Groups.[4].Value - 1
                            let r0, r1, c0, c1 = min r0 r1, max r0 r1, min c0 c1, max c0 c1
                            let next =
                                if c0 = c1 && r1 > r0 then Some { Row = r1 + 1; Col = c0 }
                                elif r0 = r1 && c1 > c0 then Some { Row = r0; Col = c1 + 1 }
                                else None
                            match next with
                            | Some cell when cell <> at -> isNumberConstant (Map.tryFind cell cells)
                            | _ -> false))
                if omits then mark at "formulaRange"
        | _ -> ()
    marked
