module Rbt.Spreadsheet.Formula

open System
open System.Collections.Generic
open System.Globalization
open Rbt.Spreadsheet.Model

/// Deliberately curated user-facing syntax. Evaluator lifting signatures below
/// describe array/scalar evaluation, NOT function arguments.
type FunctionHelp = { Name: string; Arguments: string array; Description: string; RepeatFrom: int option }

let functionHelpCatalog : FunctionHelp array =
    [| { Name = "SUM"; Arguments = [| "number1: first number or range"; "number2...: more numbers or ranges (optional)" |]; Description = "Adds numbers in values and ranges."; RepeatFrom = Some 1 }
       { Name = "IF"; Arguments = [| "test: condition to check"; "value_if_true: result when true"; "value_if_false: result when false (optional)" |]; Description = "Returns one of two values based on a condition."; RepeatFrom = None }
       { Name = "AVERAGE"; Arguments = [| "number1: first number or range"; "number2...: more numbers or ranges (optional)" |]; Description = "Arithmetic mean of numeric values."; RepeatFrom = Some 1 }
       { Name = "COUNT"; Arguments = [| "value1: first value or range"; "value2...: more values or ranges (optional)" |]; Description = "Counts numeric values."; RepeatFrom = Some 1 }
       { Name = "COUNTA"; Arguments = [| "value1: first value or range"; "value2...: more values or ranges (optional)" |]; Description = "Counts nonempty values."; RepeatFrom = Some 1 }
       { Name = "MIN"; Arguments = [| "number1: first number or range"; "number2...: more numbers or ranges (optional)" |]; Description = "Smallest numeric value."; RepeatFrom = Some 1 }
       { Name = "MAX"; Arguments = [| "number1: first number or range"; "number2...: more numbers or ranges (optional)" |]; Description = "Largest numeric value."; RepeatFrom = Some 1 }
       { Name = "ROUND"; Arguments = [| "number: value to round"; "num_digits: decimal places (optional; defaults to 0)" |]; Description = "Rounds a number to the specified digits."; RepeatFrom = None }
       { Name = "IFERROR"; Arguments = [| "value: expression to evaluate"; "value_if_error: result on error" |]; Description = "Returns a fallback when the expression has an error."; RepeatFrom = None }
       { Name = "AND"; Arguments = [| "logical1: first condition"; "logical2...: more conditions (optional)" |]; Description = "True when all conditions are true."; RepeatFrom = Some 1 }
       { Name = "OR"; Arguments = [| "logical1: first condition"; "logical2...: more conditions (optional)" |]; Description = "True when at least one condition is true."; RepeatFrom = Some 1 }
       { Name = "NOT"; Arguments = [| "logical: condition" |]; Description = "Reverses a logical result."; RepeatFrom = None }
       { Name = "CONCAT"; Arguments = [| "text1: first text or range"; "text2...: more text or ranges (optional)" |]; Description = "Joins text values."; RepeatFrom = Some 1 }
       { Name = "LEFT"; Arguments = [| "text: source text"; "num_chars: characters to take (optional)" |]; Description = "Takes characters from the beginning of text."; RepeatFrom = None }
       { Name = "RIGHT"; Arguments = [| "text: source text"; "num_chars: characters to take (optional)" |]; Description = "Takes characters from the end of text."; RepeatFrom = None }
       { Name = "MID"; Arguments = [| "text: source text"; "start_num: first character (1-based)"; "num_chars: characters to take" |]; Description = "Takes characters from inside text."; RepeatFrom = None }
       { Name = "LEN"; Arguments = [| "text: text to measure" |]; Description = "Counts characters in text."; RepeatFrom = None }
       { Name = "VLOOKUP"; Arguments = [| "lookup_value: value to find"; "table_array: table to search"; "col_index_num: result column (1-based)"; "range_lookup: approximate match (optional; defaults to TRUE)" |]; Description = "Finds a value in a table's first column."; RepeatFrom = None }
       { Name = "XLOOKUP"; Arguments = [| "lookup_value: value to find"; "lookup_array: values to search"; "return_array: matching results"; "if_not_found: fallback (optional)"; "match_mode: match method (optional)"; "search_mode: search direction (optional)" |]; Description = "Looks up a value and returns the corresponding result."; RepeatFrom = None }
       { Name = "SUMIF"; Arguments = [| "range: cells to test"; "criteria: condition"; "sum_range: cells to add (optional; defaults to range)" |]; Description = "Adds values whose cells meet a condition."; RepeatFrom = None }
       { Name = "SUMIFS"; Arguments = [| "sum_range: cells to add"; "criteria_range1: first cells to test"; "criteria1: first condition"; "criteria_range...: additional cells to test (optional)"; "criteria...: additional condition (optional)" |]; Description = "Adds values meeting every condition; additional conditions must be range/criteria pairs."; RepeatFrom = Some 3 }
       { Name = "COUNTIF"; Arguments = [| "range: cells to test"; "criteria: condition" |]; Description = "Counts cells meeting a condition."; RepeatFrom = None }
       { Name = "COUNTIFS"; Arguments = [| "criteria_range1: first cells to test"; "criteria1: first condition"; "criteria_range...: additional cells to test (optional)"; "criteria...: additional condition (optional)" |]; Description = "Counts cells meeting every condition; additional conditions must be range/criteria pairs."; RepeatFrom = Some 2 }
       { Name = "AVERAGEIF"; Arguments = [| "range: cells to test"; "criteria: condition"; "average_range: cells to average (optional; defaults to range)" |]; Description = "Averages matching numeric cells; no numbers gives #DIV/0!."; RepeatFrom = None }
       { Name = "AVERAGEIFS"; Arguments = [| "average_range: cells to average"; "criteria_range1: first cells to test"; "criteria1: first condition"; "criteria_range...: additional cells to test (optional)"; "criteria...: additional condition (optional)" |]; Description = "Averages numeric cells meeting every range/criteria pair; no numbers gives #DIV/0!."; RepeatFrom = Some 3 }
       { Name = "IFNA"; Arguments = [| "value: expression to evaluate"; "value_if_na: result on #N/A" |]; Description = "Returns a fallback only for #N/A errors."; RepeatFrom = None }
       { Name = "IFS"; Arguments = [| "test1: first condition"; "value1: result when true"; "test...: additional condition (optional)"; "value...: corresponding result (optional)" |]; Description = "Returns the result of the first true condition; unmatched conditions give #N/A. Add complete test/value pairs."; RepeatFrom = Some 2 }
       { Name = "SWITCH"; Arguments = [| "expression: value to compare"; "case1: first matching value"; "result1: first result"; "case...: additional matching value or final default"; "result...: corresponding result (optional)" |]; Description = "Returns the first matching case's result; an unmatched expression gives #N/A unless a final default follows the case/result pairs."; RepeatFrom = Some 3 }
       { Name = "INDEX"; Arguments = [| "array: range or array to read"; "row_num: 1-based row (or position in a single row)"; "column_num: 1-based column (optional; defaults to 1)" |]; Description = "Returns one cell from an array; index 0 selects the first cell, not an entire row or column."; RepeatFrom = None }
       { Name = "MATCH"; Arguments = [| "lookup_value: value to find"; "lookup_array: range to search"; "match_type: 0 exact, 1 ascending approximate (default), -1 descending approximate" |]; Description = "Returns the 1-based position of a match; exact text matches support wildcards."; RepeatFrom = None }
       { Name = "HLOOKUP"; Arguments = [| "lookup_value: value to find"; "table_array: table to search"; "row_index_num: result row (1-based)"; "range_lookup: approximate match (optional; defaults to TRUE)" |]; Description = "Finds a value in a table's first row and returns a cell in the selected row."; RepeatFrom = None }
       { Name = "FILTER"; Arguments = [| "array: values to filter"; "include: matching rows or columns"; "if_empty: result when nothing matches (optional; otherwise #CALC!)" |]; Description = "Spills values selected by a Boolean include array."; RepeatFrom = None }
       { Name = "SORT"; Arguments = [| "array: values to sort"; "sort_index: column or row to sort by (optional; defaults to 1)"; "sort_order: 1 ascending or -1 descending (optional; defaults to 1)"; "by_col: sort columns instead of rows (optional; defaults to FALSE)" |]; Description = "Spills a sorted copy of an array."; RepeatFrom = None }
       { Name = "UNIQUE"; Arguments = [| "array: values to deduplicate"; "by_col: compare columns instead of rows (optional; defaults to FALSE)" |]; Description = "Spills distinct rows or columns, retaining their first appearance; text comparisons ignore case."; RepeatFrom = None }
       { Name = "SEQUENCE"; Arguments = [| "rows: positive row count"; "columns: positive column count (optional; defaults to 1)"; "start: first number (optional; defaults to 1)"; "step: increment (optional; defaults to 1)" |]; Description = "Spills a row-major sequence of numbers."; RepeatFrom = None }
       { Name = "DATE"; Arguments = [| "year: year"; "month: month (may roll over)"; "day: day (may roll over)" |]; Description = "Returns a date serial number."; RepeatFrom = None }
       { Name = "TIME"; Arguments = [| "hour: hours"; "minute: minutes"; "second: seconds" |]; Description = "Returns a time fraction, wrapping at 24 hours."; RepeatFrom = None }
       { Name = "TODAY"; Arguments = [||]; Description = "Returns the current local date as a serial number; takes no arguments."; RepeatFrom = None }
       { Name = "NOW"; Arguments = [||]; Description = "Returns the current local date and time as a serial number; takes no arguments."; RepeatFrom = None }
       { Name = "YEAR"; Arguments = [| "serial_number: date serial" |]; Description = "Returns the year of a date."; RepeatFrom = None }
       { Name = "MONTH"; Arguments = [| "serial_number: date serial" |]; Description = "Returns the month of a date."; RepeatFrom = None }
       { Name = "DAY"; Arguments = [| "serial_number: date serial" |]; Description = "Returns the day of the month of a date."; RepeatFrom = None }
       { Name = "DAYS"; Arguments = [| "end_date: later date serial"; "start_date: earlier date serial" |]; Description = "Returns whole days between two dates (negative when reversed)."; RepeatFrom = None }
       { Name = "EDATE"; Arguments = [| "start_date: date serial"; "months: whole months to move" |]; Description = "Moves a date by a number of months, preserving the day where possible."; RepeatFrom = None }
       { Name = "EOMONTH"; Arguments = [| "start_date: date serial"; "months: whole months to move" |]; Description = "Returns the last day of the month the given number of months away."; RepeatFrom = None }
       { Name = "NETWORKDAYS"; Arguments = [| "start_date: start date serial"; "end_date: end date serial"; "holidays: date serials to omit (optional)" |]; Description = "Counts weekdays inclusively, excluding optional holidays; weekends are Saturday and Sunday."; RepeatFrom = None }
       { Name = "WORKDAY"; Arguments = [| "start_date: start date serial"; "days: working days to move"; "holidays: date serials to omit (optional)" |]; Description = "Moves by weekdays, excluding optional holidays; weekends are Saturday and Sunday."; RepeatFrom = None }
       { Name = "TEXT"; Arguments = [| "value: number or date serial"; "format_text: number format code" |]; Description = "Formats a numeric value as text using a supported number format code."; RepeatFrom = None }
       { Name = "TEXTJOIN"; Arguments = [| "delimiter: text between values"; "ignore_empty: whether to skip blank values"; "text1: first text or range"; "text...: additional text or ranges (optional)" |]; Description = "Joins text values and ranges with a delimiter."; RepeatFrom = Some 3 }
       { Name = "TRIM"; Arguments = [| "text: text to clean" |]; Description = "Removes leading, trailing and repeated spaces."; RepeatFrom = None }
       { Name = "LOWER"; Arguments = [| "text: text to convert" |]; Description = "Converts text to lowercase."; RepeatFrom = None }
       { Name = "UPPER"; Arguments = [| "text: text to convert" |]; Description = "Converts text to uppercase."; RepeatFrom = None }
       { Name = "SUBSTITUTE"; Arguments = [| "text: source text"; "old_text: text to replace"; "new_text: replacement text"; "instance_num: occurrence to replace (optional; otherwise all)" |]; Description = "Replaces matching text; comparisons are case-sensitive."; RepeatFrom = None }
       { Name = "REPLACE"; Arguments = [| "old_text: source text"; "start_num: first character (1-based)"; "num_chars: characters to replace"; "new_text: replacement text" |]; Description = "Replaces characters at a position."; RepeatFrom = None }
       { Name = "FIND"; Arguments = [| "find_text: case-sensitive text to find"; "within_text: text to search"; "start_num: first character to search (optional; defaults to 1)" |]; Description = "Returns the 1-based position of text, or #VALUE! when not found."; RepeatFrom = None }
       { Name = "SEARCH"; Arguments = [| "find_text: case-insensitive text or wildcard pattern"; "within_text: text to search"; "start_num: first character to search (optional; defaults to 1)" |]; Description = "Returns the 1-based position of text, or #VALUE! when not found; supports *, ? and ~ wildcards."; RepeatFrom = None }
       { Name = "ROUNDUP"; Arguments = [| "number: value to round away from zero"; "num_digits: decimal places (optional; defaults to 0)" |]; Description = "Rounds away from zero at the specified digits."; RepeatFrom = None }
       { Name = "ROUNDDOWN"; Arguments = [| "number: value to round toward zero"; "num_digits: decimal places (optional; defaults to 0)" |]; Description = "Rounds toward zero at the specified digits."; RepeatFrom = None }
       { Name = "INT"; Arguments = [| "number: value to round down" |]; Description = "Rounds down to an integer, including negative numbers."; RepeatFrom = None }
       { Name = "MOD"; Arguments = [| "number: dividend"; "divisor: nonzero divisor" |]; Description = "Returns the remainder with the sign of the divisor."; RepeatFrom = None }
       { Name = "ABS"; Arguments = [| "number: value" |]; Description = "Returns the absolute value of a number."; RepeatFrom = None }
       { Name = "SQRT"; Arguments = [| "number: nonnegative value" |]; Description = "Returns the square root; negative values give #NUM!."; RepeatFrom = None }
       { Name = "POWER"; Arguments = [| "number: base"; "power: exponent" |]; Description = "Raises a base to an exponent."; RepeatFrom = None }
       { Name = "PRODUCT"; Arguments = [| "number1: first value or range"; "number...: additional values or ranges (optional)" |]; Description = "Multiplies numeric values in values and ranges."; RepeatFrom = Some 1 }
       { Name = "SUMPRODUCT"; Arguments = [| "array1: first array or range"; "array...: additional equally shaped arrays (optional)" |]; Description = "Sums element-wise products; text, blanks and booleans count as zero."; RepeatFrom = Some 1 }
       { Name = "SUBTOTAL"; Arguments = [| "function_num: aggregate code (1–11; 101–111 also skip manually hidden rows)"; "ref1: first range"; "ref...: additional ranges (optional)" |]; Description = "Aggregates ranges while excluding filtered rows and nested subtotals."; RepeatFrom = Some 2 } |]

let findFunctionHelp (name: string) =
    functionHelpCatalog |> Array.tryFind (fun item -> String.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase))

/// A trailing argument group can repeat; the starting index lives alongside
/// the displayed labels rather than being inferred from their wording.
let functionArgumentHelp (info: FunctionHelp) (index: int) =
    if info.Arguments.Length = 0 then "No arguments"
    elif index < info.Arguments.Length then info.Arguments.[index]
    else
        match info.RepeatFrom with
        | Some first -> info.Arguments.[first + (index - first) % (info.Arguments.Length - first)]
        | None -> "Additional argument"

let private functionNameChar (c: char) = Char.IsLetterOrDigit c || c = '_' || c = '.'

/// A lexical caret context only; never interprets/evaluates formula expressions.
/// Parentheses in strings, quoted sheet names and structured references are ignored.
let formulaHelpAt (text: string) (caret: int) : (FunctionHelp * int) option =
    if String.IsNullOrEmpty text || text.[0] <> '=' then None
    else
        let frames = ResizeArray<string * int>()
        let mutable inText = false
        let mutable inSheetName = false
        let mutable bracketDepth = 0
        let mutable arrayDepth = 0
        let mutable i = 1
        let finish = max 1 (min caret text.Length)
        while i < finish do
            let c = text.[i]
            if inText then
                if c = '"' then
                    if i + 1 < text.Length && text.[i + 1] = '"' then i <- i + 1
                    else inText <- false
            elif inSheetName then
                if c = '\'' then
                    if i + 1 < text.Length && text.[i + 1] = '\'' then i <- i + 1
                    else inSheetName <- false
            elif c = '"' then inText <- true
            elif c = '\'' then inSheetName <- true
            elif c = '[' then bracketDepth <- bracketDepth + 1
            elif c = ']' then bracketDepth <- max 0 (bracketDepth - 1)
            elif bracketDepth = 0 && c = '{' then arrayDepth <- arrayDepth + 1
            elif bracketDepth = 0 && c = '}' then arrayDepth <- max 0 (arrayDepth - 1)
            elif bracketDepth = 0 then
                if c = '(' then
                    let mutable endName = i - 1
                    while endName >= 1 && Char.IsWhiteSpace text.[endName] do endName <- endName - 1
                    let mutable startName = endName
                    while startName >= 1 && functionNameChar text.[startName] do startName <- startName - 1
                    let name =
                        if startName < endName then text.Substring(startName + 1, endName - startName)
                        else ""
                    frames.Add((name, 0))
                elif c = ')' && frames.Count > 0 then Portable.Stack.dropLast frames
                elif c = ',' && frames.Count > 0 && arrayDepth = 0 then
                    let name, index = frames.[frames.Count - 1]
                    frames.[frames.Count - 1] <- (name, index + 1)
            i <- i + 1
        if inText || inSheetName || frames.Count = 0 then None
        else
            let name, index = frames.[frames.Count - 1]
            findFunctionHelp name |> Option.map (fun info -> info, index)

/// Completion is confined to a function token in a formula (not a cell
/// reference, quoted string, sheet name or an already-open function call).
let formulaFunctionCompletion (text: string) (caret: int) : (int * FunctionHelp) option =
    if String.IsNullOrEmpty text || text.[0] <> '=' || caret < 2 || caret > text.Length then None
    else
        let mutable start = caret
        while start > 1 && functionNameChar text.[start - 1] do start <- start - 1
        if start = caret || (start > 1 && not ("=+-*/^(,;:&<>".Contains(string text.[start - 1]))) then None
        else
            let prefix = text.Substring(start, caret - start)
            let before = text.Substring(0, start)
            // A quoted token may not be completed, even when the quote is unclosed.
            let mutable quoted = false
            let mutable sheet = false
            let mutable bracketDepth = 0
            let mutable arrayDepth = 0
            let mutable i = 1
            while i < before.Length do
                if before.[i] = '"' && not sheet then
                    if quoted && i + 1 < before.Length && before.[i + 1] = '"' then i <- i + 1
                    else quoted <- not quoted
                elif before.[i] = '\'' && not quoted then
                    if sheet && i + 1 < before.Length && before.[i + 1] = '\'' then i <- i + 1
                    else sheet <- not sheet
                elif not quoted && not sheet then
                    match before.[i] with
                    | '[' -> bracketDepth <- bracketDepth + 1
                    | ']' -> bracketDepth <- max 0 (bracketDepth - 1)
                    | '{' when bracketDepth = 0 -> arrayDepth <- arrayDepth + 1
                    | '}' when bracketDepth = 0 -> arrayDepth <- max 0 (arrayDepth - 1)
                    | _ -> ()
                i <- i + 1
            if quoted || sheet || bracketDepth > 0 || arrayDepth > 0 || (caret < text.Length && (functionNameChar text.[caret] || text.[caret] = '(')) then None
            else
                let matching = functionHelpCatalog |> Array.filter (fun item -> item.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                match findFunctionHelp prefix with
                | Some exact -> Some(start, exact)
                | None when matching.Length = 1 && prefix.Length >= 2 -> Some(start, matching.[0])
                | _ -> None

let private resizeArrayToList (items: System.Collections.Generic.List<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private hashSetToArray (items: HashSet<'T>) =
    let result = Array.zeroCreate<'T> items.Count
    let mutable index = 0
    for item in items do
        result.[index] <- item
        index <- index + 1
    result

#if FABLE_COMPILER
open Fable.Core

[<Emit("String.fromCodePoint($0)")>]
let private stringFromCodePoint (_codePoint: int) : string = jsNative
#else
let private stringFromCodePoint codePoint = Char.ConvertFromUtf32 codePoint
#endif

let private indexOfOrdinalIgnoreCase (needle: string) (haystack: string) (startIndex: int) =
#if FABLE_COMPILER
    if needle.Length = 0 then
        startIndex
    elif needle.Length > haystack.Length - startIndex then
        -1
    elif (let upper = haystack.ToUpperInvariant() in upper.Length = haystack.Length) && needle.ToUpperInvariant().Length = needle.Length then
        // Upper-casing kept every character in its place (it does but for a
        // few letters such as the German sharp s): the search is the
        // platform's own, far faster than a loop over characters here.
        haystack.ToUpperInvariant().IndexOf(needle.ToUpperInvariant(), startIndex, StringComparison.Ordinal)
    else
        let charEqualsIgnoreCase a b =
            a = b || Char.ToUpperInvariant a = Char.ToUpperInvariant b
        let mutable found = -1
        let mutable index = startIndex
        while found < 0 && index <= haystack.Length - needle.Length do
            let mutable offset = 0
            let mutable matches = true
            while matches && offset < needle.Length do
                if not (charEqualsIgnoreCase haystack.[index + offset] needle.[offset]) then
                    matches <- false
                offset <- offset + 1
            if matches then
                found <- index
            index <- index + 1
        found
#else
    haystack.IndexOf(needle, startIndex, StringComparison.OrdinalIgnoreCase)
#endif

let private equalsOrdinalIgnoreCase (left: string) (right: string) =
    String.Equals(left, right, StringComparison.OrdinalIgnoreCase)

let private trimBounds (text: string) =
    let mutable start = 0
    let mutable finish = text.Length - 1
    while start <= finish && Char.IsWhiteSpace text.[start] do
        start <- start + 1
    while finish >= start && Char.IsWhiteSpace text.[finish] do
        finish <- finish - 1
    struct (start, finish - start + 1)

let private sliceEqualsOrdinalIgnoreCase (text: string) start length (expected: string) =
    if length <> expected.Length then
        false
    else
        let mutable offset = 0
        let mutable equal = true
        while equal && offset < length do
            let left = text.[start + offset]
            let right = expected.[offset]
            if left <> right && Char.ToUpperInvariant left <> Char.ToUpperInvariant right then
                equal <- false
            offset <- offset + 1
        equal

let private trimmedNumberText (text: string) start length =
    if start = 0 && length = text.Length then text
    else text.Substring(start, length)

let private sliceEqualsOrdinal (text: string) start (expected: string) =
    if start < 0 || start + expected.Length > text.Length then
        false
    else
        let mutable offset = 0
        let mutable equal = true
        while equal && offset < expected.Length do
            if text.[start + offset] <> expected.[offset] then
                equal <- false
            offset <- offset + 1
        equal

// ---------------------------------------------------------------------------
// Excel error values. Stored as the bare code (no leading '#'); the display
// layer prepends the '#'. These are the seven Excel error literals plus the
// internal CYCLE sentinel for circular references.
// ---------------------------------------------------------------------------
module Err =
    let div0 = "DIV/0!"
    let value = "VALUE!"
    let ref = "REF!"
    let name = "NAME?"
    let num = "NUM!"
    let na = "N/A"
    let nul = "NULL!"
    let cycle = "CYCLE"
    /// Dynamic-array result errors: an empty FILTER/UNIQUE result, and a spill
    /// blocked by an occupied target cell.
    let calc = "CALC!"
    let spill = "SPILL!"
    /// A value that is still being worked out: a function whose module has not
    /// loaded yet, or a provider (`registerFunctionProvider`) that has not answered.
    let busy = "BUSY!"
    /// A service a function needs could not be reached (no provider for it).
    let connect = "CONNECT!"
    /// What a function needs is blocked or switched off (CALL, REGISTER.ID, file access).
    let blocked = "BLOCKED!"
    /// A field a data type does not have (FIELDVALUE, `A1.Field`).
    let field = "FIELD!"
    /// A data type the application does not know.
    let unknown = "UNKNOWN!"
    /// What a cube function shows while its data is fetched.
    let gettingData = "GETTING_DATA"
    /// Every recognized error literal, for the lexer and for what is typed into a cell.
    let all =
        [ "#DIV/0!"; "#VALUE!"; "#REF!"; "#NAME?"; "#NUM!"; "#N/A"; "#NULL!"; "#SPILL!"; "#CALC!"
          "#BUSY!"; "#CONNECT!"; "#BLOCKED!"; "#FIELD!"; "#UNKNOWN!"; "#GETTING_DATA" ]
    /// ERROR.TYPE's number for an error (0 when it has none). The numbers past 8 are those
    /// Excel 365 gives (#SPILL! 9 ... #CALC! 14); #BUSY! has no number Microsoft documents.
    let typeNumber (code: string) =
        match code with
        | "NULL!" -> 1 | "DIV/0!" -> 2 | "VALUE!" -> 3 | "REF!" -> 4 | "NAME?" -> 5 | "NUM!" -> 6 | "N/A" -> 7
        | "GETTING_DATA" -> 8 | "SPILL!" -> 9 | "CONNECT!" -> 10 | "BLOCKED!" -> 11 | "UNKNOWN!" -> 12
        | "FIELD!" -> 13 | "CALC!" -> 14
        | _ -> 0
    /// What a formula that does not parse evaluates to. Excel refuses to
    /// commit such a formula; here it is stored and shows a real error value
    /// rather than the parser's diagnostic text.
    let syntax = name

// ---------------------------------------------------------------------------
// What is known of every function before its code is there: the functions of the
// separately loaded modules (FunctionCatalog, generated from src/Spreadsheet/Functions)
// and the functions a host answers (providers). See "The function table" below.
// ---------------------------------------------------------------------------

/// The functions whose value comes from outside the engine: a host registers a
/// provider for them (`registerFunctionProvider`). Without one a formula keeps the
/// value the file cached for it; with no cached value it shows the error given here.
/// (name, the error without a provider, gives an array)
let providerFunctionSpecs : (string * string * bool) list =
    [ "CUBEKPIMEMBER", Err.na, false; "CUBEMEMBER", Err.na, false; "CUBEMEMBERPROPERTY", Err.na, false
      "CUBERANKEDMEMBER", Err.na, false; "CUBESET", Err.na, false; "CUBESETCOUNT", Err.na, false
      "CUBEVALUE", Err.na, false; "RTD", Err.na, false
      // (Code from outside: never run in this process. A host routes them to a sandbox.)
      "CALL", Err.blocked, false; "REGISTER.ID", Err.blocked, false
      "WEBSERVICE", Err.connect, false; "STOCKHISTORY", Err.connect, true; "FIELDVALUE", Err.connect, false
      "TRANSLATE", Err.connect, false; "DETECTLANGUAGE", Err.connect, false; "COPILOT", Err.connect, true
      "PY", Err.connect, true; "IMAGE", Err.connect, false
      "IMPORTCSV", Err.blocked, true; "IMPORTTEXT", Err.blocked, true ]

let private providerSpecByName =
    let table = Dictionary<string, string * bool>()
    for (name, error, gives) in providerFunctionSpecs do table.[name] <- (error, gives)
    table

let private catalogByName =
    let table = Dictionary<string, Rbt.Spreadsheet.FunctionCatalog.Entry>()
    for entry in Rbt.Spreadsheet.FunctionCatalog.entries do table.[entry.Name] <- entry
    table

let private upperName (name: string) =
    let upper = name.ToUpperInvariant()
    if upper = name then name else upper

/// The catalog's entry of a function a module provides (loaded or not).
let catalogEntry (name: string) : Rbt.Spreadsheet.FunctionCatalog.Entry option =
    match Portable.Dict.tryFind (upperName name) catalogByName with
    | ValueSome entry -> Some entry
    | ValueNone -> None

/// Is this one of the functions a host answers (`providerFunctionSpecs`)?
let isProviderFunction (name: string) = providerSpecByName.ContainsKey (upperName name)

let private catalogFlag (flag: int) (name: string) =
    match Portable.Dict.tryFind (upperName name) catalogByName with
    | ValueSome entry -> entry.Flags &&& flag <> 0
    | ValueNone -> false

// ---------------------------------------------------------------------------
// AST
// ---------------------------------------------------------------------------

/// A parsed cell reference. Row/Col are the 0-indexed resolved position.
/// RowAbs/ColAbs record whether the source used a `$` lock (kept for
/// copy/fill + conditional-format re-anchoring, ignored by plain eval).
/// Sheet is the qualifying sheet name when the ref was `Sheet!A1`.
type RefAst = {
    Sheet: string option
    Row: int
    Col: int
    RowAbs: bool
    ColAbs: bool
}

/// Which rows of a table a structured reference selects.
/// `[#Data]` (default), `[#All]`, `[#Headers]`, `[#Totals]`, `[@]`/`[#This Row]`.
type TableItemSpec =
    | TblData
    | TblAll
    | TblHeaders
    | TblTotals
    | TblThisRow

/// A structured (table) reference, e.g. `Sales[[#Totals],[Amount]]`.
/// `Table = ""` means an unqualified ref (`[@Col]` inside a table cell) that
/// resolves against whichever table contains the evaluating cell. `ColStart`
/// is None for a whole-row reference (all columns); `ColEnd` is Some only for
/// a column span `Table[[A]:[B]]`. Column names are matched case-insensitively.
type TableSel = {
    Table: string
    Item: TableItemSpec
    ColStart: string option
    ColEnd: string option
}

type Expr =
    | Number of float
    | Str of string
    | Bool of bool
    /// Error literal written directly in the formula (`#REF!` etc). Carries
    /// the bare code.
    | ErrorLit of string
    | CellRef of RefAst
    | RangeRef of RefAst * RefAst
    | TableRef of TableSel
    /// Spill-range operator `A1#` — the dynamic array that spilled from anchor
    /// A1. Resolved at eval time to the rectangle of A1 plus its spill members.
    | SpillRef of RefAst
    /// An inline materialized array (rows × cols of already-computed values).
    /// Not produced by the parser — used internally to feed a row/column to a
    /// lambda (BYROW/BYCOL) as a single array argument.
    | ArrayLit of CellValue[][]
    /// Named range / defined name — resolved at eval time against the
    /// workbook's NamedRanges (S3); for now evaluates to #NAME?.
    | NameRef of string
    /// Prefix `-`/`+` or postfix `%` (modeled as unary with op "%").
    | UnaryOp of op: string * operand: Expr
    | BinOp of left: Expr * op: string * right: Expr
    | Function of name: string * args: Expr list
    /// An argument left empty between commas (`IF(A1,,2)`). Evaluates as 0,
    /// as in Excel; functions whose optional argument has a different default
    /// (OFFSET's height, XLOOKUP's if_not_found) check for it explicitly.
    | MissingArg

// ---------------------------------------------------------------------------
// Lexer
// ---------------------------------------------------------------------------

type Token =
    | TNumber of float
    | TString of string
    /// Cell ref / function name / defined name / TRUE / FALSE — disambiguated
    /// by the parser. May contain `$` (e.g. "$A$1").
    | TIdent of string
    /// Quoted sheet name from `'My Sheet'!A1`.
    | TSheet of string
    /// Error literal, bare code (no '#').
    | TError of string
    /// The raw inner text of a structured-reference bracket group, e.g. for
    /// `Sales[[#Totals],[Amount]]` the token carries `[#Totals],[Amount]`.
    /// The parser hands this to `parseTableSpec`. `[` only ever begins a
    /// structured ref in our grammar (no array constants / R1C1 use it), so
    /// the lexer can greedily consume a balanced group.
    | TTableSpec of string
    | TPlus | TMinus | TStar | TSlash | TCaret | TPercent | TAmp
    | TEq | TNe | TLt | TGt | TLe | TGe
    | TLParen | TRParen
    | TLBracket | TRBracket
    /// Array-constant delimiters `{` `}` and its row separator `;`.
    | TLBrace | TRBrace | TSemicolon
    /// A bare `#` that isn't part of an error literal — the spill operator
    /// suffix in `A1#`.
    | THash
    | TComma | TColon | TBang
    /// `@` — the implicit-intersection operator (`=@A1:A9`).
    | TAt
    /// Whitespace standing between two operands: Excel's reference
    /// intersection operator (`A1:C3 B2:D4`). Any other whitespace is dropped.
    | TSpace
    | TEnd

let private futureFunctionPrefixes = [ "_xlfn."; "_xlws."; "_xlpm." ]

/// `spans` (when not null) gets, for each token, where it stands in `input`:
/// struct (start, end) - the intersection operator's is its white space, TEnd's
/// the end of the text.
/// The length of `[n]Sheet` at `i` of `input` when a `!` follows it (an external workbook's
/// sheet, unquoted), else 0.
let private externalSheetLength (input: string) (i: int) : int =
    let n = input.Length
    let mutable j = i + 1
    while j < n && Char.IsDigit input.[j] do j <- j + 1
    if j = i + 1 || j >= n || input.[j] <> ']' then 0
    else
        j <- j + 1
        let nameStart = j
        while j < n && (Char.IsLetterOrDigit input.[j] || input.[j] = '_' || input.[j] = '.') do j <- j + 1
        if j > nameStart && j < n && input.[j] = '!' && not (Char.IsDigit input.[nameStart]) then j - i else 0

let private tokenizeSpansFrom (input: string) startIndex (spans: System.Collections.Generic.List<struct (int * int)>) : Token array =
    let n = input.Length
    let tokens = System.Collections.Generic.List<Token>(max 1 (n - startIndex + 1))
    let mutable i = startIndex
    while i < n do
        let tokenStart = i
        let before = tokens.Count
        let c = input.[i]
        if Char.IsWhiteSpace c then
            while i < n && Char.IsWhiteSpace input.[i] do i <- i + 1
            // Whitespace between the end of one operand and the start of the
            // next is the intersection operator; everywhere else it is layout.
            if i < n && tokens.Count > 0 then
                let next = input.[i]
                let startsOperand = next = '$' || next = '\'' || next = '_' || Char.IsLetter next
                match tokens.[tokens.Count - 1] with
                | TIdent _ | THash when startsOperand -> tokens.Add TSpace
                | TRParen when startsOperand || next = '(' -> tokens.Add TSpace
                | _ -> ()
        elif Char.IsDigit c || (c = '.' && i + 1 < n && Char.IsDigit input.[i + 1]) then
            // Number, with optional fraction and scientific exponent.
            let start = i
            while i < n && Char.IsDigit input.[i] do i <- i + 1
            if i < n && input.[i] = '.' then
                i <- i + 1
                while i < n && Char.IsDigit input.[i] do i <- i + 1
            if i < n && (input.[i] = 'e' || input.[i] = 'E') then
                let save = i
                i <- i + 1
                if i < n && (input.[i] = '+' || input.[i] = '-') then i <- i + 1
                if i < n && Char.IsDigit input.[i] then
                    while i < n && Char.IsDigit input.[i] do i <- i + 1
                else
                    // 'E' not followed by an exponent — not part of the number.
                    i <- save
            let text = input.Substring(start, i - start)
            // A number written in a formula keeps its first 15 significant digits, the rest cut.
            tokens.Add(TNumber (Double.Parse(NumberText.cutToSignificantDigits text, CultureInfo.InvariantCulture)))
        elif c = '"' then
            // String literal. "" is an escaped quote.
            i <- i + 1
            let sb = System.Text.StringBuilder()
            let mutable closed = false
            while i < n && not closed do
                if input.[i] = '"' then
                    if i + 1 < n && input.[i + 1] = '"' then
                        sb.Append('"') |> ignore
                        i <- i + 2
                    else
                        closed <- true
                        i <- i + 1
                else
                    sb.Append(input.[i]) |> ignore
                    i <- i + 1
            if not closed then failwith "Unterminated string literal"
            tokens.Add(TString (sb.ToString()))
        elif c = '\'' then
            // Quoted sheet name. '' is an escaped apostrophe.
            i <- i + 1
            let sb = System.Text.StringBuilder()
            let mutable closed = false
            while i < n && not closed do
                if input.[i] = '\'' then
                    if i + 1 < n && input.[i + 1] = '\'' then
                        sb.Append('\'') |> ignore
                        i <- i + 2
                    else
                        closed <- true
                        i <- i + 1
                else
                    sb.Append(input.[i]) |> ignore
                    i <- i + 1
            if not closed then failwith "Unterminated sheet name"
            tokens.Add(TSheet (sb.ToString()))
        elif c = '#' then
            // Error literal: longest match against the known set.
            let rest = input.Substring(i, n - i)
            match Err.all |> List.tryFind (fun e -> rest.StartsWith(e, StringComparison.OrdinalIgnoreCase)) with
            | Some lit ->
                tokens.Add(TError (lit.Substring 1))
                i <- i + lit.Length
            | None ->
                tokens.Add(THash)  // bare '#': spill operator (A1#)
                i <- i + 1
        elif c = '$' || Char.IsLetter c || c = '_' then
            // Reference / name / function name. Allow '$', letters, digits,
            // '_', '.' so "$A$1", "Sheet_2", "_xlfn.FOO" all lex as one token.
            let start = i
            while i < n && (input.[i] = '$' || Char.IsLetterOrDigit input.[i] || input.[i] = '_' || input.[i] = '.') do
                i <- i + 1
            // XLSX files store newer functions and LAMBDA parameters behind
            // forward-compatibility prefixes (`_xlfn.XLOOKUP`,
            // `_xlfn._xlws.FILTER`, `_xlpm.x`); the name is what follows.
            let mutable ident = input.Substring(start, i - start)
            let mutable stripped = true
            while stripped do
                stripped <- false
                for prefix in futureFunctionPrefixes do
                    if ident.Length > prefix.Length && ident.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) then
                        ident <- ident.Substring prefix.Length
                        stripped <- true
            tokens.Add(TIdent ident)
        else
            // Operators / punctuation. Two-char operators checked first.
            if i + 1 < n && c = '<' && input.[i + 1] = '=' then
                tokens.Add(TLe)
                i <- i + 2
            elif i + 1 < n && c = '>' && input.[i + 1] = '=' then
                tokens.Add(TGe)
                i <- i + 2
            elif i + 1 < n && c = '<' && input.[i + 1] = '>' then
                tokens.Add(TNe)
                i <- i + 2
            else
                match c with
                | '+' -> tokens.Add(TPlus); i <- i + 1
                | '-' -> tokens.Add(TMinus); i <- i + 1
                | '*' -> tokens.Add(TStar); i <- i + 1
                | '/' -> tokens.Add(TSlash); i <- i + 1
                | '^' -> tokens.Add(TCaret); i <- i + 1
                | '%' -> tokens.Add(TPercent); i <- i + 1
                | '&' -> tokens.Add(TAmp); i <- i + 1
                | '=' -> tokens.Add(TEq); i <- i + 1
                | '<' -> tokens.Add(TLt); i <- i + 1
                | '>' -> tokens.Add(TGt); i <- i + 1
                | '(' -> tokens.Add(TLParen); i <- i + 1
                | ')' -> tokens.Add(TRParen); i <- i + 1
                | '[' when externalSheetLength input i > 0 ->
                    // An external workbook's sheet, unquoted: `[1]Data!A1`.
                    let length = externalSheetLength input i
                    tokens.Add(TSheet (input.Substring(i, length)))
                    i <- i + length
                | '[' ->
                    // Structured-reference bracket group: consume a balanced
                    // [...] (nests one level, e.g. [[#Data],[Amount]]) and emit
                    // its inner text as a single TTableSpec for the parser.
                    let start = i + 1
                    let mutable depth = 1
                    let mutable j = i + 1
                    while j < n && depth > 0 do
                        match input.[j] with
                        | '[' -> depth <- depth + 1; j <- j + 1
                        | ']' -> depth <- depth - 1; (if depth > 0 then j <- j + 1)
                        | _ -> j <- j + 1
                    if depth > 0 then failwith "Unterminated structured reference"
                    let inner = input.Substring(start, (min j n) - start)
                    tokens.Add(TTableSpec inner)
                    i <- (if j < n then j + 1 else n)
                | ']' -> tokens.Add(TRBracket); i <- i + 1
                | '{' -> tokens.Add(TLBrace); i <- i + 1
                | '}' -> tokens.Add(TRBrace); i <- i + 1
                | ';' -> tokens.Add(TSemicolon); i <- i + 1
                | ',' -> tokens.Add(TComma); i <- i + 1
                | ':' -> tokens.Add(TColon); i <- i + 1
                | '!' -> tokens.Add(TBang); i <- i + 1
                | '@' -> tokens.Add(TAt); i <- i + 1
                // Anything else is not formula syntax. Dropping it would make
                // a different formula out of the text, so the parse fails.
                | other -> failwithf "Unexpected character '%c'" other
        if not (isNull spans) then
            for _ in before .. tokens.Count - 1 do spans.Add(struct (tokenStart, i))
    tokens.Add(TEnd)
    if not (isNull spans) then spans.Add(struct (n, n))
    tokens.ToArray()

let private tokenizeFrom (input: string) startIndex : Token array =
    tokenizeSpansFrom input startIndex null

let tokenize (input: string) : Token array =
    tokenizeFrom input 0

// ---------------------------------------------------------------------------
// Reference parsing
// ---------------------------------------------------------------------------

/// Parse an A1-style reference token (possibly with `$` locks) into a RefAst
/// with no sheet. Returns None when the token isn't a well-formed cell ref.
let parseRefToken (token: string) : RefAst option =
    // [ $ ] letters [ $ ] digits
    let mutable idx = 0
    let len = token.Length
    let colAbs = idx < len && token.[idx] = '$'
    if colAbs then idx <- idx + 1
    let mutable col = 0
    let mutable hasCol = false
    let mutable valid = true
    while valid && idx < len && Char.IsLetter token.[idx] do
        hasCol <- true
        let upper = Char.ToUpperInvariant token.[idx]
        let digit = int upper - int 'A' + 1
        if col > (Int32.MaxValue - digit) / 26 then
            valid <- false
        else
            col <- col * 26 + digit
            idx <- idx + 1
    let rowAbs = idx < len && token.[idx] = '$'
    if rowAbs then idx <- idx + 1
    let mutable row = 0
    let mutable hasRow = false
    while valid && idx < len && Char.IsDigit token.[idx] do
        hasRow <- true
        let digit = int token.[idx] - int '0'
        if row > (Int32.MaxValue - digit) / 10 then
            valid <- false
        else
            row <- row * 10 + digit
            idx <- idx + 1
    if valid && idx = len && hasCol && hasRow && row >= 1 && col >= 1 then
        Some {
            Sheet = None
            Row = row - 1
            Col = col - 1
            RowAbs = rowAbs
            ColAbs = colAbs
        }
    else
        None

/// Parse the inner text of a structured-reference bracket group into a
/// `TableSel`. Handles the column shorthand (`Amount`), bracketed and
/// multi-word column names (`[Sales Amount]`), item specifiers
/// (`#All`/`#Data`/`#Headers`/`#Totals`/`#This Row`), the `@` this-row
/// shorthand, column spans (`[A]:[B]`), and combinations
/// (`[#Totals],[Amount]`). `table` is the qualifying table name ("" for an
/// unqualified `[@Col]` resolved against the containing table at eval time).
let parseTableSpec (table: string) (raw: string) : TableSel =
    // Split a string on a delimiter char that sits at bracket depth 0.
    let splitTop (delim: char) (s: string) : string list =
        let parts = System.Collections.Generic.List<string>()
        let sb = System.Text.StringBuilder()
        let mutable depth = 0
        for ch in s do
            if ch = '[' then depth <- depth + 1; sb.Append ch |> ignore
            elif ch = ']' then depth <- max 0 (depth - 1); sb.Append ch |> ignore
            elif ch = delim && depth = 0 then parts.Add(sb.ToString()); sb.Clear() |> ignore
            else sb.Append ch |> ignore
        parts.Add(sb.ToString())
        resizeArrayToList parts
    let unbracket (s: string) =
        let t = s.Trim()
        if t.StartsWith "[" && t.EndsWith "]" && t.Length >= 2 then t.Substring(1, t.Length - 2).Trim()
        else t
    let mutable item = TblData
    let mutable colStart = None
    let mutable colEnd = None
    let setItemKw (kw: string) =
        match kw.Trim().TrimStart('#').ToLowerInvariant() with
        | "all" -> item <- TblAll
        | "data" -> item <- TblData
        | "headers" -> item <- TblHeaders
        | "totals" -> item <- TblTotals
        | "this row" | "thisrow" -> item <- TblThisRow
        | _ -> ()
    let setCols (part: string) =
        match splitTop ':' part with
        | [ a; b ] -> colStart <- Some (unbracket a); colEnd <- Some (unbracket b)
        | _ -> let c = unbracket part in (if c <> "" then colStart <- Some c)
    // A leading '@' (outside any inner bracket) is the this-row shorthand.
    let s = raw.Trim()
    let s =
        if s.StartsWith "@" then item <- TblThisRow; s.Substring(1).Trim()
        else s
    for seg in splitTop ',' s do
        let t = seg.Trim()
        if t = "" then ()
        else
            let inner = unbracket t
            if inner.StartsWith "#" then setItemKw inner
            elif t.StartsWith "@" then item <- TblThisRow; setCols (t.Substring(1))
            else setCols t
    { Table = table; Item = item; ColStart = colStart; ColEnd = colEnd }

// ---------------------------------------------------------------------------
// Parser — recursive descent, Excel operator precedence.
//   loosest → tightest:
//     comparison (= <> < > <= >=)
//     concat (&)
//     add/sub (+ -)
//     mul/div (* /)
//     exponent (^)
//     percent (postfix %)
//     unary (prefix - +)
//     range (a : b)
//     primary
//   (Excel binds unary minus tighter than ^, so `-2^2` = 4.)
// ---------------------------------------------------------------------------

/// Last row / column index of a sheet (Excel's 1,048,576 × 16,384 grid). A
/// range whose rows run 0..lastSheetRow is a whole-column reference (`A:C`);
/// one whose columns run 0..lastSheetCol is a whole-row reference (`1:3`).
[<Literal>]
let lastSheetRow = 1048575
[<Literal>]
let lastSheetCol = 16383

/// A column-only reference token (`A`, `$AB`): (0-indexed column, absolute).
let private parseColumnToken (token: string) : (int * bool) option =
    let isAbs = token.Length > 0 && token.[0] = '$'
    let letters = if isAbs then token.Substring 1 else token
    if letters.Length = 0 || letters.Length > 3 then None
    else
        let mutable col = 0
        let mutable valid = true
        for ch in letters do
            if Char.IsLetter ch && ch < '\u0080' then col <- col * 26 + (int (Char.ToUpperInvariant ch) - int 'A' + 1)
            else valid <- false
        if valid && col >= 1 && col - 1 <= lastSheetCol then Some (col - 1, isAbs) else None

/// A row-only reference token written with a lock (`$12`): 0-indexed row.
let private parseAbsoluteRowToken (token: string) : int option =
    if token.Length < 2 || token.[0] <> '$' then None
    else
        let mutable row = 0
        let mutable valid = true
        for i = 1 to token.Length - 1 do
            if Char.IsDigit token.[i] && row <= lastSheetRow then row <- row * 10 + (int token.[i] - int '0')
            else valid <- false
        if valid && row >= 1 && row - 1 <= lastSheetRow then Some (row - 1) else None

/// `A:C` — whole columns, modelled as the range from row 1 to the last row.
let private wholeColumnsRef (sheet: string option) (c1: int, abs1: bool) (c2: int, abs2: bool) : Expr =
    RangeRef ({ Sheet = sheet; Row = 0; Col = c1; RowAbs = true; ColAbs = abs1 },
              { Sheet = sheet; Row = lastSheetRow; Col = c2; RowAbs = true; ColAbs = abs2 })

/// `1:3` — whole rows, modelled as the range from column A to the last column.
let private wholeRowsRef (sheet: string option) (r1: int, abs1: bool) (r2: int, abs2: bool) : Expr =
    RangeRef ({ Sheet = sheet; Row = r1; Col = 0; RowAbs = abs1; ColAbs = true },
              { Sheet = sheet; Row = r2; Col = lastSheetCol; RowAbs = abs2; ColAbs = true })

/// A row number written as a numeric token (`12` in `12:14`).
let private rowOfNumber (n: float) : int option =
    if n >= 1.0 && n <= float (lastSheetRow + 1) && n = Math.Floor n then Some (int n - 1) else None

/// Reference operators are carried in the tree as calls with names no
/// function can have, so every generic tree walk (shifting, re-pointing after
/// a row/column edit, dependency extraction) covers them without a new case.
/// ` ` intersects two references, `,` unites references, `:` spans two
/// references that are not both plain cells, `@` is implicit intersection.
[<Literal>]
let intersectionOperatorName = " "
[<Literal>]
let unionOperatorName = ","
[<Literal>]
let rangeOperatorName = ":"
[<Literal>]
let implicitIntersectionName = "SINGLE"

// What depends on a formula's parsed form alone is worked out once and kept
// with the form (weakly: it goes when the form goes) - a formula is evaluated
// many times, its form stays.
#if FABLE_COMPILER
[<Emit("new WeakMap()")>]
let private newFormTable () : obj = failwith "JS only"
[<Emit("$0.get($1)")>]
let private formTableGet (table: obj) (key: obj) : obj = failwith "JS only"
[<Emit("$0.set($1, $2)")>]
let private formTableSet (table: obj) (key: obj) (value: obj) : unit = failwith "JS only"
#else
let private newFormTable () = System.Runtime.CompilerServices.ConditionalWeakTable<obj, obj>()
let private formTableGet (table: System.Runtime.CompilerServices.ConditionalWeakTable<obj, obj>) (key: obj) : obj =
    match table.TryGetValue key with
    | true, value -> value
    | false, _ -> null
let private formTableSet (table: System.Runtime.CompilerServices.ConditionalWeakTable<obj, obj>) (key: obj) (value: obj) : unit =
    table.GetValue (key, (fun _ -> value)) |> ignore
#endif
/// The formulas written in parentheses as a whole (`=(A1-B1)`), by their parsed form: Excel takes
/// the parentheses for an operation of its own, after the formula's last addition (`evalCell`).
let private wrappedForms = newFormTable ()

type private Parser = {
    Tokens: Token[]
    mutable Pos: int
    mutable PrimaryDepth: int
    /// The formula is in parentheses as a whole: they open at its first token and close at its last.
    mutable Wrapped: bool
}

let private parserMaxPrimaryDepth = 1024

let private peekP (p: Parser) = if p.Pos < p.Tokens.Length then p.Tokens.[p.Pos] else TEnd
let private nextP (p: Parser) = let t = peekP p in p.Pos <- p.Pos + 1; t
let private expectP (p: Parser) (t: Token) =
    if peekP p = t then p.Pos <- p.Pos + 1
    else failwithf "Expected %A but got %A" t (peekP p)
let private enterPrimaryP (p: Parser) =
    p.PrimaryDepth <- p.PrimaryDepth + 1
    if p.PrimaryDepth > parserMaxPrimaryDepth then
        failwith "Formula nesting too deep"
let private leavePrimaryP (p: Parser) =
    p.PrimaryDepth <- p.PrimaryDepth - 1

let rec private parsePrimary (p: Parser) : Expr =
    enterPrimaryP p
    try
        match peekP p with
        | TNumber n -> p.Pos <- p.Pos + 1; Number n
        | TString s -> p.Pos <- p.Pos + 1; Str s
        | TError e -> p.Pos <- p.Pos + 1; ErrorLit e
        | TMinus -> p.Pos <- p.Pos + 1; UnaryOp("-", parseUnary p)
        | TPlus -> p.Pos <- p.Pos + 1; UnaryOp("+", parseUnary p)
        | TAt -> p.Pos <- p.Pos + 1; Function(implicitIntersectionName, [ parseRange p ])
        | TLParen ->
            let opensFormula = p.Pos = 0
            p.Pos <- p.Pos + 1
            let e = parseComparison p
            // `(A1:A2,C1:C2)` — a comma inside plain parentheses is the
            // reference union operator.
            if peekP p = TComma then
                let parts = ResizeArray<Expr>()
                parts.Add e
                while peekP p = TComma do
                    p.Pos <- p.Pos + 1
                    parts.Add(parseComparison p)
                expectP p TRParen
                Function(unionOperatorName, resizeArrayToList parts)
            else
                expectP p TRParen
                if opensFormula && peekP p = TEnd then p.Wrapped <- true
                e
        | TSheet sheetName ->
            p.Pos <- p.Pos + 1
            expectP p TBang
            parseSheetQualified p (Some sheetName)
        | TLBrace ->
            p.Pos <- p.Pos + 1
            parseArrayConstant p
        | TTableSpec raw ->
            // Unqualified structured ref (`[@Col]` / `[Amount]`) inside a table
            // cell — resolves against whichever table contains the formula.
            p.Pos <- p.Pos + 1
            TableRef(parseTableSpec "" raw)
        | TIdent ident ->
            p.Pos <- p.Pos + 1
            match peekP p with
            | TLParen ->
                // Function call.
                p.Pos <- p.Pos + 1
                let args = parseArgs p
                expectP p TRParen
                // Direct LAMBDA application (IIFE): LAMBDA(p1,…,pn, body)(a1,…,an).
                // Desugar to LET(p1,a1,…,pn,an, body) — same beta-reduction, no new
                // node. (LET inlines each value into the body, left to right.)
                let appliedDirectly =
                    equalsOrdinalIgnoreCase ident "LAMBDA"
                    && (peekP p = TLParen
                        || (peekP p = TSpace && p.Pos + 1 < p.Tokens.Length && p.Tokens.[p.Pos + 1] = TLParen))
                if appliedDirectly then
                    if peekP p = TSpace then p.Pos <- p.Pos + 1
                    p.Pos <- p.Pos + 1
                    let callArgs = parseArgs p
                    expectP p TRParen
                    let letArgs = ResizeArray<Expr>()
                    let rec appendLetArgs lambdaParts actualArgs =
                        match lambdaParts with
                        | [] -> false
                        | [body] ->
                            letArgs.Add body
                            true
                        | pn :: restParts ->
                            match actualArgs with
                            | arg :: restArgs ->
                                letArgs.Add pn
                                letArgs.Add arg
                                appendLetArgs restParts restArgs
                            | [] ->
                                appendLetArgs restParts []
                    if appendLetArgs args callArgs then
                        Function("LET", resizeArrayToList letArgs)
                    else
                        Function(ident, args)
                else Function(ident, args)
            | TTableSpec raw ->
                // Structured table reference: Table[Column] / Table[[#Totals],[A]] etc.
                p.Pos <- p.Pos + 1
                TableRef(parseTableSpec ident raw)
            | TBang ->
                // Unquoted sheet-qualified ref: Sheet1!A1.
                p.Pos <- p.Pos + 1
                parseSheetQualified p (Some ident)
            | _ ->
                // Cell ref, defined name, or boolean literal.
                if equalsOrdinalIgnoreCase ident "TRUE" then Bool true
                elif equalsOrdinalIgnoreCase ident "FALSE" then Bool false
                else
                    match parseRefToken ident with
                    | Some refAst ->
                        // `A1#` — spill-range operator on a cell ref.
                        if peekP p = THash then (p.Pos <- p.Pos + 1; SpillRef refAst)
                        else CellRef refAst
                    | None -> NameRef ident
        | other -> failwithf "Unexpected token %A" other
    finally
        leavePrimaryP p

/// Parse the reference portion that follows `Sheet!`. May itself be a range.
and private parseSheetQualified (p: Parser) (sheet: string option) : Expr =
    // The other end of `Sheet!A:` / `Sheet!1:`.
    let columnAfterColon () =
        match p.Tokens.[p.Pos], (if p.Pos + 1 < p.Tokens.Length then p.Tokens.[p.Pos + 1] else TEnd) with
        | TColon, TIdent other -> parseColumnToken other
        | _ -> None
    let rowAfterColon () =
        match p.Tokens.[p.Pos], (if p.Pos + 1 < p.Tokens.Length then p.Tokens.[p.Pos + 1] else TEnd) with
        | TColon, TNumber n -> rowOfNumber n |> Option.map (fun row -> row, false)
        | TColon, TIdent other -> parseAbsoluteRowToken other |> Option.map (fun row -> row, true)
        | _ -> None
    match nextP p with
    | TIdent ident ->
        match parseRefToken ident with
        | Some refAst ->
            let withSheet = { refAst with Sheet = sheet }
            // `Sheet!A1:B2` is one range of the sheet. Anything else after the
            // colon - `Sheet!A1:INDEX(Sheet!A:A,3)`, `Sheet!A1:Sheet!B2`, a
            // name - is the range operator's other operand (see parseRange).
            let plainEnd =
                match peekP p, (if p.Pos + 1 < p.Tokens.Length then p.Tokens.[p.Pos + 1] else TEnd), (if p.Pos + 2 < p.Tokens.Length then p.Tokens.[p.Pos + 2] else TEnd) with
                | TColon, TIdent endIdent, after -> after <> TLParen && after <> TBang && (parseRefToken endIdent).IsSome
                | _ -> false
            match peekP p with
            | TColon when plainEnd ->
                p.Pos <- p.Pos + 1
                match nextP p with
                | TIdent endIdent ->
                    match parseRefToken endIdent with
                    | Some endRef -> RangeRef(withSheet, { endRef with Sheet = sheet })
                    | None -> failwithf "Invalid cell ref after ':': %s" endIdent
                | other -> failwithf "Expected cell ref after ':', got %A" other
            | THash -> p.Pos <- p.Pos + 1; SpillRef withSheet
            | _ -> CellRef withSheet
        | None ->
            match parseColumnToken ident, columnAfterColon () with
            | Some first, Some last ->
                p.Pos <- p.Pos + 2
                wholeColumnsRef sheet first last
            | _ ->
                match parseAbsoluteRowToken ident, rowAfterColon () with
                | Some first, Some last ->
                    p.Pos <- p.Pos + 2
                    wholeRowsRef sheet (first, true) last
                | _ ->
                    // Sheet-qualified defined name (rare). Treat as NameRef.
                    NameRef ident
    | TNumber n ->
        match rowOfNumber n, rowAfterColon () with
        | Some first, Some last ->
            p.Pos <- p.Pos + 2
            wholeRowsRef sheet (first, false) last
        | _ -> failwith "Expected reference after sheet '!'"
    // `Sheet!#REF!`: what a deleted range of the sheet leaves; Excel's value is #REF!.
    | TError code when code = Err.ref -> ErrorLit Err.ref
    | other -> failwithf "Expected reference after sheet '!', got %A" other

/// The body of an array constant after its `{`: constants separated by `,`
/// (columns) and `;` (rows), e.g. `{1,2;3,4}`. Every row has the same width.
and private parseArrayConstant (p: Parser) : Expr =
    let rows = ResizeArray<CellValue[]>()
    let row = ResizeArray<CellValue>()
    let element () : CellValue =
        match nextP p with
        | TNumber n -> CellValue.Number n
        | TMinus ->
            (match nextP p with
             | TNumber n -> CellValue.Number (-n)
             | other -> failwithf "Expected a number in array constant, got %A" other)
        | TPlus ->
            (match nextP p with
             | TNumber n -> CellValue.Number n
             | other -> failwithf "Expected a number in array constant, got %A" other)
        | TString text -> CellValue.String text
        | TError code -> CellValue.Error code
        | TIdent ident when equalsOrdinalIgnoreCase ident "TRUE" -> CellValue.Boolean true
        | TIdent ident when equalsOrdinalIgnoreCase ident "FALSE" -> CellValue.Boolean false
        | other -> failwithf "Array constants hold only literal values, got %A" other
    let mutable finished = false
    while not finished do
        row.Add(element ())
        match nextP p with
        | TComma -> ()
        | TSemicolon ->
            rows.Add(row.ToArray())
            row.Clear()
        | TRBrace ->
            rows.Add(row.ToArray())
            finished <- true
        | other -> failwithf "Expected ',', ';' or '}' in array constant, got %A" other
    let width = rows.[0].Length
    for cells in rows do
        if cells.Length <> width then failwith "Array constant rows differ in length"
    ArrayLit (rows.ToArray())

and private parseArgs (p: Parser) : Expr list =
    if peekP p = TRParen then []
    else
        // An argument left empty between separators (`IF(A1,,2)`,
        // `OFFSET(A1,,1)`) is a MissingArg, not a syntax error.
        let argument () =
            match peekP p with
            | TComma | TRParen -> MissingArg
            | _ -> parseComparison p
        let rec loop acc =
            let e = argument ()
            match peekP p with
            | TComma -> p.Pos <- p.Pos + 1; loop (e :: acc)
            | _ -> List.rev (e :: acc)
        loop []

and private parseRange (p: Parser) : Expr =
    let next () = if p.Pos + 1 < p.Tokens.Length then p.Tokens.[p.Pos + 1] else TEnd
    // Whole-column (`A:C`, `$A:$A`) and whole-row (`1:3`, `$1:$1`) references
    // start with a token that is not a cell reference; recognise them by the
    // colon that follows and the matching token after it.
    let whole =
        match peekP p, next () with
        | TIdent first, TColon ->
            let after = if p.Pos + 2 < p.Tokens.Length then p.Tokens.[p.Pos + 2] else TEnd
            match parseColumnToken first, after with
            | Some c1, TIdent last ->
                match parseColumnToken last with
                | Some c2 when (parseRefToken first).IsNone -> Some (wholeColumnsRef None c1 c2)
                | _ -> None
            | _ ->
                match parseAbsoluteRowToken first, after with
                | Some r1, TNumber n -> rowOfNumber n |> Option.map (fun r2 -> wholeRowsRef None (r1, true) (r2, false))
                | Some r1, TIdent last -> parseAbsoluteRowToken last |> Option.map (fun r2 -> wholeRowsRef None (r1, true) (r2, true))
                | _ -> None
        | TNumber n, TColon ->
            let after = if p.Pos + 2 < p.Tokens.Length then p.Tokens.[p.Pos + 2] else TEnd
            match rowOfNumber n, after with
            | Some r1, TNumber m -> rowOfNumber m |> Option.map (fun r2 -> wholeRowsRef None (r1, false) (r2, false))
            | Some r1, TIdent last -> parseAbsoluteRowToken last |> Option.map (fun r2 -> wholeRowsRef None (r1, false) (r2, true))
            | _ -> None
        | _ -> None
    // `Sheet1:Sheet3!A1` — a 3-D reference. The sheet span travels in the
    // qualifier (a colon cannot occur in a sheet name) and is expanded to one
    // reference per sheet when the formula is evaluated.
    let threeD =
        match peekP p, next () with
        | TIdent first, TColon when p.Pos + 3 < p.Tokens.Length ->
            match p.Tokens.[p.Pos + 2], p.Tokens.[p.Pos + 3] with
            | TIdent last, TBang -> Some (first + ":" + last)
            | _ -> None
        | _ -> None
    let operand () =
        match threeD with
        | Some span ->
            p.Pos <- p.Pos + 4
            parseSheetQualified p (Some span)
        | None ->
        match whole with
        | Some range ->
            p.Pos <- p.Pos + 3
            range
        | None ->
            // The range operator joins any two references: `A1:B2`, but also
            // `A1:INDEX(A:A,3)` and `Start:End` between names.
            let isReference e =
                match e with
                | CellRef _ | RangeRef _ | NameRef _ | SpillRef _ | TableRef _ | Function _ -> true
                | _ -> false
            let mutable left = parsePrimary p
            while peekP p = TColon && isReference left do
                p.Pos <- p.Pos + 1
                let right = parsePrimary p
                left <-
                    match left, right with
                    | CellRef a, CellRef b -> RangeRef(a, b)
                    | _ when isReference right -> Function(rangeOperatorName, [ left; right ])
                    | _ -> failwith "Expected a reference after ':'"
            left
    let first = operand ()
    if peekP p <> TSpace then first
    else
        let mutable e = first
        while peekP p = TSpace do
            p.Pos <- p.Pos + 1
            e <- Function(intersectionOperatorName, [ e; parseRange p ])
        e

and private parseUnary (p: Parser) : Expr =
    match peekP p with
    | TMinus -> p.Pos <- p.Pos + 1; UnaryOp("-", parseUnary p)
    | TPlus -> p.Pos <- p.Pos + 1; UnaryOp("+", parseUnary p)
    | _ -> parsePercent p

and private parsePercent (p: Parser) : Expr =
    let mutable e = parseRange p
    while peekP p = TPercent do
        p.Pos <- p.Pos + 1
        e <- UnaryOp("%", e)
    e

and private parseExponent (p: Parser) : Expr =
    let mutable e = parseUnary p
    while peekP p = TCaret do
        p.Pos <- p.Pos + 1
        e <- BinOp(e, "^", parseUnary p)
    e

and private parseMulDiv (p: Parser) : Expr =
    let mutable e = parseExponent p
    let mutable go = true
    while go do
        match peekP p with
        | TStar -> p.Pos <- p.Pos + 1; e <- BinOp(e, "*", parseExponent p)
        | TSlash -> p.Pos <- p.Pos + 1; e <- BinOp(e, "/", parseExponent p)
        | _ -> go <- false
    e

and private parseAddSub (p: Parser) : Expr =
    let mutable e = parseMulDiv p
    let mutable go = true
    while go do
        match peekP p with
        | TPlus -> p.Pos <- p.Pos + 1; e <- BinOp(e, "+", parseMulDiv p)
        | TMinus -> p.Pos <- p.Pos + 1; e <- BinOp(e, "-", parseMulDiv p)
        | _ -> go <- false
    e

and private parseConcat (p: Parser) : Expr =
    let mutable e = parseAddSub p
    while peekP p = TAmp do
        p.Pos <- p.Pos + 1
        e <- BinOp(e, "&", parseAddSub p)
    e

and private parseComparison (p: Parser) : Expr =
    let mutable e = parseConcat p
    let mutable go = true
    while go do
        let op =
            match peekP p with
            | TEq -> Some "=" | TNe -> Some "<>"
            | TLt -> Some "<" | TGt -> Some ">"
            | TLe -> Some "<=" | TGe -> Some ">="
            | _ -> None
        match op with
        | Some o -> p.Pos <- p.Pos + 1; e <- BinOp(e, o, parseConcat p)
        | None -> go <- false
    e

/// Parse a formula string (leading `=` optional). Returns the AST or an error
/// message describing the parse failure.
let private parseFormulaText (input: string) : Result<Expr, string> =
    try
        // Legacy array formulas arrive wrapped in braces: `{=SUM(A1:A3*B1:B3)}`.
        let input =
            let trimmed = input.Trim()
            if trimmed.StartsWith "{=" && trimmed.EndsWith "}" then trimmed.Substring(1, trimmed.Length - 2)
            else input
        let start = if input.Length > 0 && input.[0] = '=' then 1 else 0
        let tokens = tokenizeFrom input start
        let p = { Tokens = tokens; Pos = 0; PrimaryDepth = 0; Wrapped = false }
        let e = parseComparison p
        match peekP p with
        | TEnd ->
            if p.Wrapped then formTableSet wrappedForms (box e) (box true)
            Ok e
        | t -> Result.Error (sprintf "Unexpected trailing token %A" t)
    with ex -> Result.Error ex.Message

/// The parsed formulas, by text. A recalculation asks for every formula of
/// the workbook, and so does building the dependency graph: parsing them
/// again each time was most of what both cost. (An AST is immutable and
/// depends on the text alone.)
let private parsedFormulas = Dictionary<string, Result<Expr, string>>()

/// Parse a formula string (leading `=` optional). Returns the AST or an error
/// message describing the parse failure.
let parseFormula (input: string) : Result<Expr, string> =
    if isNull input then parseFormulaText input
    else
        match Portable.Dict.tryFind input parsedFormulas with
        | ValueSome known -> known
        | ValueNone ->
            let parsed = parseFormulaText input
            if parsedFormulas.Count > 200000 then parsedFormulas.Clear()
            parsedFormulas.[input] <- parsed
            parsed

let private parseFormulaShared (input: string) : Result<Expr, string> =
    parseFormula input

/// Resolve a defined name to its final parsed expression, scoped to the
/// current sheet (a sheet-local name shadows a workbook-global one). A
/// name's refers-to text is itself a formula — usually a range, sometimes
/// a constant or another name; name→name hops are followed with a cycle
/// guard. Returns None when the name is unknown, won't parse, or cycles.
let private resolveNameExprUncached (name: string) (currentSheet: Rbt.Spreadsheet.Workbook.SheetId option)
                                    (workbook: Rbt.Spreadsheet.Workbook.Workbook) : Expr option =
    let rec resolve (n: string) (seen: Set<string>) (hops: int) =
        let key = n.ToUpperInvariant()
        if hops > 16 || Set.contains key seen then None
        else
            match Rbt.Spreadsheet.Workbook.getName currentSheet n workbook with
            | None -> None
            | Some refersTo ->
                let body = if refersTo.StartsWith "=" then refersTo else "=" + refersTo
                match parseFormula body with
                | Result.Error _ -> None
                | Ok (NameRef inner) -> resolve inner (Set.add key seen) (hops + 1)
                | Ok expr -> Some expr
    resolve name Set.empty 0

#if FABLE_COMPILER
// Asked for every use of a name by every formula when the dependency graph
// is built: what a name resolved to is kept for as long as the questions are
// about the same set of names.
let mutable private resolvedNamesOf : obj = null
let private resolvedNames = Dictionary<string, Expr option>()
let resolveNameExpr (name: string) (currentSheet: Rbt.Spreadsheet.Workbook.SheetId option)
                    (workbook: Rbt.Spreadsheet.Workbook.Workbook) : Expr option =
    if not (Rbt.Spreadsheet.Workbook.weakRefIs resolvedNamesOf (box workbook.Names)) then
        resolvedNamesOf <- Rbt.Spreadsheet.Workbook.weakRefOf (box workbook.Names)
        resolvedNames.Clear()
    let key = (match currentSheet with Some sheet -> sheet | None -> "") + "|" + name
    match Portable.Dict.tryFind key resolvedNames with
    | ValueSome known -> known
    | ValueNone ->
        let resolved = resolveNameExprUncached name currentSheet workbook
        resolvedNames.[key] <- resolved
        resolved
#else
let resolveNameExpr (name: string) (currentSheet: Rbt.Spreadsheet.Workbook.SheetId option)
                    (workbook: Rbt.Spreadsheet.Workbook.Workbook) : Expr option =
    resolveNameExprUncached name currentSheet workbook
#endif

// ---------------------------------------------------------------------------
// Value helpers + coercion
// ---------------------------------------------------------------------------

/// Unwrap a cached formula value to its underlying scalar.
let unwrapValue (value: CellValue) : CellValue =
    unwrapFormulaValue value

let private err code : CellValue = CellValue.Error code

/// First error among the operands, if any — for error propagation.
let private firstError (vs: CellValue list) : CellValue option =
    let mutable found = None
    let mutable remaining = vs
    while found.IsNone && not remaining.IsEmpty do
        match remaining with
        | value :: tail ->
            match unwrapValue value with
            | CellValue.Error _ as e -> found <- Some e
            | _ -> ()
            remaining <- tail
        | [] -> ()
    found

/// Format a number the way Excel coerces it to text for `&` / TEXT-less
/// contexts: integers without a decimal, otherwise the shortest round-trip.
let numberToText (n: float) : string =
    if Double.IsNaN n || Double.IsInfinity n then "#NUM!"
    elif n = Math.Floor n && abs n < 1e15 then
        (int64 n).ToString(CultureInfo.InvariantCulture)
    else
        NumberText.numberToString "G15" n

/// Coerce a value to a number using Excel's rules. Blank → 0, TRUE → 1,
/// FALSE → 0, numeric text → its value, non-numeric text → #VALUE!,
/// error → that error.
let coerceNumber (value: CellValue) : Result<float, string> =
    match unwrapValue value with
    | CellValue.Number n -> Ok n
    | CellValue.Boolean b -> Ok (if b then 1.0 else 0.0)
    | CellValue.Empty -> Ok 0.0
    | CellValue.Error e -> Result.Error e
    | CellValue.String s ->
        let t = s.Trim()
        match Double.TryParse(NumberText.cutToSignificantDigits t, NumberStyles.Float, CultureInfo.InvariantCulture) with
        | true, n -> Ok n
        | _ ->
            // Text a number would be typed as (a date "3/1/2026" or "March 2026", a time, "50%",
            // "$1,000") is that number, as typing it would enter it: ="March 2026"+0, TEXT("3/1/2026", ...).
            // Text TRUE / FALSE is not: only the boolean *value* coerces (#VALUE!).
            match (Rbt.Spreadsheet.Input.parseInput t).Value with
            | CellValue.Number n when t <> "" -> Ok n
            | _ -> Result.Error Err.value
    | CellValue.Formula _ -> Result.Error Err.value  // unreachable after unwrap

let inline private tryAggregateNumber (value: CellValue) : float voption =
    match unwrapValue value with
    | CellValue.Number n -> ValueSome n
    | CellValue.Boolean b -> ValueSome (if b then 1.0 else 0.0)
    | CellValue.Empty -> ValueSome 0.0
    | CellValue.String s ->
        let struct (start, length) = trimBounds s
        match Double.TryParse(NumberText.cutToSignificantDigits (trimmedNumberText s start length), NumberStyles.Float, CultureInfo.InvariantCulture) with
        | true, n -> ValueSome n
        | _ -> ValueNone
    | _ -> ValueNone

/// Coerce a value to text using Excel's rules.
let coerceText (value: CellValue) : Result<string, string> =
    match unwrapValue value with
    | CellValue.String s -> Ok s
    | CellValue.Number n -> Ok (numberToText n)
    | CellValue.Boolean b -> Ok (if b then "TRUE" else "FALSE")
    | CellValue.Empty -> Ok ""
    | CellValue.Error e -> Result.Error e
    | CellValue.Formula _ -> Result.Error Err.value

/// Coerce a value to a boolean for IF-style truthiness. Blank → false,
/// 0 → false, non-zero → true, "" → false, "TRUE"/"FALSE" text parsed.
let coerceBool (value: CellValue) : Result<bool, string> =
    match unwrapValue value with
    | CellValue.Boolean b -> Ok b
    | CellValue.Number n -> Ok (n <> 0.0)
    | CellValue.Empty -> Ok false
    | CellValue.Error e -> Result.Error e
    | CellValue.String s ->
        let struct (start, length) = trimBounds s
        if length = 0 then Ok false
        elif sliceEqualsOrdinalIgnoreCase s start length "TRUE" then Ok true
        elif sliceEqualsOrdinalIgnoreCase s start length "FALSE" then Ok false
        else
            match Double.TryParse(NumberText.cutToSignificantDigits (trimmedNumberText s start length), NumberStyles.Float, CultureInfo.InvariantCulture) with
            | true, n -> Ok (n <> 0.0)
            | _ -> Result.Error Err.value
    | CellValue.Formula _ -> Result.Error Err.value

// ---------------------------------------------------------------------------
// Reference resolution + evaluation
// ---------------------------------------------------------------------------

let private normalizeRange (a: RefAst) (b: RefAst) =
    (min a.Row b.Row, min a.Col b.Col, max a.Row b.Row, max a.Col b.Col)

/// The sheet and area `(r0, c0, r1, c1)` a structured reference reads, for
/// dependency tracking. A this-row reference without a known cell is widened
/// to the whole data column. None when the table or column cannot be resolved
/// or the area is empty.
let private tableRefArea (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                         (sheet: Rbt.Spreadsheet.Workbook.SheetId option)
                         (cell: CellAddress option)
                         (sel: TableSel) : (Rbt.Spreadsheet.Workbook.SheetId * (int * int * int * int)) option =
    match workbook with
    | None -> None
    | Some wb ->
        let table =
            if sel.Table <> "" then Rbt.Spreadsheet.Workbook.getTable sel.Table wb
            else cell |> Option.bind (fun c -> Rbt.Spreadsheet.Workbook.findTableContaining sheet c wb)
        match table with
        | None -> None
        | Some t ->
            let dataStart = t.StartRow + (if t.HasHeaders then 1 else 0)
            let dataEnd = t.EndRow - (if t.HasTotals then 1 else 0)
            let columnIndex name = Rbt.Spreadsheet.Workbook.tableColumnIndex t name
            let columns =
                match sel.ColStart, sel.ColEnd with
                | None, _ -> Some (t.StartCol, t.StartCol + t.Columns.Length - 1)
                | Some first, None -> columnIndex first |> Option.map (fun i -> t.StartCol + i, t.StartCol + i)
                | Some first, Some last ->
                    match columnIndex first, columnIndex last with
                    | Some i, Some j -> Some (t.StartCol + min i j, t.StartCol + max i j)
                    | _ -> None
            let rows =
                match sel.Item with
                | TblData -> Some (dataStart, dataEnd)
                | TblAll -> Some (t.StartRow, t.EndRow)
                | TblHeaders -> if t.HasHeaders then Some (t.StartRow, t.StartRow) else None
                | TblTotals -> if t.HasTotals then Some (t.EndRow, t.EndRow) else None
                | TblThisRow ->
                    match cell with
                    | Some c when c.Row >= dataStart && c.Row <= dataEnd -> Some (c.Row, c.Row)
                    | Some _ -> None
                    | None -> Some (dataStart, dataEnd)
            match columns, rows with
            | Some (c0, c1), Some (r0, r1) when r0 <= r1 -> Some (t.SheetId, (r0, c0, r1, c1))
            | _ -> None

/// Bounds of a range for dependency tracking: a whole-column / whole-row
/// reference is cut down to the rows / columns its sheet actually uses (when
/// the sheet is known), so it never expands to a million cells.
let private dependencyBounds (sheetModel: SpreadsheetModel option) (a: RefAst) (b: RefAst) : int * int * int * int =
    let sr, sc, er, ec = normalizeRange a b
    let wholeColumns = sr = 0 && er = 1048575
    let wholeRows = sc = 0 && ec = 16383
    if not wholeColumns && not wholeRows then sr, sc, er, ec
    else
        let mutable lastRow = 0
        let mutable lastCol = 0
        match sheetModel with
        | Some m ->
            for KeyValue(addr, _) in m.Cells do
                if addr.Row > lastRow then lastRow <- addr.Row
                if addr.Col > lastCol then lastCol <- addr.Col
        | None -> ()
        sr, sc, (if wholeColumns then max sr lastRow else er), (if wholeRows then max sc lastCol else ec)

/// Set of cell addresses an expression references *within the current
/// model*. Cross-sheet refs (`r.Sheet = Some _`) are deliberately excluded
/// — single-model recalc ordering can't span sheets, so they're treated as
/// external constants. A `NameRef` is resolved through the workbook (when
/// one is supplied) and its target cells contribute to the set, so
/// `=SUM(MyRange)` registers a dependency on the range's members. Defined
/// names resolve at workbook-global scope here: a missed sheet-local name
/// only affects recalc *ordering*, never the computed values, since `eval`
/// resolves names with the correct scope. Used by the dependency tracker.
let referencesInExpr (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                     (expr: Expr) : Set<CellAddress> =
    let refs = HashSet<CellAddress>()
    let stack = ResizeArray<struct (Set<string> * Expr)>()
    stack.Add(struct (Set.empty, expr))
    while stack.Count > 0 do
        let last = stack.Count - 1
        let struct (seen, e) = stack.[last]
        Portable.Stack.dropLast stack
        match e with
        | Number _ | Str _ | Bool _ | ErrorLit _ | ArrayLit _ | MissingArg -> ()
        | TableRef sel ->
            match tableRefArea workbook None None sel with
            | Some (_, (r0, c0, r1, c1)) ->
                for row in r0 .. r1 do
                    for col in c0 .. c1 do
                        refs.Add { Row = row; Col = col } |> ignore
            | None -> ()
        | NameRef name ->
            let key = name.ToUpperInvariant()
            if not (Set.contains key seen) then
                match workbook |> Option.bind (resolveNameExpr name None) with
                | Some resolved -> stack.Add(struct (Set.add key seen, resolved))
                | None -> ()
        | CellRef r when r.Sheet.IsSome -> ()
        | CellRef r -> refs.Add { Row = r.Row; Col = r.Col } |> ignore
        | SpillRef r -> refs.Add { Row = r.Row; Col = r.Col } |> ignore
        | RangeRef (a, _) when a.Sheet.IsSome -> ()
        | RangeRef (a, b) ->
            let activeModel =
                workbook |> Option.bind (fun wb -> Map.tryFind wb.ActiveSheet wb.Sheets) |> Option.map (fun sheet -> sheet.Model)
            let sr, sc, er, ec = dependencyBounds activeModel a b
            for row in sr .. er do
                for col in sc .. ec do
                    refs.Add { Row = row; Col = col } |> ignore
        | UnaryOp (_, inner) -> stack.Add(struct (seen, inner))
        | BinOp (l, _, r) ->
            stack.Add(struct (seen, l))
            stack.Add(struct (seen, r))
        | Function (_, args) ->
            for arg in args do
                stack.Add(struct (seen, arg))
    let mutable result = Set.empty
    for cell in refs do
        result <- Set.add cell result
    result

/// Cross-sheet-aware reference extraction. Every reference is qualified to
/// an absolute (sheet, cell) pair: a ref's own sheet name when it has one,
/// otherwise `homeSheet` (the sheet the formula lives on). `NameRef`s
/// resolve through the workbook with a cycle guard. A ref whose sheet name
/// can't be resolved is dropped. This is what the incremental calc engine
/// keys its dependency graph on, so dependencies — and dirty propagation —
/// span sheet boundaries.
let referencesQualified (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                        (homeSheet: Rbt.Spreadsheet.Workbook.SheetId)
                        (expr: Expr) : Set<Rbt.Spreadsheet.Workbook.SheetCell> =
    let sheetIdOf (name: string option) =
        match name with
        | None -> Some homeSheet
        | Some n -> workbook |> Option.bind (Rbt.Spreadsheet.Workbook.resolveSheetName n)
    let modelOf (sid: Rbt.Spreadsheet.Workbook.SheetId) =
        workbook |> Option.bind (fun wb -> Map.tryFind sid wb.Sheets) |> Option.map (fun sheet -> resolveSheet sheet.Model)
    let refs = HashSet<Rbt.Spreadsheet.Workbook.SheetCell>()
    let stack = ResizeArray<struct (Set<string> * Expr)>()
    stack.Add(struct (Set.empty, expr))
    while stack.Count > 0 do
        let last = stack.Count - 1
        let struct (seen, e) = stack.[last]
        Portable.Stack.dropLast stack
        match e with
        | Number _ | Str _ | Bool _ | ErrorLit _ | ArrayLit _ | MissingArg -> ()
        | TableRef sel ->
            match tableRefArea workbook (Some homeSheet) None sel with
            | Some (sid, (r0, c0, r1, c1)) ->
                for row in r0 .. r1 do
                    for col in c0 .. c1 do
                        refs.Add { Sheet = sid; Cell = { Row = row; Col = col } } |> ignore
            | None -> ()
        | NameRef name ->
            let key = name.ToUpperInvariant()
            if not (Set.contains key seen) then
                match workbook |> Option.bind (resolveNameExpr name (Some homeSheet)) with
                | Some resolved -> stack.Add(struct (Set.add key seen, resolved))
                | None -> ()
        | CellRef r | SpillRef r ->
            match sheetIdOf r.Sheet with
            | Some sid -> refs.Add { Sheet = sid; Cell = { Row = r.Row; Col = r.Col } } |> ignore
            | None -> ()
        | RangeRef (a, b) ->
            match sheetIdOf a.Sheet with
            | Some sid ->
                let sr, sc, er, ec = dependencyBounds (modelOf sid) a b
                for row in sr .. er do
                    for col in sc .. ec do
                        refs.Add { Sheet = sid; Cell = { Row = row; Col = col } } |> ignore
            | None -> ()
        | UnaryOp (_, inner) -> stack.Add(struct (seen, inner))
        | BinOp (l, _, r) ->
            stack.Add(struct (seen, l))
            stack.Add(struct (seen, r))
        | Function (_, args) ->
            for arg in args do
                stack.Add(struct (seen, arg))
    let mutable result = Set.empty
    for cell in refs do
        result <- Set.add cell result
    result

/// A rectangular cell region on a sheet (inclusive bounds), for range-granular
/// dependencies — a *large* range reference is tracked as one rectangle rather
/// than one graph edge per cell (a single `SUM(A1:A100000)` would otherwise
/// reify 100k edges + ~29 MB). Membership is tested by containment on edit.
type DepRect =
    { Sheet: Rbt.Spreadsheet.Workbook.SheetId; R0: int; C0: int; R1: int; C1: int }

let rectContainsCell (rect: DepRect) (sc: Rbt.Spreadsheet.Workbook.SheetCell) =
    sc.Sheet = rect.Sheet
    && sc.Cell.Row >= rect.R0 && sc.Cell.Row <= rect.R1
    && sc.Cell.Col >= rect.C0 && sc.Cell.Col <= rect.C1

/// Ranges with more than this many cells become a rectangle dependency rather
/// than per-cell edges. Below it, the explicit-edge path is kept (cheaper for
/// the common small range, and preserves the per-cell dependent sets the calc
/// engine + its tests rely on).
[<Literal>]
let rangeDepThreshold = 2048

/// Like `referencesQualified`, but a range larger than `rangeDepThreshold` is
/// returned as a `DepRect` (second component) instead of being expanded into
/// the cell set. Small ranges and single cells still land in the cell set, so
/// existing dependency behavior is unchanged for them.
let rec referencesRanged (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                         (homeSheet: Rbt.Spreadsheet.Workbook.SheetId)
                         (expr: Expr) : Set<Rbt.Spreadsheet.Workbook.SheetCell> * DepRect list =
    referencesRangedAt workbook homeSheet None expr

/// `referencesRanged` for a formula whose own cell is known, so this-row and
/// unqualified structured references (`[@Qty]`) resolve to the exact cells.
and referencesRangedAt (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                       (homeSheet: Rbt.Spreadsheet.Workbook.SheetId)
                       (cell: CellAddress option)
                       (expr: Expr) : Set<Rbt.Spreadsheet.Workbook.SheetCell> * DepRect list =
    let sheetIdOf (name: string option) =
        match name with
        | None -> Some homeSheet
        | Some n -> workbook |> Option.bind (Rbt.Spreadsheet.Workbook.resolveSheetName n)
    let cells = System.Collections.Generic.HashSet<Rbt.Spreadsheet.Workbook.SheetCell>()
    let rects = ResizeArray<DepRect>()
    let stack = ResizeArray<struct (Set<string> * Expr)>()
    let push seen expr = stack.Add(struct (seen, expr))
    push Set.empty expr
    while stack.Count > 0 do
        let last = stack.Count - 1
        let struct (seen, e) = stack.[last]
        Portable.Stack.dropLast stack
        let addArea sid sr sc er ec =
            if int64 (er - sr + 1) * int64 (ec - sc + 1) > int64 rangeDepThreshold then
                rects.Add { Sheet = sid; R0 = sr; C0 = sc; R1 = er; C1 = ec }
            else
                for row in sr .. er do
                    for col in sc .. ec do
                        cells.Add { Sheet = sid; Cell = { Row = row; Col = col } } |> ignore
        match e with
        | Number _ | Str _ | Bool _ | ErrorLit _ | ArrayLit _ | MissingArg -> ()
        | TableRef sel ->
            match tableRefArea workbook (Some homeSheet) cell sel with
            | Some (sid, (r0, c0, r1, c1)) -> addArea sid r0 c0 r1 c1
            | None -> ()
        | NameRef name ->
            let key = name.ToUpperInvariant()
            if not (Set.contains key seen) then
                match workbook |> Option.bind (resolveNameExpr name (Some homeSheet)) with
                | Some resolved -> push (Set.add key seen) resolved
                | None -> ()
        | CellRef r | SpillRef r ->
            match sheetIdOf r.Sheet with
            | Some sid -> cells.Add { Sheet = sid; Cell = { Row = r.Row; Col = r.Col } } |> ignore
            | None -> ()
        | RangeRef (a, b) ->
            match sheetIdOf a.Sheet with
            | Some sid ->
                let sr, sc, er, ec = normalizeRange a b
                addArea sid sr sc er ec
            | None -> ()
        | UnaryOp (_, inner) -> push seen inner
        | BinOp (l, _, r) ->
            push seen l
            push seen r
        | Function (name, args) ->
            if name.Equals("GETPIVOTDATA", StringComparison.OrdinalIgnoreCase) then
                match List.tryItem 1 args with
                | Some (CellRef anchor) | Some (RangeRef (anchor, _)) ->
                    match sheetIdOf anchor.Sheet |> Option.bind (fun sid ->
                        workbook |> Option.bind (fun wb -> Map.tryFind sid wb.Sheets)
                        |> Option.bind (fun sheet ->
                            sheet.Model.Pivots
                            |> List.tryFind (fun pivot -> anchor.Row >= pivot.Top && anchor.Row <= pivot.Bottom && anchor.Col >= pivot.Left && anchor.Col <= pivot.Right)
                            |> Option.map (fun pivot -> sid, pivot))) with
                    | Some (sid, pivot) -> addArea sid pivot.Top pivot.Left pivot.Bottom pivot.Right
                    | None -> ()
                | _ -> ()
            for arg in args do push seen arg
    let mutable cellSet = Set.empty
    for cell in cells do
        cellSet <- Set.add cell cellSet
    (cellSet, resizeArrayToList rects)

/// The sheets of a 3-D reference's qualifier (`First:Last`): every sheet from
/// the first to the last named, in tab order as it is now. Empty when one of
/// the two does not exist.
let threeDSheetIds (wb: Rbt.Spreadsheet.Workbook.Workbook) (qualifier: string) : Rbt.Spreadsheet.Workbook.SheetId list =
    let cut = qualifier.IndexOf ':'
    if cut < 0 then []
    else
        let position (sheetName: string) =
            Rbt.Spreadsheet.Workbook.resolveSheetName sheetName wb
            |> Option.bind (fun id -> wb.SheetOrder |> List.tryFindIndex (fun other -> other = id))
        match position (qualifier.Substring(0, cut)), position (qualifier.Substring(cut + 1)) with
        | Some i, Some j -> wb.SheetOrder |> List.skip (min i j) |> List.truncate (abs (i - j) + 1)
        | _ -> []

/// The areas a formula reads, one call per reference: the sheet and the
/// inclusive bounds (a single cell is a 1 x 1 area). The references are those
/// of `referencesRangedAt` - defined names and structured references resolved
/// through the workbook, a reference to a sheet that does not exist left out -
/// without building a set of cells.
let iterReferenceAreas (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                       (homeSheet: Rbt.Spreadsheet.Workbook.SheetId)
                       (cell: CellAddress option)
                       (expr: Expr)
                       (onArea: Rbt.Spreadsheet.Workbook.SheetId -> int -> int -> int -> int -> unit) : unit =
    // The sheets a reference is on: its own, the one it names, or - a 3-D
    // reference (`Sheet1:Sheet3!A1`) - every sheet from the first to the last
    // it names, in tab order.
    // (Spelled out for one area: this runs for every reference of every
    // formula when the dependency graph is built, and a closure per
    // reference shows there.)
    let onSheets (name: string option) (r0: int) (c0: int) (r1: int) (c1: int) =
        match name with
        | None -> onArea homeSheet r0 c0 r1 c1
        | Some n ->
            match workbook with
            | None -> ()
            | Some wb ->
                if n.IndexOf ':' >= 0 then
                    for sid in threeDSheetIds wb n do onArea sid r0 c0 r1 c1
                else
                    match Rbt.Spreadsheet.Workbook.resolveSheetName n wb with
                    | Some sid -> onArea sid r0 c0 r1 c1
                    | None -> ()
    let stack = ResizeArray<struct (Set<string> * Expr)>()
    stack.Add(struct (Set.empty, expr))
    while stack.Count > 0 do
        let last = stack.Count - 1
        let struct (seen, e) = stack.[last]
        Portable.Stack.dropLast stack
        match e with
        | Number _ | Str _ | Bool _ | ErrorLit _ | ArrayLit _ | MissingArg -> ()
        | TableRef sel ->
            match tableRefArea workbook (Some homeSheet) cell sel with
            | Some (sid, (r0, c0, r1, c1)) -> onArea sid r0 c0 r1 c1
            | None -> ()
        | NameRef name ->
            let key = name.ToUpperInvariant()
            if not (Set.contains key seen) then
                match workbook |> Option.bind (resolveNameExpr name (Some homeSheet)) with
                | Some resolved -> stack.Add(struct (Set.add key seen, resolved))
                | None -> ()
        | CellRef r | SpillRef r -> onSheets r.Sheet r.Row r.Col r.Row r.Col
        | RangeRef (a, b) ->
            let sr, sc, er, ec = normalizeRange a b
            onSheets a.Sheet sr sc er ec
        | UnaryOp (_, inner) -> stack.Add(struct (seen, inner))
        | BinOp (l, _, r) ->
            stack.Add(struct (seen, l))
            stack.Add(struct (seen, r))
        | Function (name, args) ->
            if name.Equals("GETPIVOTDATA", StringComparison.OrdinalIgnoreCase) then
                match List.tryItem 1 args with
                | Some (CellRef anchor) | Some (RangeRef (anchor, _)) ->
                    let pivotOn sid =
                        workbook |> Option.bind (fun wb -> Map.tryFind sid wb.Sheets)
                        |> Option.bind (fun sheet ->
                            sheet.Model.Pivots
                            |> List.tryFind (fun pivot -> anchor.Row >= pivot.Top && anchor.Row <= pivot.Bottom && anchor.Col >= pivot.Left && anchor.Col <= pivot.Right))
                        |> Option.iter (fun pivot -> onArea sid pivot.Top pivot.Left pivot.Bottom pivot.Right)
                    match anchor.Sheet with
                    | None -> pivotOn homeSheet
                    | Some qualifier -> workbook |> Option.bind (Rbt.Spreadsheet.Workbook.resolveSheetName qualifier) |> Option.iter pivotOn
                | _ -> ()
            for arg in args do stack.Add(struct (seen, arg))

/// Shift every *relative* reference in an expression by (dr, dc). Absolute
/// components (`$`-locked row/col) are left untouched. This is how a
/// formula written for one anchor cell is re-pointed for another — used by
/// fill/copy and by conditional-format `expression` rules, which store one
/// formula anchored at the range's top-left.
let rec shiftExpr (dr: int) (dc: int) (expr: Expr) : Expr =
    let shiftRef (r: RefAst) =
        { r with
            Row = if r.RowAbs then r.Row else r.Row + dr
            Col = if r.ColAbs then r.Col else r.Col + dc }
    let shiftBinOpChain expr =
        let operations = ResizeArray<struct (string * Expr)>()
        let mutable current = expr
        let mutable collecting = true
        while collecting do
            match current with
            | BinOp (left, op, right) ->
                operations.Add(struct (op, right))
                current <- left
            | _ ->
                collecting <- false
        let mutable result = shiftExpr dr dc current
        for index = operations.Count - 1 downto 0 do
            let struct (op, right) = operations.[index]
            result <- BinOp(result, op, shiftExpr dr dc right)
        result
    match expr with
    | Number _ | Str _ | Bool _ | ErrorLit _ | NameRef _ | TableRef _ | ArrayLit _ | MissingArg -> expr
    | CellRef r -> CellRef (shiftRef r)
    | SpillRef r -> SpillRef (shiftRef r)
    | RangeRef (a, b) -> RangeRef (shiftRef a, shiftRef b)
    | UnaryOp (op, e) -> UnaryOp (op, shiftExpr dr dc e)
    | BinOp _ -> shiftBinOpChain expr
    | Function (name, args) ->
        let shifted = ResizeArray<Expr>()
        for arg in args do
            shifted.Add(shiftExpr dr dc arg)
        Function (name, resizeArrayToList shifted)

// ---------------------------------------------------------------------------
// Expr → formula text. The inverse of `parseFormula`: the printed text parses
// back to the same tree, so a formula can be rewritten structurally (shifted
// for copy/fill, re-pointed after a row/column edit) and stored again.
// ---------------------------------------------------------------------------

let isWholeColumnRange (a: RefAst) (b: RefAst) =
    min a.Row b.Row = 0 && max a.Row b.Row = lastSheetRow

let isWholeRowRange (a: RefAst) (b: RefAst) =
    min a.Col b.Col = 0 && max a.Col b.Col = lastSheetCol

let private sheetNameNeedsQuotes (name: string) =
    if name.Length = 0 then true
    elif not (Char.IsLetter name.[0] || name.[0] = '_') then true
    elif equalsOrdinalIgnoreCase name "TRUE" || equalsOrdinalIgnoreCase name "FALSE" then true
    elif (match parseRefToken name with
          | Some r -> r.Row <= lastSheetRow && r.Col <= lastSheetCol
          | None -> false) then true
    else
        let mutable plain = true
        for ch in name do
            if not (Char.IsLetterOrDigit ch || ch = '_' || ch = '.') then plain <- false
        not plain

let private sheetPrefixText (sheet: string option) =
    match sheet with
    | None -> ""
    | Some name when sheetNameNeedsQuotes name -> "'" + name.Replace("'", "''") + "'!"
    | Some name -> name + "!"

let private refOnSheet (r: RefAst) =
    r.Row >= 0 && r.Row <= lastSheetRow && r.Col >= 0 && r.Col <= lastSheetCol

let private refCellText (r: RefAst) =
    (if r.ColAbs then "$" else "") + columnToLetter r.Col
    + (if r.RowAbs then "$" else "") + string (r.Row + 1)

let private numberLiteralText (n: float) =
    let magnitude = abs n
    let text =
        if magnitude = Math.Floor magnitude && magnitude < 1e15 then
            (int64 magnitude).ToString(CultureInfo.InvariantCulture)
        else
            magnitude.ToString("R", CultureInfo.InvariantCulture)
    if n < 0.0 then "-" + text else text

let private stringLiteralText (s: string) = "\"" + s.Replace("\"", "\"\"") + "\""

let private arrayLiteralText (rows: CellValue[][]) =
    let sb = Text.StringBuilder()
    sb.Append('{') |> ignore
    for r = 0 to rows.Length - 1 do
        if r > 0 then sb.Append(';') |> ignore
        for c = 0 to rows.[r].Length - 1 do
            if c > 0 then sb.Append(',') |> ignore
            match unwrapValue rows.[r].[c] with
            | CellValue.Number n -> sb.Append(numberLiteralText n) |> ignore
            | CellValue.String s -> sb.Append(stringLiteralText s) |> ignore
            | CellValue.Boolean b -> sb.Append(if b then "TRUE" else "FALSE") |> ignore
            | CellValue.Error e -> sb.Append('#').Append(e) |> ignore
            | _ -> ()
    sb.Append('}') |> ignore
    sb.ToString()

let private tableSelText (sel: TableSel) =
    let plainColumn (name: string) =
        let mutable plain = name.Length > 0
        for ch in name do
            if not (Char.IsLetterOrDigit ch || ch = '_' || ch = ' ' || ch = '.') then plain <- false
        plain && name = name.Trim()
    let columns =
        match sel.ColStart, sel.ColEnd with
        | Some first, Some last -> Some ("[" + first + "]:[" + last + "]")
        | Some only, None -> Some ("[" + only + "]")
        | None, _ -> None
    let itemText =
        match sel.Item with
        | TblAll -> "#All"
        | TblHeaders -> "#Headers"
        | TblTotals -> "#Totals"
        | TblData -> "#Data"
        | TblThisRow -> "@"
    let inner =
        match sel.Item, columns with
        | TblThisRow, None -> "@"
        | TblThisRow, Some cols ->
            match sel.ColStart, sel.ColEnd with
            | Some only, None when plainColumn only && not (only.Contains " ") -> "@" + only
            | _ -> "@" + cols
        | TblData, Some cols ->
            match sel.ColStart, sel.ColEnd with
            | Some only, None when plainColumn only -> only
            | _ -> cols
        | _, None -> itemText
        | _, Some cols -> "[" + itemText + "]," + cols
    sel.Table + "[" + inner + "]"

let private binaryPrecedence (op: string) =
    match op with
    | "=" | "<>" | "<" | ">" | "<=" | ">=" -> 1
    | "&" -> 2
    | "+" | "-" -> 3
    | "*" | "/" -> 4
    | _ -> 5   // ^

let rec private appendExpr (sb: Text.StringBuilder) (minPrecedence: int) (expr: Expr) : unit =
    match expr with
    | Number n ->
        // A negative literal reads back as unary minus, so it binds like one.
        if n < 0.0 && minPrecedence > 6 then sb.Append('(').Append(numberLiteralText n).Append(')') |> ignore
        else sb.Append(numberLiteralText n) |> ignore
    | Str s -> sb.Append(stringLiteralText s) |> ignore
    | Bool b -> sb.Append(if b then "TRUE" else "FALSE") |> ignore
    | ErrorLit e -> sb.Append('#').Append(e) |> ignore
    | MissingArg -> ()
    | NameRef name -> sb.Append(name) |> ignore
    | CellRef r ->
        if refOnSheet r then sb.Append(sheetPrefixText r.Sheet).Append(refCellText r) |> ignore
        else sb.Append("#REF!") |> ignore
    | SpillRef r ->
        if refOnSheet r then sb.Append(sheetPrefixText r.Sheet).Append(refCellText r).Append('#') |> ignore
        else sb.Append("#REF!") |> ignore
    | RangeRef (a, b) ->
        if not (refOnSheet a && refOnSheet b) then sb.Append("#REF!") |> ignore
        else
            sb.Append(sheetPrefixText a.Sheet) |> ignore
            if isWholeColumnRange a b then
                let first, last = (if a.Col <= b.Col then a, b else b, a)
                sb.Append(if first.ColAbs then "$" else "").Append(columnToLetter first.Col).Append(':')
                  .Append(if last.ColAbs then "$" else "").Append(columnToLetter last.Col) |> ignore
            elif isWholeRowRange a b then
                let first, last = (if a.Row <= b.Row then a, b else b, a)
                sb.Append(if first.RowAbs then "$" else "").Append(first.Row + 1).Append(':')
                  .Append(if last.RowAbs then "$" else "").Append(last.Row + 1) |> ignore
            else
                sb.Append(refCellText a).Append(':').Append(refCellText b) |> ignore
    | TableRef sel -> sb.Append(tableSelText sel) |> ignore
    | ArrayLit rows -> sb.Append(arrayLiteralText rows) |> ignore
    | Function (name, [ left; right ]) when name = intersectionOperatorName || name = rangeOperatorName ->
        appendExpr sb 8 left
        sb.Append(name) |> ignore
        appendExpr sb 8 right
    | Function (name, parts) when name = unionOperatorName ->
        sb.Append('(') |> ignore
        parts |> List.iteri (fun i part ->
            if i > 0 then sb.Append(',') |> ignore
            appendExpr sb 1 part)
        sb.Append(')') |> ignore
    | Function (name, [ operand ]) when name = implicitIntersectionName ->
        sb.Append('@') |> ignore
        appendExpr sb 8 operand
    | Function (name, args) ->
        sb.Append(name).Append('(') |> ignore
        let mutable first = true
        for arg in args do
            if not first then sb.Append(',') |> ignore
            first <- false
            appendExpr sb 1 arg
        sb.Append(')') |> ignore
    | UnaryOp ("%", operand) ->
        appendExpr sb 7 operand
        sb.Append('%') |> ignore
    | UnaryOp (op, operand) ->
        let wrap = minPrecedence > 6
        if wrap then sb.Append('(') |> ignore
        sb.Append(op) |> ignore
        appendExpr sb 6 operand
        if wrap then sb.Append(')') |> ignore
    | BinOp _ ->
        // Walk the left spine iteratively: `a+b+c+…` nests thousands deep on
        // the left and must not recurse once per operator.
        let levels = ResizeArray<struct (string * Expr)>()
        let mutable current = expr
        let mutable collecting = true
        while collecting do
            match current with
            | BinOp (left, op, right) ->
                levels.Add(struct (op, right))
                current <- left
            | _ -> collecting <- false
        let wrapped = Array.zeroCreate<bool> levels.Count
        for i = 0 to levels.Count - 1 do
            let struct (op, _) = levels.[i]
            let required =
                if i = 0 then minPrecedence
                else (let struct (outer, _) = levels.[i - 1] in binaryPrecedence outer)
            wrapped.[i] <- binaryPrecedence op < required
            if wrapped.[i] then sb.Append('(') |> ignore
        let struct (innermostOp, _) = levels.[levels.Count - 1]
        appendExpr sb (binaryPrecedence innermostOp) current
        for i = levels.Count - 1 downto 0 do
            let struct (op, right) = levels.[i]
            sb.Append(op) |> ignore
            appendExpr sb (binaryPrecedence op + 1) right
            if wrapped.[i] then sb.Append(')') |> ignore

/// Print an expression as formula text (no leading `=`).
let exprToFormula (expr: Expr) : string =
    let sb = Text.StringBuilder()
    appendExpr sb 1 expr
    sb.ToString()

/// A formula whose references changed (rows put in, a cell moved), as it was
/// written with only those references written anew - its brackets, spaces and
/// spelling as the user (or Excel) left them, as Excel keeps them. `printedWas`
/// and `printedNow` are the tree before and after the change printed; when they
/// differ in anything but reference tokens, or the text cannot be matched token
/// for token with what was printed of it, None (the caller prints the tree).
let private spliceChangedReferences (formula: string) (printedWas: string) (printedNow: string) : string option =
    let referenceToken t = match t with TIdent _ | TSheet _ | TError _ | TNumber _ -> true | _ -> false
    let same a b =
        match a, b with
        | TIdent x, TIdent y | TSheet x, TSheet y -> String.Equals(x, y, StringComparison.OrdinalIgnoreCase)
        | _ -> a = b
    try
        if formula.TrimStart().StartsWith "{" then None
        else
            let spans = System.Collections.Generic.List<struct (int * int)>()
            let written = tokenizeSpansFrom formula (if formula.StartsWith "=" then 1 else 0) spans
            let was = tokenize printedWas
            let nowSpans = System.Collections.Generic.List<struct (int * int)>()
            let now = tokenizeSpansFrom printedNow 0 nowSpans
            let changed = [ for k in 0 .. was.Length - 1 do if k < now.Length && not (was.[k] = now.[k]) then yield k ]
            if was.Length <> now.Length || changed.IsEmpty || changed |> List.exists (fun k -> not (referenceToken was.[k] && referenceToken now.[k]))
               || written.Length * was.Length > 1000000 then None
            else
                // Match the written tokens with the printed ones (longest common subsequence): what the
                // printing leaves out may only be brackets.
                let w, p = written.Length, was.Length
                let at i j = i * (p + 1) + j
                let lengths : int[] = Array.zeroCreate ((w + 1) * (p + 1))
                for i = w - 1 downto 0 do
                    for j = p - 1 downto 0 do
                        lengths.[at i j] <- if same written.[i] was.[j] then lengths.[at (i + 1) (j + 1)] + 1 else max lengths.[at (i + 1) j] lengths.[at i (j + 1)]
                let writtenOf : int[] = Array.create p (-1)
                let mutable i, j = 0, 0
                let mutable ok = true
                while i < w && j < p do
                    if same written.[i] was.[j] then
                        writtenOf.[j] <- i
                        i <- i + 1
                        j <- j + 1
                    elif lengths.[at (i + 1) j] >= lengths.[at i (j + 1)] then
                        (match written.[i] with TLParen | TRParen -> () | _ -> ok <- false)
                        i <- i + 1
                    else
                        ok <- false
                        j <- j + 1
                while i < w do
                    (match written.[i] with TLParen | TRParen -> () | _ -> ok <- false)
                    i <- i + 1
                if not ok || j < p || writtenOf |> Array.exists (fun index -> index < 0) then None
                else
                    let sb = System.Text.StringBuilder()
                    let mutable from = 0
                    for k in changed do
                        let struct (start, finish) = spans.[writtenOf.[k]]
                        let struct (nowStart, nowFinish) = nowSpans.[k]
                        sb.Append(formula.Substring(from, start - from)).Append(printedNow.Substring(nowStart, nowFinish - nowStart)) |> ignore
                        from <- finish
                    sb.Append(formula.Substring from) |> ignore
                    Some (sb.ToString())
    with _ -> None

/// Parse `formula`, apply `rewrite` to its tree and print it back, keeping the
/// leading `=` exactly as the original had it. Text that does not parse — or
/// that the rewrite leaves untouched — is returned as it was, so formulas that
/// need no change keep their original spelling.
let rewriteFormulaText (rewrite: Expr -> Expr) (formula: string) : string =
    match parseFormula formula with
    | Result.Error _ -> formula
    | Ok expr ->
        let rewritten = rewrite expr
        if rewritten = expr then formula
        else
            let text = exprToFormula rewritten
            match spliceChangedReferences formula (exprToFormula expr) text with
            | Some kept -> kept
            | None -> if formula.StartsWith "=" then "=" + text else text

/// Shift a formula's relative references by (rows, cols) — what copying or
/// filling it to another cell does.
let shiftFormulaText (dr: int) (dc: int) (formula: string) : string =
    rewriteFormulaText (shiftExpr dr dc) formula

/// Re-point references after rows/columns are inserted or deleted, with
/// Excel's rules: references at or after an insertion move, a range spanning
/// it grows, a reference to a deleted cell becomes #REF!, a range that loses
/// some of its rows/columns shrinks. Only references whose sheet qualifier
/// satisfies `targetsEditedSheet` (None = unqualified) are touched.
let rec adjustExprForEdit (edit: StructureEdit) (targetsEditedSheet: string option -> bool) (expr: Expr) : Expr =
    let onRows, isInsert, at, count =
        match edit with
        | InsertRows (at, count) -> true, true, at, count
        | DeleteRows (at, count) -> true, false, at, count
        | InsertCols (at, count) -> false, true, at, count
        | DeleteCols (at, count) -> false, false, at, count
    let coord (r: RefAst) = if onRows then r.Row else r.Col
    let withCoord (r: RefAst) value = if onRows then { r with Row = value } else { r with Col = value }
    let deleted index = index >= at && index < at + count
    let movePoint index =
        if isInsert then (if index >= at then index + count else index)
        elif index >= at + count then index - count
        else index
    let recurse = adjustExprForEdit edit targetsEditedSheet
    match expr with
    | Number _ | Str _ | Bool _ | ErrorLit _ | NameRef _ | TableRef _ | ArrayLit _ | MissingArg -> expr
    | CellRef r when not (targetsEditedSheet r.Sheet) -> expr
    | SpillRef r when not (targetsEditedSheet r.Sheet) -> expr
    | RangeRef (a, _) when not (targetsEditedSheet a.Sheet) -> expr
    | CellRef r ->
        if not isInsert && deleted (coord r) then ErrorLit Err.ref
        else CellRef (withCoord r (movePoint (coord r)))
    | SpillRef r ->
        if not isInsert && deleted (coord r) then ErrorLit Err.ref
        else SpillRef (withCoord r (movePoint (coord r)))
    | RangeRef (a, b) ->
        let wholeAxis = if onRows then isWholeColumnRange a b else isWholeRowRange a b
        if wholeAxis then expr
        else
            let lo = min (coord a) (coord b)
            let hi = max (coord a) (coord b)
            if isInsert then
                RangeRef (withCoord a (movePoint (coord a)), withCoord b (movePoint (coord b)))
            elif lo >= at && hi < at + count then
                ErrorLit Err.ref
            else
                // A boundary inside the deleted band snaps to the band's edge.
                let newLo = if deleted lo then at else movePoint lo
                let newHi = if deleted hi then at - 1 else movePoint hi
                let aIsLo = coord a <= coord b
                RangeRef (withCoord a (if aIsLo then newLo else newHi),
                          withCoord b (if aIsLo then newHi else newLo))
    | UnaryOp (op, operand) -> UnaryOp (op, recurse operand)
    | BinOp _ ->
        let operations = ResizeArray<struct (string * Expr)>()
        let mutable current = expr
        let mutable collecting = true
        while collecting do
            match current with
            | BinOp (left, op, right) ->
                operations.Add(struct (op, right))
                current <- left
            | _ -> collecting <- false
        let mutable result = recurse current
        for index = operations.Count - 1 downto 0 do
            let struct (op, right) = operations.[index]
            result <- BinOp (result, op, recurse right)
        result
    | Function (name, args) -> Function (name, args |> List.map recurse)

let adjustFormulaTextForEdit (edit: StructureEdit) (targetsEditedSheet: string option -> bool) (formula: string) : string =
    rewriteFormulaText (adjustExprForEdit edit targetsEditedSheet) formula

/// Re-point references after a block of cells was moved (cut and paste), with
/// Excel's rules: a reference to a moved cell follows it (onto another sheet
/// when the block went there), a range moves when the whole of it moved, and
/// a reference to a cell the block landed on is #REF! — that cell is gone.
let rec moveExprReferences (move: CellMove) (expr: Expr) : Expr =
    let inSource row col = row >= move.Top && row <= move.Bottom && col >= move.Left && col <= move.Right
    let inTarget row col = inSource (row - move.DeltaRow) (col - move.DeltaCol)
    let moved (r: RefAst) =
        { r with Row = r.Row + move.DeltaRow; Col = r.Col + move.DeltaCol; Sheet = move.Requalify r.Sheet }
    let stationary (r: RefAst) =
        match r.Sheet, move.QualifyStationary with
        | None, Some _ -> { r with Sheet = move.QualifyStationary }
        | _ -> r
    let recurse = moveExprReferences move
    match expr with
    | Number _ | Str _ | Bool _ | ErrorLit _ | NameRef _ | TableRef _ | ArrayLit _ | MissingArg -> expr
    | CellRef r ->
        if move.OnSourceSheet r.Sheet && inSource r.Row r.Col then CellRef (moved r)
        elif move.OnTargetSheet r.Sheet && inTarget r.Row r.Col then ErrorLit Err.ref
        else CellRef (stationary r)
    | SpillRef r ->
        if move.OnSourceSheet r.Sheet && inSource r.Row r.Col then SpillRef (moved r)
        elif move.OnTargetSheet r.Sheet && inTarget r.Row r.Col then ErrorLit Err.ref
        else SpillRef (stationary r)
    | RangeRef (a, b) ->
        if isWholeColumnRange a b || isWholeRowRange a b then RangeRef (stationary a, stationary b)
        elif move.OnSourceSheet a.Sheet && inSource a.Row a.Col && inSource b.Row b.Col then RangeRef (moved a, moved b)
        elif move.OnTargetSheet a.Sheet && inTarget a.Row a.Col && inTarget b.Row b.Col then ErrorLit Err.ref
        else RangeRef (stationary a, stationary b)
    | UnaryOp (op, operand) -> UnaryOp (op, recurse operand)
    | BinOp _ ->
        let operations = ResizeArray<struct (string * Expr)>()
        let mutable current = expr
        let mutable collecting = true
        while collecting do
            match current with
            | BinOp (left, op, right) ->
                operations.Add(struct (op, right))
                current <- left
            | _ -> collecting <- false
        let mutable result = recurse current
        for index = operations.Count - 1 downto 0 do
            let struct (op, right) = operations.[index]
            result <- BinOp (result, op, recurse right)
        result
    | Function (name, args) -> Function (name, args |> List.map recurse)

let moveFormulaText (move: CellMove) (formula: string) : string =
    rewriteFormulaText (moveExprReferences move) formula

/// Apply `leaf` to every reference of an expression (cell, range, spill and
/// table references), leaving the rest as it is.
let rec mapExprReferences (leaf: Expr -> Expr) (expr: Expr) : Expr =
    let recurse = mapExprReferences leaf
    match expr with
    | CellRef _ | RangeRef _ | SpillRef _ | TableRef _ -> leaf expr
    | Number _ | Str _ | Bool _ | ErrorLit _ | NameRef _ | ArrayLit _ | MissingArg -> expr
    | UnaryOp (op, operand) -> UnaryOp (op, recurse operand)
    | BinOp _ ->
        let operations = ResizeArray<struct (string * Expr)>()
        let mutable current = expr
        let mutable collecting = true
        while collecting do
            match current with
            | BinOp (left, op, right) ->
                operations.Add(struct (op, right))
                current <- left
            | _ -> collecting <- false
        let mutable result = recurse current
        for index = operations.Count - 1 downto 0 do
            let struct (op, right) = operations.[index]
            result <- BinOp (result, op, recurse right)
        result
    | Function (name, args) -> Function (name, args |> List.map recurse)

/// Re-point references after a sheet was renamed or deleted, or a table
/// renamed (see Model.SheetRewrite). Sheet names match whatever their case.
/// A 3-D reference (`First:Last!A1`) follows the rename of either end. One
/// that loses an end to a deleted sheet ends at that sheet's neighbour inside
/// the span, as in Excel (`Jan:Mar!A1` without Mar is `Jan:Feb!A1`; a span of
/// two sheets is left with one, `Jan!A1`); with no sheet left it is #REF!.
/// (A sheet deleted from the middle of a span, and sheets moved into or out
/// of one, change what the reference reads, not how it is written.)
let rewriteExprForSheets (change: SheetRewrite) (expr: Expr) : Expr =
    let same (a: string) (b: string) = String.Equals(a, b, StringComparison.OrdinalIgnoreCase)
    // The ends of a qualifier: one name, or the two of a 3-D reference.
    let ends (qualifier: string) =
        match qualifier.IndexOf ':' with
        | -1 -> [ qualifier ]
        | colon -> [ qualifier.Substring(0, colon); qualifier.Substring(colon + 1) ]
    match change with
    | RenameSheetRefs (oldName, newName) ->
        let requalify (sheet: string option) =
            match sheet with
            | Some qualifier when ends qualifier |> List.exists (same oldName) ->
                Some (ends qualifier |> List.map (fun name -> if same name oldName then newName else name) |> String.concat ":")
            | other -> other
        mapExprReferences (fun reference ->
            match reference with
            | CellRef r -> CellRef { r with Sheet = requalify r.Sheet }
            | SpillRef r -> SpillRef { r with Sheet = requalify r.Sheet }
            | RangeRef (a, b) -> RangeRef ({ a with Sheet = requalify a.Sheet }, { b with Sheet = requalify b.Sheet })
            | other -> other) expr
    | DeleteSheetRefs (name, sheets) ->
        let position (sheetName: string) = sheets |> List.tryFindIndex (same sheetName)
        // What is left of a qualifier: itself, another one, or nothing (Error).
        let remaining (sheet: string option) : Result<string option, unit> =
            match sheet with
            | None -> Ok None
            | Some qualifier ->
                match ends qualifier with
                | [ only ] -> if same only name then Result.Error () else Ok sheet
                | [ first; last ] when same first name || same last name ->
                    if same first name && same last name then Result.Error ()
                    else
                        // The end that stays, and the deleted end's neighbour towards it.
                        let kept = if same first name then last else first
                        match position name, position kept with
                        | Some gone, Some stays ->
                            let neighbour = List.item (if stays > gone then gone + 1 else gone - 1) sheets
                            if same neighbour kept then Ok (Some kept)
                            elif same first name then Ok (Some (neighbour + ":" + last))
                            else Ok (Some (first + ":" + neighbour))
                        | _ -> Result.Error ()
                | _ -> Ok sheet
        mapExprReferences (fun reference ->
            match reference with
            | CellRef r ->
                (match remaining r.Sheet with
                 | Ok sheet -> if sheet = r.Sheet then reference else CellRef { r with Sheet = sheet }
                 | Result.Error () -> ErrorLit Err.ref)
            | SpillRef r ->
                (match remaining r.Sheet with
                 | Ok sheet -> if sheet = r.Sheet then reference else SpillRef { r with Sheet = sheet }
                 | Result.Error () -> ErrorLit Err.ref)
            | RangeRef (a, b) ->
                (match remaining a.Sheet, remaining b.Sheet with
                 | Ok first, Ok second -> if first = a.Sheet && second = b.Sheet then reference else RangeRef ({ a with Sheet = first }, { b with Sheet = second })
                 | _ -> ErrorLit Err.ref)
            | other -> other) expr
    | RenameTableRefs (oldName, newName) ->
        mapExprReferences (fun reference ->
            match reference with
            | TableRef selection when same selection.Table oldName -> TableRef { selection with Table = newName }
            | other -> other) expr
    | RenameTableColumnRefs (table, oldName, newName, unqualified) ->
        let renamed (column: string option) =
            match column with
            | Some name when same name oldName -> Some newName
            | other -> other
        mapExprReferences (fun reference ->
            match reference with
            | TableRef selection when same selection.Table table || (unqualified && selection.Table = "") ->
                let moved = { selection with ColStart = renamed selection.ColStart; ColEnd = renamed selection.ColEnd }
                if moved = selection then reference else TableRef moved
            | other -> other) expr

let rewriteFormulaTextForSheets (change: SheetRewrite) (formula: string) : string =
    rewriteFormulaText (rewriteExprForSheets change) formula

/// Replace every `NameRef` whose (case-insensitive) name is bound in `subst`
/// with its bound expression. This is how LET inlines its names: a LET value
/// expression is substituted as-is (not pre-evaluated), so a name bound to a
/// range stays a range for the structural functions (SUM/SUMIFS/INDEX/…) and
/// so dependency extraction still sees the real cell references.
let rec substituteNames (subst: Map<string, Expr>) (expr: Expr) : Expr =
    let substituteBinOpChain expr =
        let operations = ResizeArray<struct (string * Expr)>()
        let mutable current = expr
        let mutable collecting = true
        while collecting do
            match current with
            | BinOp (left, op, right) ->
                operations.Add(struct (op, right))
                current <- left
            | _ ->
                collecting <- false
        let mutable result = substituteNames subst current
        for index = operations.Count - 1 downto 0 do
            let struct (op, right) = operations.[index]
            result <- BinOp(result, op, substituteNames subst right)
        result
    match expr with
    | NameRef n ->
        match Map.tryFind (n.ToUpperInvariant()) subst with
        | Some e -> e
        | None -> expr
    | Number _ | Str _ | Bool _ | ErrorLit _ | CellRef _ | RangeRef _ | TableRef _ | ArrayLit _ | SpillRef _ | MissingArg -> expr
    | UnaryOp (op, e) -> UnaryOp (op, substituteNames subst e)
    | BinOp _ -> substituteBinOpChain expr
    | Function (nm, fa) ->
        let substituted = ResizeArray<Expr>()
        for arg in fa do
            substituted.Add(substituteNames subst arg)
        let fa' = resizeArrayToList substituted
        // A call whose name is bound (via LET) to a LAMBDA is beta-reduced
        // here: `LET(f, LAMBDA(x,x*2), f(5))` → `5*2`. Other calls just
        // substitute their arguments.
        match Map.tryFind (nm.ToUpperInvariant()) subst with
        | Some (Function (ln, lparts)) when ln.ToUpperInvariant() = "LAMBDA" -> applyLambdaExpr lparts fa'
        | _ -> Function (nm, fa')

/// Beta-reduce a LAMBDA: `parts` is `[param1; …; paramN; body]` (params are
/// `NameRef`s). Bind each param to the matching argument expression and
/// substitute into the body, returning the reduced expression (then eval'd).
and applyLambdaExpr (parts: Expr list) (args: Expr list) : Expr =
    let rec bind subst remaining arguments =
        match remaining with
        | [] -> ErrorLit Err.value
        | [body] -> substituteNames subst body
        | NameRef n :: restParts ->
            match arguments with
            | argument :: restArguments ->
                bind (Map.add (n.ToUpperInvariant()) argument subst) restParts restArguments
            | [] ->
                // An omitted argument is missing (ISOMITTED is TRUE of it; as a value it is 0).
                bind (Map.add (n.ToUpperInvariant()) MissingArg subst) restParts []
        | _ :: restParts ->
            match arguments with
            | _ :: restArguments -> bind subst restParts restArguments
            | [] -> bind subst restParts []

    bind Map.empty parts args

let private lambdaBodyUsesOnlySimpleArithmetic (paramName: string) (expr: Expr) =
    let stack = ResizeArray<Expr>()
    stack.Add expr
    let mutable ok = true

    while ok && stack.Count > 0 do
        let last = stack.Count - 1
        let expr = stack.[last]
        Portable.Stack.dropLast stack

        match expr with
        | Number _ | Str _ | Bool _ | ErrorLit _ -> ()
        | NameRef name ->
            if not (equalsOrdinalIgnoreCase name paramName) then
                ok <- false
        | UnaryOp (_, inner) -> stack.Add inner
        | BinOp (left, _, right) ->
            stack.Add left
            stack.Add right
        | _ -> ok <- false

    ok

/// The body of a LAMBDA, if `e` is one (`Function("LAMBDA", parts)`).
let lambdaPartsOf (e: Expr) : Expr list option =
    match e with
    | Function (n, parts) when equalsOrdinalIgnoreCase n "LAMBDA" && not (List.isEmpty parts) -> Some parts
    | _ -> None

/// Wrap a computed value back as a literal expression, so it can be passed as
/// a LAMBDA argument (MAP/REDUCE/SCAN/MAKEARRAY feed values into lambdas).
let valueToExpr (v: CellValue) : Expr =
    match unwrapValue v with
    | CellValue.Number n -> Number n
    | CellValue.String s -> Str s
    | CellValue.Boolean b -> Bool b
    | CellValue.Error e -> ErrorLit e
    | CellValue.Empty -> Number 0.0
    | CellValue.Formula _ -> Number 0.0

/// True if the expression contains a volatile function (RAND/RANDBETWEEN/
/// NOW/TODAY/OFFSET/INDIRECT) anywhere in its tree. Such cells must be
/// recalculated on every pass regardless of dependency state.
let containsVolatile (expr: Expr) : bool =
    let stack = ResizeArray<Expr>()
    stack.Add expr
    let mutable found = false
    while not found && stack.Count > 0 do
        let last = stack.Count - 1
        let expr = stack.[last]
        Portable.Stack.dropLast stack
        match expr with
        | Function (name, args) ->
            if equalsOrdinalIgnoreCase name "RAND"
               || equalsOrdinalIgnoreCase name "RANDBETWEEN"
               || equalsOrdinalIgnoreCase name "NOW"
               || equalsOrdinalIgnoreCase name "TODAY"
               || equalsOrdinalIgnoreCase name "OFFSET"
               || equalsOrdinalIgnoreCase name "INDIRECT"
               || catalogFlag 1 name then
                found <- true
            else
                for arg in args do
                    stack.Add arg
        | UnaryOp (_, inner) -> stack.Add inner
        | BinOp (left, _, right) ->
            stack.Add left
            stack.Add right
        | _ -> ()
    found

let private containsFunctionNamed (targetName: string) (expr: Expr) =
    let stack = ResizeArray<Expr>()
    stack.Add expr
    let mutable found = false
    while not found && stack.Count > 0 do
        let last = stack.Count - 1
        let expr = stack.[last]
        Portable.Stack.dropLast stack
        match expr with
        | Function (name, args) ->
            if equalsOrdinalIgnoreCase name targetName then
                found <- true
            else
                for arg in args do
                    stack.Add arg
        | UnaryOp (_, inner) -> stack.Add inner
        | BinOp (left, _, right) ->
            stack.Add left
            stack.Add right
        | _ -> ()
    found

/// Resolve the model a reference targets. Unqualified refs use the current
/// model; sheet-qualified refs resolve through the workbook. Returns None
/// when a sheet name can't be resolved (caller surfaces #REF!).
let private resolveModel (sheet: string option) (model: SpreadsheetModel)
                         (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                         : SpreadsheetModel option =
    match sheet with
    | None -> Some model
    | Some name ->
        match workbook with
        | None -> None
        | Some wb ->
            match Rbt.Spreadsheet.Workbook.resolveSheetName name wb |> Option.bind (fun sid -> Map.tryFind sid wb.Sheets) with
            // (A sheet of an opened file is converted when a formula first reads it.)
            | Some s -> Some (resolveSheet s.Model)
            // `[1]Sheet`: what the file keeps of another workbook's sheet.
            | None when name.StartsWith "[" && not wb.External.IsEmpty -> Map.tryFind (name.ToUpperInvariant()) wb.External
            | None -> None

/// Read a single cell's value, resolving cross-sheet references.
let private readCell (r: RefAst) (model: SpreadsheetModel)
                     (workbook: Rbt.Spreadsheet.Workbook.Workbook option) : CellValue =
    match r.Sheet with
    | None -> (match getCellAt r.Row r.Col model with Some cell -> unwrapValue cell.Value | None -> CellValue.Empty)
    | Some _ ->
        match resolveModel r.Sheet model workbook with
        | Some m -> (match getCellAt r.Row r.Col m with Some cell -> unwrapValue cell.Value | None -> CellValue.Empty)
        | None -> err Err.ref

/// Resolve a spill-range operator `A1#` to its (top-left, bottom-right) RefAst
/// pair: the bounding rectangle of the anchor plus every cell that spilled from
/// it (`model.Spill` maps spilled cell → anchor). A formula anchor that didn't
/// spill resolves to the single anchor cell.
let spillRangeOf (anchor: RefAst) (model: SpreadsheetModel) : RefAst * RefAst =
    let aAddr = { Row = anchor.Row; Col = anchor.Col }
    let mutable r0 = aAddr.Row
    let mutable c0 = aAddr.Col
    let mutable r1 = aAddr.Row
    let mutable c1 = aAddr.Col
    model.Spill
    |> Map.iter (fun cell anc ->
        if anc = aAddr then
            if cell.Row < r0 then r0 <- cell.Row
            if cell.Col < c0 then c0 <- cell.Col
            if cell.Row > r1 then r1 <- cell.Row
            if cell.Col > c1 then c1 <- cell.Col)
    { anchor with Row = r0; Col = c0 }, { anchor with Row = r1; Col = c1 }

/// `spillRangeOf` for an anchor that may sit on another sheet (`Sheet2!A1#`).
let spillRangeIn (anchor: RefAst) (model: SpreadsheetModel)
                 (workbook: Rbt.Spreadsheet.Workbook.Workbook option) : RefAst * RefAst =
    match anchor.Sheet with
    | None -> spillRangeOf anchor model
    | Some _ ->
        match resolveModel anchor.Sheet model workbook with
        | Some target -> spillRangeOf anchor target
        | None -> anchor, anchor

/// Expand a RangeRef to the flat list of unwrapped cell values, row-major.
/// The sheet qualifier on the start ref selects the target sheet.
let iterRangeRef (a: RefAst) (b: RefAst) (model: SpreadsheetModel)
                 (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                 (emit: CellValue -> unit) : unit =
    match resolveModel a.Sheet model workbook with
    | Some m ->
        let sr, sc, er, ec = normalizeRange a b
        for row in sr .. er do
            for col in sc .. ec do
                emit (getCellValue { Row = row; Col = col } m |> unwrapValue)
    | None -> emit (err Err.ref)

let expandRangeRef (a: RefAst) (b: RefAst) (model: SpreadsheetModel)
                   (workbook: Rbt.Spreadsheet.Workbook.Workbook option) : CellValue list =
    let values = ResizeArray<CellValue>()
    iterRangeRef a b model workbook (fun value -> values.Add value)
    resizeArrayToList values

let firstRangeRefValue (a: RefAst) (b: RefAst) (model: SpreadsheetModel)
                       (workbook: Rbt.Spreadsheet.Workbook.Workbook option) : CellValue =
    match resolveModel a.Sheet model workbook with
    | Some m ->
        let sr, sc, _, _ = normalizeRange a b
        getCellValue { Row = sr; Col = sc } m |> unwrapValue
    | None -> err Err.ref

/// Expand a structured table reference to its column's values.
/// Resolve a structured reference to its absolute model area `(r0,c0,r1,c1)`,
/// or an error code when the table / column / item can't be resolved. Rows ×
/// columns are always contiguous for the specifiers we support. `Ok` with
/// `r0 > r1` denotes a valid-but-empty body (a table with no data rows).
let tableSelArea (sel: TableSel)
                 (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                 (sheet: Rbt.Spreadsheet.Workbook.SheetId option)
                 (currentCell: CellAddress option) : Result<int * int * int * int, string> =
    match workbook with
    | None -> Result.Error Err.ref
    | Some wb ->
        let tableOpt =
            if sel.Table <> "" then Rbt.Spreadsheet.Workbook.getTable sel.Table wb
            else currentCell |> Option.bind (fun c -> Rbt.Spreadsheet.Workbook.findTableContaining sheet c wb)
        match tableOpt with
        | None -> Result.Error Err.ref
        | Some t ->
            let headerRows = if t.HasHeaders then 1 else 0
            let totalsRows = if t.HasTotals then 1 else 0
            let dataStart = t.StartRow + headerRows
            let dataEnd = t.EndRow - totalsRows
            let colRange =
                match sel.ColStart with
                | None -> Some (t.StartCol, t.StartCol + t.Columns.Length - 1)
                | Some cs ->
                    match Rbt.Spreadsheet.Workbook.tableColumnIndex t cs with
                    | None -> None
                    | Some ci ->
                        match sel.ColEnd with
                        | None -> Some (t.StartCol + ci, t.StartCol + ci)
                        | Some ce ->
                            Rbt.Spreadsheet.Workbook.tableColumnIndex t ce
                            |> Option.map (fun cj -> t.StartCol + min ci cj, t.StartCol + max ci cj)
            let rowRange =
                match sel.Item with
                | TblData -> Some (dataStart, dataEnd)   // r0>r1 ⇒ empty body
                | TblAll -> Some (t.StartRow, t.EndRow)
                | TblHeaders -> if t.HasHeaders then Some (t.StartRow, t.StartRow) else None
                | TblTotals -> if t.HasTotals then Some (t.EndRow, t.EndRow) else None
                | TblThisRow ->
                    match currentCell with
                    | Some c when c.Row >= dataStart && c.Row <= dataEnd -> Some (c.Row, c.Row)
                    | _ -> None
            match colRange, rowRange with
            | Some (c0, c1), Some (r0, r1) -> Result.Ok (r0, c0, r1, c1)
            | Some _, None -> Result.Error (match sel.Item with TblThisRow -> Err.value | _ -> Err.ref)
            | None, _ -> Result.Error Err.ref

/// A formula with its structured references to one table written as the
/// plain references they stand for - what converting a table to a range
/// does to the formulas that named it. `homeSheet` and `cell` say where the
/// formula stands (a this-row reference is the cell of its own row; a
/// formula on another sheet gets the table's sheet name); a reference that
/// selects nothing becomes #REF!. Text that does not parse, or that does
/// not name the table, is returned as it was.
let tableRefsToRanges (workbook: Rbt.Spreadsheet.Workbook.Workbook) (tableName: string)
                      (homeSheet: Rbt.Spreadsheet.Workbook.SheetId) (cell: CellAddress option) (formula: string) : string =
    match Rbt.Spreadsheet.Workbook.getTable tableName workbook with
    | None -> formula
    | Some table ->
        let qualifier =
            if table.SheetId = homeSheet then None
            else Map.tryFind table.SheetId workbook.Sheets |> Option.map (fun sheet -> sheet.Name)
        let names (selection: TableSel) =
            if selection.Table <> "" then String.Equals(selection.Table, table.Name, StringComparison.OrdinalIgnoreCase)
            else
                match cell with
                | Some at ->
                    table.SheetId = homeSheet
                    && at.Row >= table.StartRow && at.Row <= table.EndRow && at.Col >= table.StartCol && at.Col <= table.EndCol
                | None -> false
        let corner row col : RefAst = { Sheet = qualifier; Row = row; Col = col; RowAbs = true; ColAbs = true }
        rewriteFormulaText (mapExprReferences (fun reference ->
            match reference with
            | TableRef selection when names selection ->
                (match tableSelArea selection (Some workbook) (Some homeSheet) cell with
                 | Ok (r0, c0, r1, c1) when r0 <= r1 && c0 <= c1 ->
                    if r0 = r1 && c0 = c1 then CellRef (corner r0 c0)
                    else RangeRef (corner r0 c0, corner r1 c1)
                 | _ -> ErrorLit Err.ref)
            | other -> other)) formula

/// The sheet a structured reference reads: the table's own sheet when the
/// formula stands on another one (`model` is the formula's sheet).
let private tableSelModel (sel: TableSel)
                          (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                          (sheet: Rbt.Spreadsheet.Workbook.SheetId option)
                          (model: SpreadsheetModel) : SpreadsheetModel =
    if sel.Table = "" then model
    else
        match workbook |> Option.bind (fun wb -> Rbt.Spreadsheet.Workbook.getTable sel.Table wb |> Option.map (fun table -> wb, table)) with
        | Some (wb, table) when Some table.SheetId <> sheet ->
            (match Map.tryFind table.SheetId wb.Sheets with
             | Some other -> resolveSheet other.Model
             | None -> model)
        | _ -> model

/// Resolve a structured reference (`Table[[#Totals],[Amount]]` etc.) to the
/// flat list of cell values it covers (row-major over the selected rows ×
/// columns). `sheet`/`currentCell` resolve this-row and unqualified refs.
let iterTableSel (sel: TableSel)
                 (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                 (sheet: Rbt.Spreadsheet.Workbook.SheetId option)
                 (currentCell: CellAddress option)
                 (model: SpreadsheetModel)
                 (emit: CellValue -> unit) : unit =
    match tableSelArea sel workbook sheet currentCell with
    | Result.Error e -> emit (err e)
    | Result.Ok (r0, c0, r1, c1) ->
        let model = tableSelModel sel workbook sheet model
        for row in r0 .. r1 do
            for col in c0 .. c1 do
                emit (getCellValue { Row = row; Col = col } model |> unwrapValue)

let expandTableSel (sel: TableSel)
                   (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                   (sheet: Rbt.Spreadsheet.Workbook.SheetId option)
                   (currentCell: CellAddress option)
                   (model: SpreadsheetModel) : CellValue list =
    let values = ResizeArray<CellValue>()
    iterTableSel sel workbook sheet currentCell model (fun value -> values.Add value)
    resizeArrayToList values

let firstTableSelValue (sel: TableSel)
                       (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                       (sheet: Rbt.Spreadsheet.Workbook.SheetId option)
                       (currentCell: CellAddress option)
                       (model: SpreadsheetModel) : CellValue =
    match tableSelArea sel workbook sheet currentCell with
    | Result.Error e -> err e
    | Result.Ok (r0, c0, r1, _) ->
        if r0 > r1 then err Err.ref
        else getCellValue { Row = r0; Col = c0 } (tableSelModel sel workbook sheet model) |> unwrapValue

// --- Date helpers -----------------------------------------------------------
// Excel serial dates: serial 1 = 1900-01-01. Excel's phantom 1900-02-29 (the
// 1900-leap-year bug) makes the simple offset exact for every date from
// 1900-03-01 (serial 61) onward, which covers every realistic spreadsheet.
let private utcDate year month day = DateTime(year, month, day, 0, 0, 0, DateTimeKind.Utc)
/// (A serial far outside the calendar - a formula fed something that is not a
/// date - is held at the calendar's ends: .NET's DateTime throws there, which
/// took a whole recalculation down with one cell.)
let private serialToDateIn (date1904: bool) (s: float) : DateTime =
    let s = s + dateSystemDays date1904
    let held = if System.Double.IsNaN s then 0.0 else max -657000.0 (min 2958465.0 s)
    excelEpoch.AddDays(held)
/// Reads the date's own calendar fields, so a local-kind value (a parsed
/// string, DateTime.Today) maps to the serial for the date it displays.
let private dateToSerialIn (date1904: bool) (d: DateTime) : float =
    (DateTime(d.Year, d.Month, d.Day, d.Hour, d.Minute, d.Second, d.Millisecond, DateTimeKind.Utc) - excelEpoch).TotalDays - dateSystemDays date1904

/// A whole serial moved by a number of months (EDATE) or to the end of the month it is then
/// in (EOMONTH), in whole numbers - no date object is made: these two are what a model's
/// timeline is built of, thousands of times an edit. None for a serial with a time of day,
/// one before 1 March 1900 (where Excel's calendar has a day too many) or outside the
/// calendar: the date object then says what it is.
let private monthsFromSerial (date1904: bool) (serial: float) (months: int) (endOfMonth: bool) : float voption =
    let day = serial + dateSystemDays date1904
    if day <> Math.Floor day || day < 61.0 || day > 2958465.0 || months > 120000 || months < -120000 then ValueNone
    else
        // The civil date of the day (days from 1 January 1970, serial 25569; years that begin in March).
        let z = int day - 25569 + 719468
        let era = (if z >= 0 then z else z - 146096) / 146097
        let dayOfEra = z - era * 146097
        let yearOfEra = (dayOfEra - dayOfEra / 1460 + dayOfEra / 36524 - dayOfEra / 146096) / 365
        let dayOfYear = dayOfEra - (365 * yearOfEra + yearOfEra / 4 - yearOfEra / 100)
        let marchMonth = (5 * dayOfYear + 2) / 153
        let dayOfMonth = dayOfYear - (153 * marchMonth + 2) / 5 + 1
        let month = if marchMonth < 10 then marchMonth + 3 else marchMonth - 9
        let year = yearOfEra + era * 400 + (if month <= 2 then 1 else 0)
        // The month that many months on, the day held at its last one.
        let total = year * 12 + (month - 1) + months
        let year2 = if total >= 0 then total / 12 else (total - 11) / 12
        let month2 = total - year2 * 12 + 1
        if year2 < 1900 || year2 > 9999 then ValueNone
        else
            let leap = (year2 % 4 = 0 && year2 % 100 <> 0) || year2 % 400 = 0
            let last = (match month2 with 2 -> (if leap then 29 else 28) | 4 | 6 | 9 | 11 -> 30 | _ -> 31)
            let day2 = if endOfMonth then last else min dayOfMonth last
            // ... and its day number again.
            let marchYear = if month2 <= 2 then year2 - 1 else year2
            let era2 = (if marchYear >= 0 then marchYear else marchYear - 399) / 400
            let yearOfEra2 = marchYear - era2 * 400
            let dayOfYear2 = (153 * (if month2 > 2 then month2 - 3 else month2 + 9) + 2) / 5 + day2 - 1
            let dayOfEra2 = yearOfEra2 * 365 + yearOfEra2 / 4 - yearOfEra2 / 100 + dayOfYear2
            let result = float (era2 * 146097 + dayOfEra2 - 719468 + 25569)
            if result < 61.0 then ValueNone else ValueSome (result - dateSystemDays date1904)

/// The calendar date of a whole serial as year * 10000 + month * 100 + day, in whole numbers
/// (YEAR, MONTH and DAY of a model's timeline are asked for thousands of times an edit).
/// -1 where `monthsFromSerial` gives none: the date object then says what it is.
let private datePartsOfSerial (date1904: bool) (serial: float) : int =
    let day = serial + dateSystemDays date1904
    if day <> Math.Floor day || day < 61.0 || day > 2958465.0 then -1
    else
        let z = int day - 25569 + 719468
        let era = (if z >= 0 then z else z - 146096) / 146097
        let dayOfEra = z - era * 146097
        let yearOfEra = (dayOfEra - dayOfEra / 1460 + dayOfEra / 36524 - dayOfEra / 146096) / 365
        let dayOfYear = dayOfEra - (365 * yearOfEra + yearOfEra / 4 - yearOfEra / 100)
        let marchMonth = (5 * dayOfYear + 2) / 153
        let dayOfMonth = dayOfYear - (153 * marchMonth + 2) / 5 + 1
        let month = if marchMonth < 10 then marchMonth + 3 else marchMonth - 9
        let year = yearOfEra + era * 400 + (if month <= 2 then 1 else 0)
        year * 10000 + month * 100 + dayOfMonth

/// Shared RNG for the volatile RAND / RANDBETWEEN functions.
let private rng = Random()

let private properCase (text: string) : string =
    let sb = System.Text.StringBuilder(text.Length)
    let mutable startWord = true
    for ch in text do
        if Char.IsLetter ch then
            sb.Append(if startWord then Char.ToUpperInvariant ch else Char.ToLowerInvariant ch) |> ignore
            startWord <- false
        else
            sb.Append(Char.ToLowerInvariant ch) |> ignore
            startWord <- not (Char.IsDigit ch)
    sb.ToString()

let private trimExcelText (text: string) =
    let mutable start = 0
    let mutable finish = text.Length - 1
    while start <= finish && Char.IsWhiteSpace text.[start] do
        start <- start + 1
    while finish >= start && Char.IsWhiteSpace text.[finish] do
        finish <- finish - 1

    if start > finish then ""
    else
        let mutable needsCollapse = false
        let mutable i = start
        while not needsCollapse && i <= finish do
            if text.[i] = ' ' && i < finish && text.[i + 1] = ' ' then
                needsCollapse <- true
            i <- i + 1

        if start = 0 && finish = text.Length - 1 && not needsCollapse then
            text
        else
            let sb = System.Text.StringBuilder(finish - start + 1)
            let mutable previousWasSpace = false
            for index = start to finish do
                let ch = text.[index]
                if ch = ' ' then
                    if not previousWasSpace then
                        sb.Append(' ') |> ignore
                    previousWasSpace <- true
                else
                    sb.Append(ch) |> ignore
                    previousWasSpace <- false
            sb.ToString()

let inline private compareFloats (x: float) (y: float) =
    if x < y then -1
    elif x > y then 1
    elif x = y then 0
    else
        let xIsNaN = x <> x
        let yIsNaN = y <> y
        if xIsNaN && yIsNaN then 0
        elif xIsNaN then -1
        elif yIsNaN then 1
        else 0

/// Excel's comparison of two numbers - the operators `=` `<>` `<` `>` `<=` `>=` (in a formula, an
/// array, FILTER's condition), the criteria of COUNTIF / SUMIF(S) / AVERAGEIF(S) / MAXIFS / MINIFS /
/// the database functions, SWITCH: numbers that agree to 15 significant digits are equal (each
/// rounded a half away from zero: `=0.1+0.2=0.3` is TRUE, 1+22 ulp = 1 but 1+23 ulp > 1; no
/// absolute tolerance: 1E-16 > 0). MATCH, the LOOKUPs, RANK, MODE, FREQUENCY, DELTA, GESTEP and
/// sorting compare the binary values (`compareFloats`). So does a conditional format's "cell value"
/// rule (`ConditionalFormatting`). Checked in Excel 2019.
let compareNumbers (x: float) (y: float) =
    if x = y then 0
    else
        let c = compareFloats x y
        // Numbers further apart than 2e-14 of the larger round to different 15-digit numbers.
        if c <> 0 && abs (x - y) <= 2e-14 * max (abs x) (abs y)
           && NumberText.sameToSignificantDigits 15 x y then 0
        else c

// --- Criteria parsing (SUMIF / COUNTIF / …) ---------------------------------
let private appendRegexLiteralChar (sb: Text.StringBuilder) (ch: char) =
    match ch with
    | '\\' | '^' | '$' | '.' | '|' | '?' | '*' | '+' | '(' | ')' | '[' | ']' | '{' | '}' ->
        sb.Append('\\').Append(ch) |> ignore
    | _ ->
        sb.Append(ch) |> ignore

/// Translate Excel's `*` / `?` wildcards into a case-insensitive regex.
let private wildcardRegex (pattern: string) =
    let sb = System.Text.StringBuilder()
    sb.Append('^') |> ignore
    let mutable i = 0
    while i < pattern.Length do
        match pattern.[i] with
        | '~' when i + 1 < pattern.Length ->
            appendRegexLiteralChar sb pattern.[i + 1]
            i <- i + 2
        | '*' ->
            sb.Append(".*") |> ignore
            i <- i + 1
        | '?' ->
            sb.Append('.') |> ignore
            i <- i + 1
        | c ->
            appendRegexLiteralChar sb c
            i <- i + 1
    sb.Append('$') |> ignore
    System.Text.RegularExpressions.Regex(
        sb.ToString(), System.Text.RegularExpressions.RegexOptions.IgnoreCase)

/// The same translation without anchors, for SEARCH: the pattern may match
/// anywhere in the text.
let private wildcardSearchRegex (pattern: string) =
    let sb = System.Text.StringBuilder()
    let mutable i = 0
    while i < pattern.Length do
        match pattern.[i] with
        | '~' when i + 1 < pattern.Length ->
            appendRegexLiteralChar sb pattern.[i + 1]
            i <- i + 2
        | '*' ->
            sb.Append(".*?") |> ignore
            i <- i + 1
        | '?' ->
            sb.Append('.') |> ignore
            i <- i + 1
        | c ->
            appendRegexLiteralChar sb c
            i <- i + 1
    System.Text.RegularExpressions.Regex(
        sb.ToString(),
        System.Text.RegularExpressions.RegexOptions.IgnoreCase ||| System.Text.RegularExpressions.RegexOptions.Singleline)

type private CriteriaOp =
    | CriteriaEq
    | CriteriaNe
    | CriteriaGt
    | CriteriaLt
    | CriteriaGe
    | CriteriaLe

let private parseCriteriaOperator (raw: string) =
    if raw.Length >= 2 then
        match raw.[0], raw.[1] with
        | '>', '=' -> CriteriaGe, 2
        | '<', '=' -> CriteriaLe, 2
        | '<', '>' -> CriteriaNe, 2
        | '>', _ -> CriteriaGt, 1
        | '<', _ -> CriteriaLt, 1
        | '=', _ -> CriteriaEq, 1
        | _ -> CriteriaEq, 0
    elif raw.Length = 1 then
        match raw.[0] with
        | '>' -> CriteriaGt, 1
        | '<' -> CriteriaLt, 1
        | '=' -> CriteriaEq, 1
        | _ -> CriteriaEq, 0
    else
        CriteriaEq, 0

let private hasCriteriaWildcard (text: string) =
    let mutable found = false
    let mutable index = 0
    while not found && index < text.Length do
        match text.[index] with
        | '*' | '?' | '~' -> found <- true
        | _ -> ()
        index <- index + 1
    found

/// An exact-match lookup value that uses `*`, `?` or `~` matches text cells
/// by pattern (VLOOKUP / HLOOKUP / MATCH with match type 0). None otherwise.
let private wildcardLookupPredicate (target: CellValue) : (CellValue -> bool) option =
    match unwrapValue target with
    | CellValue.String pattern when hasCriteriaWildcard pattern ->
        let rx = wildcardRegex pattern
        Some (fun v -> match unwrapValue v with CellValue.String text -> rx.IsMatch text | _ -> false)
    | _ -> None

let private tryCriteriaNumber value =
    match unwrapValue value with
    | CellValue.Number n -> ValueSome n
    | CellValue.Boolean b -> ValueSome (if b then 1.0 else 0.0)
    | CellValue.Empty -> ValueSome 0.0
    | CellValue.String s ->
        let struct (start, length) = trimBounds s
        match Double.TryParse(NumberText.cutToSignificantDigits (trimmedNumberText s start length), NumberStyles.Float, CultureInfo.InvariantCulture) with
        | true, n -> ValueSome n
        | _ -> ValueNone
    | _ -> ValueNone

let private tryCriteriaBool value =
    match unwrapValue value with
    | CellValue.Boolean b -> ValueSome b
    | CellValue.Number n -> ValueSome (n <> 0.0)
    | CellValue.Empty -> ValueSome false
    | CellValue.String s ->
        let struct (start, length) = trimBounds s
        if length = 0 then ValueSome false
        elif sliceEqualsOrdinalIgnoreCase s start length "TRUE" then ValueSome true
        elif sliceEqualsOrdinalIgnoreCase s start length "FALSE" then ValueSome false
        else
            match Double.TryParse(NumberText.cutToSignificantDigits (trimmedNumberText s start length), NumberStyles.Float, CultureInfo.InvariantCulture) with
            | true, n -> ValueSome (n <> 0.0)
            | _ -> ValueNone
    | _ -> ValueNone

let private criteriaTextCompare rest value =
    match unwrapValue value with
    | CellValue.String s -> ValueSome (String.Compare(s, rest, StringComparison.OrdinalIgnoreCase))
    | CellValue.Number n -> ValueSome (String.Compare(numberToText n, rest, StringComparison.OrdinalIgnoreCase))
    | CellValue.Boolean b -> ValueSome (String.Compare((if b then "TRUE" else "FALSE"), rest, StringComparison.OrdinalIgnoreCase))
    | CellValue.Empty -> ValueSome (String.Compare("", rest, StringComparison.OrdinalIgnoreCase))
    | _ -> ValueNone

/// Compile a criteria value (`">5"`, `"<>0"`, `"app*"`, `42`, …) into a
/// predicate over candidate cell values. Used by the *IF / *IFS family.
let private compileCriteria (criteria: CellValue) : CellValue -> bool =
    // Excel does not coerce the cells it tests: a blank is not 0 and TRUE is
    // not 1. An ordering test (`<5`) looks at real numbers only; an equality
    // test also accepts text that reads as that number (`"2"` equals 2).
    let equalsNumber (num: float) (v: CellValue) =
        match unwrapValue v with
        | CellValue.Number x -> compareNumbers x num = 0
        | CellValue.String s ->
            let struct (start, length) = trimBounds s
            (match Double.TryParse(NumberText.cutToSignificantDigits (trimmedNumberText s start length), NumberStyles.Float, CultureInfo.InvariantCulture) with
             | true, x -> compareNumbers x num = 0
             | _ -> false)
        | _ -> false
    match unwrapValue criteria with
    | CellValue.Number cn -> equalsNumber cn
    | CellValue.Boolean cb ->
        fun v -> match unwrapValue v with CellValue.Boolean x -> x = cb | _ -> false
    | CellValue.String raw ->
        let op, restStart = parseCriteriaOperator raw
        let rest = if restStart = 0 then raw else raw.Substring(restStart)
        match Double.TryParse(NumberText.cutToSignificantDigits rest, NumberStyles.Float, CultureInfo.InvariantCulture) with
        | true, num ->
            match op with
            | CriteriaEq -> equalsNumber num
            | CriteriaNe -> fun v -> not (equalsNumber num v)
            | _ ->
                fun v ->
                    match unwrapValue v with
                    | CellValue.Number x ->
                        let c = compareNumbers x num
                        match op with
                        | CriteriaGe -> c >= 0
                        | CriteriaLe -> c <= 0
                        | CriteriaGt -> c > 0
                        | _ -> c < 0
                    | _ -> false
        | _ ->
            match op with
            | CriteriaGt | CriteriaLt | CriteriaGe | CriteriaLe ->
                fun v ->
                    // Ordering against text compares text cells only.
                    match unwrapValue v with
                    | CellValue.String s ->
                        let c = String.Compare(s, rest, StringComparison.OrdinalIgnoreCase)
                        (match op with
                         | CriteriaGt -> c > 0
                         | CriteriaLt -> c < 0
                         | CriteriaGe -> c >= 0
                         | _ -> c <= 0)
                    | _ -> false
            | CriteriaNe ->
                if hasCriteriaWildcard rest then
                    let rx = wildcardRegex rest
                    fun v -> match coerceText v with Ok t -> not (rx.IsMatch t) | _ -> true
                else
                    fun v -> match criteriaTextCompare rest v with ValueSome c -> c <> 0 | ValueNone -> true
            | CriteriaEq ->
                if hasCriteriaWildcard rest then
                    let rx = wildcardRegex rest
                    fun v ->
                        // Wildcards match text; `*` does not match a number or a blank.
                        match unwrapValue v with
                        | CellValue.String t -> rx.IsMatch t
                        | _ -> false
                else
                    fun v -> match criteriaTextCompare rest v with ValueSome c -> c = 0 | ValueNone -> false
    | CellValue.Empty ->
        fun v -> match unwrapValue v with CellValue.Empty -> true | _ -> false
    | _ -> fun _ -> false

let private criteriaPredicate criteria =
    compileCriteria criteria

/// Static 0-indexed bounds of a range-shaped expression (RangeRef or a
/// single CellRef). None for anything else.
let private rangeBounds (e: Expr) : (int * int * int * int) option =
    match e with
    | RangeRef (a, b) -> Some (min a.Row b.Row, min a.Col b.Col, max a.Row b.Row, max a.Col b.Col)
    | CellRef r -> Some (r.Row, r.Col, r.Row, r.Col)
    | _ -> None

/// True if the expression tree contains a SUBTOTAL/AGGREGATE call — those
/// cells are skipped when an outer SUBTOTAL/AGGREGATE ranges over them, so
/// nested subtotals aren't double-counted (Excel's documented behaviour).
let exprUsesSubtotal (e: Expr) : bool =
    let stack = ResizeArray<Expr>()
    stack.Add e
    let mutable found = false
    while not found && stack.Count > 0 do
        let last = stack.Count - 1
        let expr = stack.[last]
        Portable.Stack.dropLast stack
        match expr with
        | Function (n, _) when equalsOrdinalIgnoreCase n "SUBTOTAL" || equalsOrdinalIgnoreCase n "AGGREGATE" ->
            found <- true
        | Function (_, args) ->
            for arg in args do
                stack.Add arg
        | UnaryOp (_, x) ->
            stack.Add x
        | BinOp (l, _, r) ->
            stack.Add r
            stack.Add l
        | _ -> ()
    found

let private formulaTextMayUseSubtotal (formula: string) =
    indexOfOrdinalIgnoreCase "SUBTOTAL" formula 0 >= 0
    || indexOfOrdinalIgnoreCase "AGGREGATE" formula 0 >= 0

let inline private compareInts (x: int) (y: int) =
    if x < y then -1
    elif x > y then 1
    else 0

let inline private compareBools (x: bool) (y: bool) =
    if x = y then 0
    elif x then 1
    else -1

let private cellValueRank v =
    match v with
    | CellValue.Number _ -> 0
    | CellValue.String _ -> 1
    | CellValue.Boolean _ -> 2
    | CellValue.Empty -> 0  // blank behaves as 0 / "" depending on the other side
    | _ -> 3

let inline private compareValuesNoErrorBy ([<InlineIfLambda>] numbers: float -> float -> int) (l: CellValue) (r: CellValue) : int voption =
    let left = unwrapValue l
    match left with
    | CellValue.Error _ -> ValueNone
    | CellValue.Number x ->
        let right = unwrapValue r
        match right with
        | CellValue.Error _ -> ValueNone
        | CellValue.Number y -> ValueSome (numbers x y)
        | CellValue.Empty -> ValueSome (numbers x 0.0)
        | _ -> ValueSome (compareInts (cellValueRank left) (cellValueRank right))
    | CellValue.String x ->
        let right = unwrapValue r
        match right with
        | CellValue.Error _ -> ValueNone
        | CellValue.String y -> ValueSome (String.Compare(x, y, StringComparison.OrdinalIgnoreCase))
        | CellValue.Empty -> ValueSome (String.Compare(x, "", StringComparison.OrdinalIgnoreCase))
        | _ -> ValueSome (compareInts (cellValueRank left) (cellValueRank right))
    | CellValue.Boolean x ->
        let right = unwrapValue r
        match right with
        | CellValue.Error _ -> ValueNone
        | CellValue.Boolean y -> ValueSome (compareBools x y)
        // A blank compares as FALSE against a boolean.
        | CellValue.Empty -> ValueSome (compareBools x false)
        | _ -> ValueSome (compareInts (cellValueRank left) (cellValueRank right))
    | CellValue.Empty ->
        let right = unwrapValue r
        match right with
        | CellValue.Error _ -> ValueNone
        | CellValue.Empty -> ValueSome 0
        | CellValue.Boolean y -> ValueSome (compareBools false y)
        | CellValue.Number y -> ValueSome (numbers 0.0 y)
        | CellValue.String y -> ValueSome (String.Compare("", y, StringComparison.OrdinalIgnoreCase))
        | _ -> ValueSome (compareInts (cellValueRank left) (cellValueRank right))
    | _ ->
        let right = unwrapValue r
        match right with
        | CellValue.Error _ -> ValueNone
        | _ -> ValueSome (compareInts (cellValueRank left) (cellValueRank right))

/// Two values compared as MATCH and the LOOKUPs compare them: numbers by their binary values.
let private compareValuesNoError (l: CellValue) (r: CellValue) : int voption =
    compareValuesNoErrorBy compareFloats l r

/// Two values compared as the comparison operators compare them: numbers to 15 significant
/// digits (`compareNumbers`).
let private compareOperandsNoError (l: CellValue) (r: CellValue) : int voption =
    compareValuesNoErrorBy compareNumbers l r

let private compareValuesForSort (l: CellValue) (r: CellValue) =
    let left = unwrapValue l
    let right = unwrapValue r
    match left with
    | CellValue.Error _ -> 0
    | CellValue.Number x ->
        match right with
        | CellValue.Error _ -> 0
        | CellValue.Number y -> compareFloats x y
        | CellValue.Empty -> compareFloats x 0.0
        | _ -> compareInts (cellValueRank left) (cellValueRank right)
    | CellValue.String x ->
        match right with
        | CellValue.Error _ -> 0
        | CellValue.String y -> String.Compare(x, y, StringComparison.OrdinalIgnoreCase)
        | CellValue.Empty -> String.Compare(x, "", StringComparison.OrdinalIgnoreCase)
        | _ -> compareInts (cellValueRank left) (cellValueRank right)
    | CellValue.Boolean x ->
        match right with
        | CellValue.Error _ -> 0
        | CellValue.Boolean y -> compareBools x y
        | _ -> compareInts (cellValueRank left) (cellValueRank right)
    | CellValue.Empty ->
        match right with
        | CellValue.Error _ -> 0
        | CellValue.Empty -> 0
        | CellValue.Number y -> compareFloats 0.0 y
        | CellValue.String y -> String.Compare("", y, StringComparison.OrdinalIgnoreCase)
        | _ -> compareInts (cellValueRank left) (cellValueRank right)
    | _ ->
        match right with
        | CellValue.Error _ -> 0
        | _ -> compareInts (cellValueRank left) (cellValueRank right)

let private valuesEqualNoError (l: CellValue) (r: CellValue) =
    let left = unwrapValue l
    match left with
    | CellValue.Error _ -> false
    | _ ->
        let right = unwrapValue r
        match right with
        | CellValue.Error _ -> false
        | _ -> compareValuesForSort left right = 0

/// Compare two scalar values as MATCH and the LOOKUPs do. Excel orders
/// number < text < FALSE < TRUE; numbers compare by their binary values, text
/// case-insensitively. Returns -1 / 0 / 1, or an Error value to propagate.
let private compareValues (l: CellValue) (r: CellValue) : Result<int, string> =
    match compareValuesNoError l r with
    | ValueSome result -> Ok result
    | ValueNone ->
        match unwrapValue l, unwrapValue r with
        | CellValue.Error e, _ -> Result.Error e
        | _, CellValue.Error e -> Result.Error e
        | _ -> Result.Error Err.value

/// `compareValues` with numbers equal to 15 significant digits (`compareNumbers`): the operators', SWITCH's.
let private compareOperands (l: CellValue) (r: CellValue) : Result<int, string> =
    match compareOperandsNoError l r with
    | ValueSome result -> Ok result
    | ValueNone ->
        match unwrapValue l, unwrapValue r with
        | CellValue.Error e, _ -> Result.Error e
        | _, CellValue.Error e -> Result.Error e
        | _ -> Result.Error Err.value

// ---------------------------------------------------------------------------
// Streaming aggregation — lazy pull `RangeView` + one-pass applicative folds.
//
// Aggregates used to materialize a `CellValue list` for every range on every
// eval (`SUM(A1:A100000)` built a 100k-element list). Instead a `RangeView` is
// a lazy rectangular window iterated by a STRUCT enumerator (no per-element
// IEnumerable boxing, no intermediate list), and a `Fold` is a one-pass
// reducer. `both` runs two folds over a single traversal — so AVERAGE is SUM
// and COUNT fused into one pass, and so are user pipelines like
// `SUM(r)+AVERAGE(r)`. Design + citations: docs/calc-streaming-fusion-design.md.
// ---------------------------------------------------------------------------

/// A lazy rectangular window over a sheet's sparse cell map. Iterating it
/// yields unwrapped `CellValue`s row-major (blanks as `CellValue.Empty`)
/// without building a list. Both the view and its enumerator are structs, so
/// `for v in view` and `run`/`foldArgs` allocate nothing per element.
[<Struct>]
type RangeView =
    { Model: SpreadsheetModel; SR: int; SC: int; ER: int; EC: int }
    member v.GetEnumerator() = RangeEnumerator(v)

and [<Struct>] RangeEnumerator =
    val mutable private R: int
    val mutable private C: int
    val mutable private Cur: CellValue
    val private View: RangeView
    new(v: RangeView) = { View = v; R = v.SR; C = v.SC - 1; Cur = CellValue.Empty }
    member e.Current = e.Cur
    member e.MoveNext() =
        if e.R > e.View.ER then false
        else
            e.C <- e.C + 1
            if e.C > e.View.EC then
                e.C <- e.View.SC
                e.R <- e.R + 1
            if e.R > e.View.ER then false
            else
                e.Cur <- getCellValue { Row = e.R; Col = e.C } e.View.Model |> unwrapValue
                true

/// A single-pass fold: thread `'S` through each element, then project to `'R`.
/// (init state, step, extract) — the applicative-fold shape from Control.Foldl.
[<Struct>]
type Fold<'T, 'S, 'R> = { Init: 'S; Step: 'S -> 'T -> 'S; Extract: 'S -> 'R }

/// Run a fold over a range view in one pass.
let inline run (view: RangeView) (f: Fold<CellValue, 'S, 'R>) : 'R =
    let mutable s = f.Init
    let mutable e = view.GetEnumerator()
    while e.MoveNext() do s <- f.Step s e.Current
    f.Extract s

/// Combine two folds so they run over a SINGLE traversal, sharing the walk.
/// State stays a struct tuple → no boxing. `both sumF countF` gives AVERAGE.
let inline both (a: Fold<'T, 'SA, 'RA>) (b: Fold<'T, 'SB, 'RB>) : Fold<'T, struct ('SA * 'SB), struct ('RA * 'RB)> =
    { Init = struct (a.Init, b.Init)
      Step = fun (struct (x, y)) t -> struct (a.Step x t, b.Step y t)
      Extract = fun (struct (x, y)) -> struct (a.Extract x, b.Extract y) }

/// Random-access view of a range: read the k-th cell (row-major) without
/// materializing a list. This is what the criteria/lookup family iterates so
/// `SUMIFS`/`MATCH`/`VLOOKUP` walk parallel ranges (and zip them) by index
/// rather than building a `CellValue list` per range. `Count` = rows × cols.
[<Struct>]
type RangeAccess = { M: SpreadsheetModel; SR: int; SC: int; W: int; Count: int }

let mutable private scratchArray : obj = null
let mutable private scratchModel : SpreadsheetModel = createEmpty 1 1

/// Resolve a range-shaped expression (RangeRef or single CellRef) to a
/// `RangeAccess`, honoring a sheet qualifier. None for non-range expressions.
let rangeAccessOf (model: SpreadsheetModel) (workbook: Rbt.Spreadsheet.Workbook.Workbook option) (e: Expr) : RangeAccess option =
    let mk sheet sr sc er ec =
        match resolveModel sheet model workbook with
        | Some m -> Some { M = m; SR = sr; SC = sc; W = ec - sc + 1; Count = (er - sr + 1) * (ec - sc + 1) }
        | None -> None
    match e with
    | RangeRef (a, b) -> let sr, sc, er, ec = normalizeRange a b in mk a.Sheet sr sc er ec
    | CellRef r -> mk r.Sheet r.Row r.Col r.Row r.Col
    | ArrayLit rows when rows.Length > 0 && rows.[0].Length > 0 ->
        // An array stands in for a range by being laid out on a scratch sheet,
        // so every positional reader (MATCH, VLOOKUP, XLOOKUP…) works on it
        // unchanged. The last array is remembered: one call reads it many times.
        let scratch =
            if obj.ReferenceEquals(scratchArray, rows) then scratchModel
            else
                let mutable cells = Map.empty
                for r in 0 .. rows.Length - 1 do
                    for c in 0 .. rows.[r].Length - 1 do
                        match unwrapValue rows.[r].[c] with
                        | CellValue.Empty -> ()
                        | value -> cells <- Map.add { Row = r; Col = c } { Value = value; Format = defaultFormat; Comment = None } cells
                let built = { createEmpty rows.Length rows.[0].Length with Cells = cells }
                scratchArray <- box rows
                scratchModel <- built
                built
        Some { M = scratch; SR = 0; SC = 0; W = rows.[0].Length; Count = rows.Length * rows.[0].Length }
    | _ -> None

/// Read the k-th cell of a range (row-major, 0-based), unwrapped.
let inline rangeRead (a: RangeAccess) (k: int) : CellValue =
    getCellValue { Row = a.SR + k / a.W; Col = a.SC + k % a.W } a.M |> unwrapValue

let private rangeEndForCount (a: RangeAccess) count =
    let last = max 0 (count - 1)
    struct (a.SR + last / a.W, a.SC + a.W - 1)

let inline private rangeOffsetWithinCount (a: RangeAccess) count er ec (addr: CellAddress) =
    if addr.Row >= a.SR && addr.Row <= er && addr.Col >= a.SC && addr.Col <= ec then
        let k = (addr.Row - a.SR) * a.W + (addr.Col - a.SC)
        if k >= 0 && k < count then ValueSome k else ValueNone
    else
        ValueNone

let private tryRangeOffsetForCount (a: RangeAccess) count (addr: CellAddress) =
    let struct (er, ec) = rangeEndForCount a count
    rangeOffsetWithinCount a count er ec addr

let private materializeRangeValuesForCount (source: RangeAccess) count =
    let count = min source.Count count
    let values = Array.create count CellValue.Empty
    if count > 0 then
        let struct (er, ec) = rangeEndForCount source count
        source.M.Cells
        |> Map.iter (fun addr cell ->
            match rangeOffsetWithinCount source count er ec addr with
            | ValueSome offset -> values.[offset] <- unwrapValue cell.Value
            | ValueNone -> ())
    values

let private materializeRangeValues (source: RangeAccess) =
    materializeRangeValuesForCount source source.Count

let private rangeReadCacheMaterializeAfter = 1024
let private rangeReadCacheMaxMaterializedValues = 200000

[<Sealed; AllowNullLiteral>]
type private RangeReadCache(source: RangeAccess, count: int, materializeAfter: int, maxMaterializedValues: int) =
    let count = min source.Count count
    let mutable reads = 0
    let mutable values: CellValue[] = null

    member _.Read(k: int) =
        if k < 0 || k >= count then
            CellValue.Empty
        elif isNull values then
            reads <- reads + 1
            if count > materializeAfter
               && count <= maxMaterializedValues
               && reads > materializeAfter then
                values <- materializeRangeValuesForCount source count
                values.[k]
            else
                rangeRead source k
        else
            values.[k]

let private rangeReadCached source count =
    RangeReadCache(source, count, rangeReadCacheMaterializeAfter, rangeReadCacheMaxMaterializedValues)

let private canSparseNonBlankScan (target: CellValue) (a: RangeAccess) =
    let area = int64 a.Count
    match compareValuesNoError target CellValue.Empty with
    | ValueSome 0 -> false
    | ValueSome _ -> area >= 1024L && int64 a.M.Cells.Count <= area * 4L
    | ValueNone -> false

let private countIfRangeAccess (a: RangeAccess) (pred: CellValue -> bool) =
    let area = int64 a.Count
    let predEmpty = pred CellValue.Empty
    if area >= 1024L && int64 a.M.Cells.Count <= area * 4L then
        let mutable storedInRange = 0L
        let mutable matchedStored = 0L
        let struct (er, ec) = rangeEndForCount a a.Count
        a.M.Cells
        |> Map.iter (fun addr cell ->
            if ValueOption.isSome (rangeOffsetWithinCount a a.Count er ec addr) then
                storedInRange <- storedInRange + 1L
                if pred (unwrapValue cell.Value) then
                    matchedStored <- matchedStored + 1L)
        matchedStored + if predEmpty then area - storedInRange else 0L
    else
        let mutable n = 0L
        let reader = rangeReadCached a a.Count
        for k in 0 .. a.Count - 1 do
            if pred (reader.Read k) then
                n <- n + 1L
        n

let private trySparseExactMatchResult (target: CellValue) (a: RangeAccess) searchMode =
    if not (canSparseNonBlankScan target a) then
        ValueNone
    else
        let reverse = searchMode = -1 || searchMode = -2
        let mutable found = ValueNone
        let considerOffset k =
            match found with
            | ValueNone -> found <- ValueSome k
            | ValueSome existing ->
                if (reverse && k > existing) || (not reverse && k < existing) then
                    found <- ValueSome k
        if reverse then
            a.M.Cells
            |> Map.iter (fun addr cell ->
                match tryRangeOffsetForCount a a.Count addr with
                | ValueSome k ->
                    if valuesEqualNoError target (unwrapValue cell.Value) then
                        considerOffset k
                | ValueNone -> ())
        else
            a.M.Cells
            |> Map.tryPick (fun addr cell ->
                match tryRangeOffsetForCount a a.Count addr with
                | ValueSome k when valuesEqualNoError target (unwrapValue cell.Value) -> Some k
                | _ -> None)
            |> Option.iter considerOffset
        match found with
        | ValueSome k -> ValueSome (Some k)
        | ValueNone -> ValueSome None

let private trySparseExactMatch (target: CellValue) (a: RangeAccess) searchMode =
    match trySparseExactMatchResult target a searchMode with
    | ValueSome (Some k) -> ValueSome k
    | _ -> ValueNone

let private trySparseClosestMatch (target: CellValue) (a: RangeAccess) matchMode searchMode =
    let area = int64 a.Count
    if matchMode <> -1 && matchMode <> 1 then
        ValueNone
    elif area < 1024L || int64 a.M.Cells.Count > area then
        ValueNone
    else
        let isOnSide value =
            match compareValuesNoError value target with
            | ValueSome c -> if matchMode = -1 then c < 0 else c > 0
            | ValueNone -> false
        let beats current candidate =
            match compareValuesNoError candidate current with
            | ValueSome c -> if matchMode = -1 then c > 0 else c < 0
            | ValueNone -> false
        let reverse = searchMode = -1 || searchMode = -2
        let preferOffset k existing =
            if reverse then k > existing else k < existing
        let mutable best = ValueNone
        a.M.Cells
        |> Map.iter (fun addr cell ->
            match tryRangeOffsetForCount a a.Count addr with
            | ValueSome k ->
                let value = unwrapValue cell.Value
                if isOnSide value then
                    match best with
                    | ValueNone -> best <- ValueSome (k, value)
                    | ValueSome (existingK, existingValue) ->
                        if beats existingValue value then
                            best <- ValueSome (k, value)
                        elif valuesEqualNoError value existingValue && preferOffset k existingK then
                            best <- ValueSome (k, value)
            | ValueNone -> ())
        match best with
        | ValueSome (k, value) ->
            if isOnSide CellValue.Empty then
                match compareValuesNoError value CellValue.Empty with
                | ValueSome c when (matchMode = -1 && c > 0) || (matchMode = 1 && c < 0) -> ValueSome k
                | _ -> ValueNone
            else
                ValueSome k
        | ValueNone -> ValueNone

/// SUMIF's total: an error among the cells it adds is the result (the first of them, in the range's
/// order), as Excel's.
let private sumIfRangeAccess (crit: RangeAccess) (sum: RangeAccess) (pred: CellValue -> bool) : CellValue =
    let count = min crit.Count sum.Count
    let area = int64 count
    let mutable total = 0.0
    // (Offset and code of the errors met; in the first way below, the ones whose criteria cell
    // fails are taken out again.)
    let errors = System.Collections.Generic.List<struct (int * string)>()
    let addNumberAt (k: int) value =
        match unwrapValue value with
        | CellValue.Error e -> errors.Add(struct (k, e))
        | value ->
            match tryAggregateNumber value with
            | ValueSome n -> total <- total + n
            | ValueNone -> ()
    let subtractNumber value =
        match tryAggregateNumber value with
        | ValueSome n -> total <- total - n
        | ValueNone -> ()
    let sumReader = rangeReadCached sum count
    let critReader = rangeReadCached crit count
    if area >= 1024L && pred CellValue.Empty && (int64 crit.M.Cells.Count <= area * 4L || int64 sum.M.Cells.Count <= area * 4L) then
        let excluded = System.Collections.Generic.HashSet<int>()
        let struct (sumEr, sumEc) = rangeEndForCount sum count
        let struct (critEr, critEc) = rangeEndForCount crit count
        if obj.ReferenceEquals(crit.M.Cells, sum.M.Cells) then
            sum.M.Cells
            |> Map.iter (fun addr cell ->
                match rangeOffsetWithinCount sum count sumEr sumEc addr with
                | ValueSome k -> addNumberAt k (unwrapValue cell.Value)
                | ValueNone -> ()
                match rangeOffsetWithinCount crit count critEr critEc addr with
                | ValueSome k when not (pred (unwrapValue cell.Value)) -> excluded.Add k |> ignore
                | _ -> ())
        else
            sum.M.Cells
            |> Map.iter (fun addr cell ->
                match rangeOffsetWithinCount sum count sumEr sumEc addr with
                | ValueSome k -> addNumberAt k (unwrapValue cell.Value)
                | ValueNone -> ())
            crit.M.Cells
            |> Map.iter (fun addr cell ->
                match rangeOffsetWithinCount crit count critEr critEc addr with
                | ValueSome k when not (pred (unwrapValue cell.Value)) -> excluded.Add k |> ignore
                | _ -> ())
        for k in excluded do
            if k < sum.Count then
                subtractNumber (sumReader.Read k)
        errors.RemoveAll(fun (struct (k, _)) -> excluded.Contains k) |> ignore
    elif area >= 1024L && int64 crit.M.Cells.Count <= area * 4L then
        crit.M.Cells
        |> Map.iter (fun addr cell ->
            match tryRangeOffsetForCount crit count addr with
            | ValueSome k when pred (unwrapValue cell.Value) -> addNumberAt k (sumReader.Read k)
            | _ -> ())
    else
        for k in 0 .. count - 1 do
            if pred (critReader.Read k) then
                addNumberAt k (sumReader.Read k)
    if errors.Count = 0 then CellValue.Number total
    else
        let mutable first = errors.[0]
        for e in errors do
            let struct (k, _) = e
            let struct (k0, _) = first
            if k < k0 then first <- e
        let struct (_, code) = first
        CellValue.Error code

let private averageIfRangeAccess (crit: RangeAccess) (avg: RangeAccess) (pred: CellValue -> bool) =
    let count = min crit.Count avg.Count
    let area = int64 count
    let mutable total = 0.0
    let mutable n = 0
    let addNumber value =
        match tryAggregateNumber value with
        | ValueSome x ->
            total <- total + x
            n <- n + 1
        | ValueNone -> ()
    let subtractNumber value =
        match tryAggregateNumber value with
        | ValueSome x ->
            total <- total - x
            n <- n - 1
        | ValueNone -> ()
    let avgReader = rangeReadCached avg count
    let critReader = rangeReadCached crit count
    if area >= 1024L && not (pred CellValue.Empty) && int64 crit.M.Cells.Count <= area * 4L then
        crit.M.Cells
        |> Map.iter (fun addr cell ->
            match tryRangeOffsetForCount crit count addr with
            | ValueSome k when pred (unwrapValue cell.Value) -> addNumber (avgReader.Read k)
            | _ -> ())
    elif area >= 1024L && pred CellValue.Empty && (int64 crit.M.Cells.Count <= area * 4L || int64 avg.M.Cells.Count <= area * 4L) then
        let excluded = System.Collections.Generic.HashSet<int>()
        let struct (avgEr, avgEc) = rangeEndForCount avg count
        let struct (critEr, critEc) = rangeEndForCount crit count
        if obj.ReferenceEquals(crit.M.Cells, avg.M.Cells) then
            avg.M.Cells
            |> Map.iter (fun addr cell ->
                match rangeOffsetWithinCount avg count avgEr avgEc addr with
                | ValueSome _ -> addNumber (unwrapValue cell.Value)
                | ValueNone -> ()
                match rangeOffsetWithinCount crit count critEr critEc addr with
                | ValueSome k when not (pred (unwrapValue cell.Value)) -> excluded.Add k |> ignore
                | _ -> ())
        else
            avg.M.Cells
            |> Map.iter (fun addr cell ->
                match rangeOffsetWithinCount avg count avgEr avgEc addr with
                | ValueSome _ -> addNumber (unwrapValue cell.Value)
                | ValueNone -> ())
            crit.M.Cells
            |> Map.iter (fun addr cell ->
                match rangeOffsetWithinCount crit count critEr critEc addr with
                | ValueSome k when not (pred (unwrapValue cell.Value)) -> excluded.Add k |> ignore
                | _ -> ())
        for k in excluded do
            if k < avg.Count then
                subtractNumber (avgReader.Read k)
    else
        for k in 0 .. count - 1 do
            if pred (critReader.Read k) then
                addNumber (avgReader.Read k)
    if n = 0 then err Err.div0 else CellValue.Number (total / float n)

let private chooseSparseIfsDriver (crits: ResizeArray<RangeAccess * (CellValue -> bool)>) count =
    let area = int64 count
    let mutable driverIndex = -1
    let mutable driverStoredCount = System.Int32.MaxValue
    let mutable eligibleCount = 0
    if area >= 1024L then
        for i = 0 to crits.Count - 1 do
            let access, pred = crits.[i]
            if not (pred CellValue.Empty) && int64 access.M.Cells.Count <= area * 4L && access.M.Cells.Count < driverStoredCount then
                driverIndex <- i
                driverStoredCount <- access.M.Cells.Count
                eligibleCount <- eligibleCount + 1
            elif not (pred CellValue.Empty) && int64 access.M.Cells.Count <= area * 4L then
                eligibleCount <- eligibleCount + 1
    if driverIndex >= 0 && eligibleCount > 1 then
        let mutable bestIndex = driverIndex
        let mutable bestMatches = System.Int32.MaxValue
        let mutable bestStoredCount = System.Int32.MaxValue
        for i = 0 to crits.Count - 1 do
            let access, pred = crits.[i]
            if not (pred CellValue.Empty) && int64 access.M.Cells.Count <= area * 4L then
                let mutable matches = 0
                let struct (er, ec) = rangeEndForCount access count
                access.M.Cells
                |> Map.iter (fun addr cell ->
                    match rangeOffsetWithinCount access count er ec addr with
                    | ValueSome _ when pred (unwrapValue cell.Value) -> matches <- matches + 1
                    | _ -> ())
                if matches < bestMatches || (matches = bestMatches && access.M.Cells.Count < bestStoredCount) then
                    bestIndex <- i
                    bestMatches <- matches
                    bestStoredCount <- access.M.Cells.Count
        bestIndex
    else
        driverIndex

let private trySparseIfs (crits: ResizeArray<RangeAccess * (CellValue -> bool)>) count (sumAcc: RangeAccess option) =
    let driverIndex = chooseSparseIfsDriver crits count
    if driverIndex < 0 then
        ValueNone
    else
        let driverAccess, driverPred = crits.[driverIndex]
        let mutable total = 0.0
        let mutable hits = 0
        let mutable sawError = false
        let critReaders =
            Array.init crits.Count (fun i ->
                let access, pred = crits.[i]
                if pred CellValue.Empty then null else rangeReadCached access count)
        let sumReader =
            match sumAcc with
            | Some s -> rangeReadCached s count
            | None -> null
        let struct (driverEr, driverEc) = rangeEndForCount driverAccess count
        driverAccess.M.Cells
        |> Map.iter (fun addr cell ->
            match rangeOffsetWithinCount driverAccess count driverEr driverEc addr with
            | ValueSome k when driverPred (unwrapValue cell.Value) ->
                let mutable matched = true
                let mutable critIndex = 0
                while matched && critIndex < crits.Count do
                    if critIndex <> driverIndex then
                        let access, pred = crits.[critIndex]
                        let reader = critReaders.[critIndex]
                        let value = if isNull reader then rangeRead access k else reader.Read k
                        if not (pred value) then
                            matched <- false
                    critIndex <- critIndex + 1
                if matched then
                    hits <- hits + 1
                    match sumAcc with
                    | Some s when k < s.Count ->
                        let value = sumReader.Read k
                        match tryAggregateNumber value with
                        | ValueSome n -> total <- total + n
                        | ValueNone -> (match unwrapValue value with CellValue.Error _ -> sawError <- true | _ -> ())
                    | _ -> ()
            | _ -> ())
        // (An error among the cells to add: the full pass says which.)
        if sawError then ValueNone else ValueSome (total, hits)

let private sparseBlankExcludedOffsets (crits: ResizeArray<RangeAccess * (CellValue -> bool)>) count =
    let excluded = System.Collections.Generic.HashSet<int>()
    if crits.Count = 0 then
        excluded
    else
        let firstAccess, _ = crits.[0]
        let accessArray = Array.zeroCreate<RangeAccess> crits.Count
        let predArray = Array.zeroCreate<CellValue -> bool> crits.Count
        let endRows = Array.zeroCreate<int> crits.Count
        let endCols = Array.zeroCreate<int> crits.Count
        let mutable sameModel = true
        let mutable i = 0
        while i < crits.Count do
            let access, pred = crits.[i]
            accessArray.[i] <- access
            predArray.[i] <- pred
            let struct (er, ec) = rangeEndForCount access count
            endRows.[i] <- er
            endCols.[i] <- ec
            if i > 0 && not (obj.ReferenceEquals(access.M.Cells, firstAccess.M.Cells)) then
                sameModel <- false
            i <- i + 1
        let inline addExcludedIfNeeded critIndex (addr: CellAddress) (value: CellValue) =
            let access = accessArray.[critIndex]
            if addr.Row >= access.SR && addr.Row <= endRows.[critIndex] && addr.Col >= access.SC && addr.Col <= endCols.[critIndex] then
                let k = (addr.Row - access.SR) * access.W + (addr.Col - access.SC)
                if k >= 0 && k < count && not (predArray.[critIndex] value) then
                    excluded.Add k |> ignore
        if sameModel then
            firstAccess.M.Cells
            |> Map.iter (fun addr cell ->
                let value = unwrapValue cell.Value
                let mutable critIndex = 0
                while critIndex < crits.Count do
                    addExcludedIfNeeded critIndex addr value
                    critIndex <- critIndex + 1)
        else
            for critIndex = 0 to crits.Count - 1 do
                let access = accessArray.[critIndex]
                access.M.Cells
                |> Map.iter (fun addr cell ->
                    addExcludedIfNeeded critIndex addr (unwrapValue cell.Value))
        excluded

let private trySparseBlankIfs (crits: ResizeArray<RangeAccess * (CellValue -> bool)>) count (sumAcc: RangeAccess option) =
    let area = int64 count
    let mutable eligible = area >= 1024L && crits.Count > 0
    let mutable i = 0
    while eligible && i < crits.Count do
        let access, pred = crits.[i]
        if not (pred CellValue.Empty) || int64 access.M.Cells.Count > area * 4L then
            eligible <- false
        i <- i + 1
    if not eligible then
        ValueNone
    else
        let excluded = sparseBlankExcludedOffsets crits count
        let hits = count - excluded.Count
        match sumAcc with
        | Some s ->
            let mutable total = 0.0
            let mutable sawError = false
            let sumReader = rangeReadCached s count
            let subtractNumber value =
                match tryAggregateNumber value with
                | ValueSome n -> total <- total - n
                | ValueNone -> ()
            let struct (er, ec) = rangeEndForCount s count
            s.M.Cells
            |> Map.iter (fun addr cell ->
                match rangeOffsetWithinCount s count er ec addr with
                | ValueSome k ->
                    match tryAggregateNumber (unwrapValue cell.Value) with
                    | ValueSome n -> total <- total + n
                    | ValueNone -> (match unwrapValue cell.Value with CellValue.Error _ when not (excluded.Contains k) -> sawError <- true | _ -> ())
                | _ -> ())
            for k in excluded do
                if k < s.Count then
                    subtractNumber (sumReader.Read k)
            // (An error among the cells to add: the full pass says which.)
            if sawError then ValueNone else ValueSome (total, hits)
        | None ->
            ValueSome (0.0, hits)

let private trySparseValueIfs fn (valueAccess: RangeAccess) (crits: ResizeArray<RangeAccess * (CellValue -> bool)>) count =
    let driverIndex = chooseSparseIfsDriver crits count
    if driverIndex < 0 then
        ValueNone
    else
        let driverAccess, driverPred = crits.[driverIndex]
        let mutable n = 0
        let mutable acc = if fn = "MINIFS" then infinity elif fn = "MAXIFS" then -infinity else 0.0
        let valueReader = rangeReadCached valueAccess count
        let critReaders =
            Array.init crits.Count (fun i ->
                let access, pred = crits.[i]
                if pred CellValue.Empty then null else rangeReadCached access count)
        let struct (driverEr, driverEc) = rangeEndForCount driverAccess count
        driverAccess.M.Cells
        |> Map.iter (fun addr cell ->
            match rangeOffsetWithinCount driverAccess count driverEr driverEc addr with
            | ValueSome k when driverPred (unwrapValue cell.Value) ->
                let mutable matched = true
                let mutable critIndex = 0
                while matched && critIndex < crits.Count do
                    if critIndex <> driverIndex then
                        let access, pred = crits.[critIndex]
                        let reader = critReaders.[critIndex]
                        let value = if isNull reader then rangeRead access k else reader.Read k
                        if not (pred value) then
                            matched <- false
                    critIndex <- critIndex + 1
                if matched then
                    match tryAggregateNumber (valueReader.Read k) with
                    | ValueSome x ->
                        n <- n + 1
                        acc <- match fn with "MINIFS" -> min acc x | "MAXIFS" -> max acc x | _ -> acc + x
                    | ValueNone -> ()
            | _ -> ())
        ValueSome (n, acc)

let private trySparseBlankValueIfs fn (valueAccess: RangeAccess) (crits: ResizeArray<RangeAccess * (CellValue -> bool)>) count =
    let area = int64 count
    let mutable eligible = area >= 1024L && crits.Count > 0
    let mutable i = 0
    while eligible && i < crits.Count do
        let access, pred = crits.[i]
        if not (pred CellValue.Empty) || int64 access.M.Cells.Count > area * 4L then
            eligible <- false
        i <- i + 1
    if not eligible then
        ValueNone
    else
        let excluded = sparseBlankExcludedOffsets crits count
        let mutable n = 0
        let mutable acc = if fn = "MINIFS" then infinity elif fn = "MAXIFS" then -infinity else 0.0
        let valueReader = rangeReadCached valueAccess count
        let struct (er, ec) = rangeEndForCount valueAccess count
        if fn = "AVERAGEIFS" then
            valueAccess.M.Cells
            |> Map.iter (fun addr cell ->
                match rangeOffsetWithinCount valueAccess count er ec addr with
                | ValueSome _ ->
                    match tryAggregateNumber (unwrapValue cell.Value) with
                    | ValueSome x ->
                        n <- n + 1
                        acc <- acc + x
                    | ValueNone -> ()
                | _ -> ())
            for k in excluded do
                if k < valueAccess.Count then
                    match tryAggregateNumber (valueReader.Read k) with
                    | ValueSome x ->
                        n <- n - 1
                        acc <- acc - x
                    | ValueNone -> ()
        else
            valueAccess.M.Cells
            |> Map.iter (fun addr cell ->
                match rangeOffsetWithinCount valueAccess count er ec addr with
                | ValueSome k when not (excluded.Contains k) ->
                    match tryAggregateNumber (unwrapValue cell.Value) with
                    | ValueSome x ->
                        n <- n + 1
                        acc <- match fn with "MINIFS" -> min acc x | "MAXIFS" -> max acc x | _ -> acc + x
                    | ValueNone -> ()
                | _ -> ())
        ValueSome (n, acc)

/// Coerce a value typed directly as an aggregate's argument: booleans and
/// numeric text become numbers, any other text is #VALUE!.
let private directNumericArgument (v: CellValue) : CellValue =
    match v with
    | CellValue.Boolean flag -> CellValue.Number (if flag then 1.0 else 0.0)
    | CellValue.String text ->
        let struct (start, length) = trimBounds text
        (match Double.TryParse(NumberText.cutToSignificantDigits (trimmedNumberText text start length), NumberStyles.Float, CultureInfo.InvariantCulture) with
         | true, n when length > 0 -> CellValue.Number n
         | _ -> CellValue.Error Err.value)
    | _ -> v

/// COUNT's version: the same coercion, but text that is not a number is
/// simply not counted.
let private directCountArgument (v: CellValue) : CellValue =
    match directNumericArgument v with
    | CellValue.Error _ when (match v with CellValue.String _ -> true | _ -> false) -> v
    | coerced -> coerced

/// Numeric reducer over already-unwrapped values: numbers and booleans fold
/// via `(z, op)`; text/blanks are skipped; the first `Error` short-circuits
/// and propagates. `project count acc` finishes — `count` is the number of
/// numeric values seen, so it can special-case the empty range or divide for
/// an average. State is `(first-error?, count, accumulator)` as a struct.
let numAggFold (z: float) (op: float -> float -> float) (project: int -> float -> CellValue)
               : Fold<CellValue, struct (string voption * int * float), CellValue> =
    { Init = struct (ValueNone, 0, z)
      Step = fun (struct (e, n, a)) v ->
        match e with
        | ValueSome _ -> struct (e, n, a)
        | ValueNone ->
            match v with
            | CellValue.Error code -> struct (ValueSome code, n, a)
            | CellValue.Number x -> struct (ValueNone, n + 1, op a x)
            | CellValue.Boolean b -> struct (ValueNone, n + 1, op a (if b then 1.0 else 0.0))
            | _ -> struct (e, n, a)
      Extract = fun (struct (e, n, a)) ->
        match e with
        | ValueSome code -> CellValue.Error code
        | ValueNone -> project n a }

let varianceFold sample sqrtResult : Fold<CellValue, struct (string voption * int * float * float), CellValue> =
    { Init = struct (ValueNone, 0, 0.0, 0.0)
      Step = fun (struct (e, n, mean, m2)) v ->
        match e with
        | ValueSome _ -> struct (e, n, mean, m2)
        | ValueNone ->
            match v with
            | CellValue.Error code -> struct (ValueSome code, n, mean, m2)
            | CellValue.Number x ->
                let n' = n + 1
                let delta = x - mean
                let mean' = mean + delta / float n'
                let m2' = m2 + delta * (x - mean')
                struct (ValueNone, n', mean', m2')
            | CellValue.Boolean b ->
                let x = if b then 1.0 else 0.0
                let n' = n + 1
                let delta = x - mean
                let mean' = mean + delta / float n'
                let m2' = m2 + delta * (x - mean')
                struct (ValueNone, n', mean', m2')
            | _ -> struct (ValueNone, n, mean, m2)
      Extract = fun (struct (e, n, _, m2)) ->
        match e with
        | ValueSome code -> CellValue.Error code
        | ValueNone ->
            if (sample && n < 2) || n < 1 then CellValue.Error Err.div0
            else
                let variance = m2 / (if sample then float (n - 1) else float n)
                CellValue.Number (if sqrtResult then sqrt variance else variance) }

let devSqFold : Fold<CellValue, struct (string voption * int * float * float), CellValue> =
    { Init = struct (ValueNone, 0, 0.0, 0.0)
      Step = fun (struct (firstErr, count, mean, m2)) value ->
          match firstErr with
          | ValueSome _ -> struct (firstErr, count, mean, m2)
          | ValueNone ->
              match unwrapValue value with
              | CellValue.Error e -> struct (ValueSome e, count, mean, m2)
              | CellValue.Number n ->
                  let count = count + 1
                  let delta = n - mean
                  let mean = mean + delta / float count
                  let m2 = m2 + delta * (n - mean)
                  struct (ValueNone, count, mean, m2)
              | CellValue.Boolean b ->
                  let n = if b then 1.0 else 0.0
                  let count = count + 1
                  let delta = n - mean
                  let mean = mean + delta / float count
                  let m2 = m2 + delta * (n - mean)
                  struct (ValueNone, count, mean, m2)
              | _ -> struct (ValueNone, count, mean, m2)
      Extract = fun (struct (firstErr, count, _, m2)) ->
          match firstErr with
          | ValueSome e -> err e
          | ValueNone -> if count = 0 then err Err.num else CellValue.Number m2 }

let aveDevFold () : Fold<CellValue, struct (string voption * ResizeArray<float> * float), CellValue> =
    { Init = struct (ValueNone, ResizeArray<float>(), 0.0)
      Step = fun (struct (firstErr, values, sum)) value ->
          match firstErr with
          | ValueSome _ -> struct (firstErr, values, sum)
          | ValueNone ->
              let addNumber n =
                  values.Add n
                  struct (ValueNone, values, sum + n)
              match unwrapValue value with
              | CellValue.Error e -> struct (ValueSome e, values, sum)
              | CellValue.Number n -> addNumber n
              | CellValue.Boolean b -> addNumber (if b then 1.0 else 0.0)
              | _ -> struct (ValueNone, values, sum)
      Extract = fun (struct (firstErr, values, sum)) ->
          match firstErr with
          | ValueSome e -> err e
          | ValueNone ->
              if values.Count = 0 then err Err.num
              else
                  let mean = sum / float values.Count
                  let mutable totalDeviation = 0.0
                  for i = 0 to values.Count - 1 do
                      totalDeviation <- totalDeviation + abs (values.[i] - mean)
                  CellValue.Number (totalDeviation / float values.Count) }

let rec private gcdInt64 a b =
    if b = 0L then a else gcdInt64 b (a % b)

let integerAggFold (z: int64) (op: int64 -> int64 -> int64)
                   : Fold<CellValue, struct (string voption * int64), CellValue> =
    { Init = struct (ValueNone, z)
      Step = fun (struct (firstErr, acc)) value ->
          match firstErr with
          | ValueSome _ -> struct (firstErr, acc)
          | ValueNone ->
              match value with
              | CellValue.Error e -> struct (ValueSome e, acc)
              | CellValue.Number n -> struct (ValueNone, op acc (int64 (floor (abs n))))
              | CellValue.Boolean b -> struct (ValueNone, op acc (if b then 1L else 0L))
              | _ -> struct (ValueNone, acc)
      Extract = fun (struct (firstErr, acc)) ->
          match firstErr with
          | ValueSome e -> err e
          | ValueNone -> CellValue.Number (float acc) }

let modeFold () : Fold<CellValue, struct (string voption * Dictionary<float, struct (int * int)> * int * float * int * int), CellValue> =
    { Init = struct (ValueNone, Dictionary<float, struct (int * int)>(), 0, 0.0, 0, System.Int32.MaxValue)
      Step = fun (struct (firstErr, counts, index, bestValue, bestCount, bestFirstIndex)) value ->
          match firstErr with
          | ValueSome _ -> struct (firstErr, counts, index, bestValue, bestCount, bestFirstIndex)
          | ValueNone ->
              let stepNumber n =
                  let mutable existing = Unchecked.defaultof<struct (int * int)>
                  let count, firstIndex =
                      if counts.TryGetValue(n, &existing) then
                          let struct (existingCount, existingFirstIndex) = existing
                          existingCount + 1, existingFirstIndex
                      else
                          1, index
                  counts.[n] <- struct (count, firstIndex)
                  let bestValue, bestCount, bestFirstIndex =
                      if count > 1 && (count > bestCount || (count = bestCount && firstIndex < bestFirstIndex)) then
                          n, count, firstIndex
                      else
                          bestValue, bestCount, bestFirstIndex
                  struct (ValueNone, counts, index + 1, bestValue, bestCount, bestFirstIndex)
              match value with
              | CellValue.Error e -> struct (ValueSome e, counts, index, bestValue, bestCount, bestFirstIndex)
              | CellValue.Number n -> stepNumber n
              | CellValue.Boolean b -> stepNumber (if b then 1.0 else 0.0)
              | _ -> struct (ValueNone, counts, index, bestValue, bestCount, bestFirstIndex)
      Extract = fun (struct (firstErr, _, _, bestValue, bestCount, _)) ->
          match firstErr with
          | ValueSome e -> err e
          | ValueNone ->
              if bestCount > 1 then CellValue.Number bestValue else err Err.na }

let rankFold (target: float) (descending: bool) : Fold<CellValue, struct (string voption * bool * int), CellValue> =
    { Init = struct (ValueNone, false, 1)
      Step = fun (struct (firstErr, found, rank)) value ->
          match firstErr with
          | ValueSome _ -> struct (firstErr, found, rank)
          | ValueNone ->
              match unwrapValue value with
              | CellValue.Error e -> struct (ValueSome e, found, rank)
              | CellValue.Number n ->
                  let found = found || n = target
                  let rank = if (if descending then n > target else n < target) then rank + 1 else rank
                  struct (ValueNone, found, rank)
              | CellValue.Boolean b ->
                  let n = if b then 1.0 else 0.0
                  let found = found || n = target
                  let rank = if (if descending then n > target else n < target) then rank + 1 else rank
                  struct (ValueNone, found, rank)
              | _ -> struct (ValueNone, found, rank)
      Extract = fun (struct (firstErr, found, rank)) ->
          match firstErr with
          | ValueSome e -> err e
          | ValueNone ->
              if found then CellValue.Number (float rank) else err Err.na }

let npvFold (rate: float) : Fold<CellValue, struct (string voption * int * float), CellValue> =
    { Init = struct (ValueNone, 0, 0.0)
      Step = fun (struct (firstErr, index, total)) value ->
          match firstErr with
          | ValueSome _ -> struct (firstErr, index, total)
          | ValueNone ->
              match unwrapValue value with
              | CellValue.Error e -> struct (ValueSome e, index, total)
              | CellValue.Number n ->
                  let index = index + 1
                  struct (ValueNone, index, total + n / ((1.0 + rate) ** float index))
              | CellValue.Boolean b ->
                  let n = if b then 1.0 else 0.0
                  let index = index + 1
                  struct (ValueNone, index, total + n / ((1.0 + rate) ** float index))
              | _ -> struct (ValueNone, index, total)
      Extract = fun (struct (firstErr, _, total)) ->
          match firstErr with
          | ValueSome e -> err e
          | ValueNone -> CellValue.Number total }

/// Counting reducer. `pred` selects which (unwrapped) values count;
/// `propagateErr` returns the first `Error` instead of a count (COUNT's
/// behavior — COUNTA/COUNTBLANK ignore errors).
let countFold (propagateErr: bool) (pred: CellValue -> bool)
              : Fold<CellValue, struct (string voption * int), CellValue> =
    { Init = struct (ValueNone, 0)
      Step = fun (struct (e, n)) v ->
        match e with
        | ValueSome _ -> struct (e, n)
        | ValueNone ->
            match v with
            | CellValue.Error code when propagateErr -> struct (ValueSome code, n)
            | _ -> struct (ValueNone, (if pred v then n + 1 else n))
      Extract = fun (struct (e, n)) ->
        match e with
        | ValueSome code -> CellValue.Error code
        | ValueNone -> CellValue.Number (float n) }

let aVariantFold (fn: string) : Fold<CellValue, struct (int * float * float * float), CellValue> =
    { Init = struct (0, 0.0, System.Double.PositiveInfinity, System.Double.NegativeInfinity)
      Step = fun (struct (count, sum, minValue, maxValue)) value ->
          let numeric =
              match unwrapValue value with
              | CellValue.Number n -> Some n
              | CellValue.Boolean b -> Some (if b then 1.0 else 0.0)
              | CellValue.String _ -> Some 0.0
              | _ -> None
          match numeric with
          | Some n -> struct (count + 1, sum + n, min minValue n, max maxValue n)
          | None -> struct (count, sum, minValue, maxValue)
      Extract = fun (struct (count, sum, minValue, maxValue)) ->
          if count = 0 then
              if fn = "AVERAGEA" then err Err.div0 else CellValue.Number 0.0
          else
              match fn with
              | "AVERAGEA" -> CellValue.Number (sum / float count)
              | "MAXA" -> CellValue.Number maxValue
              | _ -> CellValue.Number minValue }

let logicalFold (fn: string) : Fold<CellValue, struct (bool * bool * int * string option), CellValue> =
    let initialResult = if fn = "AND" then true else false
    { Init = struct (false, initialResult, 0, None)
      Step = fun (struct (seen, result, trueCount, firstError)) value ->
          match unwrapValue value with
          | CellValue.Empty -> struct (seen, result, trueCount, firstError)
          | other ->
              match coerceBool other with
              | Ok b ->
                  match fn with
                  | "AND" -> struct (true, result && b, trueCount, firstError)
                  | "OR" -> struct (true, result || b, trueCount, firstError)
                  | _ -> struct (true, result, (if b then trueCount + 1 else trueCount), firstError)
              | Result.Error e ->
                  let firstError =
                      match firstError with
                      | Some _ -> firstError
                      | None -> Some e
                  struct (seen, result, trueCount, firstError)
      Extract = fun (struct (seen, result, trueCount, firstError)) ->
          match firstError with
          | Some e -> err e
          | None when not seen -> err Err.value
          | None ->
              match fn with
              | "XOR" -> CellValue.Boolean (trueCount % 2 = 1)
              | _ -> CellValue.Boolean result }

// ---------------------------------------------------------------------------
// Dynamic arrays. `ArrVal` is a 2-D value (rows of columns) produced by the
// spilling functions (SEQUENCE/UNIQUE/SORT/FILTER/…) and by element-wise
// broadcasting of operators over ranges (`A1:A10>5`). Step 1: arrays flow
// THROUGH the engine — consumable by aggregates, collapsing to their top-left
// element in scalar context. Spill *placement* into neighbouring cells +
// rendering is a separate step. See docs/calc-streaming-fusion-design.md.
// ---------------------------------------------------------------------------

/// A rectangular array value: `arr.[row].[col]`, outer = rows.
type ArrVal = CellValue[][]

/// Functions that produce an array (a spill). Detected up front so the scalar
/// evaluator can collapse them and aggregates can consume the whole array.
let private arrayFunctionNames =
    [| "SEQUENCE"; "RANDARRAY"; "UNIQUE"; "SORT"; "SORTBY"; "FILTER"; "TRANSPOSE"
       "MAP"; "REDUCE"; "SCAN"; "MAKEARRAY"; "TEXTSPLIT"; "GROUPBY"; "PIVOTBY"
       "BYROW"; "BYCOL"
       "VSTACK"; "HSTACK"; "TAKE"; "DROP"; "CHOOSECOLS"; "CHOOSEROWS"; "TOCOL"; "TOROW"
       "MMULT"; "MINVERSE"; "FREQUENCY" |]
    // (And those of the modules and of the providers that give one.)
    |> Array.append [| for entry in Rbt.Spreadsheet.FunctionCatalog.entries do if entry.Flags &&& 2 <> 0 then yield entry.Name
                       for (name, _, gives) in providerFunctionSpecs do if gives then yield name |]

#if FABLE_COMPILER
/// (Asked for every function call of every formula: what was found for a name, as it is written, is kept.)
let private arrayFunctionKnown = Dictionary<string, bool>()
let isArrayFunction (name: string) =
    match Portable.Dict.tryFind name arrayFunctionKnown with
    | ValueSome known -> known
    | ValueNone ->
        let found = arrayFunctionNames |> Array.exists (equalsOrdinalIgnoreCase name)
        if arrayFunctionKnown.Count > 4096 then arrayFunctionKnown.Clear()
        arrayFunctionKnown.[name] <- found
        found
#else
let private arrayFunctionSet = HashSet<string>(arrayFunctionNames, StringComparer.OrdinalIgnoreCase)
let isArrayFunction (name: string) = arrayFunctionSet.Contains name
#endif

let private arrayFunctionNamesUpper = arrayFunctionNames |> Array.map (fun name -> name.ToUpperInvariant())

let private formulaTextMayContainArrayFunctionName (formula: string) =
    if isNull formula then
        false
    else
        // Upper-cased once, then the platform's own search for each name.
        let upper = formula.ToUpperInvariant()
        let mutable found = false
        let mutable index = 0
        while not found && index < arrayFunctionNamesUpper.Length do
            if upper.IndexOf(arrayFunctionNamesUpper.[index], StringComparison.Ordinal) >= 0 then
                found <- true
            index <- index + 1
        found

/// How a function takes one of its arguments.
type ParamKind =
    /// One value. A range or array given here makes the whole call an array
    /// formula: the function is applied element by element (Excel's lifting).
    | ScalarParam
    /// A reference or array consumed whole — an aggregate's operands, a
    /// criteria range, a reference whose shape is asked about.
    | RangeParam
    /// A table that is read by position and may be a reference *or* any
    /// expression yielding an array (`VLOOKUP(2,{1,"a";2,"b"},2,0)`): a
    /// non-reference argument is evaluated to an array before the call.
    | ArrayParam

/// A function's parameter kinds: `Fixed` for the leading arguments, then
/// `Repeat` cycled for the rest (`SUMIFS(sum, range, criteria, …)` is
/// Fixed = [Range], Repeat = [Range; Scalar]).
type FunctionSignature = { Fixed: ParamKind[]; Repeat: ParamKind[] }

/// The signature of every scalar-valued function the engine implements, as
/// "fixed|repeat" with S = scalar, R = range, A = array-or-range. A function
/// missing here (LET, LAMBDA, the dynamic-array functions, user-defined
/// lambdas) is never lifted: it receives its arguments exactly as written.
let private functionSignatureSpecs : (string * string) list =
    [
       "ABS", "|S"; "ACOS", "|S"; "ACOSH", "|S"; "ADDRESS", "|S"; "AGGREGATE", "SS|R"; "AND", "|R"
       "AREAS", "|R"; "ASIN", "|S"; "ASINH", "|S"; "ATAN", "|S"; "ATAN2", "|S"; "ATANH", "|S"
       "AVEDEV", "|R"; "AVERAGE", "|R"; "AVERAGEA", "|R"; "AVERAGEIF", "RSR"; "AVERAGEIFS", "R|RS"; "BIN2DEC", "|S"
       "BITAND", "|S"; "BITOR", "|S"; "BITXOR", "|S"; "CEILING", "|S"; "CHAR", "|S"; "CHOOSE", "S|R"
       "CLEAN", "|S"; "CODE", "|S"; "COLUMN", "|R"; "COLUMNS", "|R"; "COMBIN", "|S"; "CONCAT", "|R"
       "CONCATENATE", "|R"; "CORREL", "|R"; "COS", "|S"; "COSH", "|S"; "COT", "|S"; "COUNT", "|R"
       "COUNTA", "|R"; "COUNTBLANK", "|R"; "COUNTIF", "RS"; "COUNTIFS", "|RS"; "COVAR", "|R"; "COVARIANCE.P", "|R"
       "COVARIANCE.S", "|R"; "DATE", "|S"; "DATEDIF", "|S"; "DATEVALUE", "|S"; "DAVERAGE", "RSR"; "DAY", "|S"
       "DAYS", "|S"; "DAYS360", "|S"; "DCOUNT", "RSR"; "DCOUNTA", "RSR"; "DEC2BIN", "|S"; "DEC2HEX", "|S"
       "DEC2OCT", "|S"; "DEGREES", "|S"; "DEVSQ", "|R"; "DGET", "RSR"; "DMAX", "RSR"; "DMIN", "RSR"
       "DPRODUCT", "RSR"; "DSUM", "RSR"; "EDATE", "|S"; "EOMONTH", "|S"; "ERROR.TYPE", "|S"; "EVEN", "|S"
       "EXACT", "|S"; "EXP", "|S"; "FACT", "|S"; "FALSE", "|S"; "FIND", "|S"; "FIXED", "|S"
       "FLOOR", "|S"; "FORECAST", "SRR"; "FORECAST.LINEAR", "SRR"; "FV", "|S"; "GCD", "|R"; "GEOMEAN", "|R"
       "HARMEAN", "|R"; "HEX2DEC", "|S"; "HLOOKUP", "SASS"; "HOUR", "|S"; "HYPERLINK", "|S"; "IF", "|S"
       "IFERROR", "|S"; "IFNA", "|S"; "IFS", "|S"; "INDEX", "RSS"; "INT", "|S"; "INTERCEPT", "|R"
       "IRR", "RS"; "ISBLANK", "|S"; "ISERR", "|S"; "ISERROR", "|S"; "ISEVEN", "|S"; "ISFORMULA", "|R"
       "ISLOGICAL", "|S"; "ISNA", "|S"; "ISNONTEXT", "|S"; "ISNUMBER", "|S"; "ISODD", "|S"; "ISOWEEKNUM", "|S"
       "ISREF", "|R"; "ISTEXT", "|S"; "LARGE", "RS"; "LCM", "|R"; "LEFT", "|S"; "LEN", "|S"
       "LN", "|S"; "LOG", "|S"; "LOG10", "|S"; "LOOKUP", "SAA"; "LOWER", "|S"; "MATCH", "SAS"
       "MAX", "|R"; "MAXA", "|R"; "MAXIFS", "R|RS"; "MDETERM", "|R"; "MEDIAN", "|R"; "MID", "|S"
       "MIN", "|R"; "MINA", "|R"; "MINIFS", "R|RS"; "MINUTE", "|S"; "MOD", "|S"; "MODE", "|R"
       "MODE.SNGL", "|R"; "MONTH", "|S"; "MROUND", "|S"; "NA", "|S"; "NETWORKDAYS", "SSR"; "NETWORKDAYS.INTL", "SSSR"
       "NORM.DIST", "|S"; "NORM.INV", "|S"; "NORM.S.DIST", "|S"; "NORM.S.INV", "|S"; "NORMDIST", "|S"; "NORMINV", "|S"
       "NORMSDIST", "|S"; "NORMSINV", "|S"; "NOT", "|S"; "NOW", "|S"; "NPER", "|S"; "NPV", "S|R"
       "NUMBERVALUE", "|S"; "OCT2DEC", "|S"; "ODD", "|S"; "OR", "|R"; "PEARSON", "|R"; "PERCENTILE", "RS"
       "PERCENTILE.INC", "RS"; "PERCENTRANK", "RSS"; "PERCENTRANK.INC", "RSS"; "PI", "|S"; "PMT", "|S"; "IPMT", "|S"; "GETPIVOTDATA", "SR|S"; "PPMT", "|S"; "CUMIPMT", "|S"; "CUMPRINC", "|S"; "XNPV", "SRR"; "XIRR", "RRS"; "MIRR", "RSS"; "POWER", "|S"
       "PRODUCT", "|R"; "PROPER", "|S"; "PV", "|S"; "QUARTILE", "RS"; "QUARTILE.INC", "RS"; "QUOTIENT", "|S"
       "RADIANS", "|S"; "RAND", "|S"; "RANDBETWEEN", "|S"; "RANK", "SRS"; "RANK.AVG", "SRS"; "RANK.EQ", "SRS"
       "RATE", "|S"; "REPLACE", "|S"; "REPT", "|S"; "RIGHT", "|S"; "ROUND", "|S"; "ROUNDDOWN", "|S"
       "ROUNDUP", "|S"; "ROW", "|R"; "ROWS", "|R"; "RSQ", "|R"; "SEARCH", "|S"; "SECOND", "|S"
       "SIGN", "|S"; "SIN", "|S"; "SINGLE", "|R"; "SINH", "|S"; "SLN", "|S"; "SLOPE", "|R"
       "SMALL", "RS"; "SQRT", "|S"; "STDEV", "|R"; "STDEV.P", "|R"; "STDEV.S", "|R"; "STDEVP", "|R"
       "STEYX", "|R"; "SUBSTITUTE", "|S"; "SUBTOTAL", "S|R"; "SUM", "|R"; "SUMIF", "RSR"; "SUMIFS", "R|RS"
       "SUMPRODUCT", "|R"; "SUMSQ", "|R"; "SWITCH", "|S"; "TAN", "|S"; "TANH", "|S"; "TEXT", "|S"
       "TEXTAFTER", "|S"; "TEXTBEFORE", "|S"; "TEXTJOIN", "SS|R"; "TIME", "|S"; "TIMEVALUE", "|S"; "TODAY", "|S"
       "TRIM", "|S"; "TRUE", "|S"; "TRUNC", "|S"; "TYPE", "|R"; "UNICHAR", "|S"; "UNICODE", "|S"
       "UPPER", "|S"; "VALUE", "|S"; "VAR", "|R"; "VAR.P", "|R"; "VAR.S", "|R"; "VARP", "|R"
       "VLOOKUP", "SASS"; "WEEKDAY", "|S"; "WEEKNUM", "|S"; "WORKDAY", "SSR"; "WORKDAY.INTL", "SSSR"; "XLOOKUP", "SAASSS"
       "XMATCH", "SASS"; "XOR", "|R"; "YEAR", "|S"; "YEARFRAC", "|S"
    ]

let private parseSignature (spec: string) : FunctionSignature =
    let kinds (text: string) =
        text.ToCharArray()
        |> Array.map (fun ch -> match ch with 'R' -> RangeParam | 'A' -> ArrayParam | _ -> ScalarParam)
    match spec.Split('|') with
    | [| fixedPart; repeat |] -> { Fixed = kinds fixedPart; Repeat = (if repeat = "" then [| ScalarParam |] else kinds repeat) }
    | _ -> { Fixed = kinds spec; Repeat = [| ScalarParam |] }

let private functionSignatures : Dictionary<string, FunctionSignature> =
    let table = Dictionary<string, FunctionSignature>()
    for (name, spec) in functionSignatureSpecs do
        table.[name] <- parseSignature spec
    for entry in Rbt.Spreadsheet.FunctionCatalog.entries do
        if entry.Signature <> "" && not (table.ContainsKey entry.Name) then table.[entry.Name] <- parseSignature entry.Signature
    table

/// The registered signature of a function (name matched case-insensitively).
let functionSignature (name: string) : FunctionSignature voption =
    match Portable.Dict.tryFind name functionSignatures with
    | ValueSome signature -> ValueSome signature
    | ValueNone ->
        // Formula text usually carries upper-case names; fold case only on a miss.
        let upper = name.ToUpperInvariant()
        if obj.ReferenceEquals(upper, name) || upper = name then ValueNone
        else
            match Portable.Dict.tryFind upper functionSignatures with
            | ValueSome signature -> ValueSome signature
            | ValueNone -> ValueNone

let paramKindAt (signature: FunctionSignature) (index: int) : ParamKind =
    if index < signature.Fixed.Length then signature.Fixed.[index]
    else signature.Repeat.[(index - signature.Fixed.Length) % signature.Repeat.Length]

/// Does any parameter of this function accept an array in place of a range?
let private takesArrayParam (name: string) =
    name.Length >= 5 && name.Length <= 7
    && (match functionSignature name with
        | ValueSome signature -> signature.Fixed |> Array.exists (fun kind -> kind = ArrayParam)
        | ValueNone -> false)

/// OFFSET and INDIRECT return a reference, not a value; so do the range
/// (`A1:INDEX(…)`) and intersection (`A1:C3 B2:D4`) operators.
let private isReferenceFunction (name: string) =
    (name.Length = 6 && equalsOrdinalIgnoreCase name "OFFSET")
    || (name.Length = 8 && equalsOrdinalIgnoreCase name "INDIRECT")
    || (name.Length = 1 && (name = rangeOperatorName || name = intersectionOperatorName))

let private isIndexFunction (name: string) =
    name.Length = 5 && equalsOrdinalIgnoreCase name "INDEX"

/// A 3-D reference's qualifier (`Sheet1:Sheet3`) — a colon cannot occur in a
/// real sheet name.
let private isThreeDSheet (sheet: string option) =
    match sheet with
    | Some name -> name.IndexOf ':' >= 0
    | None -> false

/// Can this INDEX call select more than one cell? Only when a row or column
/// number is omitted or literally 0 (the whole line), or a lone index walks
/// something that is not visibly a single row or column.
let private indexMaySelectArea (args: Expr list) =
    let wholeLine e = match e with MissingArg | Number 0.0 -> true | _ -> false
    match args with
    | [ CellRef _; _ ] -> false
    | [ _; row ] when wholeLine row -> true
    | [ source; _ ] ->
        (match source with
         | RangeRef (a, b) -> a.Row <> b.Row && a.Col <> b.Col
         | _ -> true)
    | [ _; row; col ] -> wholeLine row || wholeLine col
    | _ -> false

/// Does this expression evaluate to an array (a spill)? An array function, a
/// range or array constant standing where one value is expected, or an
/// operator / scalar function one of whose operands is itself an array
/// (`SEQUENCE(5)*2`, `A1:A8>2`, `LEN(A1:A9)`). Aggregates consume such an
/// argument over its whole array, not its top-left. `openRanges` says whether
/// a whole-column / whole-row reference counts: it does for an aggregate's
/// operand, but not for deciding what a cell spills — there `A:A` in a scalar
/// position is implicit intersection, as in every pre-dynamic-array workbook.
let private isArrayExprCore (openRanges: bool) (dynamicReferences: bool) (e: Expr) : bool =
    let stack = ResizeArray<Expr>()
    stack.Add e
    let mutable found = false
    while not found && stack.Count > 0 do
        let last = stack.Count - 1
        let expr = stack.[last]
        Portable.Stack.dropLast stack
        match expr with
        | Function (name, _) when isArrayFunction name -> found <- true
        // OFFSET/INDIRECT and the reference operators may yield a multi-cell range.
        | Function (name, _) when name = unionOperatorName -> found <- true
        | Function (name, args) when isReferenceFunction name ->
            // What OFFSET/INDIRECT return is only known when evaluated.
            if dynamicReferences then found <- true
            elif name = rangeOperatorName || name = intersectionOperatorName then
                for arg in args do stack.Add arg
        | Function (name, args) when isIndexFunction name ->
            if indexMaySelectArea args then found <- true
            else
                // A single element: an array only when an index is one.
                match args with
                | _ :: indexes -> for index in indexes do stack.Add index
                | [] -> ()
        // ROW(A1:A5) / COLUMN(A1:C1) number every row / column of the range.
        | Function (name, [ RangeRef _ ])
            when equalsOrdinalIgnoreCase name "ROW" || equalsOrdinalIgnoreCase name "COLUMN" -> found <- true
        // A scalar function is an array when an argument in a scalar slot is.
        | Function (name, args) ->
            match functionSignature name with
            | ValueSome signature ->
                let mutable index = 0
                for arg in args do
                    if paramKindAt signature index = ScalarParam then stack.Add arg
                    index <- index + 1
            | ValueNone -> ()
        // A range/table inside an operator broadcasts element-wise (`A1:A8+B1:B8`).
        // A *bare* range arg never reaches this test — aggregates match RangeRef
        // first and stream it — so this only fires inside Unary/BinOp/array funcs.
        | RangeRef (a, b) ->
            if openRanges || not (isWholeColumnRange a b || isWholeRowRange a b) then found <- true
        | CellRef r -> if isThreeDSheet r.Sheet then found <- true
        // `[@Amount]` is one cell of the current row.
        | TableRef sel when sel.Item = TblThisRow && sel.ColEnd.IsNone -> ()
        | TableRef _ | ArrayLit _ | SpillRef _ -> found <- true
        | UnaryOp (_, inner) -> stack.Add inner
        | BinOp (left, _, right) ->
            stack.Add left
            stack.Add right
        | _ -> ()
    found

let private isArrayExprTable = newFormTable ()
let private argsNeedPreparingTable = newFormTable ()
let private argsArrayTable = newFormTable ()
let private callPlanTable = newFormTable ()

/// What a call is, worked out once for its list of arguments (`callPlanOf`): `Code` 0 for a
/// call that goes the general way, else which of the common functions it is.
type CallPlan = { PlanName: string; Code: int }

let isArrayExpr (e: Expr) : bool =
    let known = formTableGet isArrayExprTable (box e)
    if isNull known then
        let found = isArrayExprCore true true e
        formTableSet isArrayExprTable (box e) (box found)
        found
    else unbox<bool> known

let arithmeticArrayNonBlankCount model workbook e =
    let rangeShape expr =
        match rangeAccessOf model workbook expr with
        | Some access when access.W > 0 -> ValueSome (access.Count / access.W, access.W)
        | _ -> ValueNone
    let scalarShape expr =
        if isArrayExpr expr then ValueNone else ValueSome (1, 1)
    let broadcastShape left right =
        match left, right with
        // Mismatched extents pad with #N/A, which is not blank either.
        | ValueSome (lr, lc), ValueSome (rr, rc) -> ValueSome (max lr rr * max lc rc)
        | _ -> ValueNone
    match e with
    | UnaryOp (op, source) when op = "+" || op = "-" || op = "%" ->
        match rangeShape source with
        | ValueSome (rows, cols) -> ValueSome (rows * cols)
        | ValueNone -> ValueNone
    | BinOp (left, op, right) when op = "+" || op = "-" || op = "*" || op = "/" || op = "^" ->
        broadcastShape
            (match rangeShape left with ValueSome shape -> ValueSome shape | ValueNone -> scalarShape left)
            (match rangeShape right with ValueSome shape -> ValueSome shape | ValueNone -> scalarShape right)
    | _ -> ValueNone

/// Take a root that an iteration accepted (|f| small) the rest of the way:
/// Newton steps until the rate stops moving. IRR accepted a rate as soon as
/// the net present value was below 1e-8, which leaves the rate itself off in
/// its 13th digit; Excel's result is the converged one.
let private polishRoot (f: float -> float) (start: float) : float =
    let mutable r = start
    let mutable best = abs (f start)
    let mutable i = 0
    let mutable moving = true
    while moving && i < 40 do
        let h = 1e-6 * max 1.0 (abs r)
        let value = f r
        let slope = (f (r + h) - f (r - h)) / (2.0 * h)
        if slope = 0.0 || Double.IsNaN slope || Double.IsNaN value then moving <- false
        else
            let next = r - value / slope
            if Double.IsNaN next || Double.IsInfinity next then moving <- false
            else
                let residual = abs (f next)
                // Stop when the step no longer changes the rate, or no longer improves the residual.
                if abs (next - r) <= 1e-16 * max 1.0 (abs next) then
                    (if residual <= best then r <- next)
                    moving <- false
                elif residual > best then moving <- false
                else
                    best <- residual
                    r <- next
        i <- i + 1
    r

/// (1 + rate) ^ periods for the time-value-of-money functions. For a whole
/// number of periods Excel compounds period by period (repeated
/// multiplication), which differs from pow() in the last digit.
/// The payment of a loan (PMT): rate, periods, present value, future value, type (1: in advance).
/// (1 + r)^n without the rounding of 1 + r: exp(n ln(1 + r)), ln(1 + r) from r itself. (1 + r rounded
/// and raised to the 360th power is off by up to 180 units in the last place; Excel's PMT is within one.)
let private growth (r: float) (n: float) : float =
    let u = 1.0 + r
    if u <= 0.0 then u ** n
    else
        let ln1p = if u = 1.0 then r else log u * r / (u - 1.0)
        exp (n * ln1p)

/// The payment of a loan (PMT): (pv + fv / (1+r)^n) r / ((1 + r type)(1 - 1 / (1+r)^n)). (In this
/// order Excel 2019's PMT of 750,000 over 300 months at 7.25% a year is the same to the last bit:
/// its schedule ends on the same 2.5e-9. Excel's is not always: within a unit in the last place.)
let private loanPayment (r: float) (np: float) (pv: float) (fv: float) (ty: float) : float =
    if r = 0.0 then -(pv + fv) / np
    else (let p = growth r np in -(pv + fv / p) * r / ((1.0 + r * ty) * (1.0 - 1.0 / p)))

/// The future value (FV) of `periods` payments on a present value.
let private loanFuture (r: float) (periods: float) (pmt: float) (pv: float) (ty: float) : float =
    if r = 0.0 then -(pv + pmt * periods)
    else (let p = (1.0 + r) ** periods in -(pv * p + pmt * (1.0 + r * ty) * ((p - 1.0) / r)))

/// The interest part of the payment of period `per` (IPMT).
let private loanInterest (r: float) (per: float) (pmt: float) (pv: float) (ty: float) : float =
    let interest =
        if per = 1.0 then (if ty <> 0.0 then 0.0 else -pv)
        elif ty <> 0.0 then loanFuture r (per - 2.0) pmt pv 1.0 - pmt
        else loanFuture r (per - 1.0) pmt pv 0.0
    interest * r

/// The present value of cash flows on dates (XNPV): each discounted by its days after the first / 365.
let private datedPresentValue (r: float) (values: float[]) (dates: float[]) : float =
    let mutable total = 0.0
    for i in 0 .. values.Length - 1 do
        total <- total + values.[i] / ((1.0 + r) ** ((Math.Truncate dates.[i] - Math.Truncate dates.[0]) / 365.0))
    total

/// A root of `f` (a rate: above -1): Newton's method from `guess`, as Excel's IRR goes; where that does not
/// converge, a sign change looked for over the rates and halved down to the root.
let private rootOf (f: float -> float) (guess: float) : float option =
    let ok (r: float) = not (Double.IsNaN r || Double.IsInfinity r) && r > -1.0
    let mutable r = guess
    let mutable result = None
    let mutable i = 0
    while result.IsNone && i < 100 do
        let value = f r
        let h = max 1e-7 (abs r * 1e-7)
        let d = (f (r + h) - value) / h
        if value = 0.0 then result <- Some r
        elif abs d < 1e-300 || not (ok (r - value / d)) then i <- 100
        else
            let next = r - value / d
            if abs (next - r) < 1e-12 * max 1.0 (abs r) then result <- Some (polishRoot f next)
            r <- next
        i <- i + 1
    match result with
    | Some found when ok found -> Some found
    | _ ->
        // Bracketing: rates from just above -1 to 10, the first sign change, then bisection.
        let grid = [| -0.99; -0.9; -0.75; -0.5; -0.3; -0.2; -0.1; -0.05; 0.0; 0.02; 0.05; 0.1; 0.15; 0.2; 0.3; 0.5; 0.75; 1.0; 1.5; 2.0; 3.0; 5.0; 10.0 |]
        let values = grid |> Array.map f
        let finite (v: float) = not (Double.IsNaN v || Double.IsInfinity v)
        match Seq.init (grid.Length - 1) id |> Seq.tryFind (fun k -> finite values.[k] && finite values.[k + 1] && values.[k] * values.[k + 1] <= 0.0) with
        | None -> None
        | Some k ->
            let mutable lo = grid.[k]
            let mutable hi = grid.[k + 1]
            let mutable flo = values.[k]
            for _ in 1 .. 200 do
                let mid = (lo + hi) / 2.0
                let fm = f mid
                if flo * fm <= 0.0 then hi <- mid
                else
                    lo <- mid
                    flo <- fm
            Some ((lo + hi) / 2.0)

let private compound (rate: float) (periods: float) : float =
    if periods >= 0.0 && periods <= 1200.0 && periods = Math.Floor periods then
        let mutable factor = 1.0
        for _ in 1 .. int periods do
            factor <- factor * (1.0 + rate)
        factor
    else (1.0 + rate) ** periods

// Excel drops subnormal results of arithmetic that multiplies/divides (and
// powers), but not those of addition/subtraction or functions such as EXP.
// Keep the sign of zero; a subnormal operand can still produce a normal result.
let private minNormal = 2.2250738585072014e-308
// Only powers whose reciprocal approaches the normal/subnormal boundary need
// a second pow. The small band covers rounding at that binary boundary.
let private reciprocalUnderflowCutoff = (1.0 / minNormal) * (1.0 - 1e-12)

let private flushProductUnderflow (x: float) : float =
    if x > -minNormal && x < minNormal then x * 0.0 else x

let private finiteProduct (x: float) : CellValue =
    if Double.IsInfinity x || Double.IsNaN x then err Err.num
    else CellValue.Number (flushProductUnderflow x)

let private powerNumber (a: float) (b: float) : Result<float, string> =
    let p = a ** b
    // A negative power can have a subnormal positive-power denominator even
    // when its reciprocal fits a double. Most negative powers are nowhere
    // near that boundary and retain their original single-pow path.
    if Double.IsNaN p then Result.Error Err.num
    elif b < 0.0 && a = 0.0 then Result.Error Err.div0
    elif b < 0.0 && (Double.IsInfinity p || abs p >= reciprocalUnderflowCutoff) then
        let denominator = a ** (-b)
        if denominator > -minNormal && denominator < minNormal then Result.Error Err.div0
        elif Double.IsInfinity p then Result.Error Err.num
        else Ok (flushProductUnderflow p)
    elif Double.IsInfinity p then Result.Error Err.num
    else Ok (flushProductUnderflow p)

let private powerProduct (a: float) (b: float) : CellValue =
    match powerNumber a b with
    | Ok value -> CellValue.Number value
    | Result.Error e -> err e

/// Apply a prefix/postfix unary operator to one value (shared by scalar eval
/// and element-wise array broadcasting).
let applyUnary (op: string) (v: CellValue) : CellValue =
    match op with
    | "-" -> (match coerceNumber v with Ok n -> CellValue.Number (-n) | Result.Error e -> err e)
    // Unary plus does nothing at all: it does not coerce, so text stays text
    // and TRUE stays TRUE (`=+A1` is how Lotus-style formulas are typed).
    // Minus and percent do coerce: numeric text converts, other text is #VALUE!.
    | "+" -> v
    | "%" -> (match coerceNumber v with Ok n -> finiteProduct (n / 100.0) | Result.Error e -> err e)
    | _ -> err Err.value

/// `applyBin` for any two values.
let private applyBinGeneral (op: string) (lv: CellValue) (rv: CellValue) : CellValue =
    match firstError [ lv; rv ] with
    | Some e -> e
    | None ->
        match op with
        | "&" ->
            (match coerceText lv, coerceText rv with
             | Ok a, Ok b -> CellValue.String (a + b)
             | Result.Error e, _ | _, Result.Error e -> err e)
        | "=" | "<>" | "<" | ">" | "<=" | ">=" ->
            (match compareOperands lv rv with
             | Result.Error e -> err e
             | Ok cmp ->
                CellValue.Boolean
                    (match op with
                     | "=" -> cmp = 0 | "<>" -> cmp <> 0 | "<" -> cmp < 0
                     | ">" -> cmp > 0 | "<=" -> cmp <= 0 | _ -> cmp >= 0))
        | _ ->
            (match coerceNumber lv, coerceNumber rv with
             | Result.Error e, _ -> err e
             | _, Result.Error e -> err e
             | Ok a, Ok b ->
                // A result beyond the double range is #NUM!, never infinity.
                let finite (x: float) = if Double.IsInfinity x || Double.IsNaN x then err Err.num else CellValue.Number x
                match op with
                | "+" -> finite (a + b)
                | "-" -> finite (a - b)
                | "*" -> finiteProduct (a * b)
                | "/" -> if b > -minNormal && b < minNormal then err Err.div0 else finiteProduct (a / b)
                | "^" -> powerProduct a b
                | _ -> err Err.value)

/// Apply a binary operator to two values (shared by scalar eval and array
/// broadcasting). Mirrors the scalar BinOp semantics exactly.
/// (Two numbers - most of what a model calculates - without the detours of the general
/// case: no list of the operands, no result wrappers. The outcome is `applyBinGeneral`'s.)
let applyBin (op: string) (lv: CellValue) (rv: CellValue) : CellValue =
    match lv with
    | CellValue.Number a ->
        match rv with
        | CellValue.Number b ->
            match op with
            | "+" -> (let x = a + b in if Double.IsInfinity x || Double.IsNaN x then err Err.num else CellValue.Number x)
            | "-" -> (let x = a - b in if Double.IsInfinity x || Double.IsNaN x then err Err.num else CellValue.Number x)
            | "*" -> finiteProduct (a * b)
            | "/" -> if b > -minNormal && b < minNormal then err Err.div0 else finiteProduct (a / b)
            | "{o}" -> CellValue.Boolean (compareNumbers a b = 0)
            | "{o}" -> CellValue.Boolean (compareNumbers a b <> 0)
            | "{o}" -> CellValue.Boolean (compareNumbers a b < 0)
            | "{o}" -> CellValue.Boolean (compareNumbers a b > 0)
            | "{o}" -> CellValue.Boolean (compareNumbers a b <= 0)
            | "{o}" -> CellValue.Boolean (compareNumbers a b >= 0)
            | _ -> applyBinGeneral op lv rv
        | _ -> applyBinGeneral op lv rv
    | _ -> applyBinGeneral op lv rv

/// Parse R1C1-style reference text (`R2C3`, `R[-1]C`, `R1C1:R3C2`, `C2`, `R5`,
/// optionally `Sheet!`-qualified) — what INDIRECT accepts when its second
/// argument is FALSE. Bracketed offsets are relative to `current`.
let parseR1C1Reference (text: string) (current: CellAddress option) : Expr option =
    let text = text.Trim()
    let sheet, body =
        let bang = text.LastIndexOf '!'
        if bang < 0 then None, text
        else
            let rawName = text.Substring(0, bang)
            let sheetName =
                if rawName.Length >= 2 && rawName.[0] = '\'' && rawName.[rawName.Length - 1] = '\'' then
                    rawName.Substring(1, rawName.Length - 2).Replace("''", "'")
                else rawName
            Some sheetName, text.Substring(bang + 1)
    let here = defaultArg current { Row = 0; Col = 0 }
    /// One `R…C…` corner: (row, col) with None for an axis left out.
    let corner (part: string) : (int option * int option) option =
        let n = part.Length
        let mutable i = 0
        let mutable ok = true
        /// After the axis letter: nothing (same row/col), `[d]`, or a 1-based number.
        let axis (origin: int) : int =
            if i < n && part.[i] = '[' then
                let close = part.IndexOf(']', i)
                if close < 0 then (ok <- false; 0)
                else
                    let inner = part.Substring(i + 1, close - i - 1)
                    i <- close + 1
                    match Int32.TryParse(inner, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture) with
                    | true, delta -> origin + delta
                    | _ -> ok <- false; 0
            else
                let start = i
                while i < n && Char.IsDigit part.[i] do i <- i + 1
                if i = start then origin
                else
                    match Int32.TryParse(part.Substring(start, i - start), NumberStyles.None, CultureInfo.InvariantCulture) with
                    | true, value -> value - 1
                    | _ -> ok <- false; 0
        let mutable row = None
        let mutable col = None
        if i < n && (part.[i] = 'R' || part.[i] = 'r') then
            i <- i + 1
            row <- Some (axis here.Row)
        if i < n && (part.[i] = 'C' || part.[i] = 'c') then
            i <- i + 1
            col <- Some (axis here.Col)
        if ok && i = n && (row.IsSome || col.IsSome) then Some (row, col) else None
    let refAt row col = { Sheet = sheet; Row = row; Col = col; RowAbs = false; ColAbs = false }
    let inBounds row col = row >= 0 && row <= lastSheetRow && col >= 0 && col <= lastSheetCol
    let parts = body.Split(':')
    let corners = parts |> Array.map corner
    if parts.Length = 0 || parts.Length > 2 || corners |> Array.exists Option.isNone then None
    else
        let firstRow, firstCol = corners.[0].Value
        let lastRow, lastCol = corners.[corners.Length - 1].Value
        match firstRow, firstCol, lastRow, lastCol with
        | Some r0, Some c0, Some r1, Some c1 when inBounds r0 c0 && inBounds r1 c1 ->
            if r0 = r1 && c0 = c1 then Some (CellRef (refAt r0 c0)) else Some (RangeRef (refAt r0 c0, refAt r1 c1))
        // `R5` / `R2:R4`: whole rows. `C2` / `C1:C3`: whole columns.
        | Some r0, None, Some r1, None when inBounds r0 0 && inBounds r1 0 ->
            Some (RangeRef ({ refAt r0 0 with ColAbs = true }, { refAt r1 lastSheetCol with ColAbs = true }))
        | None, Some c0, None, Some c1 when inBounds 0 c0 && inBounds 0 c1 ->
            Some (RangeRef ({ refAt 0 c0 with RowAbs = true }, { refAt lastSheetRow c1 with RowAbs = true }))
        | _ -> None

/// Everything an evaluation needs beyond the expression itself: the sheet
/// model being evaluated, the workbook (for cross-sheet refs and defined
/// names), and which sheet is "current" (for scoping sheet-local names).
type EvalContext = {
    Model: SpreadsheetModel
    Workbook: Rbt.Spreadsheet.Workbook.Workbook option
    Sheet: Rbt.Spreadsheet.Workbook.SheetId option
    /// Address of the cell being evaluated, when known. Needed to resolve
    /// this-row structured refs (`Table[@Col]` / `[@Col]`) and to find the
    /// containing table for unqualified refs. None for ad-hoc evaluation.
    CurrentCell: CellAddress option
}

/// Build a context with no workbook — same-sheet formulas only, no defined
/// names. Convenient for callers that just have a bare model.
let contextOf (model: SpreadsheetModel) : EvalContext =
    { Model = model; Workbook = None; Sheet = None; CurrentCell = None }

/// The date system of the workbook an evaluation is of: what the sheet
/// being evaluated says (every sheet of a workbook says the same).
let private date1904Of (ctx: EvalContext) : bool = ctx.Model.Date1904

/// Last row that holds anything on a sheet (0 for an empty sheet). Cells are
/// keyed row-major, so this is the map's last key.
let private lastUsedRow (m: SpreadsheetModel) : int =
    if Map.isEmpty m.Cells then 0 else (Map.maxKeyValue m.Cells |> fst).Row

let mutable private lastUsedColCells : obj = null
let mutable private lastUsedColValue = 0

/// Last column that holds anything on a sheet. Needs a scan, so the answer
/// for the most recent cell map is remembered.
let private lastUsedCol (m: SpreadsheetModel) : int =
    if obj.ReferenceEquals(lastUsedColCells, m.Cells) then lastUsedColValue
    else
        let mutable last = 0
        for KeyValue(addr, _) in m.Cells do
            if addr.Col > last then last <- addr.Col
        lastUsedColCells <- box m.Cells
        lastUsedColValue <- last
        last

let private isOpenRange (a: RefAst) (b: RefAst) = isWholeColumnRange a b || isWholeRowRange a b

/// Bound a whole-column / whole-row reference by the sheet's used range, so
/// `SUM(A:A)` walks the rows that exist instead of a million blanks. Other
/// ranges are returned as they are.
let private clampOpenRange (m: SpreadsheetModel) (a: RefAst) (b: RefAst) : RefAst * RefAst =
    if not (isOpenRange a b) then a, b
    else
        let top = min a.Row b.Row
        let left = min a.Col b.Col
        let bottom = if isWholeColumnRange a b then max top (lastUsedRow m) else max a.Row b.Row
        let right = if isWholeRowRange a b then max left (lastUsedCol m) else max a.Col b.Col
        { a with Row = top; Col = left }, { b with Row = bottom; Col = right }

/// Functions that ask about a reference's shape or position, not its
/// values: a whole-column argument must reach them unbounded (ROWS(A:A) is
/// 1,048,576).
let private keepsOpenRanges (name: string) =
    equalsOrdinalIgnoreCase name "ROWS" || equalsOrdinalIgnoreCase name "COLUMNS"
    || equalsOrdinalIgnoreCase name "ROW" || equalsOrdinalIgnoreCase name "COLUMN"
    || equalsOrdinalIgnoreCase name "ISREF" || equalsOrdinalIgnoreCase name implicitIntersectionName

/// True if the expression calls OFFSET or INDIRECT anywhere: its precedents
/// are only known once it is evaluated.
let containsDynamicReference (expr: Expr) : bool =
    containsFunctionNamed "OFFSET" expr || containsFunctionNamed "INDIRECT" expr
    || containsFunctionNamed rangeOperatorName expr || containsFunctionNamed intersectionOperatorName expr

/// Defined-name lookup for evaluation: like `resolveNameExpr`, with a
/// whole-column / whole-row target bounded by its sheet's used range.
let private resolveNameForEval (name: string) (currentSheet: Rbt.Spreadsheet.Workbook.SheetId option)
                               (wb: Rbt.Spreadsheet.Workbook.Workbook) : Expr option =
    match resolveNameExpr name currentSheet wb with
    | Some (RangeRef (a, b)) when isOpenRange a b ->
        let sheetId =
            match a.Sheet with
            | Some sheetName -> Rbt.Spreadsheet.Workbook.resolveSheetName sheetName wb
            | None -> currentSheet
        match sheetId |> Option.bind (fun id -> Map.tryFind id wb.Sheets) with
        | Some sheet -> Some (RangeRef (clampOpenRange (resolveSheet sheet.Model) a b))
        | None -> Some (ErrorLit Err.ref)
    | other -> other

/// Defined names whose refers-to formula is being evaluated right now, in
/// evaluation order. A name met again while it is still on this list refers
/// to itself through other names (`A = B+1`, `B = A+1`): that is a circular
/// reference, reported as an error instead of recursing until the stack ends.
let private namesBeingEvaluated = ResizeArray<string>()

// ---------------------------------------------------------------------------
// The function table
//
// Functions added since the catalogue above was closed are not cases of its
// match: each is a `FunctionDef` of a module of src/Spreadsheet/Functions,
// registered here by the module's `install` (`registerFunctions`). In the
// browser a module is its own ES module, fetched with a dynamic `import()`
// the first time a workbook or a formula needs one of its functions
// (`requestFunctions`, or a call met while evaluating, which shows #BUSY!
// until the module is there - then `onFunctionsArrived` listeners hear of it
// and recalculate). On .NET every module is installed at start-up
// (Functions/All.fs). The generated FunctionCatalog knows every name the
// modules provide, so a function not loaded yet is told from one nobody
// provides (#NAME?).
// ---------------------------------------------------------------------------

/// What a function of a module is handed to evaluate its arguments with: the
/// evaluator's own entry points (a module compiles after this file and cannot
/// be part of its recursive group).
type FunctionEvaluator = {
    /// The value of an expression (a reference in a scalar context is its first cell).
    Eval: Expr -> EvalContext -> CellValue
    /// The values of an expression as rows of columns: a reference's cells
    /// (blank cells are Empty), an array's items, one value for anything else.
    EvalArray: Expr -> EvalContext -> CellValue[][]
    /// The reference an argument stands for - a cell or a range, also through a
    /// defined name, OFFSET, INDIRECT, INDEX or a spill (`A1#`) - when it is one.
    Reference: Expr -> EvalContext -> Expr option
    /// Does the expression give more than one value (a range, an array, a
    /// function that gives one)?
    IsArray: Expr -> bool
    /// A criteria value (COUNTIF's, the database functions') as a test of a value.
    Criteria: CellValue -> (CellValue -> bool)
    /// A call of another function, built in or of the table.
    Call: string -> Expr list -> EvalContext -> CellValue
}

type FunctionImpl =
    /// Gives one value. In an array context its scalar parameters are lifted as
    /// the signature says.
    | ScalarFunction of (FunctionEvaluator -> Expr list -> EvalContext -> CellValue)
    /// Gives an array (spills); in a scalar context its top-left value.
    | ArrayFunction of (FunctionEvaluator -> Expr list -> EvalContext -> CellValue[][])

/// A function of a module. `Signature` is the lifting signature (see
/// `functionSignatureSpecs`; "" for one that takes its arguments as written);
/// `MaxArgs` -1 for no limit. The generated FunctionCatalog repeats Name,
/// Signature and the flags (volatile, gives an array) so that they are known
/// before the module is loaded.
type FunctionDef = {
    Name: string
    MinArgs: int
    MaxArgs: int
    Signature: string
    Volatile: bool
    Impl: FunctionImpl
}

/// A call of a provider function, its arguments evaluated (an omitted one is a 1x1 of Empty).
type ProviderCall = {
    Name: string
    Arguments: CellValue[][] array
    /// The sheet (its id) and cell of the formula, when known (-1 when not).
    Sheet: string option
    Row: int
    Col: int
}

/// What a host registers for a provider function: `Invoke call answer` calls
/// `answer` with the result, then or later (a later answer shows #BUSY! until
/// it comes; answering again replaces the value, which is how a stream of
/// values - RTD - updates the cell). Held in a record: mutable function values
/// are not reliable in Fable.
type FunctionProvider = { Invoke: ProviderCall -> (CellValue[][] -> unit) -> unit }

/// Something to tell when functions arrived (a module loaded, a provider
/// answered): the formulas that showed #BUSY! are worth calculating again.
type FunctionArrival = { Arrived: unit -> unit }

let private registeredFunctions = Dictionary<string, FunctionDef>()
let private moduleStates = Dictionary<string, int>()   // 1 loading, 2 installed
let private arrivalListeners = ResizeArray<FunctionArrival>()

/// Tell the listeners that functions arrived.
let notifyFunctionsArrived () =
    for listener in arrivalListeners.ToArray() do
        try listener.Arrived () with _ -> ()

/// Listen for functions arriving; gives what stops listening.
let onFunctionsArrived (arrived: unit -> unit) : (unit -> unit) =
    let listener = { Arrived = arrived }
    arrivalListeners.Add listener
    fun () -> arrivalListeners.Remove listener |> ignore

/// Register a module's functions (its `install` does this).
let registerFunctions (moduleName: string) (defs: FunctionDef list) : unit =
    for def in defs do registeredFunctions.[upperName def.Name] <- def
    moduleStates.[moduleName] <- 2

#if !FABLE_COMPILER
let mutable private allInstalled = false
/// (.NET: reading Functions/All.fs's `installed` installs every module - the
/// assembly initialises a file when one of its values is first read.)
let private ensureAllInstalled () =
    if not allInstalled then
        allInstalled <- true
        try
            let all = typeof<Expr>.Assembly.GetType "Rbt.Spreadsheet.Functions.All"
            if not (isNull all) then
                let installed = all.GetProperty "installed"
                if not (isNull installed) then installed.GetValue(null, null) |> ignore
        with _ -> ()
#endif

/// The function of the table of that name, when its module is installed.
let tryFunctionDef (name: string) : FunctionDef option =
#if !FABLE_COMPILER
    ensureAllInstalled ()
#endif
    match Portable.Dict.tryFind (upperName name) registeredFunctions with
    | ValueSome def -> Some def
    | ValueNone -> None

/// Is the module installed?
let isFunctionModuleInstalled (moduleName: string) =
#if !FABLE_COMPILER
    ensureAllInstalled ()
#endif
    match Portable.Dict.tryFind moduleName moduleStates with
    | ValueSome 2 -> true
    | _ -> false

/// Start loading a module (browser; on .NET they are all installed). `ready`
/// is called once it is installed, at once when it is already.
let private loadFunctionModule (moduleName: string) (ready: unit -> unit) : unit =
    if isFunctionModuleInstalled moduleName then ready ()
    else
#if FABLE_COMPILER
        if not (moduleStates.ContainsKey moduleName) then moduleStates.[moduleName] <- 1
        let loading = Rbt.Spreadsheet.FunctionCatalog.load moduleName
        loading.``then``(fun () ->
            moduleStates.[moduleName] <- 2
            ready ()) |> ignore
#else
        ready ()
#endif

/// The modules that the names need and are not installed.
let private modulesFor (names: string seq) : string list =
    names
    |> Seq.choose (fun name ->
        match catalogEntry name with
        | Some entry when not (isFunctionModuleInstalled entry.Module) -> Some entry.Module
        | _ -> None)
    |> Seq.distinct
    |> List.ofSeq

/// Load the modules of these function names; `ready` once all are installed
/// (at once when nothing is missing). Opening a workbook does this with every
/// name its formulas call before calculating.
let requestFunctions (names: string seq) (ready: unit -> unit) : unit =
    match modulesFor names with
    | [] -> ready ()
    | needed ->
        let mutable left = needed.Length
        for moduleName in needed do
            loadFunctionModule moduleName (fun () ->
                left <- left - 1
                if left = 0 then ready ())

/// A function met while evaluating whose module is not installed: load it and
/// tell the listeners when it is.
let private requestFunctionWhileEvaluating (moduleName: string) =
    match Portable.Dict.tryFind moduleName moduleStates with
    | ValueSome _ -> ()   // loading already (or installed)
    | ValueNone -> loadFunctionModule moduleName notifyFunctionsArrived

let private isFormulaNameStart (c: char) = Char.IsLetter c || c = '_' || c = '\\'
let private isFormulaNameChar (c: char) = Char.IsLetterOrDigit c || c = '_' || c = '.' || c = '\\'

/// The names a formula's text calls (`NAME(`), upper-cased, prefixes
/// (`_xlfn.`) taken off; what stands in quotes is skipped. A cheap scan, for
/// finding the modules a workbook needs without parsing every formula.
let calledFunctionNames (formula: string) : string list =
    if isNull formula || formula.IndexOf '(' < 0 then []
    else
        let found = ResizeArray<string>()
        let n = formula.Length
        let mutable i = 0
        while i < n do
            let c = formula.[i]
            if c = '"' || c = '\'' then
                i <- i + 1
                while i < n && formula.[i] <> c do i <- i + 1
                i <- i + 1
            elif isFormulaNameStart c && (i = 0 || not (isFormulaNameChar formula.[i - 1])) then
                let start = i
                while i < n && isFormulaNameChar formula.[i] do i <- i + 1
                if i < n && formula.[i] = '(' then
                    let mutable word = formula.Substring(start, i - start).ToUpperInvariant()
                    for prefix in [ "_XLFN._XLWS."; "_XLFN."; "_XLWS." ] do
                        if word.StartsWith prefix then word <- word.Substring prefix.Length
                    found.Add word
            else i <- i + 1
        List.ofSeq found

/// The modules the formulas need, loaded; `ready` when they all are.
let requestFunctionsFor (formulas: string seq) (ready: unit -> unit) : unit =
    let names = HashSet<string>()
    for formula in formulas do
        for name in calledFunctionNames formula do names.Add name |> ignore
    requestFunctions names ready

// ----- providers -----------------------------------------------------------

let private providers = Dictionary<string, FunctionProvider>()
let private providerResults = Dictionary<string, CellValue[][]>()
let private providerPending = HashSet<string>()

let private forgetProviderResults (key: string) =
    let prefix = key + "|"
    for cached in [ for k in providerResults.Keys do if k.StartsWith prefix then yield k ] do
        providerResults.Remove cached |> ignore
    for pending in [ for k in providerPending do if k.StartsWith prefix then yield k ] do
        providerPending.Remove pending |> ignore

/// Register the provider of a provider function (`providerFunctionSpecs`);
/// false for a name that is none. Results the one before gave are forgotten.
let registerFunctionProvider (name: string) (provider: FunctionProvider) : bool =
    let key = upperName name
    if not (providerSpecByName.ContainsKey key) then false
    else
        providers.[key] <- provider
        forgetProviderResults key
        true

/// Take a provider away: its formulas keep the values they have.
let unregisterFunctionProvider (name: string) : unit =
    let key = upperName name
    providers.Remove key |> ignore
    forgetProviderResults key

let hasFunctionProvider (name: string) = providers.ContainsKey (upperName name)

/// The functions that have a provider now.
let functionProviderNames () : string list = [ for name in providers.Keys -> name ]

// ----- what the host says about itself (INFO, CELL("filename")) -------------

let private hostInfoFields = Dictionary<string, string>()

/// Set (Some) or clear (None) what INFO(field) gives - "directory", "origin",
/// "osversion", "system", "release", "recalc", "numfile" - and "filename", the
/// workbook's file name CELL("filename") puts with "directory". A field the
/// host does not set has INFO's default.
let setHostInfo (field: string) (value: string option) : unit =
    let key = upperName field
    match value with
    | Some text -> hostInfoFields.[key] <- text
    | None -> hostInfoFields.Remove key |> ignore

let hostInfo (field: string) : string option =
    match Portable.Dict.tryFind (upperName field) hostInfoFields with
    | ValueSome text -> Some text
    | ValueNone -> None

let private appendValueKey (sb: Text.StringBuilder) (value: CellValue) =
    match unwrapFormulaValue value with
    | CellValue.Empty -> sb.Append "e;" |> ignore
    | CellValue.Number n -> sb.Append("n").Append(n.ToString(CultureInfo.InvariantCulture)).Append(';') |> ignore
    | CellValue.String s -> sb.Append("s").Append(s.Length).Append(':').Append(s).Append(';') |> ignore
    | CellValue.Boolean b -> sb.Append(if b then "t;" else "f;") |> ignore
    | CellValue.Error e -> sb.Append("x").Append(e).Append(';') |> ignore
    | CellValue.Formula _ -> sb.Append "?;" |> ignore

/// A provider function's value for these arguments: what the provider gave,
/// #BUSY! while it has not answered, the function's error without a provider.
let callProvider (call: ProviderCall) : CellValue[][] =
    let key = upperName call.Name
    match Portable.Dict.tryFind key providers with
    | ValueNone ->
        let error = match providerSpecByName.TryGetValue key with | true, (e, _) -> e | _ -> Err.name
        [| [| CellValue.Error error |] |]
    | ValueSome provider ->
        let sb = Text.StringBuilder(key).Append('|')
        for argument in call.Arguments do
            sb.Append(argument.Length).Append('x').Append(if argument.Length > 0 then argument.[0].Length else 0).Append('[') |> ignore
            for row in argument do
                for value in row do appendValueKey sb value
            sb.Append(']') |> ignore
        let resultKey = sb.ToString()
        match Portable.Dict.tryFind resultKey providerResults with
        | ValueSome result -> result
        | ValueNone ->
            if providerPending.Contains resultKey then [| [| CellValue.Error Err.busy |] |]
            else
                let mutable invoking = true
                providerPending.Add resultKey |> ignore
                let answer (result: CellValue[][]) =
                    let result = if isNull result || result.Length = 0 || result.[0].Length = 0 then [| [| CellValue.Empty |] |] else result
                    // (An answer of a provider taken away or replaced since is dropped.)
                    match Portable.Dict.tryFind key providers with
                    | ValueSome current when obj.ReferenceEquals(current, provider) ->
                        providerResults.[resultKey] <- result
                        providerPending.Remove resultKey |> ignore
                        if not invoking then notifyFunctionsArrived ()
                    | _ -> ()
                (try provider.Invoke call answer
                 with _ -> answer [| [| CellValue.Error Err.connect |] |])
                invoking <- false
                match Portable.Dict.tryFind resultKey providerResults with
                | ValueSome result -> result
                | ValueNone -> [| [| CellValue.Error Err.busy |] |]

/// Does an expression call a function that `test` accepts?
let private callsFunctionWhere (test: string -> bool) (expr: Expr) : bool =
    let stack = ResizeArray<Expr>()
    stack.Add expr
    let mutable found = false
    while not found && stack.Count > 0 do
        let current = stack.[stack.Count - 1]
        Portable.Stack.dropLast stack
        match current with
        | Function (name, args) ->
            if test name then found <- true
            else for arg in args do stack.Add arg
        | UnaryOp (_, inner) -> stack.Add inner
        | BinOp (l, _, r) -> stack.Add l; stack.Add r
        | _ -> ()
    found

let private callsProviderTable = newFormTable ()

/// Does the formula keep the value it has instead of being calculated? Yes when
/// it calls a provider function that has no provider: what the file cached
/// for it (what Excel computed with the service) is the best value there is.
let keepsCachedValue (expr: Expr) : bool =
    let known = formTableGet callsProviderTable (box expr)
    let calls =
        if not (isNull known) then unbox<bool> known
        else
            let found = callsFunctionWhere isProviderFunction expr
            formTableSet callsProviderTable (box expr) (box found)
            found
    calls && callsFunctionWhere (fun name -> isProviderFunction name && not (hasFunctionProvider name)) expr

let mutable private sharedEvaluator : FunctionEvaluator option = None

let private noCallPlan : CallPlan = { PlanName = ""; Code = 0 }

/// Evaluate an expression to a scalar CellValue.
/// `x`, the sum of `a` and `b`, is less than 8 units in the last place of the larger of them: what
/// Excel takes for 0 at the end of an addition (`evalCell`, SUM). (Excel 2019: a sum 7 units off is 0,
/// one 8 units off is not, whatever the mantissa.)
let private cancelsWithin (x: float) (a: float) (b: float) : bool =
    let m = max (abs a) (abs b)
    if m = 0.0 || Double.IsInfinity m || Double.IsNaN m then false
    else
        let mutable e = Math.Floor(Math.Log m / Math.Log 2.0)
        if 2.0 ** e > m then e <- e - 1.0
        elif 2.0 ** (e + 1.0) <= m then e <- e + 1.0
        abs x < 2.0 ** (e - 49.0)

let rec eval (expr: Expr) (ctx: EvalContext) : CellValue =
    let model = ctx.Model
    let workbook = ctx.Workbook
    match expr with
    | Number n -> CellValue.Number n
    | Str s -> CellValue.String s
    | Bool b -> CellValue.Boolean b
    | ErrorLit e -> CellValue.Error e
    | NameRef name ->
        match workbook with
        | None -> err Err.name
        | Some wb ->
            match resolveNameForEval name ctx.Sheet wb with
            | Some e ->
                let key = name.ToUpperInvariant()
                if namesBeingEvaluated.Contains key then err Err.cycle
                else
                    namesBeingEvaluated.Add key
                    try eval e ctx
                    finally Portable.Stack.dropLast namesBeingEvaluated
            | None -> err Err.name
    | CellRef r -> readCell r model workbook
    | RangeRef (a, b) when isOpenRange a b && ctx.CurrentCell.IsSome ->
        // `A:A` where one value is expected is implicit intersection: the
        // cell of that column on the formula's own row (of that row in the
        // formula's own column), which is how workbooks written before
        // dynamic arrays read `=VLOOKUP(A:A,…)` on every row.
        let cell = ctx.CurrentCell.Value
        let sr, sc, er, ec = normalizeRange a b
        if sc = ec && isWholeColumnRange a b then readCell { a with Row = cell.Row; Col = sc } model workbook
        elif sr = er && isWholeRowRange a b then readCell { a with Row = sr; Col = cell.Col } model workbook
        else firstRangeRefValue a b model workbook
    | RangeRef (a, b) ->
        // A bare range in scalar context yields its first cell (Excel's
        // implicit-intersection collapse is close enough for our use).
        firstRangeRefValue a b model workbook
    | TableRef sel -> firstTableSelValue sel workbook ctx.Sheet ctx.CurrentCell model
    | ArrayLit a -> if a.Length > 0 && a.[0].Length > 0 then a.[0].[0] else err Err.value
    // Excel reads an omitted argument as 0 (`IF(TRUE,,2)` is 0).
    | MissingArg -> CellValue.Number 0.0
    | SpillRef anchor -> let a, b = spillRangeIn anchor model workbook in eval (RangeRef (a, b)) ctx
    | UnaryOp (op, operand) -> applyUnary op (eval operand ctx)
    | BinOp _ -> evalBinOpChain expr ctx
    | Function (name, args) ->
        // (The functions a model is mostly made of - IF, AND, OR, the parts of a date,
        // EOMONTH - with plain arguments: see `evalCommonFunction`.)
        let common =
            let plan = callPlanOf name args
            if plan.Code = 0 then Unchecked.defaultof<CellValue> else evalCommonFunction plan.Code args ctx
        if not (isNull (box common)) then common else
        // An array-producing function in scalar context collapses to its
        // top-left element (implicit intersection; spill placement is Step 2).
        let args = if argsNeedPreparing name args then prepareArgs name (keepsOpenRanges name) args ctx else args
        if name = unionOperatorName then
            // A union has no single value; in a scalar context it is its first area.
            (match args with first :: _ -> eval first ctx | [] -> err Err.value)
        elif isArrayFunction name then
            match tryEvalArrayTopLeft name args ctx with
            | ValueSome value -> value
            | ValueNone ->
                let arr = evalArray (Function (name, args)) ctx
                if arr.Length > 0 && arr.[0].Length > 0 then arr.[0].[0] else err Err.value
        elif isReferenceFunction name then
            // In a scalar context the reference collapses to its first cell.
            eval (resolveReferenceFunction name args ctx) ctx
        else
            let value =
#if FABLE_COMPILER
                evalFunction name args ctx
#else
                // (.NET's date arithmetic throws outside the calendar - EOMONTH or EDATE
                // fed something that is no date - where JavaScript gives an invalid date:
                // the cell is #NUM!, and the recalculation goes on.)
                try evalFunction name args ctx
                with :? System.ArgumentOutOfRangeException -> err Err.num
#endif
            match value with
            | CellValue.Number x when x <> 0.0 && abs x < 1e-3 && equalsOrdinalIgnoreCase name "SUM" -> sumCancelled x args ctx
            | _ -> value

/// SUM's last addition as Excel makes it: a total that is within 8 units in the last place of the
/// larger of the two it adds (the total so far and the last number), of opposite signs, is 0 -
/// whatever the SUM stands in (`=1*SUM(1.5,-1.5-2^-52)` is 0, `=SUM(1.5,-1.5-2^-52,0)` is not).
and private sumCancelled (total: float) (args: Expr list) (ctx: EvalContext) : CellValue =
    let lastNumber (arg: Expr) =
        match arg with
        | CellRef _ | RangeRef _ | NameRef _ | TableRef _ | SpillRef _ | Function _ ->
            let values = evalArray arg ctx
            let mutable found = None
            let mutable r = values.Length - 1
            while found.IsNone && r >= 0 do
                let row = values.[r]
                let mutable c = row.Length - 1
                while found.IsNone && c >= 0 do
                    (match unwrapValue row.[c] with CellValue.Number n -> found <- Some n | _ -> ())
                    c <- c - 1
                r <- r - 1
            found
        | MissingArg -> None
        | _ -> match unwrapValue (eval arg ctx) with CellValue.Number n -> Some n | _ -> None
    match args |> List.rev |> List.tryPick lastNumber with
    | Some last ->
        let before = total - last
        if (before > 0.0) <> (last > 0.0) && cancelsWithin total before last then CellValue.Number 0.0 else CellValue.Number total
    | None -> CellValue.Number total

/// Which of the common functions a call is (`CallPlan.Code`), when its arguments are plain -
/// none that has to be prepared (`argsNeedPreparing`), and for AND / OR none that is a range,
/// a name, a table or an array: 1 IF(c,t), 2 IF(c,t,f), 3 AND, 4 OR, 5 NOT, 6 IFERROR,
/// 7 YEAR, 8 MONTH, 9 DAY, 10 EOMONTH, 11 EDATE. 0 for every other call.
and callPlanOf (name: string) (args: Expr list) : CallPlan =
    if args.IsEmpty then noCallPlan
    else
        let known = formTableGet callPlanTable (box args)
        let kept = if isNull known then Unchecked.defaultof<CallPlan> else unbox<CallPlan> known
        if not (isNull (box kept)) && (obj.ReferenceEquals(kept.PlanName, name) || kept.PlanName = name) then kept
        else
            let count = List.length args
            let plainScalars () =
                args |> List.forall (fun arg ->
                    match arg with
                    | CellRef _ | BinOp _ | UnaryOp _ | Function _ | Number _ | Str _ | Bool _ -> not (isArrayExpr arg)
                    | _ -> false)
            let code =
                if name.Length < 2 || name.Length > 7 then 0
                else
                    match name.ToUpperInvariant() with
                    | "IF" -> if count = 2 then 1 elif count = 3 then 2 else 0
                    | "AND" -> if plainScalars () then 3 else 0
                    | "OR" -> if plainScalars () then 4 else 0
                    | "NOT" -> if count = 1 then 5 else 0
                    | "IFERROR" -> if count = 2 then 6 else 0
                    | "YEAR" -> if count = 1 then 7 else 0
                    | "MONTH" -> if count = 1 then 8 else 0
                    | "DAY" -> if count = 1 then 9 else 0
                    | "EOMONTH" -> if count = 2 then 10 else 0
                    | "EDATE" -> if count = 2 then 11 else 0
                    | _ -> 0
            let code = if code <> 0 && argsNeedPreparingNow name args then 0 else code
            let plan = { PlanName = name; Code = code }
            formTableSet callPlanTable (box args) (box plan)
            plan

/// A call of one of the common functions (`callPlanOf`), with the outcome `evalFunction`
/// gives for it - that function is the whole catalogue, and a call pays for its size.
/// Null where this one has no answer (a date that is not a whole day of the calendar):
/// the call then goes the general way.
and evalCommonFunction (code: int) (args: Expr list) (ctx: EvalContext) : CellValue =
    match code with
    | 1 | 2 ->
        let condition = unwrapValue (eval args.Head ctx)
        let truth =
            match condition with
            | CellValue.Boolean b -> (if b then 1 else 0)
            | CellValue.Number n -> (if n <> 0.0 then 1 else 0)
            | other -> (match coerceBool other with Ok b -> (if b then 1 else 0) | Result.Error _ -> -1)
        if truth = 1 then eval args.Tail.Head ctx
        elif truth = 0 then (if code = 1 then CellValue.Boolean false else eval args.Tail.Tail.Head ctx)
        else (match coerceBool condition with Result.Error e -> err e | Ok _ -> err Err.value)
    | 3 | 4 ->
        let isAnd = code = 3
        let mutable seen = false
        let mutable result = isAnd
        let mutable firstError : string = null
        let mutable remaining = args
        while not remaining.IsEmpty do
            (match unwrapValue (eval remaining.Head ctx) with
             | CellValue.Empty -> ()
             | CellValue.Boolean b ->
                seen <- true
                result <- (if isAnd then result && b else result || b)
             | CellValue.Number n ->
                seen <- true
                result <- (if isAnd then result && n <> 0.0 else result || n <> 0.0)
             | other ->
                match coerceBool other with
                | Ok b ->
                    seen <- true
                    result <- (if isAnd then result && b else result || b)
                | Result.Error e -> if isNull firstError then firstError <- e)
            remaining <- remaining.Tail
        if not (isNull firstError) then err firstError
        elif not seen then err Err.value
        else CellValue.Boolean result
    | 5 ->
        (match coerceBool (eval args.Head ctx) with
         | Ok b -> CellValue.Boolean (not b)
         | Result.Error e -> err e)
    | 6 ->
        let value = eval args.Head ctx
        (match unwrapValue value with
         | CellValue.Error _ -> eval args.Tail.Head ctx
         | _ -> value)
    | 7 | 8 | 9 ->
        (match unwrapValue (eval args.Head ctx) with
         | CellValue.Number serial ->
            let parts = datePartsOfSerial (date1904Of ctx) serial
            if parts < 0 then Unchecked.defaultof<CellValue>
            elif code = 7 then CellValue.Number (float (parts / 10000))
            elif code = 8 then CellValue.Number (float (parts / 100 % 100))
            else CellValue.Number (float (parts % 100))
         | CellValue.Error e -> err e
         | _ -> Unchecked.defaultof<CellValue>)
    | 10 | 11 ->
        (match unwrapValue (eval args.Head ctx) with
         | CellValue.Error e -> err e
         | CellValue.Number serial ->
            (match unwrapValue (eval args.Tail.Head ctx) with
             | CellValue.Error e -> err e
             | CellValue.Number months ->
                // (As `int months` reads it in both runtimes: toward zero, in the range the function takes.)
                if Double.IsNaN months || months > 120000.0 || months < -120000.0 then Unchecked.defaultof<CellValue>
                else
                    match monthsFromSerial (date1904Of ctx) serial (int months) (code = 10) with
                    | ValueSome moved -> CellValue.Number moved
                    | ValueNone -> Unchecked.defaultof<CellValue>
             | _ -> Unchecked.defaultof<CellValue>)
         | _ -> Unchecked.defaultof<CellValue>)
    | _ -> Unchecked.defaultof<CellValue>

/// Do any of a call's arguments need rewriting before the function sees
/// them? Functions recognise ranges by their syntax (`RangeRef`), so an
/// argument that only *evaluates* to a range has to be turned into one.
and argsNeedPreparing (name: string) (args: Expr list) : bool =
    if args.IsEmpty then false
    else
        // (The arguments are those of one call: the answer is kept with them, for the function they were asked for.)
        let known = formTableGet argsNeedPreparingTable (box args)
        let kept = if isNull known then ValueNone else (let asked, found = unbox<string * bool> known in if obj.ReferenceEquals(asked, name) || asked = name then ValueSome found else ValueNone)
        match kept with
        | ValueSome found -> found
        | ValueNone ->
            let found = argsNeedPreparingNow name args
            formTableSet argsNeedPreparingTable (box args) (box (name, found))
            found

and argsNeedPreparingNow (name: string) (args: Expr list) : bool =
    let mutable found = false
    let mutable remaining = args
    while not found && not remaining.IsEmpty do
        match remaining.Head with
        | RangeRef (a, b) when isOpenRange a b || isThreeDSheet a.Sheet -> found <- true
        | CellRef r when isThreeDSheet r.Sheet -> found <- true
        | Function (inner, _) when isReferenceFunction inner || isIndexFunction inner -> found <- true
        | _ -> ()
        remaining <- remaining.Tail
    if not found && takesArrayParam name then
        // A lookup table given as anything but a plain range has to be
        // turned into something the function can index.
        for arg in args do
            match arg with
            | RangeRef _ | CellRef _ | Number _ | Str _ | Bool _ | MissingArg | ErrorLit _ -> ()
            | _ -> found <- true
    found

/// Rewrite the arguments that stand for a range: whole-column/row references
/// are bounded by the used range, OFFSET/INDIRECT/INDEX calls and reference
/// operators are replaced by the reference they return (so
/// `SUM(OFFSET(A1,0,0,2,1))` sums two cells), a 3-D reference becomes the
/// union of its sheets, and an array handed to an array-or-range parameter is
/// evaluated so the function can read it by position.
and prepareArgs (name: string) (keepOpenRanges: bool) (args: Expr list) (ctx: EvalContext) : Expr list =
    let signature = if takesArrayParam name then functionSignature name else ValueNone
    args
    |> List.mapi (fun index arg ->
        let prepared =
            match arg with
            | RangeRef (a, _) when isThreeDSheet a.Sheet -> expandThreeD arg ctx
            | CellRef r when isThreeDSheet r.Sheet -> expandThreeD arg ctx
            | RangeRef _ when keepOpenRanges -> arg
            | RangeRef (a, b) when isOpenRange a b ->
                match resolveModel a.Sheet ctx.Model ctx.Workbook with
                | Some m -> RangeRef (clampOpenRange m a b)
                | None -> ErrorLit Err.ref
            | Function (inner, innerArgs) when isReferenceFunction inner -> resolveReferenceFunction inner innerArgs ctx
            | Function (inner, innerArgs) when isIndexFunction inner ->
                (match indexAsReference innerArgs ctx with
                 | Some reference -> reference
                 | None -> arg)
            | _ -> arg
        match signature with
        | ValueSome fnSignature when paramKindAt fnSignature index = ArrayParam ->
            match prepared with
            | RangeRef _ | CellRef _ | ArrayLit _ | MissingArg | ErrorLit _ -> prepared
            | SpillRef anchor -> RangeRef (spillRangeIn anchor ctx.Model ctx.Workbook)
            | NameRef n ->
                (match ctx.Workbook |> Option.bind (resolveNameForEval n ctx.Sheet) with
                 | Some (RangeRef _ as range) -> range
                 | Some (CellRef _ as cell) -> cell
                 | Some resolved when isArrayExpr resolved -> ArrayLit (evalArray resolved ctx)
                 | _ -> prepared)
            | _ when isArrayExpr prepared -> ArrayLit (evalArray prepared ctx)
            | _ -> prepared
        | _ -> prepared)

/// A 3-D reference (`Sheet1:Sheet3!A1:B2`) as the union of the same range on
/// every sheet from the first to the last, in tab order.
and expandThreeD (reference: Expr) (ctx: EvalContext) : Expr =
    let span =
        match reference with
        | CellRef r -> r.Sheet
        | RangeRef (a, _) -> a.Sheet
        | _ -> None
    match span, ctx.Workbook with
    | Some names, Some wb ->
        let cut = names.IndexOf ':'
        let first = names.Substring(0, cut)
        let last = names.Substring(cut + 1)
        let position (sheetName: string) =
            Rbt.Spreadsheet.Workbook.resolveSheetName sheetName wb
            |> Option.bind (fun id -> wb.SheetOrder |> List.tryFindIndex (fun other -> other = id))
        match position first, position last with
        | Some i, Some j ->
            let onSheet (r: RefAst) (sheetName: string) = { r with Sheet = Some sheetName }
            let parts =
                [ for index in min i j .. max i j do
                    match Map.tryFind (List.item index wb.SheetOrder) wb.Sheets with
                    | Some sheet ->
                        match reference with
                        | CellRef r -> yield CellRef (onSheet r sheet.Name)
                        | RangeRef (a, b) ->
                            let a = onSheet a sheet.Name
                            let b = onSheet b sheet.Name
                            yield (if isOpenRange a b then RangeRef (clampOpenRange (resolveSheet sheet.Model) a b) else RangeRef (a, b))
                        | _ -> ()
                    | None -> () ]
            Function (unionOperatorName, parts)
        | _ -> ErrorLit Err.ref
    | _ -> ErrorLit Err.ref

/// The rectangle a reference expression denotes: (sheet, top, left, bottom,
/// right), or the error it evaluates to.
and referenceRect (reference: Expr) (ctx: EvalContext) : Result<string option * int * int * int * int, string> =
    match reference with
    | CellRef r -> Ok (r.Sheet, r.Row, r.Col, r.Row, r.Col)
    | RangeRef (a, b) ->
        let sr, sc, er, ec = normalizeRange a b
        Ok (a.Sheet, sr, sc, er, ec)
    | TableRef sel ->
        (match tableSelArea sel ctx.Workbook ctx.Sheet ctx.CurrentCell with
         | Ok (r0, c0, r1, c1) when r0 <= r1 -> Ok (None, r0, c0, r1, c1)
         | Ok _ -> Result.Error Err.ref
         | Result.Error e -> Result.Error e)
    | ErrorLit e -> Result.Error e
    | _ -> Result.Error Err.ref

/// Is this argument written as a reference (so that INDEX of it is one too)?
and isReferenceSyntax (e: Expr) : bool =
    match e with
    | CellRef _ | RangeRef _ | SpillRef _ | TableRef _ | NameRef _ -> true
    | Function (name, _) -> isReferenceFunction name || isIndexFunction name
    | _ -> false

/// INDEX(reference, row, [col]) as the reference it selects — a cell, or a
/// whole row / column of the reference when the other index is 0 or omitted.
/// None when the first argument is not a reference (INDEX of an array is a
/// value, handled by `evalArray`).
and indexAsReference (args: Expr list) (ctx: EvalContext) : Expr option =
    match args with
    | source :: indexes when isReferenceSyntax source && not indexes.IsEmpty && indexes.Length <= 2 ->
        let reference =
            match source with
            | Function (inner, innerArgs) when isIndexFunction inner ->
                (match indexAsReference innerArgs ctx with Some r -> r | None -> ErrorLit Err.ref)
            | _ -> resolveReferenceFunction "" [ source ] ctx
        match referenceRect reference ctx with
        | Result.Error e -> Some (ErrorLit e)
        | Ok (sheet, top, left, bottom, right) ->
            let rows = bottom - top + 1
            let cols = right - left + 1
            let numberOf (e: Expr) =
                match e with
                | MissingArg -> Ok 0
                | _ -> (match coerceNumber (eval e ctx) with Ok n -> Ok (int (Math.Truncate n)) | Result.Error code -> Result.Error code)
            let picked =
                match indexes with
                | [ only ] ->
                    (match numberOf only with
                     | Result.Error e -> Result.Error e
                     // A lone index walks a single row or column; on a block it selects a row.
                     | Ok n -> if rows = 1 && cols > 1 then Ok (1, n) elif cols = 1 then Ok (n, 1) else Ok (n, 0))
                | [ rowIndex; colIndex ] ->
                    (match numberOf rowIndex, numberOf colIndex with
                     | Result.Error e, _ | _, Result.Error e -> Result.Error e
                     | Ok r, Ok c -> Ok (r, c))
                | _ -> Result.Error Err.value
            match picked with
            | Result.Error e -> Some (ErrorLit e)
            | Ok (r, c) ->
                if r < 0 || c < 0 then Some (ErrorLit Err.value)
                elif r > rows || c > cols then Some (ErrorLit Err.ref)
                else
                    let r0, r1 = if r = 0 then top, bottom else top + r - 1, top + r - 1
                    let c0, c1 = if c = 0 then left, right else left + c - 1, left + c - 1
                    let corner row col = { Sheet = sheet; Row = row; Col = col; RowAbs = false; ColAbs = false }
                    if r0 = r1 && c0 = c1 then Some (CellRef (corner r0 c0))
                    else Some (RangeRef (corner r0 c0, corner r1 c1))
    | _ -> None

/// INDEX in an array context: a reference when its source is one, otherwise
/// the selected element, row or column of the source array.
and evalIndexArray (args: Expr list) (ctx: EvalContext) : ArrVal =
    let scalar v : ArrVal = [| [| v |] |]
    match indexAsReference args ctx with
    | Some (ErrorLit e) -> scalar (err e)
    | Some reference -> evalArray reference ctx
    | None ->
        match args with
        | source :: indexes when not indexes.IsEmpty && indexes.Length <= 2 ->
            let arr = evalArray source ctx
            let rows = arr.Length
            let cols = if rows > 0 then arr.[0].Length else 0
            let numberOf (e: Expr) =
                match e with
                | MissingArg -> Ok 0
                | _ -> (match coerceNumber (eval e ctx) with Ok n -> Ok (int (Math.Truncate n)) | Result.Error code -> Result.Error code)
            let picked =
                match indexes with
                | [ only ] ->
                    (match numberOf only with
                     | Result.Error e -> Result.Error e
                     | Ok n -> if rows = 1 && cols > 1 then Ok (1, n) elif cols = 1 then Ok (n, 1) else Ok (n, 0))
                | [ rowIndex; colIndex ] ->
                    (match numberOf rowIndex, numberOf colIndex with
                     | Result.Error e, _ | _, Result.Error e -> Result.Error e
                     | Ok r, Ok c -> Ok (r, c))
                | _ -> Result.Error Err.value
            match picked with
            | Result.Error e -> scalar (err e)
            | Ok (r, c) ->
                if rows = 0 || cols = 0 || r < 0 || c < 0 then scalar (err Err.value)
                elif r > rows || c > cols then scalar (err Err.ref)
                else
                    let r0, r1 = if r = 0 then 0, rows - 1 else r - 1, r - 1
                    let c0, c1 = if c = 0 then 0, cols - 1 else c - 1, c - 1
                    [| for row in r0 .. r1 -> [| for col in c0 .. c1 -> arr.[row].[col] |] |]
        | _ -> scalar (err Err.value)

/// Evaluate OFFSET / INDIRECT to the reference it denotes: a `CellRef`, a
/// `RangeRef` (already bounded), a `TableRef`, or an error literal.
and resolveReferenceFunction (name: string) (args: Expr list) (ctx: EvalContext) : Expr =
    let args = List.toArray args
    let bounded (e: Expr) =
        match e with
        | RangeRef (a, b) when isOpenRange a b ->
            match resolveModel a.Sheet ctx.Model ctx.Workbook with
            | Some m -> RangeRef (clampOpenRange m a b)
            | None -> ErrorLit Err.ref
        | _ -> e
    let rec asReference (e: Expr) : Expr =
        match e with
        | CellRef _ | TableRef _ -> e
        | RangeRef _ -> bounded e
        | SpillRef anchor -> RangeRef (spillRangeIn anchor ctx.Model ctx.Workbook)
        | NameRef n ->
            match ctx.Workbook |> Option.bind (resolveNameForEval n ctx.Sheet) with
            | Some resolved -> asReference resolved
            | None -> ErrorLit Err.name
        | Function (inner, innerArgs) when isReferenceFunction inner -> resolveReferenceFunction inner innerArgs ctx
        | Function (inner, innerArgs) when isIndexFunction inner ->
            (match indexAsReference innerArgs ctx with
             | Some reference -> reference
             | None -> ErrorLit Err.ref)
        | ErrorLit _ -> e
        | _ -> ErrorLit Err.ref
    if name = "" then
        // Not a call: just the reference an expression stands for.
        if args.Length = 1 then asReference args.[0] else ErrorLit Err.ref
    elif name = rangeOperatorName || name = intersectionOperatorName then
        if args.Length <> 2 then ErrorLit Err.value
        else
            match referenceRect (asReference args.[0]) ctx, referenceRect (asReference args.[1]) ctx with
            | Result.Error e, _ | _, Result.Error e -> ErrorLit e
            | Ok (sheetA, topA, leftA, bottomA, rightA), Ok (sheetB, topB, leftB, bottomB, rightB) ->
                // The right operand inherits the left one's sheet (`S!A1:B2`).
                let sameSheet =
                    match sheetA, sheetB with
                    | _, None -> true
                    | Some a, Some b -> equalsOrdinalIgnoreCase a b
                    | None, Some _ -> false
                if not sameSheet then ErrorLit Err.value
                else
                    let corner row col = { Sheet = sheetA; Row = row; Col = col; RowAbs = false; ColAbs = false }
                    let top, left, bottom, right =
                        if name = rangeOperatorName then
                            min topA topB, min leftA leftB, max bottomA bottomB, max rightA rightB
                        else
                            max topA topB, max leftA leftB, min bottomA bottomB, min rightA rightB
                    if top > bottom || left > right then ErrorLit Err.nul
                    elif top = bottom && left = right then CellRef (corner top left)
                    else RangeRef (corner top left, corner bottom right)
    elif equalsOrdinalIgnoreCase name "INDIRECT" then
        if args.Length < 1 then ErrorLit Err.value
        else
            let a1Style =
                if args.Length < 2 then true
                else
                    match args.[1] with
                    | MissingArg -> true
                    | style -> (match coerceBool (eval style ctx) with Ok flag -> flag | _ -> true)
            match coerceText (eval args.[0] ctx) with
            | Result.Error e -> ErrorLit e
            | Ok refText when not a1Style ->
                (match parseR1C1Reference refText ctx.CurrentCell with
                 | Some target -> asReference target
                 | None -> ErrorLit Err.ref)
            | Ok refText ->
                match parseFormulaShared ("=" + refText) with
                | Ok (CellRef _ as target) | Ok (RangeRef _ as target) | Ok (TableRef _ as target) | Ok (NameRef _ as target)
                | Ok (SpillRef _ as target) ->
                    (match asReference target with
                     | ErrorLit _ -> ErrorLit Err.ref
                     | reference -> reference)
                | _ -> ErrorLit Err.ref
    else
        // OFFSET(reference, rows, cols, [height], [width])
        if args.Length < 3 then ErrorLit Err.value
        else
            let numberAt i (fallback: float) =
                if i >= args.Length then Ok fallback
                else
                    match args.[i] with
                    | MissingArg -> Ok fallback
                    | e -> coerceNumber (eval e ctx)
            let basis =
                match asReference args.[0] with
                | CellRef r -> Ok (r.Sheet, r.Row, r.Col, 1, 1)
                | RangeRef (a, b) ->
                    let sr, sc, er, ec = normalizeRange a b
                    Ok (a.Sheet, sr, sc, er - sr + 1, ec - sc + 1)
                | ErrorLit e -> Result.Error e
                | _ -> Result.Error Err.ref
            match basis with
            | Result.Error e -> ErrorLit e
            | Ok (sheet, baseRow, baseCol, baseHeight, baseWidth) ->
                match numberAt 1 0.0, numberAt 2 0.0, numberAt 3 (float baseHeight), numberAt 4 (float baseWidth) with
                | Result.Error e, _, _, _ | _, Result.Error e, _, _ | _, _, Result.Error e, _ | _, _, _, Result.Error e -> ErrorLit e
                | Ok dr, Ok dc, Ok h, Ok w ->
                    let height = int (Math.Truncate h)
                    let width = int (Math.Truncate w)
                    if height = 0 || width = 0 then ErrorLit Err.ref
                    else
                        // A negative height/width extends up/left from the anchor.
                        let r0 = baseRow + int (Math.Truncate dr)
                        let c0 = baseCol + int (Math.Truncate dc)
                        let top = if height < 0 then r0 + height + 1 else r0
                        let left = if width < 0 then c0 + width + 1 else c0
                        let bottom = top + abs height - 1
                        let right = left + abs width - 1
                        if top < 0 || left < 0 || bottom > lastSheetRow || right > lastSheetCol then ErrorLit Err.ref
                        else
                            let corner row col = { Sheet = sheet; Row = row; Col = col; RowAbs = false; ColAbs = false }
                            if top = bottom && left = right then CellRef (corner top left)
                            else RangeRef (corner top left, corner bottom right)

and evalBinOpChain (expr: Expr) (ctx: EvalContext) : CellValue =
    // (One operator, or two: no list of them is made. A long chain - `A1+A2+...` nests to
    // the left as deep as it is long - is walked, not descended into.)
    match expr with
    | BinOp (left, op, right) when (match left with BinOp _ -> false | _ -> true) ->
        let lv = eval left ctx
        applyBin op lv (eval right ctx)
    | BinOp (BinOp (first, op1, second), op2, third) when (match first with BinOp _ -> false | _ -> true) ->
        let lv = eval first ctx
        let inner = applyBin op1 lv (eval second ctx)
        applyBin op2 inner (eval third ctx)
    | _ ->
    let operations = ResizeArray<struct (string * Expr)>()
    let mutable current = expr
    let mutable collecting = true
    while collecting do
        match current with
        | BinOp (left, op, right) ->
            operations.Add(struct (op, right))
            current <- left
        | _ ->
            collecting <- false
    let mutable result = eval current ctx
    for index = operations.Count - 1 downto 0 do
        let struct (op, right) = operations.[index]
        result <- applyBin op result (eval right ctx)
    result

and tryEvalArrayTopLeft (name: string) (args: Expr list) (ctx: EvalContext) : CellValue voption =
    let argsArray = List.toArray args
    let nargs = argsArray.Length
    let argExpr i = argsArray.[i]
    let tryNumber i =
        if i < nargs then coerceNumber (eval (argExpr i) ctx) else Result.Error Err.value
    match name.ToUpperInvariant() with
    | "SEQUENCE" ->
        if nargs < 1 then ValueSome (err Err.value) else
        match tryNumber 0 with
        | Result.Error e -> ValueSome (err e)
        | Ok rowsN ->
            // (An argument left empty - `SEQUENCE(20,,0)` - is its default.)
            let colsN =
                if nargs >= 2 && argExpr 1 <> MissingArg then tryNumber 1 else Ok 1.0
            let startN =
                if nargs >= 3 && argExpr 2 <> MissingArg then tryNumber 2 else Ok 1.0
            match colsN, startN with
            | Result.Error e, _ | _, Result.Error e -> ValueSome (err e)
            | Ok colsN, Ok start ->
                let rows = int rowsN
                let cols = int colsN
                if rows < 1 || cols < 1 then ValueSome (err Err.value)
                else ValueSome (CellValue.Number start)
    | "RANDARRAY" ->
        let numberOrDefault index defaultValue =
            if index < nargs then coerceNumber (eval (argExpr index) ctx)
            else Ok defaultValue
        let boolOrDefault index defaultValue =
            if index < nargs then coerceBool (eval (argExpr index) ctx)
            else Ok defaultValue
        match numberOrDefault 0 1.0,
              numberOrDefault 1 1.0,
              numberOrDefault 2 0.0,
              numberOrDefault 3 1.0,
              boolOrDefault 4 false with
        | Result.Error e, _, _, _, _
        | _, Result.Error e, _, _, _
        | _, _, Result.Error e, _, _
        | _, _, _, Result.Error e, _
        | _, _, _, _, Result.Error e -> ValueSome (err e)
        | Ok rowsN, Ok colsN, Ok lo, Ok hi, Ok whole ->
            let rows = int rowsN
            let cols = int colsN
            if rows < 1 || cols < 1 || hi < lo then
                ValueSome (err Err.value)
            elif whole then
                ValueSome (CellValue.Number (float (rng.Next(int lo, int hi + 1))))
            else
                ValueSome (CellValue.Number (lo + rng.NextDouble() * (hi - lo)))
    | "TRANSPOSE" ->
        if nargs < 1 then ValueSome (err Err.value)
        else ValueSome (eval (argExpr 0) ctx)
    | "UNIQUE" ->
        if nargs < 1 then ValueSome (err Err.value)
        else ValueSome (eval (argExpr 0) ctx)
    | "TEXTSPLIT" ->
        if nargs < 2 then ValueSome (err Err.value) else
        let textOf i = match coerceText (eval (argExpr i) ctx) with Ok t -> t | _ -> ""
        let text = textOf 0
        let colDelimiter = textOf 1
        let rowText =
            if nargs >= 3 then
                let rowDelimiter = textOf 2
                if rowDelimiter = "" then text
                else
                    let index = text.IndexOf(rowDelimiter)
                    if index < 0 then text else text.Substring(0, index)
            else
                text
        if colDelimiter = "" then
            ValueSome (CellValue.String rowText)
        else
            let index = rowText.IndexOf(colDelimiter)
            ValueSome (CellValue.String (if index < 0 then rowText else rowText.Substring(0, index)))
    | "MAP" ->
        if nargs < 2 then ValueSome (err Err.value) else
        match lambdaPartsOf (argExpr (nargs - 1)) with
        | None -> ValueSome (err Err.value)
        | Some [ _; body ] when containsVolatile body -> ValueNone
        | Some parts ->
            let firstValue = eval (argExpr 0) ctx
            ValueSome (eval (applyLambdaExpr parts [ valueToExpr firstValue ]) ctx)
    | "BYROW" ->
        if nargs < 2 then ValueSome (err Err.value) else
        match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0), lambdaPartsOf (argExpr 1) with
        | Some source, Some parts when source.W > 0 && source.Count >= source.W ->
            let row = Array.init source.W (fun c -> rangeRead source c)
            ValueSome (eval (applyLambdaExpr parts [ ArrayLit [| row |] ]) ctx)
        | _, None -> ValueSome (err Err.value)
        | _ -> ValueNone
    | "BYCOL" ->
        if nargs < 2 then ValueSome (err Err.value) else
        match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0), lambdaPartsOf (argExpr 1) with
        | Some source, Some [ NameRef param; Function (fn, [ NameRef arg ]) ]
            when source.W > 0
                 && source.Count >= source.W
                 && equalsOrdinalIgnoreCase param arg
                 && (equalsOrdinalIgnoreCase fn "SUM" || equalsOrdinalIgnoreCase fn "COUNT" || equalsOrdinalIgnoreCase fn "AVERAGE") ->
            let isCount = equalsOrdinalIgnoreCase fn "COUNT"
            let isAverage = equalsOrdinalIgnoreCase fn "AVERAGE"
            let rows = source.Count / source.W
            let mutable firstErr = ValueNone
            let mutable total = 0.0
            let mutable numericCount = 0
            let addValue value =
                if ValueOption.isNone firstErr then
                    match value with
                    | CellValue.Error e -> firstErr <- ValueSome e
                    | CellValue.Number n ->
                        total <- total + (if isCount then 1.0 else n)
                        if isAverage then
                            numericCount <- numericCount + 1
                    | CellValue.Boolean b when not isCount ->
                        total <- total + (if b then 1.0 else 0.0)
                        if isAverage then
                            numericCount <- numericCount + 1
                    | _ -> ()
            let area = int64 source.Count
            if area >= 1024L && int64 source.M.Cells.Count <= area * 4L then
                source.M.Cells
                |> Map.iter (fun addr cell ->
                    if ValueOption.isNone firstErr then
                        match tryRangeOffsetForCount source source.Count addr with
                        | ValueSome offset when offset % source.W = 0 ->
                            addValue (unwrapValue cell.Value)
                        | _ -> ())
            else
                let mutable r = 0
                while ValueOption.isNone firstErr && r < rows do
                    addValue (rangeRead source (r * source.W))
                    r <- r + 1
            match firstErr with
            | ValueSome e -> ValueSome (CellValue.Error e)
            | ValueNone when isAverage && numericCount = 0 -> ValueSome (err Err.div0)
            | ValueNone when isAverage -> ValueSome (CellValue.Number (total / float numericCount))
            | ValueNone -> ValueSome (CellValue.Number total)
        | _, None -> ValueSome (err Err.value)
        | _ -> ValueNone
    | "PIVOTBY" ->
        if nargs < 3 then ValueSome (err Err.value) else
        match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0),
              rangeAccessOf ctx.Model ctx.Workbook (argExpr 1),
              rangeAccessOf ctx.Model ctx.Workbook (argExpr 2) with
        | Some rowAccess, Some colAccess, Some valueAccess
            when rowAccess.Count > 0 && colAccess.Count > 0 && valueAccess.Count > 0 ->
            ValueSome (CellValue.String "Row Labels")
        | _ -> ValueNone
    | "GROUPBY" ->
        if nargs < 2 then ValueSome (err Err.value) else
        match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0),
              rangeAccessOf ctx.Model ctx.Workbook (argExpr 1) with
        | Some rowAccess, Some valueAccess
            when rowAccess.W = 1 && rowAccess.Count > 0 && valueAccess.Count > 0 ->
            let mutable best = ValueNone
            rowAccess.M.Cells
            |> Map.iter (fun addr cell ->
                match tryRangeOffsetForCount rowAccess rowAccess.Count addr with
                | ValueSome _ ->
                    let value = unwrapValue cell.Value
                    match best with
                    | ValueNone -> best <- ValueSome value
                    | ValueSome current when compareCellValues value current < 0 -> best <- ValueSome value
                    | _ -> ()
                | ValueNone -> ())
            ValueSome (match best with ValueSome value -> value | ValueNone -> CellValue.Empty)
        | _ -> ValueNone
    | "SCAN" ->
        if nargs <> 3 then ValueSome (err Err.value) else
        match lambdaPartsOf (argExpr 2) with
        | None -> ValueSome (err Err.value)
        | Some [ _; _; body ] when containsVolatile body -> ValueNone
        | Some parts ->
            let initial = eval (argExpr 0) ctx
            let firstValue = eval (argExpr 1) ctx
            ValueSome (eval (applyLambdaExpr parts [ valueToExpr initial; valueToExpr firstValue ]) ctx)
    | "MAKEARRAY" ->
        if nargs <> 3 then ValueSome (err Err.value) else
        match lambdaPartsOf (argExpr 2) with
        | None -> ValueSome (err Err.value)
        | Some [ _; _; body ] when containsVolatile body -> ValueNone
        | Some parts ->
            match tryNumber 0, tryNumber 1 with
            | Result.Error e, _ | _, Result.Error e -> ValueSome (err e)
            | Ok rowsN, Ok colsN ->
                if int rowsN < 1 || int colsN < 1 then ValueSome (err Err.value)
                else ValueSome (eval (applyLambdaExpr parts [ Number 1.0; Number 1.0 ]) ctx)
    | "FILTER" ->
        if nargs < 2 then ValueSome (err Err.value) else
        let includeAccessAndPredicate includeExpr =
            let coercePredicate value =
                match coerceBool value with
                | Ok true -> true
                | _ -> false
            let comparisonPredicate op left right =
                match compareOperandsNoError left right with
                | ValueSome c ->
                    match op with
                    | "=" -> c = 0
                    | "<>" -> c <> 0
                    | "<" -> c < 0
                    | ">" -> c > 0
                    | "<=" -> c <= 0
                    | ">=" -> c >= 0
                    | _ -> coercePredicate (applyBin op left right)
                | ValueNone -> false
            let readerFor access predicate =
                if access.W <> 1 then ValueNone
                else
                    if predicate CellValue.Empty then ValueNone
                    else ValueSome (access, fun rowIndex -> predicate (rangeRead access rowIndex))
            match includeExpr with
            | BinOp (leftExpr, op, rightExpr) ->
                match rangeAccessOf ctx.Model ctx.Workbook leftExpr, rangeAccessOf ctx.Model ctx.Workbook rightExpr with
                | Some access, None ->
                    let scalar = eval rightExpr ctx
                    readerFor access (fun value -> comparisonPredicate op value scalar)
                | None, Some access ->
                    let scalar = eval leftExpr ctx
                    readerFor access (fun value -> comparisonPredicate op scalar value)
                | _ -> ValueNone
            | _ ->
                match rangeAccessOf ctx.Model ctx.Workbook includeExpr with
                | Some access -> readerFor access coercePredicate
                | None -> ValueNone
        match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0), includeAccessAndPredicate (argExpr 1) with
        | Some source, ValueSome (includeAccess, includeAt) when source.Count >= source.W && includeAccess.W = 1 ->
            let rows = min (source.Count / source.W) includeAccess.Count
            if rows <= 0 then
                ValueSome (if nargs >= 3 then eval (argExpr 2) ctx else err Err.calc)
            else
                let mutable firstRow = ValueNone
                let area = int64 includeAccess.Count
                if area >= 1024L && int64 includeAccess.M.Cells.Count < area then
                    includeAccess.M.Cells
                    |> Map.iter (fun addr _ ->
                        match tryRangeOffsetForCount includeAccess includeAccess.Count addr with
                        | ValueSome rowIndex when rowIndex < rows ->
                            match firstRow, includeAt rowIndex with
                            | ValueNone, true -> firstRow <- ValueSome rowIndex
                            | ValueSome existing, true when rowIndex < existing -> firstRow <- ValueSome rowIndex
                            | _ -> ()
                        | _ -> ())
                else
                    let mutable r = 0
                    while ValueOption.isNone firstRow && r < rows do
                        if includeAt r then
                            firstRow <- ValueSome r
                        r <- r + 1
                match firstRow with
                | ValueSome rowIndex -> ValueSome (rangeRead source (rowIndex * source.W))
                | ValueNone -> ValueSome (if nargs >= 3 then eval (argExpr 2) ctx else err Err.calc)
        | _ -> ValueNone
    | "SORT" ->
        if nargs < 1 then ValueSome (err Err.value) else
        match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0) with
        | Some source when source.W = 1 ->
            let idx =
                if nargs >= 2 then
                    match tryNumber 1 with
                    | Ok n -> int n
                    | Result.Error _ -> 1
                else
                    1
            let order =
                if nargs >= 3 then
                    match tryNumber 2 with
                    | Ok n -> int n
                    | Result.Error _ -> 1
                else
                    1
            let byCol =
                if nargs >= 4 then
                    match coerceBool (eval (argExpr 3) ctx) with
                    | Ok b -> b
                    | _ -> false
                else
                    false
            if idx <> 1 || byCol || source.Count <= 0 then
                ValueNone
            else
                let descending = order < 0
                let mutable best = ValueNone
                let mutable ambiguous = false
                let consider value =
                    if not ambiguous then
                        match unwrapValue value with
                        | CellValue.Error _ -> ambiguous <- true
                        | v ->
                            match best with
                            | ValueNone -> best <- ValueSome v
                            | ValueSome current ->
                                match compareValues v current with
                                | Result.Error _ -> ambiguous <- true
                                | Ok c when (descending && c > 0) || (not descending && c < 0) -> best <- ValueSome v
                                | Ok 0 when v <> current -> ambiguous <- true
                                | _ -> ()
                let area = int64 source.Count
                if area >= 1024L && int64 source.M.Cells.Count <= area * 4L then
                    let mutable storedInRange = 0
                    source.M.Cells
                    |> Map.iter (fun addr cell ->
                        match tryRangeOffsetForCount source source.Count addr with
                        | ValueSome _ ->
                            storedInRange <- storedInRange + 1
                            consider (unwrapValue cell.Value)
                        | ValueNone -> ())
                    if storedInRange < source.Count then
                        consider CellValue.Empty
                else
                    for k = 0 to source.Count - 1 do
                        consider (rangeRead source k)
                match best with
                | ValueSome value when not ambiguous -> ValueSome value
                | _ -> ValueNone
        | _ -> ValueNone
    | "SORTBY" ->
        if nargs < 2 then ValueSome (err Err.value) else
        match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0) with
        | Some source when source.W > 0 && source.Count >= source.W ->
            let rows = source.Count / source.W
            let keys = ResizeArray<struct (RangeAccess * int)>()
            let mutable valid = true
            let mutable i = 1
            while valid && i < nargs do
                match rangeAccessOf ctx.Model ctx.Workbook (argExpr i) with
                | Some key when key.W = 1 ->
                    let order =
                        if i + 1 < nargs then
                            match tryNumber (i + 1) with
                            | Ok n -> int n
                            | Result.Error _ ->
                                valid <- false
                                1
                        else
                            1
                    keys.Add(struct (key, order))
                    i <- i + 2
                | _ ->
                    valid <- false
            if not valid || rows <= 0 || keys.Count = 0 then
                ValueNone
            else
                let keyAt (key: RangeAccess) row =
                    if row < key.Count then rangeRead key row else CellValue.Empty
                let bestKeyValues = Array.zeroCreate<CellValue> keys.Count
                for keyIndex = 0 to keys.Count - 1 do
                    let struct (key, _) = keys.[keyIndex]
                    bestKeyValues.[keyIndex] <- keyAt key 0
                let candidateKeyValues = Array.zeroCreate<CellValue> keys.Count
                let mutable bestRow = 0
                for row = 1 to rows - 1 do
                    let mutable result = 0
                    let mutable keyIndex = 0
                    let mutable readCount = 0
                    while result = 0 && keyIndex < keys.Count do
                        let struct (key, order) = keys.[keyIndex]
                        let candidateValue = keyAt key row
                        candidateKeyValues.[keyIndex] <- candidateValue
                        readCount <- keyIndex + 1
                        let c =
                            match compareValuesNoError candidateValue bestKeyValues.[keyIndex] with
                            | ValueSome c -> c
                            | ValueNone -> 0
                        result <- if order < 0 then -c else c
                        keyIndex <- keyIndex + 1
                    if result < 0 then
                        bestRow <- row
                        for keyIndex = readCount to keys.Count - 1 do
                            let struct (key, _) = keys.[keyIndex]
                            candidateKeyValues.[keyIndex] <- keyAt key row
                        Array.Copy(candidateKeyValues, bestKeyValues, keys.Count)
                ValueSome (rangeRead source (bestRow * source.W))
        | _ -> ValueNone
    | _ -> ValueNone

/// Evaluate every argument, flattening ranges/tables into their cell values
/// (the form most aggregate functions want).
and iterFlattenedArg (arg: Expr) (ctx: EvalContext) (emit: CellValue -> unit) : unit =
    let model = ctx.Model
    let workbook = ctx.Workbook
    match arg with
    | RangeRef (a, b) ->
        iterRangeRef a b model workbook emit
    | SpillRef anchor ->
        let a, b = spillRangeIn anchor model workbook
        iterRangeRef a b model workbook emit
    | TableRef sel ->
        iterTableSel sel workbook ctx.Sheet ctx.CurrentCell model emit
    | _ when isArrayExpr arg ->
        for row in evalArray arg ctx do
            for value in row do
                emit (unwrapValue value)
    | NameRef name ->
        // A name standing in for a range must flatten like the range
        // would — otherwise SUM(MyRange) would see only its first cell.
        match workbook |> Option.bind (resolveNameForEval name ctx.Sheet) with
        | Some (RangeRef (a, b)) ->
            iterRangeRef a b model workbook emit
        | Some (TableRef sel) ->
            iterTableSel sel workbook ctx.Sheet ctx.CurrentCell model emit
        | Some e -> emit (eval e ctx)
        | None -> emit (err Err.name)
    | _ -> emit (eval arg ctx)

and iterFlattenedArgs (args: Expr list) (ctx: EvalContext) (emit: CellValue -> unit) : unit =
    for arg in args do
        iterFlattenedArg arg ctx emit

and iterFlattenedArgIgnoringMissingCells (arg: Expr) (ctx: EvalContext) (emit: CellValue -> unit) : unit =
    let model = ctx.Model
    let workbook = ctx.Workbook
    let iterRange (a: RefAst) (b: RefAst) =
        match resolveModel a.Sheet model workbook with
        | Some m ->
            let sr, sc, er, ec = normalizeRange a b
            let area = int64 (er - sr + 1) * int64 (ec - sc + 1)
            if area >= 1024L && int64 m.Cells.Count <= area * 4L then
                m.Cells
                |> Map.iter (fun addr cell ->
                    if addr.Row >= sr && addr.Row <= er && addr.Col >= sc && addr.Col <= ec then
                        emit (unwrapValue cell.Value))
            else
                let mutable e = ({ Model = m; SR = sr; SC = sc; ER = er; EC = ec }: RangeView).GetEnumerator()
                while e.MoveNext() do
                    match e.Current with
                    | CellValue.Empty -> ()
                    | value -> emit value
        | None -> emit (err Err.ref)
    match arg with
    | RangeRef (a, b) -> iterRange a b
    | SpillRef anchor -> let a, b = spillRangeIn anchor model workbook in iterRange a b
    | TableRef sel ->
        iterTableSel sel workbook ctx.Sheet ctx.CurrentCell model (fun value ->
            match value with
            | CellValue.Empty -> ()
            | _ -> emit value)
    | _ when isArrayExpr arg ->
        for row in evalArray arg ctx do
            for value in row do
                match unwrapValue value with
                | CellValue.Empty -> ()
                | v -> emit v
    | NameRef name ->
        match workbook |> Option.bind (resolveNameForEval name ctx.Sheet) with
        | Some (RangeRef (a, b)) -> iterRange a b
        | Some (TableRef sel) ->
            iterTableSel sel workbook ctx.Sheet ctx.CurrentCell model (fun value ->
                match value with
                | CellValue.Empty -> ()
                | _ -> emit value)
        | Some inner -> emit (eval inner ctx)
        | None -> emit (err Err.name)
    | _ -> emit (eval arg ctx)

and iterFlattenedArgsIgnoringMissingCells (args: Expr list) (ctx: EvalContext) (emit: CellValue -> unit) : unit =
    for arg in args do
        iterFlattenedArgIgnoringMissingCells arg ctx emit

and flattenArgs (args: Expr list) (ctx: EvalContext) : CellValue list =
    let values = ResizeArray<CellValue>()
    iterFlattenedArgs args ctx (fun value -> values.Add value)
    resizeArrayToList values

/// Evaluate an expression to a 2-D array (a "spill"). Array functions build
/// their result; a range materializes to its grid; operators broadcast
/// element-wise; anything else is a 1×1 of its scalar value.
and evalArray (expr: Expr) (ctx: EvalContext) : ArrVal =
    let scalar v : ArrVal = [| [| v |] |]
    match expr with
    | Function (name, args) when isArrayFunction name ->
        let args = if argsNeedPreparing name args then prepareArgs name false args ctx else args
        evalArrayFunction (name.ToUpperInvariant()) args ctx
    | Function (name, args) when isReferenceFunction name ->
        match resolveReferenceFunction name args ctx with
        | ErrorLit e -> scalar (err e)
        | reference -> evalArray reference ctx
    | Function (name, parts) when name = unionOperatorName ->
        // A union is several areas; as a value it is all their cells, one
        // after another, which is what every aggregate wants from it.
        let values = ResizeArray<CellValue[]>()
        for part in parts do
            for row in evalArray part ctx do
                for value in row do
                    values.Add [| value |]
        if values.Count = 0 then scalar (err Err.value) else values.ToArray()
    | Function (name, args) when isIndexFunction name -> evalIndexArray args ctx
    | Function (name, [ RangeRef _ ])
        when equalsOrdinalIgnoreCase name "ROW" || equalsOrdinalIgnoreCase name "COLUMN" ->
        // Deliberately not bounded: ROW(A:A) is rarely meant, and shape matters here.
        (match expr with
         | Function (_, args) -> evalExtendedArrayFunction (name.ToUpperInvariant()) args ctx
         | _ -> scalar (eval expr ctx))
    | Function (name, args) when isArrayExpr expr ->
        match functionSignature name with
        | ValueNone -> scalar (eval expr ctx)
        | ValueSome signature ->
        // Lifting: a scalar function applied element by element over the
        // arrays given in its scalar slots, broadcast to a common shape.
        // Other arguments stay unevaluated, so IF remains lazy per element
        // and range parameters are still handed over whole.
        let args = (if argsNeedPreparing name args then prepareArgs name false args ctx else args) |> List.toArray
        let arrays =
            args |> Array.mapi (fun i arg ->
                if paramKindAt signature i = ScalarParam && isArrayExpr arg then Some (evalArray arg ctx) else None)
        let mutable rows = 1
        let mutable cols = 1
        for arr in arrays do
            match arr with
            | Some a ->
                rows <- max rows a.Length
                cols <- max cols (if a.Length > 0 then a.[0].Length else 0)
            | None -> ()
        // Broadcasting: a single row or column repeats along the other axis,
        // a 1×1 everywhere; a position an array does not cover is #N/A.
        let element (a: ArrVal) r c =
            let ar = a.Length
            let ac = if ar > 0 then a.[0].Length else 0
            if ar = 0 || ac = 0 then err Err.value
            elif (ar <> 1 && r >= ar) || (ac <> 1 && c >= ac) then err Err.na
            else a.[(if ar = 1 then 0 else r)].[(if ac = 1 then 0 else c)]
        Array.init rows (fun r ->
            Array.init cols (fun c ->
                let elementArgs =
                    [ for i in 0 .. args.Length - 1 ->
                        match arrays.[i] with
                        | Some a -> ArrayLit [| [| element a r c |] |]
                        | None -> args.[i] ]
                evalFunction name elementArgs ctx))
    | CellRef r when isThreeDSheet r.Sheet -> evalArray (expandThreeD expr ctx) ctx
    | RangeRef (a, _) when isThreeDSheet a.Sheet -> evalArray (expandThreeD expr ctx) ctx
    | CellRef _ -> scalar (eval expr ctx)
    | RangeRef (a, b) ->
        match resolveModel a.Sheet ctx.Model ctx.Workbook with
        | Some m ->
            let a, b = clampOpenRange m a b
            let sr, sc, er, ec = normalizeRange a b
            [| for r in sr .. er -> [| for c in sc .. ec -> getCellValue { Row = r; Col = c } m |> unwrapValue |] |]
        | None -> scalar (err Err.ref)
    | TableRef sel -> [| for v in expandTableSel sel ctx.Workbook ctx.Sheet ctx.CurrentCell ctx.Model -> [| v |] |]
    | ArrayLit a -> a
    | SpillRef anchor -> let a, b = spillRangeIn anchor ctx.Model ctx.Workbook in evalArray (RangeRef (a, b)) ctx
    | NameRef name ->
        match ctx.Workbook |> Option.bind (resolveNameForEval name ctx.Sheet) with
        | Some (RangeRef (a, b)) -> evalArray (RangeRef (a, b)) ctx
        | Some (TableRef sel) -> evalArray (TableRef sel) ctx
        | Some e when isArrayExpr e ->
            let key = name.ToUpperInvariant()
            if namesBeingEvaluated.Contains key then scalar (err Err.cycle)
            else
                namesBeingEvaluated.Add key
                try evalArray e ctx
                finally Portable.Stack.dropLast namesBeingEvaluated
        | Some e -> scalar (eval e ctx)
        | None -> scalar (err Err.name)
    | UnaryOp (op, inner) -> evalArray inner ctx |> Array.map (Array.map (applyUnary op))
    | BinOp (l, op, r) -> broadcastArr op (evalArray l ctx) (evalArray r ctx)
    | _ -> scalar (eval expr ctx)

/// Element-wise binary op with Excel's broadcasting: a 1×1 operand is reused
/// for every element, a single row or column repeats along the other axis,
/// and the result takes the larger extent on each axis — a position one of
/// the operands does not cover is #N/A.
and broadcastArr (op: string) (la: ArrVal) (ra: ArrVal) : ArrVal =
    let lr = la.Length
    let lc = if lr > 0 then la.[0].Length else 0
    let rr = ra.Length
    let rc = if rr > 0 then ra.[0].Length else 0
    if lr = 0 || lc = 0 || rr = 0 || rc = 0 then [| [| err Err.value |] |]
    else
        let nr = max lr rr
        let nc = max lc rc
        let pick (a: ArrVal) rows cols r c =
            if (rows <> 1 && r >= rows) || (cols <> 1 && c >= cols) then ValueNone
            else ValueSome a.[(if rows = 1 then 0 else r)].[(if cols = 1 then 0 else c)]
        [| for r in 0 .. nr - 1 ->
            [| for c in 0 .. nc - 1 ->
                match pick la lr lc r c, pick ra rr rc r c with
                | ValueSome l, ValueSome v -> applyBin op l v
                | _ -> err Err.na |] |]

/// The dynamic-array (spilling) functions. Operate row-wise by default;
/// `by_col` variants and multi-key SORTBY beyond the first key are deferred.
and evalArrayFunction (name: string) (args: Expr list) (ctx: EvalContext) : ArrVal =
    let argsArray = List.toArray args
    let nargs = argsArray.Length
    let argExpr i = argsArray.[i]
    let num i = match coerceNumber (eval (argExpr i) ctx) with Ok n -> n | _ -> nan
    let boolAt i = match coerceBool (eval (argExpr i) ctx) with Ok b -> b | _ -> false
    let errArr e : ArrVal = [| [| err e |] |]
    let firstCol (a: ArrVal) = a |> Array.map (fun row -> if row.Length > 0 then row.[0] else CellValue.Empty)
    let cmpVal = compareValuesForSort
    let transposeArr (a: ArrVal) : ArrVal =
        let nr = a.Length
        let nc = if nr > 0 then a.[0].Length else 0
        [| for c in 0 .. nc - 1 -> [| for r in 0 .. nr - 1 -> a.[r].[c] |] |]
    match name with
    | "SEQUENCE" ->
        if nargs < 1 then errArr Err.value else
        // (An argument left empty - `SEQUENCE(20,,0)` - is its default.)
        let given i = nargs > i && argExpr i <> MissingArg
        let rows = int (num 0)
        let cols = if given 1 then int (num 1) else 1
        let start = if given 2 then num 2 else 1.0
        let step = if given 3 then num 3 else 1.0
        if rows < 1 || cols < 1 then errArr Err.value
        else [| for r in 0 .. rows - 1 -> [| for c in 0 .. cols - 1 -> CellValue.Number (start + float (r * cols + c) * step) |] |]
    | "RANDARRAY" ->
        let given i = nargs > i && argExpr i <> MissingArg
        let rows = if given 0 then int (num 0) else 1
        let cols = if given 1 then int (num 1) else 1
        let lo = if given 2 then num 2 else 0.0
        let hi = if given 3 then num 3 else 1.0
        let whole = given 4 && boolAt 4
        if rows < 1 || cols < 1 then errArr Err.value
        else
            let gen () = if whole then float (rng.Next(int lo, int hi + 1)) else lo + rng.NextDouble() * (hi - lo)
            Array.init rows (fun _ -> Array.init cols (fun _ -> CellValue.Number (gen ())))
    | "UNIQUE" ->
        if nargs < 1 then errArr Err.value else
        let byCol = nargs >= 2 && boolAt 1
        // Distinct rows, first occurrence kept; case-insensitive (engine-wide).
        let keyOf (row: CellValue[]) =
            let sb = Text.StringBuilder()
            for i = 0 to row.Length - 1 do
                if i > 0 then
                    sb.Append('\u0001') |> ignore
                match row.[i] with
                | CellValue.Number n ->
                    sb.Append("n:").Append(n.ToString("R", CultureInfo.InvariantCulture)) |> ignore
                | CellValue.String s ->
                    sb.Append("s:") |> ignore
#if FABLE_COMPILER
                    sb.Append(s.ToLowerInvariant()) |> ignore
#else
                    sb.Append(s) |> ignore
#endif
                | CellValue.Boolean b ->
                    sb.Append("b:").Append(string b) |> ignore
                | CellValue.Error e ->
                    sb.Append("x:").Append(e) |> ignore
                | _ ->
                    sb.Append("e") |> ignore
            sb.ToString()
        let makeSeen () =
#if FABLE_COMPILER
            HashSet<string>()
#else
            HashSet<string>(StringComparer.OrdinalIgnoreCase)
#endif
        let sparseUniqueRows () =
            match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0) with
            | Some source when not byCol ->
                let rows = source.Count / source.W
                let area = int64 source.Count
                if rows <= 0 || area < 1024L || int64 source.M.Cells.Count > area then
                    ValueNone
                else
                    let rowSet = HashSet<int>()
                    let rowIndexes = ResizeArray<int>()
                    source.M.Cells
                    |> Map.iter (fun addr _ ->
                        match tryRangeOffsetForCount source source.Count addr with
                        | ValueSome offset ->
                            let rowIndex = offset / source.W
                            if rowSet.Add rowIndex then
                                rowIndexes.Add rowIndex
                        | _ -> ())
                    rowIndexes.Sort()
                    let seen = makeSeen ()
                    let outRows = ResizeArray<CellValue[]>()
                    let addRow row =
                        if seen.Add(keyOf row) then
                            outRows.Add row
                    let blankRow = Array.create source.W CellValue.Empty
                    let mutable nextRow = 0
                    for i = 0 to rowIndexes.Count - 1 do
                        let rowIndex = rowIndexes.[i]
                        if rowIndex >= nextRow then
                            if rowIndex > nextRow then
                                addRow blankRow
                            addRow (Array.init source.W (fun c -> rangeRead source (rowIndex * source.W + c)))
                            nextRow <- rowIndex + 1
                    if nextRow < rows then
                        addRow blankRow
                    ValueSome (outRows.ToArray())
            | _ -> ValueNone
        match sparseUniqueRows () with
        | ValueSome res when res.Length > 0 -> res
        | ValueSome _ -> errArr Err.calc
        | ValueNone ->
            let arr0 = evalArray (argExpr 0) ctx
            let arr = if byCol then transposeArr arr0 else arr0
            let seen = makeSeen ()
            let outRows = ResizeArray<CellValue[]>()
            for row in arr do if seen.Add(keyOf row) then outRows.Add row
            if outRows.Count = 0 then errArr Err.calc
            else let res = outRows.ToArray() in (if byCol then transposeArr res else res)
    | "SORT" ->
        if nargs < 1 then errArr Err.value else
        let idx = (if nargs >= 2 then int (num 1) else 1) - 1
        let order = if nargs >= 3 then int (num 2) else 1
        let byCol = nargs >= 4 && boolAt 3
        let trySortRangeRows () =
            match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0) with
            | Some source when not byCol && source.Count >= 1024 && source.Count >= source.W ->
                let values = materializeRangeValues source
                let rows = source.Count / source.W
                let rowEntries =
                    Array.init rows (fun row ->
                        let key =
                            if idx >= 0 && idx < source.W then
                                values.[row * source.W + idx]
                            else
                                CellValue.Empty
                        struct (key, row))
                Array.sortInPlaceWith
                    (fun (struct (key1, _)) (struct (key2, _)) ->
                        let c = cmpVal key1 key2
                        if order < 0 then -c else c)
                    rowEntries
                let result =
                    Array.init rows (fun outRow ->
                        let struct (_, sourceRow) = rowEntries.[outRow]
                        let baseOffset = sourceRow * source.W
                        Array.init source.W (fun c -> values.[baseOffset + c]))
                ValueSome result
            | _ -> ValueNone
        match trySortRangeRows () with
        | ValueSome result -> result
        | ValueNone ->
            let arr0 = evalArray (argExpr 0) ctx
            // by_col → sort columns by row `idx` (transpose, sort rows, transpose back).
            let arr = if byCol then transposeArr arr0 else arr0
            let sorted =
                arr
                |> Array.sortWith (fun r1 r2 ->
                    let a = if idx >= 0 && idx < r1.Length then r1.[idx] else CellValue.Empty
                    let b = if idx >= 0 && idx < r2.Length then r2.[idx] else CellValue.Empty
                    let c = cmpVal a b in if order < 0 then -c else c)
            if byCol then transposeArr sorted else sorted
    | "SORTBY" ->
        // SORTBY(array, by1, [order1], [by2, order2], …) — stable multi-key sort.
        if nargs < 2 then errArr Err.value else
        let trySortByRangeRows () =
            match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0) with
            | Some source when source.Count >= 1024 && source.Count >= source.W ->
                let rows = source.Count / source.W
                let keySpecs = ResizeArray<struct (RangeAccess * int)>()
                let mutable ok = true
                let mutable i = 1
                while ok && i < nargs do
                    match rangeAccessOf ctx.Model ctx.Workbook (argExpr i) with
                    | Some keyAccess when keyAccess.Count >= keyAccess.W ->
                        let order = if i + 1 < nargs then int (num (i + 1)) else 1
                        keySpecs.Add(struct (keyAccess, order))
                        i <- i + 2
                    | _ ->
                        ok <- false
                if not ok || keySpecs.Count = 0 then
                    ValueNone
                else
                    let values = materializeRangeValues source
                    let keys =
                        Array.init keySpecs.Count (fun keyIndex ->
                            let struct (keyAccess, order) = keySpecs.[keyIndex]
                            let keyValues = materializeRangeValues keyAccess
                            let keyRows = keyAccess.Count / keyAccess.W
                            struct (keyValues, keyAccess.W, keyRows, order))
                    let indexes = Array.init rows id
                    Array.sortInPlaceWith
                        (fun leftRow rightRow ->
                            let mutable result = 0
                            let mutable keyIndex = 0
                            while result = 0 && keyIndex < keys.Length do
                                let struct (keyValues, keyWidth, keyRows, order) = keys.[keyIndex]
                                let left =
                                    if leftRow < keyRows then keyValues.[leftRow * keyWidth]
                                    else CellValue.Empty
                                let right =
                                    if rightRow < keyRows then keyValues.[rightRow * keyWidth]
                                    else CellValue.Empty
                                let c = cmpVal left right
                                result <- if order < 0 then -c else c
                                keyIndex <- keyIndex + 1
                            if result <> 0 then result else compareInts leftRow rightRow)
                        indexes
                    let result =
                        Array.init rows (fun outRow ->
                            let sourceRow = indexes.[outRow]
                            let baseOffset = sourceRow * source.W
                            Array.init source.W (fun c -> values.[baseOffset + c]))
                    ValueSome result
            | _ -> ValueNone
        match trySortByRangeRows () with
        | ValueSome result -> result
        | ValueNone ->
            let arr = evalArray (argExpr 0) ctx
            let keys =
                [ let mutable i = 1
                  while i < nargs do
                      let key = firstCol (evalArray (argExpr i) ctx)
                      let order = if i + 1 < nargs then int (num (i + 1)) else 1
                      yield key, order
                      i <- i + 2 ]
            Array.init arr.Length id
            |> Array.sortWith (fun i j ->
                let rec go ks =
                    match ks with
                    | (key: CellValue[], order) :: rest ->
                        let a = if i < key.Length then key.[i] else CellValue.Empty
                        let b = if j < key.Length then key.[j] else CellValue.Empty
                        let c = cmpVal a b
                        if c <> 0 then (if order < 0 then -c else c) else go rest
                    | [] -> compareInts i j
                go keys)
            |> Array.map (fun i -> arr.[i])
    | "FILTER" ->
        if nargs < 2 then errArr Err.value else
        let includeAccessAndPredicate includeExpr =
            let coercePredicate value =
                match coerceBool value with
                | Ok true -> true
                | _ -> false
            let comparisonPredicate op left right =
                match compareOperandsNoError left right with
                | ValueSome c ->
                    match op with
                    | "=" -> c = 0
                    | "<>" -> c <> 0
                    | "<" -> c < 0
                    | ">" -> c > 0
                    | "<=" -> c <= 0
                    | ">=" -> c >= 0
                    | _ -> coercePredicate (applyBin op left right)
                | ValueNone -> false
            let readerFor (access: RangeAccess) predicate =
                if access.W <> 1 then ValueNone
                else
                    if predicate CellValue.Empty then ValueNone
                    else ValueSome (access, predicate)
            match includeExpr with
            | BinOp (leftExpr, op, rightExpr) ->
                match rangeAccessOf ctx.Model ctx.Workbook leftExpr, rangeAccessOf ctx.Model ctx.Workbook rightExpr with
                | Some access, None ->
                    let scalar = eval rightExpr ctx
                    readerFor access (fun value -> comparisonPredicate op value scalar)
                | None, Some access ->
                    let scalar = eval leftExpr ctx
                    readerFor access (fun value -> comparisonPredicate op scalar value)
                | _ -> ValueNone
            | _ ->
                match rangeAccessOf ctx.Model ctx.Workbook includeExpr with
                | Some access -> readerFor access coercePredicate
                | None -> ValueNone
        let sparseFilterRows () =
            match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0), includeAccessAndPredicate (argExpr 1) with
            | Some source, ValueSome (includeAccess, includeReader) ->
                let rows = source.Count / source.W
                let area = int64 includeAccess.Count
                if rows <= 0 || area < 1024L || int64 includeAccess.M.Cells.Count > area then
                    ValueNone
                else
                    let keptRows = ResizeArray<CellValue[]>()
                    includeAccess.M.Cells
                    |> Map.iter (fun addr _ ->
                        match tryRangeOffsetForCount includeAccess includeAccess.Count addr with
                        | ValueSome rowIndex when rowIndex < rows ->
                            if includeReader (rangeRead includeAccess rowIndex) then
                                keptRows.Add(Array.init source.W (fun c -> rangeRead source (rowIndex * source.W + c)))
                        | _ -> ())
                    ValueSome (keptRows.ToArray())
            | _ -> ValueNone
        let directFilterRows () =
            match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0), includeAccessAndPredicate (argExpr 1) with
            | Some source, ValueSome (includeAccess, includeReader) when source.Count >= source.W ->
                let rows = source.Count / source.W
                let keptRows = ResizeArray<CellValue[]>()
                let includeRows = includeAccess.Count / includeAccess.W
                let sourceValues = materializeRangeValues source
                let includeValues = materializeRangeValues includeAccess
                let mutable rowIndex = 0
                while rowIndex < rows do
                    let keep =
                        if rowIndex < includeRows then
                            includeReader includeValues.[rowIndex]
                        else
                            false
                    if keep then
                        let baseOffset = rowIndex * source.W
                        keptRows.Add(Array.init source.W (fun c -> sourceValues.[baseOffset + c]))
                    rowIndex <- rowIndex + 1
                ValueSome (keptRows.ToArray())
            | _ -> ValueNone
        match sparseFilterRows () with
        | ValueSome kept ->
            if kept.Length > 0 then kept
            elif nargs >= 3 then [| [| eval (argExpr 2) ctx |] |]
            else errArr Err.calc
        | ValueNone ->
            match directFilterRows () with
            | ValueSome kept ->
                if kept.Length > 0 then kept
                elif nargs >= 3 then [| [| eval (argExpr 2) ctx |] |]
                else errArr Err.calc
            | ValueNone ->
                let arr = evalArray (argExpr 0) ctx
                let incl = evalArray (argExpr 1) ctx
                let includeAt i =
                    if i < incl.Length && incl.[i].Length > 0 then
                        match coerceBool incl.[i].[0] with
                        | Ok true -> true
                        | _ -> false
                    else
                        false
                let keptRows = ResizeArray<CellValue[]>()
                for i in 0 .. arr.Length - 1 do
                    if includeAt i then
                        keptRows.Add arr.[i]
                let kept = keptRows.ToArray()
                if kept.Length > 0 then kept
                elif nargs >= 3 then [| [| eval (argExpr 2) ctx |] |]
                else errArr Err.calc
    | "TRANSPOSE" ->
        if nargs < 1 then errArr Err.value else
        match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0) with
        | Some source when source.Count >= source.W ->
            let rows = source.Count / source.W
            let area = int64 source.Count
            if area >= 1024L && int64 source.M.Cells.Count <= area * 4L then
                let result = Array.init source.W (fun _ -> Array.create rows CellValue.Empty)
                source.M.Cells
                |> Map.iter (fun addr cell ->
                    match tryRangeOffsetForCount source source.Count addr with
                    | ValueSome offset ->
                        result.[offset % source.W].[offset / source.W] <- unwrapValue cell.Value
                    | ValueNone -> ())
                result
            else
                let result = Array.init source.W (fun _ -> Array.zeroCreate<CellValue> rows)
                for r = 0 to rows - 1 do
                    let baseOffset = r * source.W
                    for c = 0 to source.W - 1 do
                        result.[c].[r] <- rangeRead source (baseOffset + c)
                result
        | _ ->
            let arr = evalArray (argExpr 0) ctx
            let nr = arr.Length
            let nc = if nr > 0 then arr.[0].Length else 0
            [| for c in 0 .. nc - 1 -> [| for r in 0 .. nr - 1 -> arr.[r].[c] |] |]
    | "TEXTSPLIT" ->
        // TEXTSPLIT(text, col_delim, [row_delim]) → 2-D array.
        if nargs < 2 then errArr Err.value else
        let textOf i = match coerceText (eval (argExpr i) ctx) with Ok t -> t | _ -> ""
        let text = textOf 0
        let colD = textOf 1
        let rowD = if nargs >= 3 then textOf 2 else ""
        let rowsOf = if rowD <> "" then text.Split([| rowD |], StringSplitOptions.None) else [| text |]
        rowsOf
        |> Array.map (fun r ->
            (if colD <> "" then r.Split([| colD |], StringSplitOptions.None) else [| r |])
            |> Array.map CellValue.String)
    | "MAP" ->
        // MAP(array, lambda) — apply the lambda to each element. (Excel allows
        // several parallel arrays; v1 supports one.)
        if nargs < 2 then errArr Err.value else
        match lambdaPartsOf (argExpr (nargs - 1)) with
        | None -> errArr Err.value
        | Some parts ->
            let trySparseMap () =
                match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0), parts with
                | Some source, [ NameRef paramName; body ]
                    when source.Count > 0
                         && source.Count >= source.W
                         && lambdaBodyUsesOnlySimpleArithmetic paramName body
                         && not (containsVolatile body) ->
                    let area = int64 source.Count
                    if area < 1024L || int64 source.M.Cells.Count > area * 4L then
                        ValueNone
                    else
                        let rows = source.Count / source.W
                        let result = Array.init rows (fun _ -> Array.zeroCreate<CellValue> source.W)
                        let lambdaParts = [ NameRef paramName; body ]
                        let setAt offset value =
                            result.[offset / source.W].[offset % source.W] <- value
                        let fill repeats value startOffset =
                            let mapped = eval (applyLambdaExpr lambdaParts [ valueToExpr value ]) ctx
                            let mutable offset = startOffset
                            let mutable remaining = repeats
                            while remaining > 0L do
                                setAt offset mapped
                                offset <- offset + 1
                                remaining <- remaining - 1L
                            offset
                        let mutable nextOffset = 0
                        source.M.Cells
                        |> Map.iter (fun addr cell ->
                            match tryRangeOffsetForCount source source.Count addr with
                            | ValueSome offset when offset >= nextOffset ->
                                nextOffset <- fill (int64 (offset - nextOffset)) CellValue.Empty nextOffset
                                nextOffset <- fill 1L (unwrapValue cell.Value) nextOffset
                            | _ -> ())
                        let _ = fill (int64 (source.Count - nextOffset)) CellValue.Empty nextOffset
                        ValueSome result
                | _ -> ValueNone
            match trySparseMap () with
            | ValueSome result -> result
            | ValueNone ->
                evalArray (argExpr 0) ctx
                |> Array.map (Array.map (fun v -> eval (applyLambdaExpr parts [ valueToExpr v ]) ctx))
    | "REDUCE" ->
        // REDUCE(initial, array, lambda(acc, value)) → scalar accumulator.
        if nargs <> 3 then errArr Err.value else
        match lambdaPartsOf (argExpr 2) with
        | None -> errArr Err.value
        | Some parts ->
            let trySparseAdditiveReduce () =
                let isName expected expr =
                    match expr with
                    | NameRef name -> equalsOrdinalIgnoreCase name expected
                    | _ -> false
                match rangeAccessOf ctx.Model ctx.Workbook (argExpr 1), parts with
                | Some source, [ NameRef accName; NameRef valueName; BinOp (left, "+", right) ]
                    when source.Count > 0 ->
                    let increment =
                        if isName accName left && lambdaBodyUsesOnlySimpleArithmetic valueName right && not (containsVolatile right) then
                            ValueSome right
                        elif isName accName right && lambdaBodyUsesOnlySimpleArithmetic valueName left && not (containsVolatile left) then
                            ValueSome left
                        else
                            ValueNone
                    match increment with
                    | ValueNone -> ValueNone
                    | ValueSome inc ->
                        let area = int64 source.Count
                        if area < 1024L || int64 source.M.Cells.Count > area * 4L then
                            ValueNone
                        else
                            match coerceNumber (eval (argExpr 0) ctx) with
                            | Result.Error _ -> ValueNone
                            | Ok initial ->
                                let valueLambda = [ NameRef valueName; inc ]
                                let mutable firstErr = ValueNone
                                let mutable total = initial
                                let addIncrement repeats value =
                                    if repeats > 0L && ValueOption.isNone firstErr then
                                        let mapped = eval (applyLambdaExpr valueLambda [ valueToExpr value ]) ctx
                                        match coerceNumber mapped with
                                        | Ok n -> total <- total + n * float repeats
                                        | Result.Error e -> firstErr <- ValueSome e
                                let mutable nextOffset = 0
                                source.M.Cells
                                |> Map.iter (fun addr cell ->
                                    match tryRangeOffsetForCount source source.Count addr with
                                    | ValueSome offset when offset >= nextOffset ->
                                        addIncrement (int64 (offset - nextOffset)) CellValue.Empty
                                        addIncrement 1L (unwrapValue cell.Value)
                                        nextOffset <- offset + 1
                                    | _ -> ())
                                addIncrement (int64 (source.Count - nextOffset)) CellValue.Empty
                                match firstErr with
                                | ValueSome e -> ValueSome [| [| err e |] |]
                                | ValueNone -> ValueSome [| [| CellValue.Number total |] |]
                | _ -> ValueNone
            match trySparseAdditiveReduce () with
            | ValueSome result -> result
            | ValueNone ->
                let mutable acc = eval (argExpr 0) ctx
                let step v =
                    acc <- eval (applyLambdaExpr parts [ valueToExpr acc; valueToExpr v ]) ctx
                match rangeAccessOf ctx.Model ctx.Workbook (argExpr 1) with
                | Some source ->
                    for k = 0 to source.Count - 1 do
                        step (rangeRead source k)
                | None ->
                    for row in evalArray (argExpr 1) ctx do
                        for v in row do step v
                [| [| acc |] |]
    | "SCAN" ->
        // SCAN(initial, array, lambda(acc, value)) → array of running accumulators.
        if nargs <> 3 then errArr Err.value else
        match lambdaPartsOf (argExpr 2) with
        | None -> errArr Err.value
        | Some parts ->
            let trySparseAdditiveScan () =
                let isName expected expr =
                    match expr with
                    | NameRef name -> equalsOrdinalIgnoreCase name expected
                    | _ -> false
                match rangeAccessOf ctx.Model ctx.Workbook (argExpr 1), parts with
                | Some source, [ NameRef accName; NameRef valueName; BinOp (left, "+", right) ]
                    when source.Count > 0 && source.Count >= source.W ->
                    let increment =
                        if isName accName left && lambdaBodyUsesOnlySimpleArithmetic valueName right && not (containsVolatile right) then
                            ValueSome right
                        elif isName accName right && lambdaBodyUsesOnlySimpleArithmetic valueName left && not (containsVolatile left) then
                            ValueSome left
                        else
                            ValueNone
                    match increment with
                    | ValueNone -> ValueNone
                    | ValueSome inc ->
                        let area = int64 source.Count
                        if area < 1024L || int64 source.M.Cells.Count > area * 4L then
                            ValueNone
                        else
                            match coerceNumber (eval (argExpr 0) ctx) with
                            | Result.Error _ -> ValueNone
                            | Ok initial ->
                                let valueLambda = [ NameRef valueName; inc ]
                                let rows = source.Count / source.W
                                let result = Array.init rows (fun _ -> Array.zeroCreate<CellValue> source.W)
                                let mutable total = initial
                                let mutable firstErr = ValueNone
                                let setAt offset value =
                                    result.[offset / source.W].[offset % source.W] <- value
                                let addIncrement repeats value startOffset =
                                    let mutable offset = startOffset
                                    if repeats > 0L then
                                        match firstErr with
                                        | ValueSome e ->
                                            let mutable remaining = repeats
                                            while remaining > 0L do
                                                setAt offset (err e)
                                                offset <- offset + 1
                                                remaining <- remaining - 1L
                                        | ValueNone ->
                                            let mapped = eval (applyLambdaExpr valueLambda [ valueToExpr value ]) ctx
                                            match coerceNumber mapped with
                                            | Result.Error e ->
                                                firstErr <- ValueSome e
                                                let mutable remaining = repeats
                                                while remaining > 0L do
                                                    setAt offset (err e)
                                                    offset <- offset + 1
                                                    remaining <- remaining - 1L
                                            | Ok n ->
                                                let mutable remaining = repeats
                                                while remaining > 0L do
                                                    total <- total + n
                                                    setAt offset (CellValue.Number total)
                                                    offset <- offset + 1
                                                    remaining <- remaining - 1L
                                    offset
                                let mutable nextOffset = 0
                                source.M.Cells
                                |> Map.iter (fun addr cell ->
                                    match tryRangeOffsetForCount source source.Count addr with
                                    | ValueSome offset when offset >= nextOffset ->
                                        nextOffset <- addIncrement (int64 (offset - nextOffset)) CellValue.Empty nextOffset
                                        nextOffset <- addIncrement 1L (unwrapValue cell.Value) nextOffset
                                    | _ -> ())
                                let _ = addIncrement (int64 (source.Count - nextOffset)) CellValue.Empty nextOffset
                                ValueSome result
                | _ -> ValueNone
            match trySparseAdditiveScan () with
            | ValueSome result -> result
            | ValueNone ->
                let mutable acc = eval (argExpr 0) ctx
                match rangeAccessOf ctx.Model ctx.Workbook (argExpr 1) with
                | Some source when source.Count >= source.W ->
                    let rows = source.Count / source.W
                    Array.init rows (fun r ->
                        Array.init source.W (fun c ->
                            acc <- eval (applyLambdaExpr parts [ valueToExpr acc; valueToExpr (rangeRead source (r * source.W + c)) ]) ctx
                            acc))
                | _ ->
                    evalArray (argExpr 1) ctx
                    |> Array.map (Array.map (fun v ->
                        acc <- eval (applyLambdaExpr parts [ valueToExpr acc; valueToExpr v ]) ctx
                        acc))
    | "MAKEARRAY" ->
        // MAKEARRAY(rows, cols, lambda(r, c)) — r,c are 1-based.
        if nargs <> 3 then errArr Err.value else
        match lambdaPartsOf (argExpr 2) with
        | None -> errArr Err.value
        | Some parts ->
            let rows = int (num 0)
            let cols = int (num 1)
            if rows < 1 || cols < 1 then errArr Err.value
            else Array.init rows (fun r -> Array.init cols (fun c ->
                    eval (applyLambdaExpr parts [ Number (float (r + 1)); Number (float (c + 1)) ]) ctx))
    | "BYROW" ->
        // BYROW(array, lambda(row)) → one value per row (single column). The
        // whole row is handed to the lambda as one array argument.
        if nargs < 2 then errArr Err.value else
        match lambdaPartsOf (argExpr 1) with
        | None -> errArr Err.value
        | Some parts ->
            let tryByRowAggregate () =
                match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0), parts with
                | Some source, [ NameRef param; Function (fn, [ NameRef arg ]) ]
                    when source.Count >= source.W
                         && (equalsOrdinalIgnoreCase fn "SUM" || equalsOrdinalIgnoreCase fn "COUNT" || equalsOrdinalIgnoreCase fn "AVERAGE")
                         && equalsOrdinalIgnoreCase param arg ->
                    let isCount = equalsOrdinalIgnoreCase fn "COUNT"
                    let isAverage = equalsOrdinalIgnoreCase fn "AVERAGE"
                    let rows = source.Count / source.W
                    let result = Array.zeroCreate<CellValue[]> rows
                    let totals = Array.zeroCreate<float> rows
                    let numericCounts = Array.zeroCreate<int> rows
                    let firstErrs = Array.create rows ValueNone
                    let addValue row value =
                        if ValueOption.isNone firstErrs.[row] then
                            match value with
                            | CellValue.Error e -> firstErrs.[row] <- ValueSome e
                            | CellValue.Number n ->
                                totals.[row] <- totals.[row] + (if isCount then 1.0 else n)
                                if isAverage then
                                    numericCounts.[row] <- numericCounts.[row] + 1
                            | CellValue.Boolean b when not isCount ->
                                totals.[row] <- totals.[row] + (if b then 1.0 else 0.0)
                                if isAverage then
                                    numericCounts.[row] <- numericCounts.[row] + 1
                            | _ -> ()
                    let area = int64 source.Count
                    if area >= 1024L && int64 source.M.Cells.Count <= area * 4L then
                        source.M.Cells
                        |> Map.iter (fun addr cell ->
                            match tryRangeOffsetForCount source source.Count addr with
                            | ValueSome offset ->
                                addValue (offset / source.W) (unwrapValue cell.Value)
                            | ValueNone -> ())
                    else
                        for r = 0 to rows - 1 do
                            let baseOffset = r * source.W
                            let mutable c = 0
                            while ValueOption.isNone firstErrs.[r] && c < source.W do
                                addValue r (rangeRead source (baseOffset + c))
                                c <- c + 1
                    for r = 0 to rows - 1 do
                        result.[r] <-
                            [| match firstErrs.[r] with
                               | ValueSome e -> CellValue.Error e
                               | ValueNone when isAverage && numericCounts.[r] = 0 -> CellValue.Error Err.div0
                               | ValueNone when isAverage -> CellValue.Number (totals.[r] / float numericCounts.[r])
                               | ValueNone -> CellValue.Number totals.[r] |]
                    ValueSome result
                | _ -> ValueNone
            match tryByRowAggregate () with
            | ValueSome result -> result
            | _ ->
                evalArray (argExpr 0) ctx
                |> Array.map (fun row -> [| eval (applyLambdaExpr parts [ ArrayLit [| row |] ]) ctx |])
    | "BYCOL" ->
        // BYCOL(array, lambda(col)) → one value per column (single row).
        if nargs < 2 then errArr Err.value else
        match lambdaPartsOf (argExpr 1) with
        | None -> errArr Err.value
        | Some parts ->
            let tryByColSum () =
                match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0), parts with
                | Some source, [ NameRef param; Function (fn, [ NameRef arg ]) ]
                    when source.Count >= source.W
                         && (equalsOrdinalIgnoreCase fn "SUM" || equalsOrdinalIgnoreCase fn "COUNT" || equalsOrdinalIgnoreCase fn "AVERAGE")
                         && equalsOrdinalIgnoreCase param arg ->
                    let isCount = equalsOrdinalIgnoreCase fn "COUNT"
                    let isAverage = equalsOrdinalIgnoreCase fn "AVERAGE"
                    let rows = source.Count / source.W
                    let result = Array.zeroCreate<CellValue> source.W
                    let totals = Array.zeroCreate<float> source.W
                    let numericCounts = Array.zeroCreate<int> source.W
                    let firstErrs = Array.create source.W ValueNone
                    let addValue col value =
                        if ValueOption.isNone firstErrs.[col] then
                            match value with
                            | CellValue.Error e -> firstErrs.[col] <- ValueSome e
                            | CellValue.Number n ->
                                totals.[col] <- totals.[col] + (if isCount then 1.0 else n)
                                if isAverage then
                                    numericCounts.[col] <- numericCounts.[col] + 1
                            | CellValue.Boolean b when not isCount ->
                                totals.[col] <- totals.[col] + (if b then 1.0 else 0.0)
                                if isAverage then
                                    numericCounts.[col] <- numericCounts.[col] + 1
                            | _ -> ()
                    let area = int64 source.Count
                    if area >= 1024L && int64 source.M.Cells.Count <= area * 4L then
                        source.M.Cells
                        |> Map.iter (fun addr cell ->
                            match tryRangeOffsetForCount source source.Count addr with
                            | ValueSome offset ->
                                addValue (offset % source.W) (unwrapValue cell.Value)
                            | ValueNone -> ())
                    else
                        for c = 0 to source.W - 1 do
                            let mutable r = 0
                            while ValueOption.isNone firstErrs.[c] && r < rows do
                                addValue c (rangeRead source (r * source.W + c))
                                r <- r + 1
                    for c = 0 to source.W - 1 do
                        result.[c] <-
                            match firstErrs.[c] with
                            | ValueSome e -> CellValue.Error e
                            | ValueNone when isAverage && numericCounts.[c] = 0 -> CellValue.Error Err.div0
                            | ValueNone when isAverage -> CellValue.Number (totals.[c] / float numericCounts.[c])
                            | ValueNone -> CellValue.Number totals.[c]
                    ValueSome [| result |]
                | _ -> ValueNone
            match tryByColSum () with
            | ValueSome result -> result
            | _ ->
                let arr = evalArray (argExpr 0) ctx
                let nc = if arr.Length > 0 then arr.[0].Length else 0
                [| [| for c in 0 .. nc - 1 ->
                        let col = [| for r in 0 .. arr.Length - 1 -> [| arr.[r].[c] |] |]
                        eval (applyLambdaExpr parts [ ArrayLit col ]) ctx |] |]
    | "GROUPBY" ->
        // GROUPBY(row_range, value_range, [agg]) — a LIVE spilling pivot:
        // groups the row range (1+ columns, nested) and aggregates the value
        // range, with subtotals + a Grand Total. Recomputes when the source
        // changes (it's a formula), and the result is a normal range you can
        // FILTER/SORT/chart. (Better than a static PivotTable: no Refresh.)
        if nargs < 2 then errArr Err.value else
        let parseAggArg i = match coerceText (eval (argExpr i) ctx) with Ok s -> Rbt.Spreadsheet.Pivot.parseAgg s |> Option.defaultValue Rbt.Spreadsheet.Pivot.Sum | _ -> Rbt.Spreadsheet.Pivot.Sum
        let agg = if nargs >= 3 then parseAggArg 2 else Rbt.Spreadsheet.Pivot.Sum
        let sparseGroupBy1 () =
            match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0), rangeAccessOf ctx.Model ctx.Workbook (argExpr 1) with
            | Some rowAccess, Some valueAccess when rowAccess.W = 1 && valueAccess.W = 1 && obj.ReferenceEquals(rowAccess.M, valueAccess.M) ->
                let count = min rowAccess.Count valueAccess.Count
                let area = int64 count
                if count <= 0 || area < 1024L || int64 rowAccess.M.Cells.Count > area * 4L then
                    ValueNone
                else
                    let pivotTextOf value =
                        match unwrapValue value with
                        | CellValue.String s -> s
                        | CellValue.Number n -> n.ToString("R", CultureInfo.InvariantCulture)
                        | CellValue.Boolean b -> if b then "TRUE" else "FALSE"
                        | _ -> ""
                    let groupKey value =
#if FABLE_COMPILER
                        (pivotTextOf value).ToLowerInvariant()
#else
                        pivotTextOf value
#endif
                    let states =
#if FABLE_COMPILER
                        Dictionary<string, struct (int64 * int * float * float * float)>()
                    let valuesByKey = Dictionary<string, CellValue>()
#else
                        Dictionary<string, struct (int64 * int * float * float * float)>(StringComparer.OrdinalIgnoreCase)
                    let valuesByKey = Dictionary<string, CellValue>(StringComparer.OrdinalIgnoreCase)
#endif
                    let rowIndexes = ResizeArray<int>()
                    let rowSet = HashSet<int>()
                    let addRowIndex access addr =
                        match tryRangeOffsetForCount access count addr with
                        | ValueSome offset ->
                            let rowIndex = offset / access.W
                            if rowSet.Add rowIndex then
                                rowIndexes.Add rowIndex
                        | _ -> ()
                    rowAccess.M.Cells
                    |> Map.iter (fun addr _ ->
                        addRowIndex rowAccess addr
                        addRowIndex valueAccess addr)
                    rowIndexes.Sort()
                    let getState key =
                        let mutable state = Unchecked.defaultof<struct (int64 * int * float * float * float)>
                        if states.TryGetValue(key, &state) then state
                        else struct (0L, 0, 0.0, Double.NegativeInfinity, Double.PositiveInfinity)
                    let addValueKey key value =
                        let mutable existing = Unchecked.defaultof<CellValue>
                        if not (valuesByKey.TryGetValue(key, &existing)) then
                            valuesByKey.[key] <- value
                    let addGroupValue groupValue value rowCount =
                        let key = groupKey groupValue
                        addValueKey key groupValue
                        let struct (count, numericCount, sum, maxValue, minValue) = getState key
                        match unwrapValue value with
                        | CellValue.Number n ->
                            states.[key] <- struct (count + rowCount, numericCount + 1, sum + n, max maxValue n, min minValue n)
                        | CellValue.Boolean b ->
                            let n = if b then 1.0 else 0.0
                            states.[key] <- struct (count + rowCount, numericCount + 1, sum + n, max maxValue n, min minValue n)
                        | _ ->
                            states.[key] <- struct (count + rowCount, numericCount, sum, maxValue, minValue)
                    let addBlankRows rowCount =
                        if rowCount > 0L then
                            let key = groupKey CellValue.Empty
                            addValueKey key CellValue.Empty
                            let struct (count, numericCount, sum, maxValue, minValue) = getState key
                            states.[key] <- struct (count + rowCount, numericCount, sum, maxValue, minValue)
                    let mutable nextRow = 0
                    for i = 0 to rowIndexes.Count - 1 do
                        let rowIndex = rowIndexes.[i]
                        if rowIndex >= nextRow then
                            addBlankRows (int64 (rowIndex - nextRow))
                            addGroupValue (rangeRead rowAccess rowIndex) (rangeRead valueAccess rowIndex) 1L
                            nextRow <- rowIndex + 1
                    addBlankRows (int64 (count - nextRow))
                    let finishState (struct (count, numericCount, sum, maxValue, minValue)) =
                        match agg with
                        | Rbt.Spreadsheet.Pivot.Count -> CellValue.Number(float count)
                        | Rbt.Spreadsheet.Pivot.Sum -> CellValue.Number sum
                        | Rbt.Spreadsheet.Pivot.Average -> if numericCount = 0 then CellValue.Empty else CellValue.Number(sum / float numericCount)
                        | Rbt.Spreadsheet.Pivot.Max -> if numericCount = 0 then CellValue.Empty else CellValue.Number maxValue
                        | Rbt.Spreadsheet.Pivot.Min -> if numericCount = 0 then CellValue.Empty else CellValue.Number minValue
                    let sortedValues = ResizeArray<CellValue>(valuesByKey.Count)
                    for pair in valuesByKey do
                        sortedValues.Add pair.Value
                    let rowValues = sortedValues.ToArray()
                    Array.sortInPlaceWith compareCellValues rowValues
                    let result = ResizeArray<CellValue[]>(rowValues.Length + 1)
                    for value in rowValues do
                        let key = groupKey value
                        result.Add [| value; finishState states.[key] |]
                    let mutable grandCount = 0L
                    let mutable grandNumericCount = 0
                    let mutable grandSum = 0.0
                    let mutable grandMax = Double.NegativeInfinity
                    let mutable grandMin = Double.PositiveInfinity
                    for pair in states do
                        let struct (count, numericCount, sum, maxValue, minValue) = pair.Value
                        grandCount <- grandCount + count
                        grandNumericCount <- grandNumericCount + numericCount
                        grandSum <- grandSum + sum
                        if numericCount > 0 then
                            grandMax <- max grandMax maxValue
                            grandMin <- min grandMin minValue
                    result.Add [| CellValue.String "Grand Total"; finishState (struct (grandCount, grandNumericCount, grandSum, grandMax, grandMin)) |]
                    ValueSome (result.ToArray())
            | _ -> ValueNone
        match sparseGroupBy1 () with
        | ValueSome result -> result
        | ValueNone -> Rbt.Spreadsheet.Pivot.groupBy (evalArray (argExpr 0) ctx) (evalArray (argExpr 1) ctx) agg
    | "PIVOTBY" ->
        // PIVOTBY(row_range, col_range, value_range, [agg]) — live spilling
        // cross-tab.
        if nargs < 3 then errArr Err.value else
        let parseAggArg i = match coerceText (eval (argExpr i) ctx) with Ok s -> Rbt.Spreadsheet.Pivot.parseAgg s |> Option.defaultValue Rbt.Spreadsheet.Pivot.Sum | _ -> Rbt.Spreadsheet.Pivot.Sum
        let agg = if nargs >= 4 then parseAggArg 3 else Rbt.Spreadsheet.Pivot.Sum
        let sparsePivotBy1 () =
            match rangeAccessOf ctx.Model ctx.Workbook (argExpr 0), rangeAccessOf ctx.Model ctx.Workbook (argExpr 1), rangeAccessOf ctx.Model ctx.Workbook (argExpr 2) with
            | Some rowAccess, Some colAccess, Some valueAccess
                when rowAccess.W = 1 && colAccess.W = 1 && valueAccess.W = 1
                     && obj.ReferenceEquals(rowAccess.M, colAccess.M)
                     && obj.ReferenceEquals(rowAccess.M, valueAccess.M) ->
                let count = min rowAccess.Count (min colAccess.Count valueAccess.Count)
                let area = int64 count
                if count <= 0 || area < 1024L || int64 rowAccess.M.Cells.Count > area * 4L then
                    ValueNone
                else
                    let pivotTextOf value =
                        match unwrapValue value with
                        | CellValue.String s -> s
                        | CellValue.Number n -> n.ToString("R", CultureInfo.InvariantCulture)
                        | CellValue.Boolean b -> if b then "TRUE" else "FALSE"
                        | _ -> ""
                    let groupKey value =
#if FABLE_COMPILER
                        (pivotTextOf value).ToLowerInvariant()
#else
                        pivotTextOf value
#endif
                    let newState () = struct (0L, 0, 0.0, Double.NegativeInfinity, Double.PositiveInfinity)
                    let addStateValue (struct (count, numericCount, sum, maxValue, minValue)) value rowCount =
                        let valueNumber value =
                            match unwrapValue value with
                            | CellValue.Number n -> ValueSome n
                            | CellValue.Boolean b -> ValueSome(if b then 1.0 else 0.0)
                            | _ -> ValueNone
                        match valueNumber value with
                        | ValueSome n -> struct (count + rowCount, numericCount + 1, sum + n, max maxValue n, min minValue n)
                        | ValueNone -> struct (count + rowCount, numericCount, sum, maxValue, minValue)
                    let finishState (struct (count, numericCount, sum, maxValue, minValue)) =
                        match agg with
                        | Rbt.Spreadsheet.Pivot.Count -> CellValue.Number(float count)
                        | Rbt.Spreadsheet.Pivot.Sum -> CellValue.Number sum
                        | Rbt.Spreadsheet.Pivot.Average -> if numericCount = 0 then CellValue.Empty else CellValue.Number(sum / float numericCount)
                        | Rbt.Spreadsheet.Pivot.Max -> if numericCount = 0 then CellValue.Empty else CellValue.Number maxValue
                        | Rbt.Spreadsheet.Pivot.Min -> if numericCount = 0 then CellValue.Empty else CellValue.Number minValue
                    let rowValuesByKey =
#if FABLE_COMPILER
                        Dictionary<string, CellValue>()
#else
                        Dictionary<string, CellValue>(StringComparer.OrdinalIgnoreCase)
#endif
                    let colValuesByKey =
#if FABLE_COMPILER
                        Dictionary<string, CellValue>()
#else
                        Dictionary<string, CellValue>(StringComparer.OrdinalIgnoreCase)
#endif
                    let rowTotals =
#if FABLE_COMPILER
                        Dictionary<string, struct (int64 * int * float * float * float)>()
#else
                        Dictionary<string, struct (int64 * int * float * float * float)>(StringComparer.OrdinalIgnoreCase)
#endif
                    let colTotals =
#if FABLE_COMPILER
                        Dictionary<string, struct (int64 * int * float * float * float)>()
#else
                        Dictionary<string, struct (int64 * int * float * float * float)>(StringComparer.OrdinalIgnoreCase)
#endif
                    let cells =
#if FABLE_COMPILER
                        Dictionary<string, Dictionary<string, struct (int64 * int * float * float * float)>>()
#else
                        Dictionary<string, Dictionary<string, struct (int64 * int * float * float * float)>>(StringComparer.OrdinalIgnoreCase)
#endif
                    let getState key (dict: Dictionary<string, struct (int64 * int * float * float * float)>) =
                        let mutable state = Unchecked.defaultof<struct (int64 * int * float * float * float)>
                        if dict.TryGetValue(key, &state) then state else newState ()
                    let getRowCells rowKey =
                        let mutable rowCells = Unchecked.defaultof<Dictionary<string, struct (int64 * int * float * float * float)>>
                        if cells.TryGetValue(rowKey, &rowCells) then
                            rowCells
                        else
                            let rowCells =
#if FABLE_COMPILER
                                Dictionary<string, struct (int64 * int * float * float * float)>()
#else
                                Dictionary<string, struct (int64 * int * float * float * float)>(StringComparer.OrdinalIgnoreCase)
#endif
                            cells.[rowKey] <- rowCells
                            rowCells
                    let addDistinctValue key value (dict: Dictionary<string, CellValue>) =
                        let mutable existing = Unchecked.defaultof<CellValue>
                        if not (dict.TryGetValue(key, &existing)) then
                            dict.[key] <- value
                    let mutable grand = newState ()
                    let addRecord rowValue colValue value rowCount =
                        let rowKey = groupKey rowValue
                        let colKey = groupKey colValue
                        addDistinctValue rowKey rowValue rowValuesByKey
                        addDistinctValue colKey colValue colValuesByKey
                        rowTotals.[rowKey] <- addStateValue (getState rowKey rowTotals) value rowCount
                        colTotals.[colKey] <- addStateValue (getState colKey colTotals) value rowCount
                        let rowCells = getRowCells rowKey
                        rowCells.[colKey] <- addStateValue (getState colKey rowCells) value rowCount
                        grand <- addStateValue grand value rowCount
                    let rowIndexes = ResizeArray<int>()
                    let rowSet = HashSet<int>()
                    let addRowIndex access addr =
                        match tryRangeOffsetForCount access count addr with
                        | ValueSome offset ->
                            let rowIndex = offset / access.W
                            if rowSet.Add rowIndex then
                                rowIndexes.Add rowIndex
                        | _ -> ()
                    rowAccess.M.Cells
                    |> Map.iter (fun addr _ ->
                        addRowIndex rowAccess addr
                        addRowIndex colAccess addr
                        addRowIndex valueAccess addr)
                    rowIndexes.Sort()
                    let mutable nextRow = 0
                    for i = 0 to rowIndexes.Count - 1 do
                        let rowIndex = rowIndexes.[i]
                        if rowIndex >= nextRow then
                            if rowIndex > nextRow then
                                addRecord CellValue.Empty CellValue.Empty CellValue.Empty (int64 (rowIndex - nextRow))
                            addRecord (rangeRead rowAccess rowIndex) (rangeRead colAccess rowIndex) (rangeRead valueAccess rowIndex) 1L
                            nextRow <- rowIndex + 1
                    if count > nextRow then
                        addRecord CellValue.Empty CellValue.Empty CellValue.Empty (int64 (count - nextRow))
                    let sortedValues (valuesByKey: Dictionary<string, CellValue>) =
                        let values = ResizeArray<CellValue>(valuesByKey.Count)
                        for pair in valuesByKey do
                            values.Add pair.Value
                        let result = values.ToArray()
                        Array.sortInPlaceWith compareCellValues result
                        result
                    let rowValues = sortedValues rowValuesByKey
                    let colValues = sortedValues colValuesByKey
                    let colKeys = Array.zeroCreate<string> colValues.Length
                    for i = 0 to colValues.Length - 1 do
                        colKeys.[i] <- groupKey colValues.[i]
                    let result = ResizeArray<CellValue[]>(rowValues.Length + 2)
                    let header = Array.create (colValues.Length + 2) CellValue.Empty
                    header.[0] <- CellValue.String "Row Labels"
                    for i = 0 to colValues.Length - 1 do
                        header.[i + 1] <- colValues.[i]
                    header.[colValues.Length + 1] <- CellValue.String "Grand Total"
                    result.Add header
                    for rowValue in rowValues do
                        let rowKey = groupKey rowValue
                        let rowCells =
                            let mutable value = Unchecked.defaultof<Dictionary<string, struct (int64 * int * float * float * float)>>
                            if cells.TryGetValue(rowKey, &value) then ValueSome value else ValueNone
                        let row = Array.create (colValues.Length + 2) CellValue.Empty
                        row.[0] <- rowValue
                        for i = 0 to colKeys.Length - 1 do
                            row.[i + 1] <-
                                match rowCells with
                                | ValueSome cellDict -> finishState (getState colKeys.[i] cellDict)
                                | ValueNone -> finishState (newState ())
                        row.[colValues.Length + 1] <- finishState (getState rowKey rowTotals)
                        result.Add row
                    let totalRow = Array.create (colValues.Length + 2) CellValue.Empty
                    totalRow.[0] <- CellValue.String "Grand Total"
                    for i = 0 to colKeys.Length - 1 do
                        totalRow.[i + 1] <- finishState (getState colKeys.[i] colTotals)
                    totalRow.[colValues.Length + 1] <- finishState grand
                    result.Add totalRow
                    ValueSome (result.ToArray())
            | _ -> ValueNone
        match sparsePivotBy1 () with
        | ValueSome result -> result
        | ValueNone -> Rbt.Spreadsheet.Pivot.pivotBy (evalArray (argExpr 0) ctx) (evalArray (argExpr 1) ctx) (evalArray (argExpr 2) ctx) agg
    | _ -> evalExtendedArrayFunction name args ctx

/// Run a one-pass `Fold` across an aggregate's whole argument list WITHOUT
/// materializing it: each range arg is streamed cell-by-cell through a
/// `RangeView` enumerator, each scalar arg is stepped once. This is the
/// streaming replacement for `numbersOf (flattenArgs …)` and preserves its
/// semantics (same row-major order, same first-error propagation).
and foldArgs (f: Fold<CellValue, 'S, 'R>) (args: Expr list) (ctx: EvalContext) : 'R =
    let model = ctx.Model
    let workbook = ctx.Workbook
    let mutable s = f.Init
    let iterRange (a: RefAst) (b: RefAst) =
        match resolveModel a.Sheet model workbook with
        | Some m ->
            let sr, sc, er, ec = normalizeRange a b
            let mutable e = ({ Model = m; SR = sr; SC = sc; ER = er; EC = ec }: RangeView).GetEnumerator()
            while e.MoveNext() do s <- f.Step s e.Current
        | None -> s <- f.Step s (err Err.ref)
    for arg in args do
        match arg with
        | RangeRef (a, b) -> iterRange a b
        | SpillRef anchor -> let a, b = spillRangeIn anchor model workbook in iterRange a b
        | TableRef sel -> iterTableSel sel workbook ctx.Sheet ctx.CurrentCell model (fun v -> s <- f.Step s v)
        | _ when isArrayExpr arg ->
            for row in evalArray arg ctx do
                for v in row do s <- f.Step s (unwrapValue v)
        | NameRef n ->
            match workbook |> Option.bind (resolveNameForEval n ctx.Sheet) with
            | Some (RangeRef (a, b)) -> iterRange a b
            | Some (TableRef sel) -> iterTableSel sel workbook ctx.Sheet ctx.CurrentCell model (fun v -> s <- f.Step s v)
            | Some inner -> s <- f.Step s (unwrapValue (eval inner ctx))
            | None -> s <- f.Step s (err Err.name)
        | _ -> s <- f.Step s (unwrapValue (eval arg ctx))
    f.Extract s

/// Numeric/statistical aggregate ranges ignore missing cells. For broad
/// sparse ranges, scanning stored cells avoids allocating one temporary
/// CellAddress per grid coordinate.
and foldArgsIgnoringMissingCells (f: Fold<CellValue, 'S, 'R>) (args: Expr list) (ctx: EvalContext) : 'R =
    foldArgsIgnoringMissingCellsWith id f args ctx

/// Numeric aggregates (SUM, AVERAGE, MIN, STDEV, …) with Excel's two rules:
/// a value typed directly as an argument is coerced — `"3"` and TRUE count,
/// other text is #VALUE! — while values read from a range or an array are
/// not: text and booleans there are ignored.
and foldNumericArgs (f: Fold<CellValue, 'S, 'R>) (args: Expr list) (ctx: EvalContext) : 'R =
    let rangeFold = { f with Step = fun state v -> match v with CellValue.Boolean _ -> state | _ -> f.Step state v }
    foldArgsIgnoringMissingCellsWith directNumericArgument rangeFold args ctx

/// `direct` is applied to the value of an argument that is neither a
/// reference nor an array — a literal or a computed scalar.
and foldArgsIgnoringMissingCellsWith (direct: CellValue -> CellValue) (f: Fold<CellValue, 'S, 'R>) (args: Expr list) (ctx: EvalContext) : 'R =
    let model = ctx.Model
    let workbook = ctx.Workbook
    let mutable s = f.Init
    let iterRange (a: RefAst) (b: RefAst) =
        match resolveModel a.Sheet model workbook with
        | Some m ->
            let sr, sc, er, ec = normalizeRange a b
            let area = int64 (er - sr + 1) * int64 (ec - sc + 1)
            if area >= 1024L && int64 m.Cells.Count <= area * 4L then
                s <-
                    m.Cells
                    |> Map.fold (fun state addr cell ->
                        if addr.Row >= sr && addr.Row <= er && addr.Col >= sc && addr.Col <= ec then
                            f.Step state (unwrapValue cell.Value)
                        else
                            state) s
            else
                let mutable e = ({ Model = m; SR = sr; SC = sc; ER = er; EC = ec }: RangeView).GetEnumerator()
                while e.MoveNext() do s <- f.Step s e.Current
        | None -> s <- f.Step s (err Err.ref)
    for arg in args do
        match arg with
        | RangeRef (a, b) -> iterRange a b
        | SpillRef anchor -> let a, b = spillRangeIn anchor model workbook in iterRange a b
        | TableRef sel -> iterTableSel sel workbook ctx.Sheet ctx.CurrentCell model (fun v -> s <- f.Step s v)
        | _ when isArrayExpr arg ->
            for row in evalArray arg ctx do
                for v in row do s <- f.Step s (unwrapValue v)
        | NameRef n ->
            match workbook |> Option.bind (resolveNameForEval n ctx.Sheet) with
            | Some (RangeRef (a, b)) -> iterRange a b
            | Some (TableRef sel) -> iterTableSel sel workbook ctx.Sheet ctx.CurrentCell model (fun v -> s <- f.Step s v)
            | Some inner -> s <- f.Step s (unwrapValue (eval inner ctx))
            | None -> s <- f.Step s (err Err.name)
        | CellRef _ -> s <- f.Step s (unwrapValue (eval arg ctx))
        | _ -> s <- f.Step s (direct (unwrapValue (eval arg ctx)))
    f.Extract s

and countBlankArgs (args: Expr list) (ctx: EvalContext) : CellValue =
    let model = ctx.Model
    let workbook = ctx.Workbook
    let mutable count = 0L
    let addValue value =
        match unwrapValue value with
        | CellValue.Empty
        | CellValue.String "" -> count <- count + 1L
        | _ -> ()
    let iterRange (a: RefAst) (b: RefAst) =
        match resolveModel a.Sheet model workbook with
        | Some m ->
            let sr, sc, er, ec = normalizeRange a b
            let area = int64 (er - sr + 1) * int64 (ec - sc + 1)
            if area >= 1024L && int64 m.Cells.Count <= area * 4L then
                let mutable storedInRange = 0L
                let mutable storedBlank = 0L
                m.Cells
                |> Map.iter (fun addr cell ->
                    if addr.Row >= sr && addr.Row <= er && addr.Col >= sc && addr.Col <= ec then
                        storedInRange <- storedInRange + 1L
                        match unwrapValue cell.Value with
                        | CellValue.Empty
                        | CellValue.String "" -> storedBlank <- storedBlank + 1L
                        | _ -> ())
                count <- count + (area - storedInRange) + storedBlank
            else
                let mutable e = ({ Model = m; SR = sr; SC = sc; ER = er; EC = ec }: RangeView).GetEnumerator()
                while e.MoveNext() do addValue e.Current
        | None -> ()
    for arg in args do
        match arg with
        | RangeRef (a, b) -> iterRange a b
        | SpillRef anchor -> let a, b = spillRangeIn anchor model workbook in iterRange a b
        | TableRef sel -> iterTableSel sel workbook ctx.Sheet ctx.CurrentCell model addValue
        | _ when ValueOption.isSome (arithmeticArrayNonBlankCount model workbook arg) -> ()
        | _ when isArrayExpr arg ->
            for row in evalArray arg ctx do
                for value in row do addValue value
        | NameRef n ->
            match workbook |> Option.bind (resolveNameForEval n ctx.Sheet) with
            | Some (RangeRef (a, b)) -> iterRange a b
            | Some (TableRef sel) -> iterTableSel sel workbook ctx.Sheet ctx.CurrentCell model addValue
            | Some inner -> addValue (eval inner ctx)
            | None -> ()
        | _ -> addValue (eval arg ctx)
    CellValue.Number (float count)

/// Numeric values from a flattened argument list — text and blanks dropped,
/// matching how SUM/AVERAGE/etc. ignore non-numeric cells. The first error
/// encountered is returned via the error channel so it can propagate.
and numbersOf (values: CellValue list) : Result<float list, string> =
    let numbers = ResizeArray<float>()
    let mutable firstErr = ValueNone
    let mutable remaining = values
    while ValueOption.isNone firstErr && not remaining.IsEmpty do
        match remaining with
        | v :: tail ->
            match unwrapValue v with
            | CellValue.Error e -> firstErr <- ValueSome e
            | CellValue.Number n -> numbers.Add n
            | CellValue.Boolean b -> numbers.Add(if b then 1.0 else 0.0)
            | _ -> ()
            remaining <- tail
        | [] -> ()
    match firstErr with
    | ValueSome e -> Result.Error e
    | ValueNone ->
        let mutable result = []
        for i = numbers.Count - 1 downto 0 do
            result <- numbers.[i] :: result
        Ok result

and numbersOfArgs (args: Expr list) (ctx: EvalContext) : Result<float list, string> =
    foldNumericArgs
        { Init = struct (ValueNone, ResizeArray<float>())
          Step = fun (struct (firstErr, numbers)) value ->
            match firstErr with
            | ValueSome _ -> struct (firstErr, numbers)
            | ValueNone ->
                match unwrapValue value with
                | CellValue.Error e -> struct (ValueSome e, numbers)
                | CellValue.Number n ->
                    numbers.Add n
                    struct (ValueNone, numbers)
                | CellValue.Boolean b ->
                    numbers.Add(if b then 1.0 else 0.0)
                    struct (ValueNone, numbers)
                | _ -> struct (ValueNone, numbers)
          Extract = fun (struct (firstErr, numbers)) ->
            match firstErr with
            | ValueSome e -> Result.Error e
            | ValueNone ->
                let mutable result = []
                for i = numbers.Count - 1 downto 0 do
                    result <- numbers.[i] :: result
                Ok result }
        args
        ctx

and numbersArrayOfArgs (args: Expr list) (ctx: EvalContext) : Result<float array, string> =
    foldNumericArgs
        { Init = struct (ValueNone, ResizeArray<float>())
          Step = fun (struct (firstErr, numbers)) value ->
            match firstErr with
            | ValueSome _ -> struct (firstErr, numbers)
            | ValueNone ->
                match unwrapValue value with
                | CellValue.Error e -> struct (ValueSome e, numbers)
                | CellValue.Number n ->
                    numbers.Add n
                    struct (ValueNone, numbers)
                | CellValue.Boolean b ->
                    numbers.Add(if b then 1.0 else 0.0)
                    struct (ValueNone, numbers)
                | _ -> struct (ValueNone, numbers)
          Extract = fun (struct (firstErr, numbers)) ->
            match firstErr with
            | ValueSome e -> Result.Error e
            | ValueNone -> Ok (numbers.ToArray()) }
        args
        ctx

and private geoMeanFold : Fold<CellValue, struct (string voption * int * float * bool), CellValue> =
    { Init = struct (ValueNone, 0, 0.0, false)
      Step = fun (struct (firstErr, count, logSum, invalid)) value ->
          match firstErr with
          | ValueSome _ -> struct (firstErr, count, logSum, invalid)
          | ValueNone ->
              match unwrapValue value with
              | CellValue.Error e -> struct (ValueSome e, count, logSum, invalid)
              | CellValue.Number n ->
                  if n <= 0.0 then struct (ValueNone, count, logSum, true)
                  else struct (ValueNone, count + 1, logSum + log n, invalid)
              | CellValue.Boolean b ->
                  let n = if b then 1.0 else 0.0
                  if n <= 0.0 then struct (ValueNone, count, logSum, true)
                  else struct (ValueNone, count + 1, logSum + log n, invalid)
              | _ -> struct (ValueNone, count, logSum, invalid)
      Extract = fun (struct (firstErr, count, logSum, invalid)) ->
          match firstErr with
          | ValueSome e -> err e
          | ValueNone ->
              if invalid || count = 0 then err Err.num
              else CellValue.Number (exp (logSum / float count)) }

and private harmonicMeanFold : Fold<CellValue, struct (string voption * int * float * bool), CellValue> =
    { Init = struct (ValueNone, 0, 0.0, false)
      Step = fun (struct (firstErr, count, reciprocalSum, invalid)) value ->
          match firstErr with
          | ValueSome _ -> struct (firstErr, count, reciprocalSum, invalid)
          | ValueNone ->
              match unwrapValue value with
              | CellValue.Error e -> struct (ValueSome e, count, reciprocalSum, invalid)
              | CellValue.Number n ->
                  if n <= 0.0 then struct (ValueNone, count, reciprocalSum, true)
                  else struct (ValueNone, count + 1, reciprocalSum + 1.0 / n, invalid)
              | CellValue.Boolean b ->
                  let n = if b then 1.0 else 0.0
                  if n <= 0.0 then struct (ValueNone, count, reciprocalSum, true)
                  else struct (ValueNone, count + 1, reciprocalSum + 1.0 / n, invalid)
              | _ -> struct (ValueNone, count, reciprocalSum, invalid)
      Extract = fun (struct (firstErr, count, reciprocalSum, invalid)) ->
          match firstErr with
          | ValueSome e -> err e
          | ValueNone ->
              if invalid || count = 0 then err Err.num
              else CellValue.Number (float count / reciprocalSum) }

/// Read a cell straight from the model (used by lookup functions that
/// index into a range rectangle).
and private gridCell (model: SpreadsheetModel) row col : CellValue =
    getCellValue { Row = row; Col = col } model |> unwrapValue

and evalFunction (name: string) (args: Expr list) (ctx: EvalContext) : CellValue =
    let date1904 = date1904Of ctx
    let serialToDate (s: float) = serialToDateIn date1904 s
    let dateToSerial (d: DateTime) = dateToSerialIn date1904 d
    let model = ctx.Model
    let workbook = ctx.Workbook
    let fnName = name.ToUpperInvariant()
    let argsArray =
        match fnName with
        | "SUM" | "PRODUCT" | "AVERAGE" | "MIN" | "MAX" | "MEDIAN"
        | "COUNT" | "COUNTA" | "COUNTBLANK"
        | "STDEV" | "STDEV.S" | "STDEVP" | "STDEV.P"
        | "VAR" | "VAR.S" | "VARP" | "VAR.P" -> Array.empty
        | _ ->
            // (The arguments by position, made once for the call's list of them; it is only read.)
            if args.IsEmpty then Array.empty
            else
                let known = formTableGet argsArrayTable (box args)
                if isNull known then
                    let made = List.toArray args
                    formTableSet argsArrayTable (box args) (box made)
                    made
                else unbox<Expr array> known
    let nargs = argsArray.Length
    let argExpr i = argsArray.[i]
    let argRaw i = eval (argExpr i) ctx
    let argNum i = coerceNumber (argRaw i)
    let argText i = coerceText (argRaw i)
    let argBool i = coerceBool (argRaw i)
    // An optional number: `dflt` when it is left out or empty; an error in it is the call's.
    let optNum i (dflt: float) = if nargs > i then (match argExpr i with MissingArg -> Ok dflt | _ -> argNum i) else Ok dflt
    // Apply a unary numeric function to arg 0 with error propagation.
    let unaryNum f =
        match argNum 0 with
        | Ok x ->
            // Overflow (EXP(1000)) and undefined results are #NUM!.
            (match f x with
             | CellValue.Number result when Double.IsInfinity result || Double.IsNaN result -> err Err.num
             | value -> value)
        | Result.Error e -> err e
    let numericBufferOfRangeValue (e: Expr) =
        let bounds = rangeBounds e
        let capacity =
            match bounds with
            | Some (sr, sc, er, ec) ->
                let area = int64 (er - sr + 1) * int64 (ec - sc + 1)
                let maxValues = min area (int64 model.Cells.Count)
                if maxValues > int64 Int32.MaxValue then Int32.MaxValue else int maxValues
            | None -> 1
        let values = Array.zeroCreate<float> capacity
        let mutable count = 0
        let mutable firstErr = ValueNone
        let addValue value =
            match firstErr with
            | ValueSome _ -> ()
            | ValueNone ->
                match unwrapValue value with
                | CellValue.Error e -> firstErr <- ValueSome e
                | CellValue.Number n ->
                    values.[count] <- n
                    count <- count + 1
                // Booleans and text read from a range are not numbers.
                | _ -> ()
        match bounds with
        | Some (sr, sc, er, ec) ->
            let area = int64 (er - sr + 1) * int64 (ec - sc + 1)
            if area >= 1024L && int64 model.Cells.Count <= area * 4L then
                model.Cells
                |> Map.iter (fun addr cell ->
                    if addr.Row >= sr && addr.Row <= er && addr.Col >= sc && addr.Col <= ec then
                        addValue cell.Value)
            else
                for r in sr .. er do
                    for c in sc .. ec do
                        addValue (gridCell model r c)
        | None -> addValue (directNumericArgument (unwrapValue (eval e ctx)))
        match firstErr with
        | ValueSome e -> Result.Error e
        | ValueNone -> Ok (struct (values, count))
    let numericBufferOfArithmeticValue (e: Expr) =
        let collectNumeric count readNumber =
            let values = Array.zeroCreate<float> count
            let mutable firstErr = ValueNone
            for offset in 0 .. count - 1 do
                match firstErr with
                | ValueSome _ -> ()
                | ValueNone ->
                    match readNumber offset with
                    | Result.Error e -> firstErr <- ValueSome e
                    | Ok n -> values.[offset] <- n
            match firstErr with
            | ValueSome e -> Result.Error e
            | ValueNone -> Ok (struct (values, count))
        let collectStoredDenseRangeNumeric (source: RangeAccess) readNumber =
            let values = Array.zeroCreate<float> source.Count
            let mutable storedInRange = 0
            let mutable firstErr = ValueNone
            let struct (er, ec) = rangeEndForCount source source.Count
            source.M.Cells
            |> Map.iter (fun addr cell ->
                if addr.Row >= source.SR && addr.Row <= er && addr.Col >= source.SC && addr.Col <= ec then
                    storedInRange <- storedInRange + 1
                    match firstErr with
                    | ValueSome _ -> ()
                    | ValueNone ->
                        match readNumber cell.Value with
                        | Result.Error e -> firstErr <- ValueSome e
                        | Ok n ->
                            let offset = (addr.Row - source.SR) * source.W + (addr.Col - source.SC)
                            values.[offset] <- n)
            if storedInRange = source.Count then
                match firstErr with
                | ValueSome e -> Result.Error e
                | ValueNone -> Ok (struct (values, source.Count))
                |> ValueSome
            else
                ValueNone
        let unaryNumber op value =
            match coerceNumber (unwrapValue value) with
            | Result.Error e -> Result.Error e
            | Ok n ->
                match op with
                | "-" -> Ok -n
                | "+" -> Ok n
                | "%" -> Ok (flushProductUnderflow (n / 100.0))
                | _ -> Result.Error Err.value
        let arithmeticNumber op a b =
            match op with
            | "+" -> Ok (a + b)
            | "-" -> Ok (a - b)
            | "*" -> Ok (flushProductUnderflow (a * b))
            | "/" -> if b > -minNormal && b < minNormal then Result.Error Err.div0 else Ok (flushProductUnderflow (a / b))
            | "^" -> powerNumber a b
            | _ -> Result.Error Err.value
        let collectStoredDenseRangePairNumeric (left: RangeAccess) (right: RangeAccess) op =
            let count = left.Count
            let leftValues = Array.zeroCreate<float> count
            let rightValues = Array.zeroCreate<float> count
            let seen = Array.zeroCreate<byte> count
            let mutable unsupported = false
            let addValue side (access: RangeAccess) addr value =
                if not unsupported then
                    match tryRangeOffsetForCount access count addr with
                    | ValueSome offset ->
                        match coerceNumber (unwrapValue value) with
                        | Ok n ->
                            if side = 1uy then
                                leftValues.[offset] <- n
                            else
                                rightValues.[offset] <- n
                            seen.[offset] <- seen.[offset] ||| side
                        | Result.Error _ -> unsupported <- true
                    | ValueNone -> ()
            left.M.Cells |> Map.iter (fun addr cell -> addValue 1uy left addr cell.Value)
            right.M.Cells |> Map.iter (fun addr cell -> addValue 2uy right addr cell.Value)
            if unsupported then
                ValueNone
            else
                let mutable allStored = true
                let mutable i = 0
                while allStored && i < count do
                    if seen.[i] <> 3uy then
                        allStored <- false
                    i <- i + 1
                if not allStored then
                    ValueNone
                else
                    let values = Array.zeroCreate<float> count
                    let mutable firstErr = ValueNone
                    let mutable offset = 0
                    while ValueOption.isNone firstErr && offset < count do
                        match arithmeticNumber op leftValues.[offset] rightValues.[offset] with
                        | Ok n -> values.[offset] <- n
                        | Result.Error e -> firstErr <- ValueSome e
                        offset <- offset + 1
                    match firstErr with
                    | ValueSome e -> Result.Error e
                    | ValueNone -> Ok (struct (values, count))
                    |> ValueSome
        let rangeScalarNumber op value scalarResult =
            match unwrapValue value with
            | CellValue.Error e -> Result.Error e
            | unwrapped ->
                match scalarResult with
                | Result.Error e -> Result.Error e
                | Ok scalar ->
                    match coerceNumber unwrapped with
                    | Result.Error e -> Result.Error e
                    | Ok n -> arithmeticNumber op n scalar
        let scalarRangeNumber op scalarResult value =
            match scalarResult with
            | Result.Error e -> Result.Error e
            | Ok scalar ->
                match unwrapValue value with
                | CellValue.Error e -> Result.Error e
                | unwrapped ->
                    match coerceNumber unwrapped with
                    | Result.Error e -> Result.Error e
                    | Ok n -> arithmeticNumber op scalar n
        let rangeRangeNumber op leftValue rightValue =
            match unwrapValue leftValue, unwrapValue rightValue with
            | CellValue.Error e, _ -> Result.Error e
            | _, CellValue.Error e -> Result.Error e
            | left, right ->
                match coerceNumber left, coerceNumber right with
                | Result.Error e, _ -> Result.Error e
                | _, Result.Error e -> Result.Error e
                | Ok a, Ok b -> arithmeticNumber op a b
        match e with
        | UnaryOp (op, sourceExpr) when op = "+" || op = "-" || op = "%" ->
            match rangeAccessOf model workbook sourceExpr with
            | Some source when source.Count > 0 && int64 source.M.Cells.Count * 4L >= int64 source.Count ->
                match collectStoredDenseRangeNumeric source (unaryNumber op) with
                | ValueSome result -> ValueSome result
                | ValueNone -> collectNumeric source.Count (fun offset -> unaryNumber op (rangeRead source offset)) |> ValueSome
            | _ -> ValueNone
        | BinOp (leftExpr, op, rightExpr) when op = "+" || op = "-" || op = "*" || op = "/" || op = "^" ->
            match rangeAccessOf model workbook leftExpr, rangeAccessOf model workbook rightExpr with
            | Some source, None when source.Count > 0 && int64 source.M.Cells.Count * 4L >= int64 source.Count && not (isArrayExpr rightExpr) ->
                let scalar = eval rightExpr ctx
                let scalarNumber = coerceNumber (unwrapValue scalar)
                match collectStoredDenseRangeNumeric source (fun value -> rangeScalarNumber op value scalarNumber) with
                | ValueSome result -> ValueSome result
                | ValueNone -> collectNumeric source.Count (fun offset -> rangeScalarNumber op (rangeRead source offset) scalarNumber) |> ValueSome
            | None, Some source when source.Count > 0 && int64 source.M.Cells.Count * 4L >= int64 source.Count && not (isArrayExpr leftExpr) ->
                let scalar = eval leftExpr ctx
                let scalarNumber = coerceNumber (unwrapValue scalar)
                match collectStoredDenseRangeNumeric source (scalarRangeNumber op scalarNumber) with
                | ValueSome result -> ValueSome result
                | ValueNone -> collectNumeric source.Count (fun offset -> scalarRangeNumber op scalarNumber (rangeRead source offset)) |> ValueSome
            | Some left, Some right
                when left.Count = right.Count
                     && left.W = right.W
                     && left.Count > 0
                     && (int64 left.M.Cells.Count + int64 right.M.Cells.Count) * 4L >= int64 left.Count ->
                match collectStoredDenseRangePairNumeric left right op with
                | ValueSome result -> ValueSome result
                | ValueNone ->
                    collectNumeric left.Count (fun offset -> rangeRangeNumber op (rangeRead left offset) (rangeRead right offset))
                    |> ValueSome
            | _ -> ValueNone
        | _ -> ValueNone
    let sortedNumberArray (values: float array) =
        Array.sortInPlace values
        values

    let selectNumberAtCount (count: int) (index: int) (values: float array) =
        let swap i j =
            let tmp = values.[i]
            values.[i] <- values.[j]
            values.[j] <- tmp
        if count <= 0 then
            nan
        elif index = 0 then
            let mutable best = values.[0]
            for i = 1 to count - 1 do
                if compare values.[i] best < 0 then
                    best <- values.[i]
            best
        elif index = count - 1 then
            let mutable best = values.[0]
            for i = 1 to count - 1 do
                if compare values.[i] best > 0 then
                    best <- values.[i]
            best
        else
            let mutable left = 0
            let mutable right = count - 1
            while left < right do
                let pivot = values.[left + (right - left) / 2]
                let mutable i = left
                let mutable j = right
                while i <= j do
                    while compare values.[i] pivot < 0 do
                        i <- i + 1
                    while compare values.[j] pivot > 0 do
                        j <- j - 1
                    if i <= j then
                        swap i j
                        i <- i + 1
                        j <- j - 1
                if index <= j then
                    right <- j
                elif index >= i then
                    left <- i
                else
                    left <- right
            values.[index]

    let selectNumberAt (index: int) (values: float array) =
        selectNumberAtCount values.Length index values

    let rankOneRangeValue largest (source: RangeAccess) =
        let mutable firstErr = ValueNone
        let mutable found = false
        let mutable best = 0.0
        let addValue value =
            match firstErr with
            | ValueSome _ -> ()
            | ValueNone ->
                match unwrapValue value with
                | CellValue.Error e -> firstErr <- ValueSome e
                | CellValue.Number n ->
                    if not found then
                        found <- true
                        best <- n
                    elif largest then
                        if compare n best > 0 then best <- n
                    else
                        if compare n best < 0 then best <- n
                | CellValue.Boolean b ->
                    let n = if b then 1.0 else 0.0
                    if not found then
                        found <- true
                        best <- n
                    elif largest then
                        if compare n best > 0 then best <- n
                    else
                        if compare n best < 0 then best <- n
                | _ -> ()
        let area = int64 source.Count
        if area >= 1024L && int64 source.M.Cells.Count <= area * 4L then
            let struct (er, ec) = rangeEndForCount source source.Count
            source.M.Cells
            |> Map.iter (fun addr cell ->
                if addr.Row >= source.SR && addr.Row <= er && addr.Col >= source.SC && addr.Col <= ec then
                    addValue cell.Value)
        else
            for k in 0 .. source.Count - 1 do
                addValue (rangeRead source k)
        match firstErr with
        | ValueSome e -> Result.Error e
        | ValueNone ->
            if found then Ok best else Result.Error Err.num

    let percentileSelected (values: float array) count k =
        if count <= 0 || k < 0.0 || k > 1.0 then
            err Err.num
        else
            let pos = k * float (count - 1)
            let lo = int (floor pos)
            let frac = pos - float lo
            if lo + 1 >= count || abs frac < 1e-12 then
                CellValue.Number (selectNumberAtCount count lo values)
            else
                let lower = selectNumberAtCount count lo values
                let upper = selectNumberAtCount count (lo + 1) values
                CellValue.Number (lower + frac * (upper - lower))

    let medianSelected values (count: int) =
        percentileSelected values count 0.5

    let modeSelected (values: float array) (count: int) =
        let counts = Dictionary<float, struct (int * int)>(count)
        let mutable bestValue = 0.0
        let mutable bestCount = 0
        let mutable bestFirstIndex = System.Int32.MaxValue
        for i = 0 to count - 1 do
            let value = values.[i]
            let mutable existing = Unchecked.defaultof<struct (int * int)>
            let count, firstIndex =
                if counts.TryGetValue(value, &existing) then
                    let struct (existingCount, existingFirstIndex) = existing
                    existingCount + 1, existingFirstIndex
                else
                    1, i
            counts.[value] <- struct (count, firstIndex)
            if count > 1 && (count > bestCount || (count = bestCount && firstIndex < bestFirstIndex)) then
                bestValue <- value
                bestCount <- count
                bestFirstIndex <- firstIndex
        if bestCount > 1 then CellValue.Number bestValue else err Err.na

    let rankSelected target descending (values: float array) count =
        let mutable found = false
        let mutable rank = 1
        for i = 0 to count - 1 do
            let value = values.[i]
            found <- found || value = target
            if (if descending then value > target else value < target) then
                rank <- rank + 1
        if found then CellValue.Number (float rank) else err Err.na

    let arithmeticNumericAggregate (values: float array) count =
        let mutable total = 0.0
        let mutable productValue = 1.0
        let mutable minValue = infinity
        let mutable maxValue = -infinity
        let mutable varianceMean = 0.0
        let mutable varianceM2 = 0.0
        let mutable geoLogSum = 0.0
        let mutable harmonicReciprocalSum = 0.0
        let mutable invalidExtended = false
        for i = 0 to count - 1 do
            let n = values.[i]
            match fnName with
            | "SUM" | "AVERAGE" | "AVERAGEA" ->
                total <- total + n
            | "PRODUCT" ->
                productValue <- flushProductUnderflow (productValue * n)
            | "SUMSQ" ->
                total <- total + n * n
            | "MIN" | "MINA" ->
                if n < minValue then minValue <- n
            | "MAX" | "MAXA" ->
                if n > maxValue then maxValue <- n
            | "STDEV" | "STDEV.S" | "STDEVP" | "STDEV.P"
            | "VAR" | "VAR.S" | "VARP" | "VAR.P"
            | "DEVSQ" ->
                if i = 0 then
                    varianceMean <- n
                    varianceM2 <- 0.0
                else
                    let newCount = float (i + 1)
                    let delta = n - varianceMean
                    varianceMean <- varianceMean + delta / newCount
                    varianceM2 <- varianceM2 + delta * (n - varianceMean)
            | "AVEDEV" ->
                total <- total + n
            | "GEOMEAN" ->
                if n <= 0.0 then invalidExtended <- true
                else geoLogSum <- geoLogSum + log n
            | "HARMEAN" ->
                if n <= 0.0 then invalidExtended <- true
                else harmonicReciprocalSum <- harmonicReciprocalSum + (1.0 / n)
            | _ -> ()
        let count64 = int64 count
        match fnName with
        | "SUM" -> CellValue.Number total
        | "AVERAGE" | "AVERAGEA" -> if count = 0 then err Err.div0 else CellValue.Number (total / float count)
        | "COUNT" -> CellValue.Number (float count)
        | "PRODUCT" -> CellValue.Number (if count = 0 then 0.0 else productValue)
        | "SUMSQ" -> CellValue.Number total
        | "MIN" | "MINA" -> CellValue.Number (if count = 0 then 0.0 else minValue)
        | "MAX" | "MAXA" -> CellValue.Number (if count = 0 then 0.0 else maxValue)
        | "STDEV" | "STDEV.S" | "STDEVP" | "STDEV.P"
        | "VAR" | "VAR.S" | "VARP" | "VAR.P" ->
            let sample = fnName = "STDEV" || fnName = "STDEV.S" || fnName = "VAR" || fnName = "VAR.S"
            let sqrtResult = fnName = "STDEV" || fnName = "STDEV.S" || fnName = "STDEVP" || fnName = "STDEV.P"
            if (sample && count64 < 2L) || count64 < 1L then err Err.div0
            else
                let variance = varianceM2 / (if sample then float (count64 - 1L) else float count64)
                CellValue.Number (if sqrtResult then sqrt variance else variance)
        | "DEVSQ" -> if count = 0 then err Err.num else CellValue.Number varianceM2
        | "AVEDEV" ->
            if count = 0 then err Err.num
            else
                let mean = total / float count
                let mutable totalDeviation = 0.0
                for i = 0 to count - 1 do
                    totalDeviation <- totalDeviation + abs (values.[i] - mean)
                CellValue.Number (totalDeviation / float count)
        | "GEOMEAN" ->
            if invalidExtended || count = 0 then err Err.num
            else CellValue.Number (exp (geoLogSum / float count))
        | "HARMEAN" ->
            if invalidExtended || count = 0 then err Err.num
            else CellValue.Number (float count / harmonicReciprocalSum)
        | _ -> CellValue.Number total

    let medianArray (values: float array) =
        let a = sortedNumberArray values
        let n = a.Length
        if n = 0 then nan
        elif n % 2 = 1 then a.[n / 2]
        else (a.[n / 2 - 1] + a.[n / 2]) / 2.0
    let sortedCompressedNumbers (values: ResizeArray<struct (float * int64 * int)>) =
        let items = values.ToArray()
        Array.sortInPlaceBy (fun struct (value, _, _) -> value) items
        items
    let compressedCount (values: ResizeArray<struct (float * int64 * int)>) =
        let mutable count = 0L
        for struct (_, repeats, _) in values do
            count <- count + repeats
        count
    let selectCompressedSorted index (values: struct (float * int64 * int) array) =
        let mutable remaining = index
        let mutable selected = 0.0
        let mutable found = false
        let mutable i = 0
        while not found && i < values.Length do
            let struct (value, repeats, _) = values.[i]
            if remaining < repeats then
                selected <- value
                found <- true
            else
                remaining <- remaining - repeats
                i <- i + 1
        selected
    let medianCompressed (values: ResizeArray<struct (float * int64 * int)>) =
        let count = compressedCount values
        if count = 0L then err Err.num
        else
            let sorted = sortedCompressedNumbers values
            if count % 2L = 1L then
                CellValue.Number (selectCompressedSorted (count / 2L) sorted)
            else
                let left = selectCompressedSorted (count / 2L - 1L) sorted
                let right = selectCompressedSorted (count / 2L) sorted
                CellValue.Number ((left + right) / 2.0)
    let rankedCompressed largest k (values: ResizeArray<struct (float * int64 * int)>) =
        let count = compressedCount values
        if count = 0L || k < 1 || int64 k > count then
            err Err.num
        else
            let sorted = sortedCompressedNumbers values
            let index = if largest then count - int64 k else int64 (k - 1)
            CellValue.Number (selectCompressedSorted index sorted)
    let percentileCompressed p (values: ResizeArray<struct (float * int64 * int)>) =
        let count = compressedCount values
        if count = 0L || p < 0.0 || p > 1.0 then
            err Err.num
        else
            let sorted = sortedCompressedNumbers values
            let pos = p * float (count - 1L)
            let lo = int64 (floor pos)
            let frac = pos - float lo
            let left = selectCompressedSorted lo sorted
            let right = if lo + 1L < count then selectCompressedSorted (lo + 1L) sorted else left
            CellValue.Number (left + frac * (right - left))
    let modeCompressed (values: ResizeArray<struct (float * int64 * int)>) =
        let counts = Dictionary<float, struct (int64 * int)>(values.Count)
        let mutable bestValue = 0.0
        let mutable bestCount = 0L
        let mutable bestFirstIndex = System.Int32.MaxValue
        for struct (value, repeats, firstIndex) in values do
            let mutable existing = Unchecked.defaultof<struct (int64 * int)>
            let count, first =
                if counts.TryGetValue(value, &existing) then
                    let struct (existingCount, existingFirstIndex) = existing
                    existingCount + repeats, min existingFirstIndex firstIndex
                else
                    repeats, firstIndex
            counts.[value] <- struct (count, first)
            if count > 1L && (count > bestCount || (count = bestCount && first < bestFirstIndex)) then
                bestValue <- value
                bestCount <- count
                bestFirstIndex <- first
        if bestCount > 1L then CellValue.Number bestValue else err Err.na
    let rankCompressed target descending (values: ResizeArray<struct (float * int64 * int)>) =
        let mutable found = false
        let mutable rank = 1L
        for struct (value, repeats, _) in values do
            found <- found || value = target
            if (if descending then value > target else value < target) then
                rank <- rank + repeats
        if found then CellValue.Number (float rank) else err Err.na
    let isErr v = match unwrapValue v with CellValue.Error _ -> true | _ -> false
    let repeatedProduct product factor repeats =
        if repeats <= 0L then product
        elif factor = 0.0 then flushProductUnderflow (product * factor)
        elif factor = 1.0 then product
        elif factor = -1.0 then if repeats % 2L = 0L then product else -product
        else
            let endValue = product * Math.Pow(factor, float repeats)
            // Keep compressed repetition when no step can cross the normal
            // boundary; otherwise multiply in order until zero is reached.
            if abs factor >= 1.0 && (product <= -minNormal || product >= minNormal)
               || abs factor < 1.0 && (endValue <= -minNormal || endValue >= minNormal) then
                endValue
            elif abs factor < 1.0 && abs product <= 1.0 && abs endValue < minNormal / 2.0 then
                endValue * 0.0
            else
                let mutable value = product
                let mutable remaining = repeats
                while remaining > 0L && value <> 0.0 do
                    value <- flushProductUnderflow (value * factor)
                    remaining <- remaining - 1L
                value
    let varianceOptions =
        match fnName with
        | "STDEV" | "STDEV.S" -> ValueSome (true, true)
        | "STDEVP" | "STDEV.P" -> ValueSome (false, true)
        | "VAR" | "VAR.S" -> ValueSome (true, false)
        | "VARP" | "VAR.P" -> ValueSome (false, false)
        | _ -> ValueNone
    let finishVariance count m2 =
        match varianceOptions with
        | ValueSome (sample, sqrtResult) ->
            if (sample && count < 2L) || count < 1L then err Err.div0
            else
                let variance = m2 / (if sample then float (count - 1L) else float count)
                CellValue.Number (if sqrtResult then sqrt variance else variance)
        | ValueNone -> CellValue.Number 0.0
    let tryArithmeticNumericAggregate sourceExpr =
        match numericBufferOfArithmeticValue sourceExpr with
        | ValueSome (Result.Error e) -> ValueSome (err e)
        | ValueSome (Ok (struct (values, count))) -> ValueSome (arithmeticNumericAggregate values count)
        | ValueNone -> ValueNone
    let tryRangeBinOpAggregate () =
        match args with
        | [ sourceExpr ] ->
            match tryArithmeticNumericAggregate sourceExpr with
            | ValueSome result -> ValueSome result
            | ValueNone ->
                match sourceExpr with
                | BinOp (leftExpr, op, rightExpr) when op = "+" || op = "-" || op = "*" || op = "/" ->
                        let finish count total productValue minValue maxValue varianceM2 geoLogSum harmonicReciprocalSum invalidExtended (aveDevValues: ResizeArray<int64 * float>) =
                            match fnName with
                            | "AVERAGE" -> if count = 0L then err Err.div0 else CellValue.Number (total / float count)
                            | "AVERAGEA" -> if count = 0L then err Err.div0 else CellValue.Number (total / float count)
                            | "COUNT" -> CellValue.Number (float count)
                            | "MIN" -> CellValue.Number (if count = 0L then 0.0 else minValue)
                            | "MINA" -> CellValue.Number (if count = 0L then 0.0 else minValue)
                            | "MAX" -> CellValue.Number (if count = 0L then 0.0 else maxValue)
                            | "MAXA" -> CellValue.Number (if count = 0L then 0.0 else maxValue)
                            | "PRODUCT" -> CellValue.Number (if count = 0L then 0.0 else productValue)
                            | "SUMSQ" -> CellValue.Number total
                            | "STDEV" | "STDEV.S" | "STDEVP" | "STDEV.P"
                            | "VAR" | "VAR.S" | "VARP" | "VAR.P" -> finishVariance count varianceM2
                            | "DEVSQ" -> if count = 0L then err Err.num else CellValue.Number varianceM2
                            | "AVEDEV" ->
                                if count = 0L then err Err.num
                                else
                                    let mean = total / float count
                                    let mutable totalDeviation = 0.0
                                    for repeats, n in aveDevValues do
                                        totalDeviation <- totalDeviation + abs (n - mean) * float repeats
                                    CellValue.Number (totalDeviation / float count)
                            | "GEOMEAN" ->
                                if invalidExtended || count = 0L then err Err.num
                                else CellValue.Number (exp (geoLogSum / float count))
                            | "HARMEAN" ->
                                if invalidExtended || count = 0L then err Err.num
                                else CellValue.Number (float count / harmonicReciprocalSum)
                            | _ -> CellValue.Number total
                        let aggregate count sparseEligible iterOffsets readAt gapValue =
                            let mutable total = 0.0
                            let mutable productValue = 1.0
                            let mutable numericCount = 0L
                            let mutable minValue = infinity
                            let mutable maxValue = -infinity
                            let mutable varianceMean = 0.0
                            let mutable varianceM2 = 0.0
                            let mutable geoLogSum = 0.0
                            let mutable harmonicReciprocalSum = 0.0
                            let mutable invalidExtended = false
                            let aveDevValues = ResizeArray<int64 * float>()
                            let mutable firstErr = ValueNone
                            let addNumber repeats n =
                                match fnName with
                                | "SUM" -> total <- total + n * float repeats
                                | "AVERAGE" ->
                                    total <- total + n * float repeats
                                    numericCount <- numericCount + repeats
                                | "AVERAGEA" ->
                                    total <- total + n * float repeats
                                    numericCount <- numericCount + repeats
                                | "COUNT" -> numericCount <- numericCount + repeats
                                | "PRODUCT" ->
                                    productValue <- repeatedProduct productValue n repeats
                                    numericCount <- numericCount + repeats
                                | "SUMSQ" ->
                                    total <- total + n * n * float repeats
                                    numericCount <- numericCount + repeats
                                | "MIN" ->
                                    minValue <- min minValue n
                                    numericCount <- numericCount + repeats
                                | "MINA" ->
                                    minValue <- min minValue n
                                    numericCount <- numericCount + repeats
                                | "MAX" ->
                                    maxValue <- max maxValue n
                                    numericCount <- numericCount + repeats
                                | "MAXA" ->
                                    maxValue <- max maxValue n
                                    numericCount <- numericCount + repeats
                                | "STDEV" | "STDEV.S" | "STDEVP" | "STDEV.P"
                                | "VAR" | "VAR.S" | "VARP" | "VAR.P" ->
                                    if repeats > 0L then
                                        if numericCount = 0L then
                                            varianceMean <- n
                                            varianceM2 <- 0.0
                                            numericCount <- repeats
                                        else
                                            let oldCount = numericCount
                                            let newCount = numericCount + repeats
                                            let delta = n - varianceMean
                                            varianceMean <- varianceMean + delta * (float repeats / float newCount)
                                            varianceM2 <- varianceM2 + delta * delta * float oldCount * float repeats / float newCount
                                            numericCount <- newCount
                                | "DEVSQ" ->
                                    if repeats > 0L then
                                        if numericCount = 0L then
                                            varianceMean <- n
                                            varianceM2 <- 0.0
                                            numericCount <- repeats
                                        else
                                            let oldCount = numericCount
                                            let newCount = numericCount + repeats
                                            let delta = n - varianceMean
                                            varianceMean <- varianceMean + delta * (float repeats / float newCount)
                                            varianceM2 <- varianceM2 + delta * delta * float oldCount * float repeats / float newCount
                                            numericCount <- newCount
                                | "AVEDEV" ->
                                    if repeats > 0L then
                                        total <- total + n * float repeats
                                        numericCount <- numericCount + repeats
                                        aveDevValues.Add(repeats, n)
                                | "GEOMEAN" ->
                                    if repeats > 0L then
                                        if n <= 0.0 then
                                            invalidExtended <- true
                                        else
                                            geoLogSum <- geoLogSum + log n * float repeats
                                            numericCount <- numericCount + repeats
                                | "HARMEAN" ->
                                    if repeats > 0L then
                                        if n <= 0.0 then
                                            invalidExtended <- true
                                        else
                                            harmonicReciprocalSum <- harmonicReciprocalSum + (1.0 / n) * float repeats
                                            numericCount <- numericCount + repeats
                                | _ -> ()
                            let consume offset repeats value =
                                match unwrapValue value with
                                | CellValue.Error e ->
                                    match firstErr with
                                    | ValueNone -> firstErr <- ValueSome (offset, e)
                                    | ValueSome (existingOffset, _) when offset < existingOffset -> firstErr <- ValueSome (offset, e)
                                    | _ -> ()
                                | CellValue.Number n -> addNumber repeats n
                                | CellValue.Boolean b when fnName <> "COUNT" ->
                                    let n = if b then 1.0 else 0.0
                                    addNumber repeats n
                                | _ -> ()
                            if sparseEligible then
                                let visited = HashSet<int>()
                                if fnName = "PRODUCT" then
                                    // Products are order-dependent at the normal boundary.
                                    // Process the stored offsets and intervening blank runs in row order.
                                    let offsets = ResizeArray<int>()
                                    iterOffsets (fun offset -> if visited.Add offset then offsets.Add offset)
                                    offsets.Sort()
                                    let mutable nextOffset = 0
                                    for offset in offsets do
                                        if offset > nextOffset then consume nextOffset (int64 (offset - nextOffset)) gapValue
                                        consume offset 1L (readAt offset)
                                        nextOffset <- offset + 1
                                    if nextOffset < count then consume nextOffset (int64 (count - nextOffset)) gapValue
                                else
                                    let visit offset =
                                        if visited.Add offset then consume offset 1L (readAt offset)
                                    iterOffsets visit
                                    if visited.Count < count then
                                        consume 0 (int64 (count - visited.Count)) gapValue
                            else
                                let mutable offset = 0
                                while ValueOption.isNone firstErr && offset < count do
                                    consume offset 1L (readAt offset)
                                    offset <- offset + 1
                            match firstErr with
                            | ValueSome (_, e) -> ValueSome (err e)
                            | ValueNone -> ValueSome (finish numericCount total productValue minValue maxValue varianceM2 geoLogSum harmonicReciprocalSum invalidExtended aveDevValues)
                        let sparseOffsets (source: RangeAccess) (emit: int -> unit) =
                            source.M.Cells
                            |> Map.iter (fun addr _ ->
                                match tryRangeOffsetForCount source source.Count addr with
                                | ValueSome offset -> emit offset
                                | ValueNone -> ())
                        match rangeAccessOf model workbook leftExpr, rangeAccessOf model workbook rightExpr with
                        | Some left, Some right when op <> "/" && left.Count = right.Count && left.W = right.W && left.Count > 0 ->
                            let area = int64 left.Count
                            let sparseEligible = area >= 1024L && int64 left.M.Cells.Count + int64 right.M.Cells.Count <= area * 4L
                            let iterOffsets emit =
                                sparseOffsets left emit
                                sparseOffsets right emit
                            aggregate
                                left.Count
                                sparseEligible
                                iterOffsets
                                (fun offset -> applyBin op (rangeRead left offset) (rangeRead right offset))
                                (applyBin op CellValue.Empty CellValue.Empty)
                        | Some source, None when source.Count > 0 && not (isArrayExpr rightExpr) ->
                            let scalar = eval rightExpr ctx
                            let aggregateSource () =
                                let area = int64 source.Count
                                let sparseEligible = area >= 1024L && int64 source.M.Cells.Count <= area * 4L
                                aggregate
                                    source.Count
                                    sparseEligible
                                    (sparseOffsets source)
                                    (fun offset -> applyBin op (rangeRead source offset) scalar)
                                    (applyBin op CellValue.Empty scalar)
                            match op, coerceNumber scalar with
                            | "/", Ok denominator when denominator <> 0.0 -> aggregateSource ()
                            | ("+" | "-" | "*"), _ -> aggregateSource ()
                            | _ -> ValueNone
                        | None, Some source when op <> "/" && source.Count > 0 && not (isArrayExpr leftExpr) ->
                            let scalar = eval leftExpr ctx
                            let area = int64 source.Count
                            let sparseEligible = area >= 1024L && int64 source.M.Cells.Count <= area * 4L
                            aggregate
                                source.Count
                                sparseEligible
                                (sparseOffsets source)
                                (fun offset -> applyBin op scalar (rangeRead source offset))
                                (applyBin op scalar CellValue.Empty)
                        | _ -> ValueNone
                    | _ -> ValueNone
        | _ -> ValueNone
    let tryRangeUnaryAggregate () =
        match args with
        | [ sourceExpr ] ->
            match tryArithmeticNumericAggregate sourceExpr with
            | ValueSome result -> ValueSome result
            | ValueNone ->
                match sourceExpr with
                | UnaryOp (op, sourceExpr) when op = "+" || op = "-" || op = "%" ->
                    let finish count total productValue minValue maxValue varianceM2 geoLogSum harmonicReciprocalSum invalidExtended (aveDevValues: ResizeArray<int64 * float>) =
                        match fnName with
                        | "AVERAGE" -> if count = 0L then err Err.div0 else CellValue.Number (total / float count)
                        | "AVERAGEA" -> if count = 0L then err Err.div0 else CellValue.Number (total / float count)
                        | "COUNT" -> CellValue.Number (float count)
                        | "MIN" -> CellValue.Number (if count = 0L then 0.0 else minValue)
                        | "MINA" -> CellValue.Number (if count = 0L then 0.0 else minValue)
                        | "MAX" -> CellValue.Number (if count = 0L then 0.0 else maxValue)
                        | "MAXA" -> CellValue.Number (if count = 0L then 0.0 else maxValue)
                        | "PRODUCT" -> CellValue.Number (if count = 0L then 0.0 else productValue)
                        | "SUMSQ" -> CellValue.Number total
                        | "STDEV" | "STDEV.S" | "STDEVP" | "STDEV.P"
                        | "VAR" | "VAR.S" | "VARP" | "VAR.P" -> finishVariance count varianceM2
                        | "DEVSQ" -> if count = 0L then err Err.num else CellValue.Number varianceM2
                        | "AVEDEV" ->
                            if count = 0L then err Err.num
                            else
                                let mean = total / float count
                                let mutable totalDeviation = 0.0
                                for repeats, n in aveDevValues do
                                    totalDeviation <- totalDeviation + abs (n - mean) * float repeats
                                CellValue.Number (totalDeviation / float count)
                        | "GEOMEAN" ->
                            if invalidExtended || count = 0L then err Err.num
                            else CellValue.Number (exp (geoLogSum / float count))
                        | "HARMEAN" ->
                            if invalidExtended || count = 0L then err Err.num
                            else CellValue.Number (float count / harmonicReciprocalSum)
                        | _ -> CellValue.Number total
                    let aggregate count sparseEligible iterOffsets readAt gapValue =
                        let mutable total = 0.0
                        let mutable productValue = 1.0
                        let mutable numericCount = 0L
                        let mutable minValue = infinity
                        let mutable maxValue = -infinity
                        let mutable varianceMean = 0.0
                        let mutable varianceM2 = 0.0
                        let mutable geoLogSum = 0.0
                        let mutable harmonicReciprocalSum = 0.0
                        let mutable invalidExtended = false
                        let aveDevValues = ResizeArray<int64 * float>()
                        let mutable firstErr = ValueNone
                        let addNumber repeats n =
                            match fnName with
                            | "SUM" -> total <- total + n * float repeats
                            | "AVERAGE" ->
                                total <- total + n * float repeats
                                numericCount <- numericCount + repeats
                            | "AVERAGEA" ->
                                total <- total + n * float repeats
                                numericCount <- numericCount + repeats
                            | "COUNT" -> numericCount <- numericCount + repeats
                            | "PRODUCT" ->
                                productValue <- repeatedProduct productValue n repeats
                                numericCount <- numericCount + repeats
                            | "SUMSQ" ->
                                total <- total + n * n * float repeats
                                numericCount <- numericCount + repeats
                            | "MIN" ->
                                minValue <- min minValue n
                                numericCount <- numericCount + repeats
                            | "MINA" ->
                                minValue <- min minValue n
                                numericCount <- numericCount + repeats
                            | "MAX" ->
                                maxValue <- max maxValue n
                                numericCount <- numericCount + repeats
                            | "MAXA" ->
                                maxValue <- max maxValue n
                                numericCount <- numericCount + repeats
                            | "STDEV" | "STDEV.S" | "STDEVP" | "STDEV.P"
                            | "VAR" | "VAR.S" | "VARP" | "VAR.P" ->
                                if repeats > 0L then
                                    if numericCount = 0L then
                                        varianceMean <- n
                                        varianceM2 <- 0.0
                                        numericCount <- repeats
                                    else
                                        let oldCount = numericCount
                                        let newCount = numericCount + repeats
                                        let delta = n - varianceMean
                                        varianceMean <- varianceMean + delta * (float repeats / float newCount)
                                        varianceM2 <- varianceM2 + delta * delta * float oldCount * float repeats / float newCount
                                        numericCount <- newCount
                            | "DEVSQ" ->
                                if repeats > 0L then
                                    if numericCount = 0L then
                                        varianceMean <- n
                                        varianceM2 <- 0.0
                                        numericCount <- repeats
                                    else
                                        let oldCount = numericCount
                                        let newCount = numericCount + repeats
                                        let delta = n - varianceMean
                                        varianceMean <- varianceMean + delta * (float repeats / float newCount)
                                        varianceM2 <- varianceM2 + delta * delta * float oldCount * float repeats / float newCount
                                        numericCount <- newCount
                            | "AVEDEV" ->
                                if repeats > 0L then
                                    total <- total + n * float repeats
                                    numericCount <- numericCount + repeats
                                    aveDevValues.Add(repeats, n)
                            | "GEOMEAN" ->
                                if repeats > 0L then
                                    if n <= 0.0 then
                                        invalidExtended <- true
                                    else
                                        geoLogSum <- geoLogSum + log n * float repeats
                                        numericCount <- numericCount + repeats
                            | "HARMEAN" ->
                                if repeats > 0L then
                                    if n <= 0.0 then
                                        invalidExtended <- true
                                    else
                                        harmonicReciprocalSum <- harmonicReciprocalSum + (1.0 / n) * float repeats
                                        numericCount <- numericCount + repeats
                            | _ -> ()
                        let consume offset repeats value =
                            match unwrapValue value with
                            | CellValue.Error e ->
                                match firstErr with
                                | ValueNone -> firstErr <- ValueSome (offset, e)
                                | ValueSome (existingOffset, _) when offset < existingOffset -> firstErr <- ValueSome (offset, e)
                                | _ -> ()
                            | CellValue.Number n -> addNumber repeats n
                            | CellValue.Boolean b when fnName <> "COUNT" ->
                                let n = if b then 1.0 else 0.0
                                addNumber repeats n
                            | _ -> ()
                        if sparseEligible then
                            let visited = HashSet<int>()
                            if fnName = "PRODUCT" then
                                let offsets = ResizeArray<int>()
                                iterOffsets (fun offset -> if visited.Add offset then offsets.Add offset)
                                offsets.Sort()
                                let mutable nextOffset = 0
                                for offset in offsets do
                                    if offset > nextOffset then consume nextOffset (int64 (offset - nextOffset)) gapValue
                                    consume offset 1L (readAt offset)
                                    nextOffset <- offset + 1
                                if nextOffset < count then consume nextOffset (int64 (count - nextOffset)) gapValue
                            else
                                let visit offset =
                                    if visited.Add offset then consume offset 1L (readAt offset)
                                iterOffsets visit
                                if visited.Count < count then
                                    consume 0 (int64 (count - visited.Count)) gapValue
                        else
                            let mutable offset = 0
                            while ValueOption.isNone firstErr && offset < count do
                                consume offset 1L (readAt offset)
                                offset <- offset + 1
                        match firstErr with
                        | ValueSome (_, e) -> ValueSome (err e)
                        | ValueNone -> ValueSome (finish numericCount total productValue minValue maxValue varianceM2 geoLogSum harmonicReciprocalSum invalidExtended aveDevValues)
                    match rangeAccessOf model workbook sourceExpr with
                    | Some source when source.Count > 0 ->
                        let area = int64 source.Count
                        let sparseEligible = area >= 1024L && int64 source.M.Cells.Count <= area * 4L
                        let sparseOffsets emit =
                            source.M.Cells
                            |> Map.iter (fun addr _ ->
                                match tryRangeOffsetForCount source source.Count addr with
                                | ValueSome offset -> emit offset
                                | ValueNone -> ())
                        aggregate
                            source.Count
                            sparseEligible
                            sparseOffsets
                            (fun offset -> applyUnary op (rangeRead source offset))
                            (applyUnary op CellValue.Empty)
                    | _ -> ValueNone
                | _ -> ValueNone
        | _ -> ValueNone
    let trySparseMapAggregate () =
        match args with
        | [ Function (mapName, [ sourceExpr; lambdaExpr ]) ] when equalsOrdinalIgnoreCase mapName "MAP" ->
            match rangeAccessOf model workbook sourceExpr, lambdaPartsOf lambdaExpr with
            | Some source, Some [ NameRef paramName; body ]
                when source.Count > 0
                     && (fnName = "SUM" || fnName = "AVERAGE" || fnName = "COUNT" || fnName = "MIN" || fnName = "MAX")
                     && lambdaBodyUsesOnlySimpleArithmetic paramName body
                     && not (containsVolatile body) ->
                let area = int64 source.Count
                if area < 1024L || int64 source.M.Cells.Count > area * 4L then
                    ValueNone
                else
                    let lambdaParts = [ NameRef paramName; body ]
                    let evalMapValue value root =
                        let frames = ResizeArray<struct (int * Expr * string)>()
                        let results = ResizeArray<ValueOption<CellValue>>()
                        frames.Add(struct (0, root, ""))

                        let popResult () =
                            let last = results.Count - 1
                            let result = results.[last]
                            Portable.Stack.dropLast results
                            result

                        while frames.Count > 0 do
                            let last = frames.Count - 1
                            let struct (kind, expr, op) = frames.[last]
                            Portable.Stack.dropLast frames

                            match kind with
                            | 1 ->
                                match popResult () with
                                | ValueSome mapped -> results.Add(ValueSome (applyUnary op mapped))
                                | ValueNone -> results.Add ValueNone
                            | 2 ->
                                let right = popResult ()
                                let left = popResult ()
                                match left, right with
                                | ValueSome leftValue, ValueSome rightValue ->
                                    results.Add(ValueSome (applyBin op leftValue rightValue))
                                | _ -> results.Add ValueNone
                            | _ ->
                                match expr with
                                | Number n -> results.Add(ValueSome (CellValue.Number n))
                                | Str s -> results.Add(ValueSome (CellValue.String s))
                                | Bool b -> results.Add(ValueSome (CellValue.Boolean b))
                                | ErrorLit e -> results.Add(ValueSome (err e))
                                | NameRef name when equalsOrdinalIgnoreCase name paramName ->
                                    match value with
                                    | CellValue.Empty -> results.Add(ValueSome (CellValue.Number 0.0))
                                    | _ -> results.Add(ValueSome value)
                                | UnaryOp (op, inner) ->
                                    frames.Add(struct (1, Number 0.0, op))
                                    frames.Add(struct (0, inner, ""))
                                | BinOp (left, op, right) ->
                                    frames.Add(struct (2, Number 0.0, op))
                                    frames.Add(struct (0, right, ""))
                                    frames.Add(struct (0, left, ""))
                                | _ -> results.Add ValueNone

                        if results.Count = 1 then results.[0] else ValueNone
                    let inline numberOk n = struct (0, n, "")
                    let inline numberError e = struct (1, 0.0, e)
                    let numberUnsupported = struct (2, 0.0, "")
                    let numberOfMapValue value =
                        match unwrapValue value with
                        | CellValue.Number n -> numberOk n
                        | CellValue.Boolean b -> numberOk (if b then 1.0 else 0.0)
                        | CellValue.Empty -> numberOk 0.0
                        | CellValue.Error e -> numberError e
                        | CellValue.String s ->
                            let struct (start, length) = trimBounds s
                            match Double.TryParse(NumberText.cutToSignificantDigits (trimmedNumberText s start length), NumberStyles.Float, CultureInfo.InvariantCulture) with
                            | true, n -> numberOk n
                            | _ -> numberUnsupported
                        | _ -> numberUnsupported
                    let evalMapNumber value root =
                        let frames = ResizeArray<struct (int * Expr * string)>()
                        let results = ResizeArray<struct (int * float * string)>()
                        frames.Add(struct (0, root, ""))

                        let popResult () =
                            let last = results.Count - 1
                            let result = results.[last]
                            Portable.Stack.dropLast results
                            result

                        while frames.Count > 0 do
                            let last = frames.Count - 1
                            let struct (kind, expr, op) = frames.[last]
                            Portable.Stack.dropLast frames

                            match kind with
                            | 1 ->
                                let struct (innerKind, n, e) = popResult ()
                                if innerKind = 0 then
                                    match op with
                                    | "+" -> results.Add(numberOk n)
                                    | "-" -> results.Add(numberOk -n)
                                    | "%" -> results.Add(numberOk (flushProductUnderflow (n / 100.0)))
                                    | _ -> results.Add numberUnsupported
                                elif innerKind = 1 then
                                    results.Add(numberError e)
                                else
                                    results.Add numberUnsupported
                            | 2 ->
                                let struct (rightKind, rightNumber, rightError) = popResult ()
                                let struct (leftKind, leftNumber, leftError) = popResult ()
                                if leftKind = 1 then
                                    results.Add(numberError leftError)
                                elif rightKind = 1 then
                                    results.Add(numberError rightError)
                                elif leftKind <> 0 || rightKind <> 0 then
                                    results.Add numberUnsupported
                                else
                                    match op with
                                    | "+" -> results.Add(numberOk (leftNumber + rightNumber))
                                    | "-" -> results.Add(numberOk (leftNumber - rightNumber))
                                    | "*" -> results.Add(numberOk (flushProductUnderflow (leftNumber * rightNumber)))
                                    | "/" when rightNumber > -minNormal && rightNumber < minNormal -> results.Add(numberError Err.div0)
                                    | "/" -> results.Add(numberOk (flushProductUnderflow (leftNumber / rightNumber)))
                                    | "^" ->
                                        match powerNumber leftNumber rightNumber with
                                        | Ok n -> results.Add(numberOk n)
                                        | Result.Error e -> results.Add(numberError e)
                                    | _ -> results.Add numberUnsupported
                            | _ ->
                                match expr with
                                | Number n -> results.Add(numberOk n)
                                | Bool b -> results.Add(numberOk (if b then 1.0 else 0.0))
                                | ErrorLit e -> results.Add(numberError e)
                                | Str s ->
                                    let struct (start, length) = trimBounds s
                                    match Double.TryParse(NumberText.cutToSignificantDigits (trimmedNumberText s start length), NumberStyles.Float, CultureInfo.InvariantCulture) with
                                    | true, n -> results.Add(numberOk n)
                                    | _ -> results.Add numberUnsupported
                                | NameRef name when equalsOrdinalIgnoreCase name paramName -> results.Add(numberOfMapValue value)
                                | UnaryOp (op, inner) ->
                                    frames.Add(struct (1, Number 0.0, op))
                                    frames.Add(struct (0, inner, ""))
                                | BinOp (left, op, right) ->
                                    match op with
                                    | "+" | "-" | "*" | "/" | "^" ->
                                        frames.Add(struct (2, Number 0.0, op))
                                        frames.Add(struct (0, right, ""))
                                        frames.Add(struct (0, left, ""))
                                    | _ -> results.Add numberUnsupported
                                | _ -> results.Add numberUnsupported

                        if results.Count = 1 then results.[0] else numberUnsupported
                    let tryEvalMapComparison value expr =
                        match expr with
                        | BinOp (left, op, right) when op = "=" || op = "<>" || op = "<" || op = ">" || op = "<=" || op = ">=" ->
                            let struct (leftKind, leftNumber, leftError) = evalMapNumber value left
                            let struct (rightKind, rightNumber, rightError) = evalMapNumber value right
                            if leftKind = 1 then ValueSome (struct (1, false, leftError))
                            elif rightKind = 1 then ValueSome (struct (1, false, rightError))
                            elif leftKind <> 0 || rightKind <> 0 then ValueNone
                            else
                                let c = compareNumbers leftNumber rightNumber
                                let matched =
                                    match op with
                                    | "=" -> c = 0
                                    | "<>" -> c <> 0
                                    | "<" -> c < 0
                                    | ">" -> c > 0
                                    | "<=" -> c <= 0
                                    | _ -> c >= 0
                                ValueSome (struct (0, matched, ""))
                        | _ -> ValueNone
                    let mutable firstErr = ValueNone
                    let mutable total = 0.0
                    let mutable count = 0L
                    let mutable minValue = infinity
                    let mutable maxValue = -infinity
                    let addNumeric repeats n =
                        match fnName with
                        | "SUM" ->
                            total <- total + n * float repeats
                        | "AVERAGE" ->
                            total <- total + n * float repeats
                            count <- count + repeats
                        | "COUNT" ->
                            count <- count + repeats
                        | "MIN" ->
                            minValue <- min minValue n
                            count <- count + repeats
                        | "MAX" ->
                            maxValue <- max maxValue n
                            count <- count + repeats
                        | _ -> ()
                    let consumeMapped repeats mapped =
                        match mapped with
                        | CellValue.Error e -> firstErr <- ValueSome e
                        | CellValue.Number n -> addNumeric repeats n
                        | CellValue.Boolean b when fnName <> "COUNT" -> addNumeric repeats (if b then 1.0 else 0.0)
                        | _ -> ()
                    let addMapped repeats value =
                        if repeats > 0L && ValueOption.isNone firstErr then
                            let struct (numberKind, number, numberErrorText) = evalMapNumber value body
                            if numberKind = 0 then
                                addNumeric repeats number
                            elif numberKind = 1 then
                                firstErr <- ValueSome numberErrorText
                            else
                                match tryEvalMapComparison value body with
                                | ValueSome (struct (0, matched, _)) ->
                                    if fnName <> "COUNT" then
                                        addNumeric repeats (if matched then 1.0 else 0.0)
                                | ValueSome (struct (1, _, errorText)) ->
                                    firstErr <- ValueSome errorText
                                | _ ->
                                    let mapped =
                                        match evalMapValue value body with
                                        | ValueSome mapped -> unwrapValue mapped
                                        | ValueNone -> eval (applyLambdaExpr lambdaParts [ valueToExpr value ]) ctx |> unwrapValue
                                    consumeMapped repeats mapped
                    let mutable nextOffset = 0
                    source.M.Cells
                    |> Map.iter (fun addr cell ->
                        match tryRangeOffsetForCount source source.Count addr with
                        | ValueSome offset when offset >= nextOffset ->
                            addMapped (int64 (offset - nextOffset)) CellValue.Empty
                            addMapped 1L (unwrapValue cell.Value)
                            nextOffset <- offset + 1
                        | _ -> ())
                    addMapped (int64 (source.Count - nextOffset)) CellValue.Empty
                    match firstErr with
                    | ValueSome e -> ValueSome (err e)
                    | ValueNone ->
                        let result =
                            match fnName with
                            | "AVERAGE" -> if count = 0L then err Err.div0 else CellValue.Number (total / float count)
                            | "COUNT" -> CellValue.Number (float count)
                            | "MIN" -> CellValue.Number (if count = 0L then 0.0 else minValue)
                            | "MAX" -> CellValue.Number (if count = 0L then 0.0 else maxValue)
                            | _ -> CellValue.Number total
                        ValueSome result
            | _ -> ValueNone
        | _ -> ValueNone
    let tryMakeArrayAggregate () =
        let isSupportedAggregate =
            match fnName with
            | "SUM" | "PRODUCT" | "AVERAGE" | "COUNT" | "MIN" | "MAX" -> true
            | _ -> false
        let evalSimple row col rowName colName root =
            let frames = ResizeArray<struct (int * Expr * string)>()
            let results = ResizeArray<ValueOption<CellValue>>()
            frames.Add(struct (0, root, ""))

            let popResult () =
                let last = results.Count - 1
                let result = results.[last]
                Portable.Stack.dropLast results
                result

            while frames.Count > 0 do
                let last = frames.Count - 1
                let struct (kind, expr, op) = frames.[last]
                Portable.Stack.dropLast frames

                match kind with
                | 1 ->
                    match popResult () with
                    | ValueSome value -> results.Add(ValueSome (applyUnary op value))
                    | ValueNone -> results.Add ValueNone
                | 2 ->
                    let right = popResult ()
                    let left = popResult ()
                    match left, right with
                    | ValueSome leftValue, ValueSome rightValue ->
                        results.Add(ValueSome (applyBin op leftValue rightValue))
                    | _ -> results.Add ValueNone
                | _ ->
                    match expr with
                    | Number n -> results.Add(ValueSome (CellValue.Number n))
                    | Bool b -> results.Add(ValueSome (CellValue.Boolean b))
                    | Str s -> results.Add(ValueSome (CellValue.String s))
                    | ErrorLit e -> results.Add(ValueSome (err e))
                    | NameRef name when equalsOrdinalIgnoreCase name rowName -> results.Add(ValueSome (CellValue.Number row))
                    | NameRef name when equalsOrdinalIgnoreCase name colName -> results.Add(ValueSome (CellValue.Number col))
                    | UnaryOp (op, inner) ->
                        frames.Add(struct (1, Number 0.0, op))
                        frames.Add(struct (0, inner, ""))
                    | BinOp (left, op, right) ->
                        frames.Add(struct (2, Number 0.0, op))
                        frames.Add(struct (0, right, ""))
                        frames.Add(struct (0, left, ""))
                    | _ -> results.Add ValueNone

            if results.Count = 1 then results.[0] else ValueNone
        let evalNumeric row col rowName colName root =
            let inline numericOk n = struct (0, n, "")
            let inline numericError e = struct (1, 0.0, e)
            let numericUnsupported = struct (2, 0.0, "")
            let inline numericBool b = struct (3, (if b then 1.0 else 0.0), "")
            let frames = ResizeArray<struct (int * Expr * string)>()
            let results = ResizeArray<struct (int * float * string)>()
            frames.Add(struct (0, root, ""))

            let popResult () =
                let last = results.Count - 1
                let result = results.[last]
                Portable.Stack.dropLast results
                result

            while frames.Count > 0 do
                let last = frames.Count - 1
                let struct (kind, expr, op) = frames.[last]
                Portable.Stack.dropLast frames

                match kind with
                | 1 ->
                    let struct (innerKind, n, e) = popResult ()
                    if innerKind = 0 then
                        match op with
                        | "+" -> results.Add(numericOk n)
                        | "-" -> results.Add(numericOk -n)
                        | "%" -> results.Add(numericOk (flushProductUnderflow (n / 100.0)))
                        | _ -> results.Add numericUnsupported
                    elif innerKind = 1 then
                        results.Add(numericError e)
                    else
                        results.Add numericUnsupported
                | 2 ->
                    let struct (rightKind, rightValue, rightError) = popResult ()
                    let struct (leftKind, leftValue, leftError) = popResult ()
                    if leftKind = 1 then
                        results.Add(numericError leftError)
                    elif rightKind = 1 then
                        results.Add(numericError rightError)
                    elif leftKind = 2 || rightKind = 2 then
                        results.Add numericUnsupported
                    else
                        match op with
                        | "+" -> results.Add(numericOk (leftValue + rightValue))
                        | "-" -> results.Add(numericOk (leftValue - rightValue))
                        | "*" -> results.Add(numericOk (flushProductUnderflow (leftValue * rightValue)))
                        | "/" when rightValue > -minNormal && rightValue < minNormal -> results.Add(numericError Err.div0)
                        | "/" -> results.Add(numericOk (flushProductUnderflow (leftValue / rightValue)))
                        | "^" ->
                            match powerNumber leftValue rightValue with
                            | Ok n -> results.Add(numericOk n)
                            | Result.Error e -> results.Add(numericError e)
                        | "=" | "<>" | "<" | ">" | "<=" | ">=" ->
                            let c = compareNumbers leftValue rightValue
                            let matched =
                                match op with
                                | "=" -> c = 0
                                | "<>" -> c <> 0
                                | "<" -> c < 0
                                | ">" -> c > 0
                                | "<=" -> c <= 0
                                | _ -> c >= 0
                            results.Add(numericBool matched)
                        | _ -> results.Add numericUnsupported
                | _ ->
                    match expr with
                    | Number n -> results.Add(numericOk n)
                    | ErrorLit e -> results.Add(numericError e)
                    | NameRef name when equalsOrdinalIgnoreCase name rowName -> results.Add(numericOk row)
                    | NameRef name when equalsOrdinalIgnoreCase name colName -> results.Add(numericOk col)
                    | UnaryOp (op, inner) ->
                        frames.Add(struct (1, Number 0.0, op))
                        frames.Add(struct (0, inner, ""))
                    | BinOp (left, op, right) ->
                        frames.Add(struct (2, Number 0.0, op))
                        frames.Add(struct (0, right, ""))
                        frames.Add(struct (0, left, ""))
                    | _ -> results.Add numericUnsupported

            if results.Count <> 1 then
                ValueNone
            else
                let struct (kind, n, e) = results.[0]
                if kind = 0 then ValueSome (Ok n)
                elif kind = 1 then ValueSome (Result.Error e)
                elif kind = 3 && fnName <> "COUNT" then ValueSome (Ok n)
                else ValueNone
        let rec tryAffine rowName colName expr =
            let scale factor (struct (rowCoeff, colCoeff, constant)) =
                struct (rowCoeff * factor, colCoeff * factor, constant * factor)
            let add (struct (lr, lc, lk)) (struct (rr, rc, rk)) =
                struct (lr + rr, lc + rc, lk + rk)
            let sub (struct (lr, lc, lk)) (struct (rr, rc, rk)) =
                struct (lr - rr, lc - rc, lk - rk)
            let isConstant (struct (rowCoeff, colCoeff, constant)) =
                if rowCoeff = 0.0 && colCoeff = 0.0 then ValueSome constant else ValueNone
            let comparisonValue op c =
                let matched =
                    match op with
                    | "=" -> c = 0
                    | "<>" -> c <> 0
                    | "<" -> c < 0
                    | ">" -> c > 0
                    | "<=" -> c <= 0
                    | _ -> c >= 0
                if matched then 1.0 else 0.0
            let tryConstantComparison expr =
                match expr with
                | BinOp (left, op, right) when op = "=" || op = "<>" || op = "<" || op = ">" || op = "<=" || op = ">=" ->
                    match tryAffine rowName colName left, tryAffine rowName colName right with
                    // Equal terms are equal; terms a constant apart need not differ at 15 digits.
                    | ValueSome (Ok (struct (lr, lc, lk))), ValueSome (Ok (struct (rr, rc, rk))) when lr = rr && lc = rc && lk = rk ->
                        ValueSome (Ok (comparisonValue op 0))
                    | ValueSome (Result.Error e), _ | _, ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                    | _ -> ValueNone
                | _ -> ValueNone
            let tryAddSubChain root =
                let isAddSub expr =
                    match expr with
                    | BinOp (_, op, _) when op = "+" || op = "-" -> true
                    | _ -> false

                if not (isAddSub root) then
                    struct (false, ValueNone)
                else
                    let stack = ResizeArray<struct (float * Expr)>()
                    stack.Add(struct (1.0, root))
                    let mutable sawChain = false
                    let mutable unsupported = false
                    let mutable firstErr = ValueNone
                    let mutable total = struct (0.0, 0.0, 0.0)

                    let addScaled factor (struct (rowCoeff, colCoeff, constant)) =
                        let struct (tr, tc, tk) = total
                        total <- struct (tr + factor * rowCoeff, tc + factor * colCoeff, tk + factor * constant)

                    while not unsupported && ValueOption.isNone firstErr && stack.Count > 0 do
                        let last = stack.Count - 1
                        let struct (factor, expr) = stack.[last]
                        Portable.Stack.dropLast stack

                        match expr with
                        | BinOp (left, op, right) when op = "+" || op = "-" ->
                            sawChain <- true
                            stack.Add(struct ((if op = "-" then -factor else factor), right))
                            stack.Add(struct (factor, left))
                        | _ ->
                            match tryConstantComparison expr with
                            | ValueSome (Ok value) -> addScaled factor (struct (0.0, 0.0, value))
                            | ValueSome (Result.Error e) -> firstErr <- ValueSome e
                            | ValueNone ->
                                match tryAffine rowName colName expr with
                                | ValueSome (Ok value) -> addScaled factor value
                                | ValueSome (Result.Error e) -> firstErr <- ValueSome e
                                | ValueNone -> unsupported <- true

                    if not sawChain then
                        struct (false, ValueNone)
                    elif unsupported then
                        struct (true, ValueNone)
                    else
                        match firstErr with
                        | ValueSome e -> struct (true, ValueSome (Result.Error e))
                        | ValueNone -> struct (true, ValueSome (Ok total))

            let struct (isAddSubChain, addSubResult) = tryAddSubChain expr
            if isAddSubChain then
                addSubResult
            else
                match expr with
                | Number n -> ValueSome (Ok (struct (0.0, 0.0, n)))
                | ErrorLit e -> ValueSome (Result.Error e)
                | NameRef name when equalsOrdinalIgnoreCase name rowName -> ValueSome (Ok (struct (1.0, 0.0, 0.0)))
                | NameRef name when equalsOrdinalIgnoreCase name colName -> ValueSome (Ok (struct (0.0, 1.0, 0.0)))
                | UnaryOp (op, inner) ->
                    match tryAffine rowName colName inner with
                    | ValueSome (Ok affine) ->
                        match op with
                        | "+" -> ValueSome (Ok affine)
                        | "-" -> ValueSome (Ok (scale -1.0 affine))
                        | "%" -> ValueSome (Ok (scale 0.01 affine))
                        | _ -> ValueNone
                    | ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                    | ValueNone -> ValueNone
                | BinOp (left, op, right) ->
                    match tryAffine rowName colName left, tryAffine rowName colName right with
                    | ValueSome (Ok leftAffine), ValueSome (Ok rightAffine) ->
                        match op with
                        | "+" -> ValueSome (Ok (add leftAffine rightAffine))
                        | "-" -> ValueSome (Ok (sub leftAffine rightAffine))
                        | "*" ->
                            match isConstant leftAffine, isConstant rightAffine with
                            | ValueSome leftConstant, _ -> ValueSome (Ok (scale leftConstant rightAffine))
                            | _, ValueSome rightConstant -> ValueSome (Ok (scale rightConstant leftAffine))
                            | _ -> ValueNone
                        | "/" ->
                            match isConstant rightAffine with
                            | ValueSome 0.0 -> ValueSome (Result.Error Err.div0)
                            | ValueSome divisor -> ValueSome (Ok (scale (1.0 / divisor) leftAffine))
                            | ValueNone -> ValueNone
                        | "^" ->
                            match isConstant leftAffine, isConstant rightAffine with
                            | _, ValueSome 0.0 -> ValueSome (Ok (struct (0.0, 0.0, 1.0)))
                            | _, ValueSome 1.0 -> ValueSome (Ok leftAffine)
                            | ValueSome baseValue, ValueSome exponent -> ValueSome (Ok (struct (0.0, 0.0, Math.Pow(baseValue, exponent))))
                            | _ -> ValueNone
                        | _ -> ValueNone
                    | ValueSome (Result.Error e), _ | _, ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                    | _ -> ValueNone
                | _ -> ValueNone
        let rec tryBilinear rowName colName expr =
            let scale factor (struct (rowColCoeff, rowCoeff, colCoeff, constant)) =
                struct (rowColCoeff * factor, rowCoeff * factor, colCoeff * factor, constant * factor)
            let add (struct (lrc, lr, lc, lk)) (struct (rrc, rr, rc, rk)) =
                struct (lrc + rrc, lr + rr, lc + rc, lk + rk)
            let sub (struct (lrc, lr, lc, lk)) (struct (rrc, rr, rc, rk)) =
                struct (lrc - rrc, lr - rr, lc - rc, lk - rk)
            let ofAffine (struct (rowCoeff, colCoeff, constant)) =
                struct (0.0, rowCoeff, colCoeff, constant)
            let isConstant (struct (rowCoeff, colCoeff, constant)) =
                if rowCoeff = 0.0 && colCoeff = 0.0 then ValueSome constant else ValueNone
            let isBilinearConstant (struct (rowColCoeff, rowCoeff, colCoeff, constant)) =
                if rowColCoeff = 0.0 && rowCoeff = 0.0 && colCoeff = 0.0 then ValueSome constant else ValueNone
            let comparisonValue op c =
                let matched =
                    match op with
                    | "=" -> c = 0
                    | "<>" -> c <> 0
                    | "<" -> c < 0
                    | ">" -> c > 0
                    | "<=" -> c <= 0
                    | _ -> c >= 0
                if matched then 1.0 else 0.0
            let tryConstantComparison expr =
                match expr with
                | BinOp (left, op, right) when op = "=" || op = "<>" || op = "<" || op = ">" || op = "<=" || op = ">=" ->
                    match tryBilinear rowName colName left, tryBilinear rowName colName right with
                    | ValueSome (Ok (struct (lrc, lr, lc, lk))), ValueSome (Ok (struct (rrc, rr, rc, rk)))
                        when lrc = rrc && lr = rr && lc = rc && lk = rk ->
                        ValueSome (Ok (comparisonValue op 0))
                    | ValueSome (Result.Error e), _ | _, ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                    | _ -> ValueNone
                | _ -> ValueNone
            let tryAddSubChain root =
                let isAddSub expr =
                    match expr with
                    | BinOp (_, op, _) when op = "+" || op = "-" -> true
                    | _ -> false

                if not (isAddSub root) then
                    struct (false, ValueNone)
                else
                    let stack = ResizeArray<struct (float * Expr)>()
                    stack.Add(struct (1.0, root))
                    let mutable sawChain = false
                    let mutable unsupported = false
                    let mutable firstErr = ValueNone
                    let mutable total = struct (0.0, 0.0, 0.0, 0.0)

                    let addScaled factor (struct (rowColCoeff, rowCoeff, colCoeff, constant)) =
                        let struct (trc, tr, tc, tk) = total
                        total <- struct (trc + factor * rowColCoeff, tr + factor * rowCoeff, tc + factor * colCoeff, tk + factor * constant)

                    while not unsupported && ValueOption.isNone firstErr && stack.Count > 0 do
                        let last = stack.Count - 1
                        let struct (factor, expr) = stack.[last]
                        Portable.Stack.dropLast stack

                        match expr with
                        | BinOp (left, op, right) when op = "+" || op = "-" ->
                            sawChain <- true
                            stack.Add(struct ((if op = "-" then -factor else factor), right))
                            stack.Add(struct (factor, left))
                        | _ ->
                            match tryConstantComparison expr with
                            | ValueSome (Ok value) -> addScaled factor (struct (0.0, 0.0, 0.0, value))
                            | ValueSome (Result.Error e) -> firstErr <- ValueSome e
                            | ValueNone ->
                                match tryBilinear rowName colName expr with
                                | ValueSome (Ok value) -> addScaled factor value
                                | ValueSome (Result.Error e) -> firstErr <- ValueSome e
                                | ValueNone -> unsupported <- true

                    if not sawChain then
                        struct (false, ValueNone)
                    elif unsupported then
                        struct (true, ValueNone)
                    else
                        match firstErr with
                        | ValueSome e -> struct (true, ValueSome (Result.Error e))
                        | ValueNone -> struct (true, ValueSome (Ok total))

            let struct (isAddSubChain, addSubResult) = tryAddSubChain expr
            if isAddSubChain then
                addSubResult
            else
                match expr with
                | Number n -> ValueSome (Ok (struct (0.0, 0.0, 0.0, n)))
                | ErrorLit e -> ValueSome (Result.Error e)
                | NameRef name when equalsOrdinalIgnoreCase name rowName -> ValueSome (Ok (struct (0.0, 1.0, 0.0, 0.0)))
                | NameRef name when equalsOrdinalIgnoreCase name colName -> ValueSome (Ok (struct (0.0, 0.0, 1.0, 0.0)))
                | UnaryOp (op, inner) ->
                    match tryBilinear rowName colName inner with
                    | ValueSome (Ok value) ->
                        match op with
                        | "+" -> ValueSome (Ok value)
                        | "-" -> ValueSome (Ok (scale -1.0 value))
                        | "%" -> ValueSome (Ok (scale 0.01 value))
                        | _ -> ValueNone
                    | ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                    | ValueNone -> ValueNone
                | BinOp (left, op, right) ->
                    match op with
                    | "+" | "-" ->
                        match tryBilinear rowName colName left, tryBilinear rowName colName right with
                        | ValueSome (Ok leftValue), ValueSome (Ok rightValue) ->
                            ValueSome (Ok (if op = "+" then add leftValue rightValue else sub leftValue rightValue))
                        | ValueSome (Result.Error e), _ | _, ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                        | _ -> ValueNone
                    | "*" ->
                        match tryAffine rowName colName left, tryAffine rowName colName right with
                        | ValueSome (Ok leftAffine), ValueSome (Ok rightAffine) ->
                            let struct (lr, lc, lk) = leftAffine
                            let struct (rr, rc, rk) = rightAffine
                            if lr <> 0.0 && rr <> 0.0 then
                                ValueNone
                            elif lc <> 0.0 && rc <> 0.0 then
                                ValueNone
                            else
                                ValueSome (Ok (struct (lr * rc + lc * rr, lr * rk + lk * rr, lc * rk + lk * rc, lk * rk)))
                        | ValueSome (Result.Error e), _ | _, ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                        | _ -> ValueNone
                    | "/" ->
                        match tryBilinear rowName colName left, tryAffine rowName colName right with
                        | ValueSome (Ok leftValue), ValueSome (Ok rightAffine) ->
                            match isConstant rightAffine with
                            | ValueSome 0.0 -> ValueSome (Result.Error Err.div0)
                            | ValueSome divisor -> ValueSome (Ok (scale (1.0 / divisor) leftValue))
                            | ValueNone -> ValueNone
                        | ValueSome (Result.Error e), _ | _, ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                        | _ -> ValueNone
                    | "^" ->
                        match tryBilinear rowName colName left, tryAffine rowName colName right with
                        | ValueSome (Ok leftValue), ValueSome (Ok rightAffine) ->
                            match isBilinearConstant leftValue, isConstant rightAffine with
                            | _, ValueSome 0.0 -> ValueSome (Ok (struct (0.0, 0.0, 0.0, 1.0)))
                            | _, ValueSome 1.0 -> ValueSome (Ok leftValue)
                            | ValueSome baseValue, ValueSome exponent -> ValueSome (Ok (struct (0.0, 0.0, 0.0, Math.Pow(baseValue, exponent))))
                            | _ -> ValueNone
                        | ValueSome (Result.Error e), _ | _, ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                        | _ -> ValueNone
                    | _ -> ValueNone
                | _ -> ValueNone
        // A symbolic bilinear form cannot encode the zero produced at an
        // intermediate multiplication/division/power. Decline it for tiny or
        // extreme coefficients; the numeric MAKEARRAY stream still runs
        // without materializing the array.
        let sensitiveNumber n = n <> 0.0 && (abs n < 1e-150 || abs n > 1e150)
        let sensitiveBilinear (struct (rc, r, c, k)) =
            sensitiveNumber rc || sensitiveNumber r || sensitiveNumber c || sensitiveNumber k
        let rec unsafeSymbolic rowName colName expr =
            let resultSensitive =
                match expr with
                | BinOp (_, ("*" | "/"), _) | UnaryOp ("%", _) ->
                    match tryBilinear rowName colName expr with
                    | ValueSome (Ok value) -> sensitiveBilinear value
                    | _ -> false
                | _ -> false
            resultSensitive ||
                (match expr with
                 | Number n -> sensitiveNumber n
                 | UnaryOp (_, inner) -> unsafeSymbolic rowName colName inner
                 | BinOp (_, "^", _) -> true
                 | BinOp (left, _, right) -> unsafeSymbolic rowName colName left || unsafeSymbolic rowName colName right
                 | _ -> false)
        let makeArrayStreamOfArgs makeArgs =
            match makeArgs with
            | [ rowsExpr; colsExpr; lambdaExpr ] ->
                match coerceNumber (eval rowsExpr ctx), coerceNumber (eval colsExpr ctx), lambdaPartsOf lambdaExpr with
                | Result.Error e, _, _ | _, Result.Error e, _ -> ValueSome (Result.Error e)
                | Ok rowsN, Ok colsN, Some [ NameRef rowName; NameRef colName; body ] when not (containsVolatile body) ->
                    let rows = int rowsN
                    let cols = int colsN
                    if rows < 1 || cols < 1 then
                        ValueSome (Result.Error Err.value)
                    else
                        let count = int64 rows * int64 cols
                        let mutable offset = 0L
                        let nextNumber () =
                            let r = float (offset / int64 cols + 1L)
                            let c = float (offset % int64 cols + 1L)
                            offset <- offset + 1L
                            match evalNumeric r c rowName colName body with
                            | ValueSome result -> result
                            | ValueNone -> Result.Error Err.value
                        ValueSome (Ok (count, nextNumber, true))
                | _ -> ValueNone
            | _ -> ValueNone
        let rec makeArrayStreamOfExpr expr =
            let tryUnaryChain root =
                let isUnary expr =
                    match expr with
                    | UnaryOp (op, _) when op = "+" || op = "-" || op = "%" -> true
                    | _ -> false

                if not (isUnary root) then
                    struct (false, ValueNone)
                else
                    let mutable current = root
                    let mutable factor = 1.0
                    let mutable peeling = true
                    while peeling do
                        match current with
                        | UnaryOp (op, inner) when op = "+" || op = "-" || op = "%" ->
                            match op with
                            | "-" -> factor <- -factor
                            | "%" -> factor <- factor * 0.01
                            | _ -> ()
                            current <- inner
                        | _ -> peeling <- false

                    match makeArrayStreamOfExpr current with
                    | ValueSome (Ok (count, nextNumber, _)) ->
                        let mapped () =
                            match nextNumber () with
                            | Ok n -> Ok (if factor = 1.0 || factor = -1.0 then factor * n else flushProductUnderflow (factor * n))
                            | Result.Error e -> Result.Error e
                        struct (true, ValueSome (Ok (count, mapped, true)))
                    | ValueSome (Result.Error e) -> struct (true, ValueSome (Result.Error e))
                    | ValueNone -> struct (true, ValueNone)

            let tryAddSubChain root =
                let isAddSub expr =
                    match expr with
                    | BinOp (_, op, _) when op = "+" || op = "-" -> true
                    | _ -> false

                if not (isAddSub root) then
                    struct (false, ValueNone)
                else
                    let stack = ResizeArray<struct (float * Expr)>()
                    stack.Add(struct (1.0, root))
                    let mutable unsupported = false
                    let mutable firstErr = ValueNone
                    let mutable foundStream = false
                    let mutable streamCount = 0L
                    let mutable streamFactor = 1.0
                    let mutable streamNext = Unchecked.defaultof<unit -> Result<float, string>>
                    let mutable constant = 0.0

                    while not unsupported && ValueOption.isNone firstErr && stack.Count > 0 do
                        let last = stack.Count - 1
                        let struct (factor, expr) = stack.[last]
                        Portable.Stack.dropLast stack

                        match expr with
                        | BinOp (left, op, right) when op = "+" || op = "-" ->
                            stack.Add(struct ((if op = "-" then -factor else factor), right))
                            stack.Add(struct (factor, left))
                        | _ ->
                            match makeArrayStreamOfExpr expr with
                            | ValueSome (Ok (count, nextNumber, _)) ->
                                if foundStream || factor <> 1.0 && factor <> -1.0 then
                                    unsupported <- true
                                else
                                    foundStream <- true
                                    streamCount <- count
                                    streamFactor <- factor
                                    streamNext <- nextNumber
                            | ValueSome (Result.Error e) -> firstErr <- ValueSome e
                            | ValueNone ->
                                if isArrayExpr expr then
                                    unsupported <- true
                                else
                                    match coerceNumber (eval expr ctx) with
                                    | Ok scalar -> constant <- constant + factor * scalar
                                    | Result.Error e -> firstErr <- ValueSome e

                    if unsupported then
                        struct (true, ValueNone)
                    else
                        match firstErr with
                        | ValueSome e -> struct (true, ValueSome (Result.Error e))
                        | ValueNone ->
                            if foundStream then
                                let mapped () =
                                    match streamNext () with
                                    | Ok n -> Ok (constant + streamFactor * n)
                                    | Result.Error e -> Result.Error e
                                struct (true, ValueSome (Ok (streamCount, mapped, true)))
                            else
                                struct (true, ValueNone)

            let tryMulDivChain root =
                let isMulDiv expr =
                    match expr with
                    | BinOp (_, op, _) when op = "*" || op = "/" -> true
                    | _ -> false

                if not (isMulDiv root) || (match root with BinOp (left, _, right) -> isMulDiv left || isMulDiv right | _ -> false) then
                    struct (false, ValueNone)
                else
                    let stack = ResizeArray<struct (bool * Expr)>()
                    stack.Add(struct (true, root))
                    let mutable unsupported = false
                    let mutable firstErr = ValueNone
                    let mutable foundStream = false
                    let mutable streamCount = 0L
                    let mutable streamNext = Unchecked.defaultof<unit -> Result<float, string>>
                    let mutable scalar = 1.0

                    while not unsupported && ValueOption.isNone firstErr && stack.Count > 0 do
                        let last = stack.Count - 1
                        let struct (numerator, expr) = stack.[last]
                        Portable.Stack.dropLast stack

                        match expr with
                        | BinOp (left, op, right) when op = "*" || op = "/" ->
                            stack.Add(struct ((op = "*" && numerator) || (op = "/" && not numerator), right))
                            stack.Add(struct (numerator, left))
                        | _ ->
                            match makeArrayStreamOfExpr expr with
                            | ValueSome (Ok (count, nextNumber, _)) ->
                                if foundStream || not numerator then
                                    unsupported <- true
                                else
                                    foundStream <- true
                                    streamCount <- count
                                    streamNext <- nextNumber
                            | ValueSome (Result.Error e) -> firstErr <- ValueSome e
                            | ValueNone ->
                                if isArrayExpr expr then
                                    unsupported <- true
                                else
                                    match coerceNumber (eval expr ctx) with
                                    | Ok value ->
                                        if numerator then
                                            scalar <- scalar * value
                                        elif value > -minNormal && value < minNormal then
                                            firstErr <- ValueSome Err.div0
                                        else
                                            scalar <- scalar / value
                                    | Result.Error e -> firstErr <- ValueSome e

                    if unsupported then
                        struct (true, ValueNone)
                    else
                        match firstErr with
                        | ValueSome e -> struct (true, ValueSome (Result.Error e))
                        | ValueNone ->
                            if foundStream then
                                let mapped () =
                                    match streamNext () with
                                    | Ok n -> Ok (flushProductUnderflow (scalar * n))
                                    | Result.Error e -> Result.Error e
                                struct (true, ValueSome (Ok (streamCount, mapped, true)))
                            else
                                struct (true, ValueNone)

            let struct (isUnaryChain, unaryResult) = tryUnaryChain expr
            if isUnaryChain then
                unaryResult
            else
                let struct (isAddSubChain, addSubResult) = tryAddSubChain expr
                if isAddSubChain then
                    addSubResult
                else
                    let struct (isMulDivChain, mulDivResult) = tryMulDivChain expr
                    if isMulDivChain then
                        mulDivResult
                    else
                        match expr with
                        | Function (makeName, makeArgs) when equalsOrdinalIgnoreCase makeName "MAKEARRAY" ->
                            makeArrayStreamOfArgs makeArgs
                        | Function (transposeName, [ sourceExpr ]) when equalsOrdinalIgnoreCase transposeName "TRANSPOSE" ->
                            makeArrayStreamOfExpr sourceExpr
                        | Function (sortName, [ sourceExpr ]) when equalsOrdinalIgnoreCase sortName "SORT" ->
                            makeArrayStreamOfExpr sourceExpr
                        | UnaryOp (op, inner) when op = "+" || op = "-" || op = "%" ->
                            match makeArrayStreamOfExpr inner with
                            | ValueSome (Ok (count, nextNumber, _)) ->
                                let mapped () =
                                    match nextNumber () with
                                    | Ok n ->
                                        match op with
                                        | "-" -> Ok -n
                                        | "%" -> Ok (flushProductUnderflow (n / 100.0))
                                        | _ -> Ok n
                                    | Result.Error e -> Result.Error e
                                ValueSome (Ok (count, mapped, true))
                            | other -> other
                        | BinOp (left, op, right) when op = "+" || op = "-" || op = "*" || op = "/" ->
                            match makeArrayStreamOfExpr left with
                            | ValueSome (Ok (count, nextNumber, _)) when not (isArrayExpr right) ->
                                match coerceNumber (eval right ctx) with
                                | Result.Error e -> ValueSome (Result.Error e)
                                | Ok scalar ->
                                    let mapped () =
                                        match nextNumber () with
                                        | Result.Error e -> Result.Error e
                                        | Ok n ->
                                            match op with
                                            | "+" -> Ok (n + scalar)
                                            | "-" -> Ok (n - scalar)
                                            | "*" -> Ok (flushProductUnderflow (n * scalar))
                                            | "/" when scalar > -minNormal && scalar < minNormal -> Result.Error Err.div0
                                            | "/" -> Ok (flushProductUnderflow (n / scalar))
                                            | _ -> Result.Error Err.value
                                    ValueSome (Ok (count, mapped, true))
                            | ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                            | _ ->
                                match makeArrayStreamOfExpr right with
                                | ValueSome (Ok (count, nextNumber, _)) when not (isArrayExpr left) ->
                                    match coerceNumber (eval left ctx) with
                                    | Result.Error e -> ValueSome (Result.Error e)
                                    | Ok scalar ->
                                        let mapped () =
                                            match nextNumber () with
                                            | Result.Error e -> Result.Error e
                                            | Ok n ->
                                                match op with
                                                | "+" -> Ok (scalar + n)
                                                | "-" -> Ok (scalar - n)
                                                | "*" -> Ok (flushProductUnderflow (scalar * n))
                                                | _ -> Result.Error Err.value
                                        ValueSome (Ok (count, mapped, true))
                                | other -> other
                        | BinOp (left, op, right) when op = "=" || op = "<>" || op = "<" || op = ">" || op = "<=" || op = ">=" ->
                            let comparisonValue cmp =
                                let matches =
                                    match op with
                                    | "=" -> cmp = 0
                                    | "<>" -> cmp <> 0
                                    | "<" -> cmp < 0
                                    | ">" -> cmp > 0
                                    | "<=" -> cmp <= 0
                                    | _ -> cmp >= 0
                                if matches then 1.0 else 0.0
                            match makeArrayStreamOfExpr left with
                            | ValueSome (Ok (count, nextNumber, _)) when not (isArrayExpr right) ->
                                match coerceNumber (eval right ctx) with
                                | Result.Error e -> ValueSome (Result.Error e)
                                | Ok scalar ->
                                    let mapped () =
                                        match nextNumber () with
                                        | Ok n -> Ok (comparisonValue (compareNumbers n scalar))
                                        | Result.Error e -> Result.Error e
                                    ValueSome (Ok (count, mapped, false))
                            | ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                            | _ ->
                                match makeArrayStreamOfExpr right with
                                | ValueSome (Ok (count, nextNumber, _)) when not (isArrayExpr left) ->
                                    match coerceNumber (eval left ctx) with
                                    | Result.Error e -> ValueSome (Result.Error e)
                                    | Ok scalar ->
                                        let mapped () =
                                            match nextNumber () with
                                            | Ok n -> Ok (comparisonValue (compareNumbers scalar n))
                                            | Result.Error e -> Result.Error e
                                        ValueSome (Ok (count, mapped, false))
                                | other -> other
                        | _ -> ValueNone
        let aggregateStream count nextNumber countAsNumber =
            if count < 1024L then
                ValueNone
            elif fnName = "COUNT" && not countAsNumber then
                ValueSome (CellValue.Number 0.0)
            else
                let mutable firstErr = ValueNone
                let mutable total = 0.0
                let mutable productValue = 1.0
                let mutable numericCount = 0L
                let mutable minValue = infinity
                let mutable maxValue = -infinity
                let mutable remaining = count
                while ValueOption.isNone firstErr && remaining > 0L do
                    match nextNumber () with
                    | Result.Error e -> firstErr <- ValueSome e
                    | Ok n ->
                        numericCount <- numericCount + 1L
                        match fnName with
                        | "SUM" | "AVERAGE" ->
                            total <- total + n
                        | "PRODUCT" ->
                            productValue <- flushProductUnderflow (productValue * n)
                        | "MIN" ->
                            if n < minValue then minValue <- n
                        | "MAX" ->
                            if n > maxValue then maxValue <- n
                        | _ -> ()
                    remaining <- remaining - 1L
                match firstErr with
                | ValueSome e -> ValueSome (err e)
                | ValueNone ->
                    let result =
                        match fnName with
                        | "PRODUCT" -> CellValue.Number (if numericCount = 0L then 0.0 else productValue)
                        | "AVERAGE" -> if numericCount = 0L then err Err.div0 else CellValue.Number (total / float numericCount)
                        | "COUNT" -> CellValue.Number (float numericCount)
                        | "MIN" -> CellValue.Number (if numericCount = 0L then 0.0 else minValue)
                        | "MAX" -> CellValue.Number (if numericCount = 0L then 0.0 else maxValue)
                        | _ -> CellValue.Number total
                    ValueSome result
        let aggregateNumberStream count nextNumber countAsNumber =
            if count < 1024L then
                ValueNone
            elif fnName = "COUNT" && not countAsNumber then
                ValueSome (CellValue.Number 0.0)
            elif fnName = "COUNT" then
                ValueSome (CellValue.Number (float count))
            else
                let mutable total = 0.0
                let mutable productValue = 1.0
                let mutable minValue = infinity
                let mutable maxValue = -infinity
                let mutable remaining = count
                while remaining > 0L do
                    let n = nextNumber ()
                    match fnName with
                    | "SUM" | "AVERAGE" ->
                        total <- total + n
                    | "PRODUCT" ->
                        productValue <- flushProductUnderflow (productValue * n)
                    | "MIN" ->
                        if n < minValue then minValue <- n
                    | "MAX" ->
                        if n > maxValue then maxValue <- n
                    | _ -> ()
                    remaining <- remaining - 1L
                let result =
                    match fnName with
                    | "PRODUCT" -> CellValue.Number productValue
                    | "AVERAGE" -> CellValue.Number (total / float count)
                    | "MIN" -> CellValue.Number minValue
                    | "MAX" -> CellValue.Number maxValue
                    | _ -> CellValue.Number total
                ValueSome result
        let aggregateBilinear rows cols (struct (rowColCoeff, rowCoeff, colCoeff, constant)) =
            let totalCells = int64 rows * int64 cols
            if totalCells < 1024L then
                ValueNone
            elif fnName = "PRODUCT" && (rowColCoeff <> 0.0 || rowCoeff <> 0.0 || colCoeff <> 0.0) then
                ValueNone
            else
                let total = float totalCells
                let rowTotal = float rows * float (rows + 1) / 2.0
                let colTotal = float cols * float (cols + 1) / 2.0
                let sumRows = float cols * float rows * float (rows + 1) / 2.0
                let sumCols = float rows * float cols * float (cols + 1) / 2.0
                let sumRowCols = rowTotal * colTotal
                let sum = rowColCoeff * sumRowCols + rowCoeff * sumRows + colCoeff * sumCols + constant * total
                let valueAt r c = rowColCoeff * float r * float c + rowCoeff * float r + colCoeff * float c + constant
                let result =
                    match fnName with
                    | "PRODUCT" -> CellValue.Number (flushProductUnderflow (Math.Pow(constant, total)))
                    | "SUM" -> CellValue.Number sum
                    | "AVERAGE" -> CellValue.Number (sum / total)
                    | "COUNT" -> CellValue.Number total
                    | "MIN" ->
                        let v1 = valueAt 1 1
                        let v2 = valueAt rows 1
                        let v3 = valueAt 1 cols
                        let v4 = valueAt rows cols
                        CellValue.Number (min (min v1 v2) (min v3 v4))
                    | "MAX" ->
                        let v1 = valueAt 1 1
                        let v2 = valueAt rows 1
                        let v3 = valueAt 1 cols
                        let v4 = valueAt rows cols
                        CellValue.Number (max (max v1 v2) (max v3 v4))
                    | _ -> CellValue.Number sum
                ValueSome result
        let aggregateConstantComparison count value =
            if count < 1024L then
                ValueNone
            else
                let result =
                    match fnName with
                    | "PRODUCT" -> CellValue.Number (if value = 0.0 then 0.0 else 1.0)
                    | "AVERAGE" -> CellValue.Number value
                    | "COUNT" -> CellValue.Number 0.0
                    | "MIN" -> CellValue.Number value
                    | "MAX" -> CellValue.Number value
                    | _ -> CellValue.Number (value * float count)
                ValueSome result
        let scaleBilinear factor (struct (rowColCoeff, rowCoeff, colCoeff, constant)) =
            struct (rowColCoeff * factor, rowCoeff * factor, colCoeff * factor, constant * factor)
        let addBilinear (struct (lrc, lr, lc, lk)) (struct (rrc, rr, rc, rk)) =
            struct (lrc + rrc, lr + rr, lc + rc, lk + rk)
        let subBilinear (struct (lrc, lr, lc, lk)) (struct (rrc, rr, rc, rk)) =
            struct (lrc - rrc, lr - rr, lc - rc, lk - rk)
        let addBilinearConstant constant value =
            addBilinear value (struct (0.0, 0.0, 0.0, constant))
        let rec tryBilinearMakeArrayExpr expr =
            let tryUnaryChain root =
                let isUnary expr =
                    match expr with
                    | UnaryOp (op, _) when op = "+" || op = "-" || op = "%" -> true
                    | _ -> false

                if not (isUnary root) then
                    struct (false, ValueNone)
                else
                    let mutable current = root
                    let mutable factor = 1.0
                    let mutable peeling = true
                    while peeling do
                        match current with
                        | UnaryOp (op, inner) when op = "+" || op = "-" || op = "%" ->
                            match op with
                            | "-" -> factor <- -factor
                            | "%" -> factor <- factor * 0.01
                            | _ -> ()
                            current <- inner
                        | _ -> peeling <- false

                    match tryBilinearMakeArrayExpr current with
                    | ValueSome (Ok (rows, cols, value)) ->
                        if factor <> 1.0 && factor <> -1.0
                           && (sensitiveBilinear value || sensitiveBilinear (scaleBilinear factor value)) then
                            struct (true, ValueNone)
                        else struct (true, ValueSome (Ok (rows, cols, scaleBilinear factor value)))
                    | ValueSome (Result.Error e) -> struct (true, ValueSome (Result.Error e))
                    | ValueNone -> struct (true, ValueNone)

            let tryAddSubChain root =
                let isAddSub expr =
                    match expr with
                    | BinOp (_, op, _) when op = "+" || op = "-" -> true
                    | _ -> false

                if not (isAddSub root) then
                    struct (false, ValueNone)
                else
                    let stack = ResizeArray<struct (float * Expr)>()
                    stack.Add(struct (1.0, root))
                    let mutable unsupported = false
                    let mutable firstErr = ValueNone
                    let mutable foundArray = false
                    let mutable rows = 0
                    let mutable cols = 0
                    let mutable totalValue = struct (0.0, 0.0, 0.0, 0.0)
                    let mutable constant = 0.0

                    let addArray factor leafRows leafCols leafValue =
                        let scaled = scaleBilinear factor leafValue
                        if not foundArray then
                            foundArray <- true
                            rows <- leafRows
                            cols <- leafCols
                            totalValue <- scaled
                        elif rows = leafRows && cols = leafCols then
                            totalValue <- addBilinear totalValue scaled
                        else
                            unsupported <- true

                    while not unsupported && ValueOption.isNone firstErr && stack.Count > 0 do
                        let last = stack.Count - 1
                        let struct (factor, expr) = stack.[last]
                        Portable.Stack.dropLast stack

                        match expr with
                        | BinOp (left, op, right) when op = "+" || op = "-" ->
                            stack.Add(struct ((if op = "-" then -factor else factor), right))
                            stack.Add(struct (factor, left))
                        | _ ->
                            match tryBilinearMakeArrayExpr expr with
                            | ValueSome (Ok (leafRows, leafCols, leafValue)) -> addArray factor leafRows leafCols leafValue
                            | ValueSome (Result.Error e) -> firstErr <- ValueSome e
                            | ValueNone ->
                                if isArrayExpr expr then
                                    unsupported <- true
                                else
                                    match coerceNumber (eval expr ctx) with
                                    | Ok scalar -> constant <- constant + factor * scalar
                                    | Result.Error e -> firstErr <- ValueSome e

                    if unsupported then
                        struct (true, ValueNone)
                    else
                        match firstErr with
                        | ValueSome e -> struct (true, ValueSome (Result.Error e))
                        | ValueNone ->
                            if foundArray then
                                struct (true, ValueSome (Ok (rows, cols, addBilinearConstant constant totalValue)))
                            else
                                struct (true, ValueNone)

            let tryMulDivChain root =
                let isMulDiv expr =
                    match expr with
                    | BinOp (_, op, _) when op = "*" || op = "/" -> true
                    | _ -> false

                if not (isMulDiv root) || (match root with BinOp (left, _, right) -> isMulDiv left || isMulDiv right | _ -> false) then
                    struct (false, ValueNone)
                else
                    let stack = ResizeArray<struct (bool * Expr)>()
                    stack.Add(struct (true, root))
                    let mutable unsupported = false
                    let mutable firstErr = ValueNone
                    let mutable foundArray = false
                    let mutable rows = 0
                    let mutable cols = 0
                    let mutable arrayValue = struct (0.0, 0.0, 0.0, 0.0)
                    let mutable scalar = 1.0

                    while not unsupported && ValueOption.isNone firstErr && stack.Count > 0 do
                        let last = stack.Count - 1
                        let struct (numerator, expr) = stack.[last]
                        Portable.Stack.dropLast stack

                        match expr with
                        | BinOp (left, op, right) when op = "*" || op = "/" ->
                            stack.Add(struct ((op = "*" && numerator) || (op = "/" && not numerator), right))
                            stack.Add(struct (numerator, left))
                        | _ ->
                            match tryBilinearMakeArrayExpr expr with
                            | ValueSome (Ok (leafRows, leafCols, leafValue)) ->
                                if foundArray || not numerator then
                                    unsupported <- true
                                else
                                    foundArray <- true
                                    rows <- leafRows
                                    cols <- leafCols
                                    arrayValue <- leafValue
                            | ValueSome (Result.Error e) -> firstErr <- ValueSome e
                            | ValueNone ->
                                if isArrayExpr expr then
                                    unsupported <- true
                                else
                                    match coerceNumber (eval expr ctx) with
                                    | Ok value ->
                                        if numerator then
                                            scalar <- scalar * value
                                        elif value > -minNormal && value < minNormal then
                                            firstErr <- ValueSome Err.div0
                                        else
                                            scalar <- scalar / value
                                    | Result.Error e -> firstErr <- ValueSome e

                    if unsupported then
                        struct (true, ValueNone)
                    else
                        match firstErr with
                        | ValueSome e -> struct (true, ValueSome (Result.Error e))
                        | ValueNone ->
                            if foundArray then
                                if sensitiveNumber scalar || sensitiveBilinear arrayValue || sensitiveBilinear (scaleBilinear scalar arrayValue) then
                                    struct (false, ValueNone)
                                else struct (true, ValueSome (Ok (rows, cols, scaleBilinear scalar arrayValue)))
                            else
                                struct (true, ValueNone)

            let struct (isUnaryChain, unaryResult) = tryUnaryChain expr
            if isUnaryChain then
                unaryResult
            else
                let struct (isAddSubChain, addSubResult) = tryAddSubChain expr
                if isAddSubChain then
                    addSubResult
                else
                    let struct (isMulDivChain, mulDivResult) = tryMulDivChain expr
                    if isMulDivChain then
                        mulDivResult
                    else
                        match expr with
                        | Function (makeName, makeArgs) when equalsOrdinalIgnoreCase makeName "MAKEARRAY" ->
                            match makeArgs with
                            | [ rowsExpr; colsExpr; lambdaExpr ] ->
                                match coerceNumber (eval rowsExpr ctx), coerceNumber (eval colsExpr ctx), lambdaPartsOf lambdaExpr with
                                | Result.Error e, _, _ | _, Result.Error e, _ -> ValueSome (Result.Error e)
                                | Ok rowsN, Ok colsN, Some [ NameRef rowName; NameRef colName; body ] when not (containsVolatile body) ->
                                    let rows = int rowsN
                                    let cols = int colsN
                                    if rows < 1 || cols < 1 then
                                        ValueSome (Result.Error Err.value)
                                    else
                                        match (if unsafeSymbolic rowName colName body then ValueNone else tryBilinear rowName colName body) with
                                        | ValueSome (Ok value) -> ValueSome (Ok (rows, cols, value))
                                        | ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                                        | ValueNone -> ValueNone
                                | _ -> ValueNone
                            | _ -> ValueNone
                        | Function (transposeName, [ sourceExpr ]) when equalsOrdinalIgnoreCase transposeName "TRANSPOSE" ->
                            tryBilinearMakeArrayExpr sourceExpr
                        | Function (sortName, [ sourceExpr ]) when equalsOrdinalIgnoreCase sortName "SORT" ->
                            tryBilinearMakeArrayExpr sourceExpr
                        | UnaryOp (op, inner) when op = "+" || op = "-" || op = "%" ->
                            match tryBilinearMakeArrayExpr inner with
                            | ValueSome (Ok (rows, cols, value)) ->
                                if op = "%" && (sensitiveBilinear value || sensitiveBilinear (scaleBilinear 0.01 value)) then ValueNone
                                else
                                    let mapped =
                                        match op with
                                        | "-" -> scaleBilinear -1.0 value
                                        | "%" -> scaleBilinear 0.01 value
                                        | _ -> value
                                    ValueSome (Ok (rows, cols, mapped))
                            | other -> other
                        | BinOp (left, op, right) when op = "+" || op = "-" || op = "*" || op = "/" ->
                            match tryBilinearMakeArrayExpr left with
                            | ValueSome (Ok (rows, cols, value)) when not (isArrayExpr right) ->
                                match coerceNumber (eval right ctx) with
                                | Result.Error e -> ValueSome (Result.Error e)
                                | Ok scalar ->
                                    match op with
                                    | "+" -> ValueSome (Ok (rows, cols, addBilinearConstant scalar value))
                                    | "-" -> ValueSome (Ok (rows, cols, addBilinearConstant -scalar value))
                                    | "*" when sensitiveNumber scalar || sensitiveBilinear value || sensitiveBilinear (scaleBilinear scalar value) -> ValueNone
                                    | "*" -> ValueSome (Ok (rows, cols, scaleBilinear scalar value))
                                    | "/" when scalar > -minNormal && scalar < minNormal -> ValueNone
                                    | "/" when sensitiveBilinear value || sensitiveBilinear (scaleBilinear (1.0 / scalar) value) -> ValueNone
                                    | "/" -> ValueSome (Ok (rows, cols, scaleBilinear (1.0 / scalar) value))
                                    | _ -> ValueNone
                            | ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                            | _ ->
                                match tryBilinearMakeArrayExpr right with
                                | ValueSome (Ok (rows, cols, value)) when not (isArrayExpr left) ->
                                    match coerceNumber (eval left ctx) with
                                    | Result.Error e -> ValueSome (Result.Error e)
                                    | Ok scalar ->
                                        match op with
                                        | "+" -> ValueSome (Ok (rows, cols, addBilinearConstant scalar value))
                                        | "-" -> ValueSome (Ok (rows, cols, subBilinear (struct (0.0, 0.0, 0.0, scalar)) value))
                                        | "*" when sensitiveNumber scalar || sensitiveBilinear value || sensitiveBilinear (scaleBilinear scalar value) -> ValueNone
                                        | "*" -> ValueSome (Ok (rows, cols, scaleBilinear scalar value))
                                        | _ -> ValueNone
                                | other -> other
                        | _ -> ValueNone
        let bilinearMinMax rows cols (struct (rowColCoeff, rowCoeff, colCoeff, constant)) =
            let valueAt r c = rowColCoeff * float r * float c + rowCoeff * float r + colCoeff * float c + constant
            let v1 = valueAt 1 1
            let v2 = valueAt rows 1
            let v3 = valueAt 1 cols
            let v4 = valueAt rows cols
            struct (min (min v1 v2) (min v3 v4), max (max v1 v2) (max v3 v4))
        let comparisonConstantValue op minValue maxValue scalar =
            // As the operators compare (15 digits; rounding keeps the order, so what holds of
            // the least and the greatest holds of every value between).
            let lo = compareNumbers minValue scalar
            let hi = compareNumbers maxValue scalar
            let allTrue, allFalse =
                match op with
                | "=" -> lo = 0 && hi = 0, lo > 0 || hi < 0
                | "<>" -> lo > 0 || hi < 0, lo = 0 && hi = 0
                | "<" -> hi < 0, lo >= 0
                | ">" -> lo > 0, hi <= 0
                | "<=" -> hi <= 0, lo > 0
                | _ -> lo >= 0, hi < 0
            if allTrue then ValueSome 1.0
            elif allFalse then ValueSome 0.0
            else ValueNone
        let reversedComparisonOp op =
            match op with
            | "<" -> ">"
            | ">" -> "<"
            | "<=" -> ">="
            | ">=" -> "<="
            | _ -> op
        let comparisonMakeArrayConstantAggregate expr =
            match expr with
            | BinOp (left, op, right) when op = "=" || op = "<>" || op = "<" || op = ">" || op = "<=" || op = ">=" ->
                match tryBilinearMakeArrayExpr left with
                | ValueSome (Ok (rows, cols, value)) when not (isArrayExpr right) ->
                    match coerceNumber (eval right ctx) with
                    | Result.Error e -> ValueSome (Result.Error e)
                    | Ok scalar ->
                        let struct (minValue, maxValue) = bilinearMinMax rows cols value
                        match comparisonConstantValue op minValue maxValue scalar with
                        | ValueSome comparisonValue -> ValueSome (Ok (aggregateConstantComparison (int64 rows * int64 cols) comparisonValue))
                        | ValueNone -> ValueNone
                | ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                | _ ->
                    match tryBilinearMakeArrayExpr right with
                    | ValueSome (Ok (rows, cols, value)) when not (isArrayExpr left) ->
                        match coerceNumber (eval left ctx) with
                        | Result.Error e -> ValueSome (Result.Error e)
                        | Ok scalar ->
                            let struct (minValue, maxValue) = bilinearMinMax rows cols value
                            match comparisonConstantValue (reversedComparisonOp op) minValue maxValue scalar with
                            | ValueSome comparisonValue -> ValueSome (Ok (aggregateConstantComparison (int64 rows * int64 cols) comparisonValue))
                            | ValueNone -> ValueNone
                    | ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                    | _ -> ValueNone
            | _ -> ValueNone
        let comparisonMakeArrayStreamOfExpr expr =
            match expr with
            | BinOp (left, op, right) when op = "=" || op = "<>" || op = "<" || op = ">" || op = "<=" || op = ">=" ->
                let comparisonValue cmp =
                    let matches =
                        match op with
                        | "=" -> cmp = 0
                        | "<>" -> cmp <> 0
                        | "<" -> cmp < 0
                        | ">" -> cmp > 0
                        | "<=" -> cmp <= 0
                        | _ -> cmp >= 0
                    if matches then 1.0 else 0.0
                let makeStream rows cols (struct (rowColCoeff, rowCoeff, colCoeff, constant)) scalar compareValues =
                    let mutable offset = 0L
                    let cols64 = int64 cols
                    let nextNumber () =
                        let r = float (offset / cols64 + 1L)
                        let c = float (offset % cols64 + 1L)
                        offset <- offset + 1L
                        let value = rowColCoeff * r * c + rowCoeff * r + colCoeff * c + constant
                        comparisonValue (compareValues value scalar)
                    Ok (int64 rows * int64 cols, nextNumber, false)
                match tryBilinearMakeArrayExpr left with
                | ValueSome (Ok (rows, cols, value)) when not (isArrayExpr right) ->
                    match coerceNumber (eval right ctx) with
                    | Result.Error e -> ValueSome (Result.Error e)
                    | Ok scalar -> ValueSome (makeStream rows cols value scalar compareNumbers)
                | ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                | _ ->
                    match tryBilinearMakeArrayExpr right with
                    | ValueSome (Ok (rows, cols, value)) when not (isArrayExpr left) ->
                        match coerceNumber (eval left ctx) with
                        | Result.Error e -> ValueSome (Result.Error e)
                        | Ok scalar -> ValueSome (makeStream rows cols value scalar (fun value scalar -> compareNumbers scalar value))
                    | ValueSome (Ok _) -> ValueNone
                    | ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                    | ValueNone -> ValueNone
            | _ -> ValueNone
        match args with
        | [ Function (makeName, [ rowsExpr; colsExpr; lambdaExpr ]) ] when isSupportedAggregate && equalsOrdinalIgnoreCase makeName "MAKEARRAY" ->
            match coerceNumber (eval rowsExpr ctx), coerceNumber (eval colsExpr ctx), lambdaPartsOf lambdaExpr with
            | Result.Error e, _, _ | _, Result.Error e, _ -> ValueSome (err e)
            | Ok rowsN, Ok colsN, Some [ NameRef rowName; NameRef colName; body ] when not (containsVolatile body) ->
                let rows = int rowsN
                let cols = int colsN
                if rows < 1 || cols < 1 then
                    ValueSome (err Err.value)
                else
                    let totalCells = int64 rows * int64 cols
                    if totalCells < 1024L then
                        ValueNone
                    else
                        let affineResult =
                            match (if unsafeSymbolic rowName colName body then ValueNone else tryBilinear rowName colName body) with
                            | ValueSome (Result.Error e) -> ValueSome (err e)
                            | ValueSome (Ok (struct (rowColCoeff, rowCoeff, colCoeff, constant))) ->
                                if fnName = "PRODUCT" then
                                    if rowColCoeff = 0.0 && rowCoeff = 0.0 && colCoeff = 0.0 then
                                        ValueSome (CellValue.Number (flushProductUnderflow (Math.Pow(constant, float totalCells))))
                                    else
                                        ValueNone
                                else
                                    let total = float totalCells
                                    let rowTotal = float rows * float (rows + 1) / 2.0
                                    let colTotal = float cols * float (cols + 1) / 2.0
                                    let sumRows = float cols * float rows * float (rows + 1) / 2.0
                                    let sumCols = float rows * float cols * float (cols + 1) / 2.0
                                    let sumRowCols = rowTotal * colTotal
                                    let sum = rowColCoeff * sumRowCols + rowCoeff * sumRows + colCoeff * sumCols + constant * total
                                    let valueAt r c = rowColCoeff * float r * float c + rowCoeff * float r + colCoeff * float c + constant
                                    let result =
                                        match fnName with
                                        | "SUM" -> CellValue.Number sum
                                        | "AVERAGE" -> CellValue.Number (sum / total)
                                        | "COUNT" -> CellValue.Number total
                                        | "MIN" ->
                                            let v1 = valueAt 1 1
                                            let v2 = valueAt rows 1
                                            let v3 = valueAt 1 cols
                                            let v4 = valueAt rows cols
                                            CellValue.Number (min (min v1 v2) (min v3 v4))
                                        | "MAX" ->
                                            let v1 = valueAt 1 1
                                            let v2 = valueAt rows 1
                                            let v3 = valueAt 1 cols
                                            let v4 = valueAt rows cols
                                            CellValue.Number (max (max v1 v2) (max v3 v4))
                                        | _ -> CellValue.Number sum
                                    ValueSome result
                            | ValueNone -> ValueNone
                        match affineResult with
                        | ValueSome result -> ValueSome result
                        | ValueNone ->
                            let mutable unsupported = false
                            let mutable firstErr = ValueNone
                            let mutable total = 0.0
                            let mutable productValue = 1.0
                            let mutable numericCount = 0L
                            let mutable minValue = infinity
                            let mutable maxValue = -infinity
                            let addNumeric n =
                                match fnName with
                                | "SUM" ->
                                    total <- total + n
                                    numericCount <- numericCount + 1L
                                | "PRODUCT" ->
                                    productValue <- flushProductUnderflow (productValue * n)
                                    numericCount <- numericCount + 1L
                                | "AVERAGE" ->
                                    total <- total + n
                                    numericCount <- numericCount + 1L
                                | "COUNT" ->
                                    numericCount <- numericCount + 1L
                                | "MIN" ->
                                    minValue <- min minValue n
                                    numericCount <- numericCount + 1L
                                | "MAX" ->
                                    maxValue <- max maxValue n
                                    numericCount <- numericCount + 1L
                                | _ -> ()
                            let consume value =
                                match unwrapValue value with
                                | CellValue.Error e -> firstErr <- ValueSome e
                                | CellValue.Number n -> addNumeric n
                                | CellValue.Boolean b when fnName <> "COUNT" -> addNumeric (if b then 1.0 else 0.0)
                                | _ -> ()
                            let consumeNumeric n =
                                addNumeric n
                            let mutable r = 1
                            match evalNumeric 1.0 1.0 rowName colName body with
                            | ValueSome _ ->
                                while not unsupported && ValueOption.isNone firstErr && r <= rows do
                                    let mutable c = 1
                                    while not unsupported && ValueOption.isNone firstErr && c <= cols do
                                        match evalNumeric (float r) (float c) rowName colName body with
                                        | ValueSome (Ok n) -> consumeNumeric n
                                        | ValueSome (Result.Error e) -> firstErr <- ValueSome e
                                        | ValueNone -> unsupported <- true
                                        c <- c + 1
                                    r <- r + 1
                            | ValueNone ->
                                while not unsupported && ValueOption.isNone firstErr && r <= rows do
                                    let mutable c = 1
                                    while not unsupported && ValueOption.isNone firstErr && c <= cols do
                                        match evalSimple (float r) (float c) rowName colName body with
                                        | ValueSome value -> consume value
                                        | ValueNone -> unsupported <- true
                                        c <- c + 1
                                    r <- r + 1
                            if unsupported then
                                ValueNone
                            else
                                match firstErr with
                                | ValueSome e -> ValueSome (err e)
                                | ValueNone ->
                                    let result =
                                        match fnName with
                                        | "PRODUCT" -> CellValue.Number (if numericCount = 0L then 0.0 else productValue)
                                        | "AVERAGE" -> if numericCount = 0L then err Err.div0 else CellValue.Number (total / float numericCount)
                                        | "COUNT" -> CellValue.Number (float numericCount)
                                        | "MIN" -> CellValue.Number (if numericCount = 0L then 0.0 else minValue)
                                        | "MAX" -> CellValue.Number (if numericCount = 0L then 0.0 else maxValue)
                                        | _ -> CellValue.Number total
                                    ValueSome result
            | _ -> ValueNone
        | [ sourceExpr ] when isSupportedAggregate ->
            match tryBilinearMakeArrayExpr sourceExpr with
            | ValueSome (Result.Error e) -> ValueSome (err e)
            | ValueSome (Ok (rows, cols, value)) -> aggregateBilinear rows cols value
            | ValueNone ->
                match comparisonMakeArrayConstantAggregate sourceExpr with
                | ValueSome (Result.Error e) -> ValueSome (err e)
                | ValueSome (Ok (ValueSome result)) -> ValueSome result
                | ValueSome (Ok ValueNone) | ValueNone ->
                    match comparisonMakeArrayStreamOfExpr sourceExpr with
                    | ValueSome (Result.Error e) -> ValueSome (err e)
                    | ValueSome (Ok (count, nextNumber, countAsNumber)) -> aggregateNumberStream count nextNumber countAsNumber
                    | ValueNone ->
                        match makeArrayStreamOfExpr sourceExpr with
                        | ValueSome (Result.Error e) -> ValueSome (err e)
                        | ValueSome (Ok (count, nextNumber, countAsNumber)) -> aggregateStream count nextNumber countAsNumber
                        | ValueNone -> ValueNone
        | _ -> ValueNone
    let trySequenceAggregate () =
        let isSupportedAggregate =
            match fnName with
            | "SUM" | "AVERAGE" | "COUNT" | "MIN" | "MAX" -> true
            | _ -> false
        let argOrDefault index defaultValue seqArgs =
            match List.tryItem index seqArgs with
            | Some MissingArg | None -> Ok defaultValue
            | Some arg -> eval arg ctx |> coerceNumber
        match args with
        | [ Function (sequenceName, seqArgs) ] when isSupportedAggregate && equalsOrdinalIgnoreCase sequenceName "SEQUENCE" ->
            match seqArgs with
            | [] -> ValueSome (err Err.value)
            | _ ->
                match argOrDefault 0 nan seqArgs, argOrDefault 1 1.0 seqArgs, argOrDefault 2 1.0 seqArgs, argOrDefault 3 1.0 seqArgs with
                | Result.Error e, _, _, _
                | _, Result.Error e, _, _
                | _, _, Result.Error e, _
                | _, _, _, Result.Error e -> ValueSome (err e)
                | Ok rowsN, Ok colsN, Ok start, Ok step ->
                    let rows = int rowsN
                    let cols = int colsN
                    if rows < 1 || cols < 1 then
                        ValueSome (err Err.value)
                    else
                        let count = int64 rows * int64 cols
                        if count < 1024L then
                            ValueNone
                        else
                            let countF = float count
                            let last = start + float (count - 1L) * step
                            let sum = countF * (start + last) / 2.0
                            let result =
                                match fnName with
                                | "SUM" -> CellValue.Number sum
                                | "AVERAGE" -> CellValue.Number (sum / countF)
                                | "COUNT" -> CellValue.Number countF
                                | "MIN" -> CellValue.Number (if step >= 0.0 then start else last)
                                | "MAX" -> CellValue.Number (if step >= 0.0 then last else start)
                                | _ -> CellValue.Number sum
                            ValueSome result
        | _ -> ValueNone
    let trySequenceStreamAggregate () =
        let isSupportedAggregate =
            match fnName with
            | "SUM" | "PRODUCT" | "AVERAGE" | "COUNT" | "MIN" | "MAX" -> true
            | _ -> false
        let argOrDefault index defaultValue seqArgs =
            match List.tryItem index seqArgs with
            | Some MissingArg | None -> Ok defaultValue
            | Some arg -> eval arg ctx |> coerceNumber
        let sequenceStreamOfArgs seqArgs =
            match seqArgs with
            | [] -> ValueSome (Result.Error Err.value)
            | _ ->
                match argOrDefault 0 nan seqArgs, argOrDefault 1 1.0 seqArgs, argOrDefault 2 1.0 seqArgs, argOrDefault 3 1.0 seqArgs with
                | Result.Error e, _, _, _
                | _, Result.Error e, _, _
                | _, _, Result.Error e, _
                | _, _, _, Result.Error e -> ValueSome (Result.Error e)
                | Ok rowsN, Ok colsN, Ok start, Ok step ->
                    let rows = int rowsN
                    let cols = int colsN
                    if rows < 1 || cols < 1 then
                        ValueSome (Result.Error Err.value)
                    else
                        let count = int64 rows * int64 cols
                        let mutable offset = 0L
                        let nextNumber () =
                            let value = start + float offset * step
                            offset <- offset + 1L
                            value
                        ValueSome (Ok (count, nextNumber, true))
        // Collapse a SEQUENCE mul/div chain only when its progression and
        // every intermediate result AND accumulated scale stay well within
        // normal finite doubles. Inspect the AST once, never generated cells.
        let safeSequenceMulDivChain root =
            let margin = minNormal * 4.0
            let safe lo hi =
                not (Double.IsNaN lo || Double.IsNaN hi || Double.IsInfinity lo || Double.IsInfinity hi)
                && ((lo > margin && hi > margin) || (lo < -margin && hi < -margin))
            let literalValue = function
                | Number n -> ValueSome n
                | UnaryOp ("-", Number n) -> ValueSome -n
                | _ -> ValueNone
            let literalDefault index fallback args =
                match List.tryItem index args with
                | None | Some MissingArg -> ValueSome fallback
                | Some arg -> literalValue arg
            let rec bounds expr =
                match expr with
                | Function (name, args) when equalsOrdinalIgnoreCase name "SEQUENCE" ->
                    match literalDefault 0 nan args, literalDefault 1 1.0 args,
                          literalDefault 2 1.0 args, literalDefault 3 1.0 args with
                    | ValueSome rows, ValueSome cols, ValueSome start, ValueSome step
                        when rows >= 1.0 && cols >= 1.0 && rows <= 2147483647.0 && cols <= 2147483647.0
                             && rows = floor rows && cols = floor cols ->
                        let count = int64 rows * int64 cols
                        let last = start + float (count - 1L) * step
                        let lo, hi = min start last, max start last
                        if safe lo hi then ValueSome (struct (count, start, step, lo, hi, 1.0))
                        else ValueNone
                    | _ -> ValueNone
                | BinOp (left, op, right) when op = "*" || op = "/" ->
                    match bounds left, literalValue right with
                    | ValueSome (struct (count, start, step, lo, hi, scale)), ValueSome factor ->
                        if op = "/" && factor > -minNormal && factor < minNormal then ValueNone
                        else
                            let nextScale = if op = "*" then scale * factor else scale / factor
                            let first = if op = "*" then lo * factor else lo / factor
                            let last = if op = "*" then hi * factor else hi / factor
                            let lower, upper = min first last, max first last
                            if safe lower upper && safe nextScale nextScale then
                                ValueSome (struct (count, start, step, lower, upper, nextScale))
                            else ValueNone
                    | _ -> ValueNone
                | _ -> ValueNone
            match bounds root with
            | ValueSome (struct (count, start, step, _, _, scale)) -> ValueSome (struct (count, start, step, scale))
            | ValueNone -> ValueNone
        let rec sequenceStreamOfExpr expr =
            let tryUnaryChain root =
                let isUnary expr =
                    match expr with
                    | UnaryOp (op, _) when op = "+" || op = "-" || op = "%" -> true
                    | _ -> false

                if not (isUnary root) then
                    struct (false, ValueNone)
                else
                    let mutable current = root
                    let mutable factor = 1.0
                    let mutable peeling = true
                    while peeling do
                        match current with
                        | UnaryOp (op, inner) when op = "+" || op = "-" || op = "%" ->
                            match op with
                            | "-" -> factor <- -factor
                            | "%" -> factor <- factor * 0.01
                            | _ -> ()
                            current <- inner
                        | _ -> peeling <- false

                    match sequenceStreamOfExpr current with
                    | ValueSome (Ok (count, nextNumber, countAsNumber)) ->
                        ValueSome (Ok (count, (fun () -> let n = factor * nextNumber () in if factor = 1.0 || factor = -1.0 then n else flushProductUnderflow n), countAsNumber))
                        |> fun result -> struct (true, result)
                    | ValueSome (Result.Error e) -> struct (true, ValueSome (Result.Error e))
                    | ValueNone -> struct (true, ValueNone)

            let tryAddSubChain root =
                let isAddSub expr =
                    match expr with
                    | BinOp (_, op, _) when op = "+" || op = "-" -> true
                    | _ -> false

                if not (isAddSub root) then
                    struct (false, ValueNone)
                else
                    let stack = ResizeArray<struct (float * Expr)>()
                    stack.Add(struct (1.0, root))
                    let mutable unsupported = false
                    let mutable firstErr = ValueNone
                    let mutable foundStream = false
                    let mutable streamCount = 0L
                    let mutable streamFactor = 1.0
                    let mutable streamNext = Unchecked.defaultof<unit -> float>
                    let mutable constant = 0.0

                    while not unsupported && ValueOption.isNone firstErr && stack.Count > 0 do
                        let last = stack.Count - 1
                        let struct (factor, expr) = stack.[last]
                        Portable.Stack.dropLast stack

                        match expr with
                        | BinOp (left, op, right) when op = "+" || op = "-" ->
                            stack.Add(struct ((if op = "-" then -factor else factor), right))
                            stack.Add(struct (factor, left))
                        | _ ->
                            match sequenceStreamOfExpr expr with
                            | ValueSome (Ok (count, nextNumber, _)) ->
                                if foundStream || factor <> 1.0 && factor <> -1.0 then
                                    unsupported <- true
                                else
                                    foundStream <- true
                                    streamCount <- count
                                    streamFactor <- factor
                                    streamNext <- nextNumber
                            | ValueSome (Result.Error e) -> firstErr <- ValueSome e
                            | ValueNone ->
                                if isArrayExpr expr then
                                    unsupported <- true
                                else
                                    match coerceNumber (eval expr ctx) with
                                    | Ok scalar -> constant <- constant + factor * scalar
                                    | Result.Error e -> firstErr <- ValueSome e

                    if unsupported then
                        struct (true, ValueNone)
                    else
                        match firstErr with
                        | ValueSome e -> struct (true, ValueSome (Result.Error e))
                        | ValueNone ->
                            if foundStream then
                                ValueSome (Ok (streamCount, (fun () -> constant + streamFactor * streamNext ()), true))
                                |> fun result -> struct (true, result)
                            else
                                struct (true, ValueNone)

            let tryMulDivChain root =
                let isMulDiv expr =
                    match expr with
                    | BinOp (_, op, _) when op = "*" || op = "/" -> true
                    | _ -> false

                if not (isMulDiv root)
                   || (match root with BinOp (left, _, right) -> isMulDiv left || isMulDiv right | _ -> false) then
                    struct (false, ValueNone)
                else
                    let stack = ResizeArray<struct (bool * Expr)>()
                    stack.Add(struct (true, root))
                    let mutable unsupported = false
                    let mutable firstErr = ValueNone
                    let mutable foundStream = false
                    let mutable streamCount = 0L
                    let mutable streamNext = Unchecked.defaultof<unit -> float>
                    let mutable streamCountAsNumber = true
                    let mutable scalar = 1.0

                    while not unsupported && ValueOption.isNone firstErr && stack.Count > 0 do
                        let last = stack.Count - 1
                        let struct (numerator, expr) = stack.[last]
                        Portable.Stack.dropLast stack

                        match expr with
                        | BinOp (left, op, right) when op = "*" || op = "/" ->
                            stack.Add(struct ((op = "*" && numerator) || (op = "/" && not numerator), right))
                            stack.Add(struct (numerator, left))
                        | _ ->
                            match sequenceStreamOfExpr expr with
                            | ValueSome (Ok (count, nextNumber, countAsNumber)) ->
                                if foundStream || not numerator then
                                    unsupported <- true
                                else
                                    foundStream <- true
                                    streamCount <- count
                                    streamNext <- nextNumber
                                    streamCountAsNumber <- countAsNumber
                            | ValueSome (Result.Error e) -> firstErr <- ValueSome e
                            | ValueNone ->
                                if isArrayExpr expr then
                                    unsupported <- true
                                else
                                    match coerceNumber (eval expr ctx) with
                                    | Ok value ->
                                        if numerator then
                                            scalar <- scalar * value
                                        elif value > -minNormal && value < minNormal then
                                            firstErr <- ValueSome Err.div0
                                        else
                                            scalar <- scalar / value
                                    | Result.Error e -> firstErr <- ValueSome e

                    if unsupported then
                        struct (true, ValueNone)
                    else
                        match firstErr with
                        | ValueSome e -> struct (true, ValueSome (Result.Error e))
                        | ValueNone ->
                            if foundStream then
                                ValueSome (Ok (streamCount, (fun () -> flushProductUnderflow (scalar * streamNext ())), streamCountAsNumber))
                                |> fun result -> struct (true, result)
                            else
                                struct (true, ValueNone)

            let struct (isUnaryChain, unaryResult) = tryUnaryChain expr
            if isUnaryChain then
                unaryResult
            else
                let struct (isAddSubChain, addSubResult) = tryAddSubChain expr
                if isAddSubChain then
                    addSubResult
                else
                    let struct (isMulDivChain, mulDivResult) =
                        match expr with
                        | BinOp (BinOp (_, ("*" | "/"), _), ("*" | "/"), _) ->
                            match safeSequenceMulDivChain expr with
                            | ValueSome (struct (count, start, step, scale)) ->
                                let mutable offset = 0L
                                let nextNumber () =
                                    let value = start + float offset * step
                                    offset <- offset + 1L
                                    scale * value
                                struct (true, ValueSome (Ok (count, nextNumber, true)))
                            | ValueNone -> tryMulDivChain expr
                        | _ -> tryMulDivChain expr
                    if isMulDivChain then
                        mulDivResult
                    else
                        match expr with
                        | Function (sequenceName, seqArgs) when equalsOrdinalIgnoreCase sequenceName "SEQUENCE" ->
                            sequenceStreamOfArgs seqArgs
                        | Function (transposeName, [ sourceExpr ]) when equalsOrdinalIgnoreCase transposeName "TRANSPOSE" ->
                            sequenceStreamOfExpr sourceExpr
                        | Function (sortName, [ sourceExpr ]) when equalsOrdinalIgnoreCase sortName "SORT" ->
                            sequenceStreamOfExpr sourceExpr
                        | UnaryOp (op, inner) when op = "+" || op = "-" || op = "%" ->
                            match sequenceStreamOfExpr inner with
                            | ValueSome (Ok (count, nextNumber, countAsNumber)) ->
                                let mapped =
                                    match op with
                                    | "-" -> fun () -> -(nextNumber ())
                                    | "%" -> fun () -> flushProductUnderflow (nextNumber () / 100.0)
                                    | _ -> nextNumber
                                ValueSome (Ok (count, mapped, countAsNumber))
                            | other -> other
                        | BinOp (left, op, right) when op = "+" || op = "-" || op = "*" || op = "/" ->
                            match sequenceStreamOfExpr left with
                            | ValueSome (Ok (count, nextNumber, countAsNumber)) when not (isArrayExpr right) ->
                                match coerceNumber (eval right ctx) with
                                | Result.Error e -> ValueSome (Result.Error e)
                                | Ok scalar ->
                                    match op with
                                    | "+" -> ValueSome (Ok (count, (fun () -> nextNumber () + scalar), countAsNumber))
                                    | "-" -> ValueSome (Ok (count, (fun () -> nextNumber () - scalar), countAsNumber))
                                    | "*" -> ValueSome (Ok (count, (fun () -> flushProductUnderflow (nextNumber () * scalar)), countAsNumber))
                                    | "/" when scalar > -minNormal && scalar < minNormal -> ValueSome (Result.Error Err.div0)
                                    | "/" -> ValueSome (Ok (count, (fun () -> flushProductUnderflow (nextNumber () / scalar)), countAsNumber))
                                    | _ -> ValueNone
                            | ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                            | _ ->
                                match sequenceStreamOfExpr right with
                                | ValueSome (Ok (count, nextNumber, countAsNumber)) when not (isArrayExpr left) ->
                                    match coerceNumber (eval left ctx) with
                                    | Result.Error e -> ValueSome (Result.Error e)
                                    | Ok scalar ->
                                        match op with
                                        | "+" -> ValueSome (Ok (count, (fun () -> scalar + nextNumber ()), countAsNumber))
                                        | "-" -> ValueSome (Ok (count, (fun () -> scalar - nextNumber ()), countAsNumber))
                                        | "*" -> ValueSome (Ok (count, (fun () -> flushProductUnderflow (scalar * nextNumber ())), countAsNumber))
                                        | _ -> ValueNone
                                | other -> other
                        | BinOp (left, op, right) when op = "=" || op = "<>" || op = "<" || op = ">" || op = "<=" || op = ">=" ->
                            let comparisonValue cmp =
                                let matches =
                                    match op with
                                    | "=" -> cmp = 0
                                    | "<>" -> cmp <> 0
                                    | "<" -> cmp < 0
                                    | ">" -> cmp > 0
                                    | "<=" -> cmp <= 0
                                    | _ -> cmp >= 0
                                if matches then 1.0 else 0.0
                            match sequenceStreamOfExpr left with
                            | ValueSome (Ok (count, nextNumber, _)) when not (isArrayExpr right) ->
                                match coerceNumber (eval right ctx) with
                                | Result.Error e -> ValueSome (Result.Error e)
                                | Ok scalar ->
                                    ValueSome (Ok (count, (fun () -> comparisonValue (compareNumbers (nextNumber ()) scalar)), false))
                            | ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                            | _ ->
                                match sequenceStreamOfExpr right with
                                | ValueSome (Ok (count, nextNumber, _)) when not (isArrayExpr left) ->
                                    match coerceNumber (eval left ctx) with
                                    | Result.Error e -> ValueSome (Result.Error e)
                                    | Ok scalar ->
                                        ValueSome (Ok (count, (fun () -> comparisonValue (compareNumbers scalar (nextNumber ()))), false))
                                | other -> other
                        | _ -> ValueNone
        match args with
        | [ sourceExpr ] when isSupportedAggregate && containsFunctionNamed "SEQUENCE" sourceExpr ->
            match sequenceStreamOfExpr sourceExpr with
            | ValueSome (Result.Error e) -> ValueSome (err e)
            | ValueSome (Ok (count, nextNumber, countAsNumber)) ->
                if count < 1024L then
                    ValueNone
                elif fnName = "COUNT" then
                    ValueSome (CellValue.Number (if countAsNumber then float count else 0.0))
                else
                    let mutable total = 0.0
                    let mutable productValue = 1.0
                    let mutable minValue = infinity
                    let mutable maxValue = -infinity
                    let mutable remaining = count
                    while remaining > 0L do
                        let n = nextNumber ()
                        match fnName with
                        | "SUM" | "AVERAGE" ->
                            total <- total + n
                        | "PRODUCT" ->
                            productValue <- flushProductUnderflow (productValue * n)
                        | "MIN" ->
                            if n < minValue then minValue <- n
                        | "MAX" ->
                            if n > maxValue then maxValue <- n
                        | _ -> ()
                        remaining <- remaining - 1L
                    let result =
                        match fnName with
                        | "PRODUCT" -> CellValue.Number productValue
                        | "AVERAGE" -> CellValue.Number (total / float count)
                        | "MIN" -> CellValue.Number minValue
                        | "MAX" -> CellValue.Number maxValue
                        | _ -> CellValue.Number total
                    ValueSome result
            | ValueNone -> ValueNone
        | _ -> ValueNone
    let tryRandArrayAggregate () =
        let isSupportedAggregate =
            match fnName with
            | "SUM" | "PRODUCT" | "AVERAGE" | "COUNT" | "MIN" | "MAX" -> true
            | _ -> false
        let argOrDefault index defaultValue randArgs =
            match List.tryItem index randArgs with
            | Some MissingArg | None -> Ok defaultValue
            | Some arg -> eval arg ctx |> coerceNumber
        let boolArgOrDefault index defaultValue randArgs =
            match List.tryItem index randArgs with
            | Some arg ->
                match coerceBool (eval arg ctx) with
                | Ok value -> Ok value
                | Result.Error e -> Result.Error e
            | None -> Ok defaultValue
        let randStreamOfArgs randArgs =
            match argOrDefault 0 1.0 randArgs, argOrDefault 1 1.0 randArgs, argOrDefault 2 0.0 randArgs, argOrDefault 3 1.0 randArgs, boolArgOrDefault 4 false randArgs with
            | Result.Error e, _, _, _, _
            | _, Result.Error e, _, _, _
            | _, _, Result.Error e, _, _
            | _, _, _, Result.Error e, _
            | _, _, _, _, Result.Error e -> ValueSome (Result.Error e)
            | Ok rowsN, Ok colsN, Ok lo, Ok hi, Ok whole ->
                let rows = int rowsN
                let cols = int colsN
                if rows < 1 || cols < 1 then
                    ValueSome (Result.Error Err.value)
                elif hi < lo then
                    ValueSome (Result.Error Err.value)
                else
                    let count = int64 rows * int64 cols
                    let nextNumber () =
                        if whole then float (rng.Next(int lo, int hi + 1))
                        else lo + rng.NextDouble() * (hi - lo)
                    ValueSome (Ok (count, nextNumber, true))
        let rec randStreamOfExpr expr =
            let tryUnaryChain root =
                let isUnary expr =
                    match expr with
                    | UnaryOp (op, _) when op = "+" || op = "-" || op = "%" -> true
                    | _ -> false

                if not (isUnary root) then
                    struct (false, ValueNone)
                else
                    let mutable current = root
                    let mutable factor = 1.0
                    let mutable peeling = true
                    while peeling do
                        match current with
                        | UnaryOp (op, inner) when op = "+" || op = "-" || op = "%" ->
                            match op with
                            | "-" -> factor <- -factor
                            | "%" -> factor <- factor * 0.01
                            | _ -> ()
                            current <- inner
                        | _ -> peeling <- false

                    match randStreamOfExpr current with
                    | ValueSome (Ok (count, nextNumber, countAsNumber)) ->
                        ValueSome (Ok (count, (fun () -> let n = factor * nextNumber () in if factor = 1.0 || factor = -1.0 then n else flushProductUnderflow n), countAsNumber))
                        |> fun result -> struct (true, result)
                    | ValueSome (Result.Error e) -> struct (true, ValueSome (Result.Error e))
                    | ValueNone -> struct (true, ValueNone)

            let tryAddSubChain root =
                let isAddSub expr =
                    match expr with
                    | BinOp (_, op, _) when op = "+" || op = "-" -> true
                    | _ -> false

                if not (isAddSub root) then
                    struct (false, ValueNone)
                else
                    let stack = ResizeArray<struct (float * Expr)>()
                    stack.Add(struct (1.0, root))
                    let mutable unsupported = false
                    let mutable firstErr = ValueNone
                    let mutable foundStream = false
                    let mutable streamCount = 0L
                    let mutable streamFactor = 1.0
                    let mutable streamNext = Unchecked.defaultof<unit -> float>
                    let mutable constant = 0.0

                    while not unsupported && ValueOption.isNone firstErr && stack.Count > 0 do
                        let last = stack.Count - 1
                        let struct (factor, expr) = stack.[last]
                        Portable.Stack.dropLast stack

                        match expr with
                        | BinOp (left, op, right) when op = "+" || op = "-" ->
                            stack.Add(struct ((if op = "-" then -factor else factor), right))
                            stack.Add(struct (factor, left))
                        | _ ->
                            match randStreamOfExpr expr with
                            | ValueSome (Ok (count, nextNumber, _)) ->
                                if foundStream || factor <> 1.0 && factor <> -1.0 then
                                    unsupported <- true
                                else
                                    foundStream <- true
                                    streamCount <- count
                                    streamFactor <- factor
                                    streamNext <- nextNumber
                            | ValueSome (Result.Error e) -> firstErr <- ValueSome e
                            | ValueNone ->
                                if isArrayExpr expr then
                                    unsupported <- true
                                else
                                    match coerceNumber (eval expr ctx) with
                                    | Ok scalar -> constant <- constant + factor * scalar
                                    | Result.Error e -> firstErr <- ValueSome e

                    if unsupported then
                        struct (true, ValueNone)
                    else
                        match firstErr with
                        | ValueSome e -> struct (true, ValueSome (Result.Error e))
                        | ValueNone ->
                            if foundStream then
                                ValueSome (Ok (streamCount, (fun () -> constant + streamFactor * streamNext ()), true))
                                |> fun result -> struct (true, result)
                            else
                                struct (true, ValueNone)

            let tryMulDivChain root =
                let isMulDiv expr =
                    match expr with
                    | BinOp (_, op, _) when op = "*" || op = "/" -> true
                    | _ -> false

                if not (isMulDiv root) || (match root with BinOp (left, _, right) -> isMulDiv left || isMulDiv right | _ -> false) then
                    struct (false, ValueNone)
                else
                    let stack = ResizeArray<struct (bool * Expr)>()
                    stack.Add(struct (true, root))
                    let mutable unsupported = false
                    let mutable firstErr = ValueNone
                    let mutable foundStream = false
                    let mutable streamCount = 0L
                    let mutable streamNext = Unchecked.defaultof<unit -> float>
                    let mutable streamCountAsNumber = true
                    let mutable scalar = 1.0

                    while not unsupported && ValueOption.isNone firstErr && stack.Count > 0 do
                        let last = stack.Count - 1
                        let struct (numerator, expr) = stack.[last]
                        Portable.Stack.dropLast stack

                        match expr with
                        | BinOp (left, op, right) when op = "*" || op = "/" ->
                            stack.Add(struct ((op = "*" && numerator) || (op = "/" && not numerator), right))
                            stack.Add(struct (numerator, left))
                        | _ ->
                            match randStreamOfExpr expr with
                            | ValueSome (Ok (count, nextNumber, countAsNumber)) ->
                                if foundStream || not numerator then
                                    unsupported <- true
                                else
                                    foundStream <- true
                                    streamCount <- count
                                    streamNext <- nextNumber
                                    streamCountAsNumber <- countAsNumber
                            | ValueSome (Result.Error e) -> firstErr <- ValueSome e
                            | ValueNone ->
                                if isArrayExpr expr then
                                    unsupported <- true
                                else
                                    match coerceNumber (eval expr ctx) with
                                    | Ok value ->
                                        if numerator then
                                            scalar <- scalar * value
                                        elif value > -minNormal && value < minNormal then
                                            firstErr <- ValueSome Err.div0
                                        else
                                            scalar <- scalar / value
                                    | Result.Error e -> firstErr <- ValueSome e

                    if unsupported then
                        struct (true, ValueNone)
                    else
                        match firstErr with
                        | ValueSome e -> struct (true, ValueSome (Result.Error e))
                        | ValueNone ->
                            if foundStream then
                                ValueSome (Ok (streamCount, (fun () -> flushProductUnderflow (scalar * streamNext ())), streamCountAsNumber))
                                |> fun result -> struct (true, result)
                            else
                                struct (true, ValueNone)

            let struct (isUnaryChain, unaryResult) = tryUnaryChain expr
            if isUnaryChain then
                unaryResult
            else
                let struct (isAddSubChain, addSubResult) = tryAddSubChain expr
                if isAddSubChain then
                    addSubResult
                else
                    let struct (isMulDivChain, mulDivResult) = tryMulDivChain expr
                    if isMulDivChain then
                        mulDivResult
                    else
                        match expr with
                        | Function (randName, randArgs) when equalsOrdinalIgnoreCase randName "RANDARRAY" ->
                            randStreamOfArgs randArgs
                        | Function (transposeName, [ sourceExpr ]) when equalsOrdinalIgnoreCase transposeName "TRANSPOSE" ->
                            randStreamOfExpr sourceExpr
                        | Function (sortName, [ sourceExpr ]) when equalsOrdinalIgnoreCase sortName "SORT" ->
                            randStreamOfExpr sourceExpr
                        | UnaryOp (op, inner) when op = "+" || op = "-" || op = "%" ->
                            match randStreamOfExpr inner with
                            | ValueSome (Ok (count, nextNumber, countAsNumber)) ->
                                let mapped =
                                    match op with
                                    | "-" -> fun () -> -(nextNumber ())
                                    | "%" -> fun () -> flushProductUnderflow (nextNumber () / 100.0)
                                    | _ -> nextNumber
                                ValueSome (Ok (count, mapped, countAsNumber))
                            | other -> other
                        | BinOp (left, op, right) when op = "+" || op = "-" || op = "*" || op = "/" ->
                            match randStreamOfExpr left with
                            | ValueSome (Ok (count, nextNumber, countAsNumber)) when not (isArrayExpr right) ->
                                match coerceNumber (eval right ctx) with
                                | Result.Error e -> ValueSome (Result.Error e)
                                | Ok scalar ->
                                    match op with
                                    | "+" -> ValueSome (Ok (count, (fun () -> nextNumber () + scalar), countAsNumber))
                                    | "-" -> ValueSome (Ok (count, (fun () -> nextNumber () - scalar), countAsNumber))
                                    | "*" -> ValueSome (Ok (count, (fun () -> flushProductUnderflow (nextNumber () * scalar)), countAsNumber))
                                    | "/" when scalar > -minNormal && scalar < minNormal -> ValueSome (Result.Error Err.div0)
                                    | "/" -> ValueSome (Ok (count, (fun () -> flushProductUnderflow (nextNumber () / scalar)), countAsNumber))
                                    | _ -> ValueNone
                            | ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                            | _ ->
                                match randStreamOfExpr right with
                                | ValueSome (Ok (count, nextNumber, countAsNumber)) when not (isArrayExpr left) ->
                                    match coerceNumber (eval left ctx) with
                                    | Result.Error e -> ValueSome (Result.Error e)
                                    | Ok scalar ->
                                        match op with
                                        | "+" -> ValueSome (Ok (count, (fun () -> scalar + nextNumber ()), countAsNumber))
                                        | "-" -> ValueSome (Ok (count, (fun () -> scalar - nextNumber ()), countAsNumber))
                                        | "*" -> ValueSome (Ok (count, (fun () -> flushProductUnderflow (scalar * nextNumber ())), countAsNumber))
                                        | _ -> ValueNone
                                | other -> other
                        | BinOp (left, op, right) when op = "=" || op = "<>" || op = "<" || op = ">" || op = "<=" || op = ">=" ->
                            let comparisonValue cmp =
                                let matches =
                                    match op with
                                    | "=" -> cmp = 0
                                    | "<>" -> cmp <> 0
                                    | "<" -> cmp < 0
                                    | ">" -> cmp > 0
                                    | "<=" -> cmp <= 0
                                    | _ -> cmp >= 0
                                if matches then 1.0 else 0.0
                            match randStreamOfExpr left with
                            | ValueSome (Ok (count, nextNumber, _)) when not (isArrayExpr right) ->
                                match coerceNumber (eval right ctx) with
                                | Result.Error e -> ValueSome (Result.Error e)
                                | Ok scalar ->
                                    ValueSome (Ok (count, (fun () -> comparisonValue (compareNumbers (nextNumber ()) scalar)), false))
                            | ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                            | _ ->
                                match randStreamOfExpr right with
                                | ValueSome (Ok (count, nextNumber, _)) when not (isArrayExpr left) ->
                                    match coerceNumber (eval left ctx) with
                                    | Result.Error e -> ValueSome (Result.Error e)
                                    | Ok scalar ->
                                        ValueSome (Ok (count, (fun () -> comparisonValue (compareNumbers scalar (nextNumber ()))), false))
                                | other -> other
                        | _ -> ValueNone
        match args with
        | [ sourceExpr ] when isSupportedAggregate && containsFunctionNamed "RANDARRAY" sourceExpr ->
            match randStreamOfExpr sourceExpr with
            | ValueSome (Result.Error e) -> ValueSome (err e)
            | ValueSome (Ok (count, nextNumber, countAsNumber)) ->
                if count < 1024L then
                    ValueNone
                elif fnName = "COUNT" then
                    ValueSome (CellValue.Number (if countAsNumber then float count else 0.0))
                else
                    let mutable total = 0.0
                    let mutable productValue = 1.0
                    let mutable minValue = infinity
                    let mutable maxValue = -infinity
                    let mutable remaining = count
                    while remaining > 0L do
                        let n = nextNumber ()
                        match fnName with
                        | "SUM" | "AVERAGE" ->
                            total <- total + n
                        | "PRODUCT" ->
                            productValue <- flushProductUnderflow (productValue * n)
                        | "MIN" ->
                            if n < minValue then minValue <- n
                        | "MAX" ->
                            if n > maxValue then maxValue <- n
                        | _ -> ()
                        remaining <- remaining - 1L
                    let result =
                        match fnName with
                        | "PRODUCT" -> CellValue.Number productValue
                        | "AVERAGE" -> CellValue.Number (total / float count)
                        | "MIN" -> CellValue.Number minValue
                        | "MAX" -> CellValue.Number maxValue
                        | _ -> CellValue.Number total
                    ValueSome result
            | ValueNone -> ValueNone
        | _ -> ValueNone
    let trySparseSortPassthroughAggregate () =
        let isSupportedAggregate =
            match fnName with
            | "SUM" | "PRODUCT" | "AVERAGE" | "COUNT" | "MIN" | "MAX" -> true
            | _ -> false
        match args with
        | [ Function (sortName, sortArgs) ] when isSupportedAggregate && (equalsOrdinalIgnoreCase sortName "SORT" || equalsOrdinalIgnoreCase sortName "SORTBY") ->
            let sourceArg =
                match sortArgs with
                | [ sourceExpr ] when equalsOrdinalIgnoreCase sortName "SORT" -> ValueSome sourceExpr
                | sourceExpr :: _ :: _ when equalsOrdinalIgnoreCase sortName "SORTBY" -> ValueSome sourceExpr
                | _ -> ValueNone
            match sourceArg with
            | ValueNone -> ValueNone
            | ValueSome sourceExpr ->
                match rangeAccessOf model workbook sourceExpr with
                | Some source when source.Count > 0 ->
                    let area = int64 source.Count
                    if area < 1024L || int64 source.M.Cells.Count > area * 4L then
                        ValueNone
                    else
                        let mutable hasStoredError = false
                        source.M.Cells
                        |> Map.iter (fun addr cell ->
                            if not hasStoredError then
                                match tryRangeOffsetForCount source source.Count addr with
                                | ValueSome _ ->
                                    match unwrapValue cell.Value with
                                    | CellValue.Error _ -> hasStoredError <- true
                                    | _ -> ()
                                | ValueNone -> ())
                        if hasStoredError then ValueNone
                        else ValueSome (evalFunction fnName [ sourceExpr ] ctx)
                | _ -> ValueNone
        | _ -> ValueNone
    let trySparseTransposePassthroughAggregate () =
        let isSupportedAggregate =
            match fnName with
            | "SUM" | "PRODUCT" | "AVERAGE" | "COUNT" | "MIN" | "MAX" -> true
            | _ -> false
        let rec unwrapTranspose expr =
            match expr with
            | Function (transposeName, [ inner ]) when equalsOrdinalIgnoreCase transposeName "TRANSPOSE" ->
                unwrapTranspose inner
            | _ -> expr
        match args with
        | [ Function (transposeName, [ sourceExpr ]) ] when isSupportedAggregate && equalsOrdinalIgnoreCase transposeName "TRANSPOSE" ->
            ValueSome (evalFunction fnName [ unwrapTranspose sourceExpr ] ctx)
        | _ -> ValueNone
    let trySparseFilterAggregate () =
        let isSupportedAggregate =
            match fnName with
            | "SUM" | "PRODUCT" | "AVERAGE" | "COUNT" | "MIN" | "MAX" -> true
            | _ -> false
        let includeAccessAndPredicate includeExpr =
            let coercePredicate value =
                match coerceBool value with
                | Ok true -> true
                | _ -> false
            let comparisonPredicate op left right =
                match compareOperandsNoError left right with
                | ValueSome c ->
                    match op with
                    | "=" -> c = 0
                    | "<>" -> c <> 0
                    | "<" -> c < 0
                    | ">" -> c > 0
                    | "<=" -> c <= 0
                    | ">=" -> c >= 0
                    | _ -> coercePredicate (applyBin op left right)
                | ValueNone -> false
            let readerFor (access: RangeAccess) predicate =
                if access.W <> 1 then ValueNone
                else
                    if predicate CellValue.Empty then ValueNone
                    else ValueSome (access, predicate)
            match includeExpr with
            | BinOp (leftExpr, op, rightExpr) ->
                match rangeAccessOf model workbook leftExpr, rangeAccessOf model workbook rightExpr with
                | Some access, None ->
                    let scalar = eval rightExpr ctx
                    readerFor access (fun value -> comparisonPredicate op value scalar)
                | None, Some access ->
                    let scalar = eval leftExpr ctx
                    readerFor access (fun value -> comparisonPredicate op scalar value)
                | _ -> ValueNone
            | _ ->
                match rangeAccessOf model workbook includeExpr with
                | Some access -> readerFor access coercePredicate
                | None -> ValueNone
        match args with
        | Function (filterName, filterArgs) :: [] when isSupportedAggregate && equalsOrdinalIgnoreCase filterName "FILTER" ->
            match filterArgs with
            | sourceExpr :: includeExpr :: _ ->
                match rangeAccessOf model workbook sourceExpr, includeAccessAndPredicate includeExpr with
                | Some source, ValueSome (includeAccess, includeReader) when source.Count >= source.W && source.W > 0 ->
                    let rows = min (source.Count / source.W) includeAccess.Count
                    let area = int64 includeAccess.Count
                    if rows <= 0 then
                        match filterArgs with
                        | _ :: _ :: ifEmptyExpr :: _ -> ValueSome (evalFunction fnName [ ifEmptyExpr ] ctx)
                        | _ -> ValueSome (err Err.calc)
                    else
                        let mutable firstErr = ValueNone
                        let mutable count = 0
                        let mutable total = 0.0
                        let mutable productValue = 1.0
                        let mutable minValue = infinity
                        let mutable maxValue = -infinity
                        let mutable keptAny = false
                        let addNumeric n =
                            match fnName with
                            | "SUM" ->
                                total <- total + n
                                count <- count + 1
                            | "PRODUCT" ->
                                productValue <- flushProductUnderflow (productValue * n)
                                count <- count + 1
                            | "AVERAGE" ->
                                total <- total + n
                                count <- count + 1
                            | "COUNT" ->
                                count <- count + 1
                            | "MIN" ->
                                minValue <- min minValue n
                                count <- count + 1
                            | "MAX" ->
                                maxValue <- max maxValue n
                                count <- count + 1
                            | _ -> ()
                        let consume value =
                            match firstErr with
                            | ValueSome _ -> ()
                            | ValueNone ->
                                match unwrapValue value with
                                | CellValue.Error e -> firstErr <- ValueSome e
                                | CellValue.Number n -> addNumeric n
                                | CellValue.Boolean b when fnName <> "COUNT" -> addNumeric (if b then 1.0 else 0.0)
                                | _ -> ()
                        let consumeRow rowIndex =
                            if includeReader (rangeRead includeAccess rowIndex) then
                                keptAny <- true
                                let baseOffset = rowIndex * source.W
                                for c = 0 to source.W - 1 do
                                    consume (rangeRead source (baseOffset + c))
                        let consumeDenseRows () =
                            let sourceValues = materializeRangeValues source
                            let includeValues = materializeRangeValues includeAccess
                            for rowIndex = 0 to rows - 1 do
                                if includeReader includeValues.[rowIndex] then
                                    keptAny <- true
                                    let baseOffset = rowIndex * source.W
                                    for c = 0 to source.W - 1 do
                                        consume sourceValues.[baseOffset + c]
                        if area >= 1024L && int64 includeAccess.M.Cells.Count <= area then
                            includeAccess.M.Cells
                            |> Map.iter (fun addr _ ->
                                match tryRangeOffsetForCount includeAccess includeAccess.Count addr with
                                | ValueSome rowIndex when rowIndex < rows -> consumeRow rowIndex
                                | _ -> ())
                        else
                            consumeDenseRows ()
                        match firstErr with
                        | ValueSome e -> ValueSome (err e)
                        | ValueNone ->
                            if not keptAny then
                                match filterArgs with
                                | _ :: _ :: ifEmptyExpr :: _ -> ValueSome (evalFunction fnName [ ifEmptyExpr ] ctx)
                                | _ -> ValueSome (err Err.calc)
                            else
                                let result =
                                    match fnName with
                                    | "PRODUCT" -> CellValue.Number (if count = 0 then 0.0 else productValue)
                                    | "AVERAGE" -> if count = 0 then err Err.div0 else CellValue.Number (total / float count)
                                    | "COUNT" -> CellValue.Number (float count)
                                    | "MIN" -> CellValue.Number (if count = 0 then 0.0 else minValue)
                                    | "MAX" -> CellValue.Number (if count = 0 then 0.0 else maxValue)
                                    | _ -> CellValue.Number total
                                ValueSome result
                | _ -> ValueNone
            | _ -> ValueNone
        | _ -> ValueNone
    let trySparseByRowAggregate () =
        match args with
        | [ Function (byRowName, [ sourceExpr; lambdaExpr ]) ] when (fnName = "SUM" || fnName = "COUNT") && equalsOrdinalIgnoreCase byRowName "BYROW" ->
            match rangeAccessOf model workbook sourceExpr, lambdaPartsOf lambdaExpr with
            | Some source, Some [ NameRef param; Function (innerName, [ NameRef arg ]) ]
                when source.Count >= source.W
                     && source.W > 0
                     && equalsOrdinalIgnoreCase param arg
                     && (equalsOrdinalIgnoreCase innerName "SUM" || equalsOrdinalIgnoreCase innerName "COUNT") ->
                let area = int64 source.Count
                if area < 1024L || int64 source.M.Cells.Count > area * 4L then
                    ValueNone
                else
                    let rows = source.Count / source.W
                    let innerIsCount = equalsOrdinalIgnoreCase innerName "COUNT"
                    let mutable total = 0.0
                    let mutable firstErr = ValueNone
                    source.M.Cells
                    |> Map.iter (fun addr cell ->
                        match tryRangeOffsetForCount source source.Count addr with
                        | ValueSome offset ->
                            match unwrapValue cell.Value with
                            | CellValue.Error e ->
                                match firstErr with
                                | ValueNone -> firstErr <- ValueSome (offset, e)
                                | ValueSome (existingOffset, _) when offset < existingOffset -> firstErr <- ValueSome (offset, e)
                                | _ -> ()
                            | CellValue.Number n ->
                                if innerIsCount then total <- total + 1.0
                                else total <- total + n
                            | CellValue.Boolean b when not innerIsCount ->
                                total <- total + if b then 1.0 else 0.0
                            | _ -> ()
                        | ValueNone -> ())
                    match firstErr with
                    | ValueSome (_, e) -> ValueSome (err e)
                    | ValueNone ->
                        if fnName = "COUNT" then ValueSome (CellValue.Number (float rows))
                        else ValueSome (CellValue.Number total)
            | _ -> ValueNone
        | _ -> ValueNone
    let trySparseScanAggregate () =
        let isName expected expr =
            match expr with
            | NameRef name -> equalsOrdinalIgnoreCase name expected
            | _ -> false
        match args with
        | [ Function (scanName, [ initialExpr; sourceExpr; lambdaExpr ]) ]
            when (fnName = "SUM" || fnName = "COUNT" || fnName = "AVERAGE" || fnName = "MIN" || fnName = "MAX")
                 && equalsOrdinalIgnoreCase scanName "SCAN" ->
            match rangeAccessOf model workbook sourceExpr, lambdaPartsOf lambdaExpr with
            | Some source, Some [ NameRef accName; NameRef valueName; BinOp (left, "+", right) ]
                when source.Count > 0 && source.Count >= source.W ->
                let increment =
                    if isName accName left && lambdaBodyUsesOnlySimpleArithmetic valueName right && not (containsVolatile right) then
                        ValueSome right
                    elif isName accName right && lambdaBodyUsesOnlySimpleArithmetic valueName left && not (containsVolatile left) then
                        ValueSome left
                    else
                        ValueNone
                match increment with
                | ValueNone -> ValueNone
                | ValueSome inc ->
                    let area = int64 source.Count
                    if area < 1024L || int64 source.M.Cells.Count > area * 4L then
                        ValueNone
                    else
                        match coerceNumber (eval initialExpr ctx) with
                        | Result.Error _ -> ValueNone
                        | Ok initial ->
                            let valueLambda = [ NameRef valueName; inc ]
                            let incrementFor value =
                                eval (applyLambdaExpr valueLambda [ valueToExpr value ]) ctx
                                |> coerceNumber
                            match incrementFor CellValue.Empty with
                            | Result.Error e -> ValueSome (err e)
                            | Ok blankIncrement ->
                                let totalCount = int64 source.Count
                                let mutable running = initial
                                let mutable scanSum = 0.0
                                let mutable scanMin = infinity
                                let mutable scanMax = -infinity
                                let mutable scanCount = 0L
                                let mutable nextOffset = 0
                                let mutable firstErr = ValueNone
                                let addRun repeats increment =
                                    if repeats > 0L then
                                        let repeatsF = float repeats
                                        let firstValue = running + increment
                                        let lastValue = running + increment * repeatsF
                                        scanSum <- scanSum + repeatsF * running + increment * repeatsF * (repeatsF + 1.0) / 2.0
                                        if increment >= 0.0 then
                                            if firstValue < scanMin then scanMin <- firstValue
                                            if lastValue > scanMax then scanMax <- lastValue
                                        else
                                            if lastValue < scanMin then scanMin <- lastValue
                                            if firstValue > scanMax then scanMax <- firstValue
                                        scanCount <- scanCount + repeats
                                        running <- lastValue
                                let addValue offset value =
                                    if ValueOption.isNone firstErr then
                                        if offset > nextOffset then
                                            addRun (int64 (offset - nextOffset)) blankIncrement
                                            nextOffset <- offset
                                        match incrementFor value with
                                        | Result.Error e -> firstErr <- ValueSome e
                                        | Ok increment ->
                                            addRun 1L increment
                                            nextOffset <- offset + 1
                                source.M.Cells
                                |> Map.iter (fun addr cell ->
                                    match tryRangeOffsetForCount source source.Count addr with
                                    | ValueSome offset when offset >= nextOffset -> addValue offset (unwrapValue cell.Value)
                                    | _ -> ())
                                match firstErr with
                                | ValueSome e -> ValueSome (err e)
                                | ValueNone ->
                                    if nextOffset < source.Count then
                                        addRun (int64 (source.Count - nextOffset)) blankIncrement
                                    match fnName with
                                    | "COUNT" -> ValueSome (CellValue.Number (float totalCount))
                                    | "AVERAGE" ->
                                        if scanCount = 0L then ValueSome (err Err.div0)
                                        else ValueSome (CellValue.Number (scanSum / float scanCount))
                                    | "MIN" -> ValueSome (CellValue.Number (if scanCount = 0L then 0.0 else scanMin))
                                    | "MAX" -> ValueSome (CellValue.Number (if scanCount = 0L then 0.0 else scanMax))
                                    | _ -> ValueSome (CellValue.Number scanSum)
            | _ -> ValueNone
        | _ -> ValueNone
    let trySparseOrderedArithmeticValues sourceExpr =
        let addNumeric offset repeats value (values: ResizeArray<struct (float * int64 * int)>) =
            if repeats > 0L then
                match unwrapValue value with
                | CellValue.Error e -> ValueSome e
                | CellValue.Number n ->
                    values.Add(struct (n, repeats, offset))
                    ValueNone
                | CellValue.Boolean b ->
                    values.Add(struct ((if b then 1.0 else 0.0), repeats, offset))
                    ValueNone
                | _ -> ValueNone
            else
                ValueNone
        let sparseOffsets (source: RangeAccess) (emit: int -> unit) =
            source.M.Cells
            |> Map.iter (fun addr _ ->
                match tryRangeOffsetForCount source source.Count addr with
                | ValueSome offset -> emit offset
                | ValueNone -> ())
        let collect count sparseEligible iterOffsets readAt gapValue =
            if not sparseEligible then
                ValueNone
            else
                let values = ResizeArray<struct (float * int64 * int)>()
                let visited = HashSet<int>()
                let mutable firstErr = ValueNone
                let consume offset repeats value =
                    match addNumeric offset repeats value values with
                    | ValueSome e ->
                        match firstErr with
                        | ValueNone -> firstErr <- ValueSome (offset, e)
                        | ValueSome (existingOffset, _) when offset < existingOffset -> firstErr <- ValueSome (offset, e)
                        | _ -> ()
                    | ValueNone -> ()
                let visit offset =
                    if visited.Add offset then
                        consume offset 1L (readAt offset)
                iterOffsets visit
                if visited.Count < count then
                    let mutable gapOffset = 0
                    while gapOffset < count && visited.Contains gapOffset do
                        gapOffset <- gapOffset + 1
                    consume gapOffset (int64 (count - visited.Count)) gapValue
                match firstErr with
                | ValueSome (_, e) -> ValueSome (Result.Error e)
                | ValueNone -> ValueSome (Ok values)
        match sourceExpr with
        | BinOp (leftExpr, op, rightExpr) when op = "+" || op = "-" || op = "*" || op = "/" ->
            match rangeAccessOf model workbook leftExpr, rangeAccessOf model workbook rightExpr with
            | Some left, Some right when op <> "/" && left.Count = right.Count && left.W = right.W && left.Count > 0 ->
                let area = int64 left.Count
                let sparseEligible = area >= 1024L && int64 left.M.Cells.Count + int64 right.M.Cells.Count <= area * 4L
                let iterOffsets emit =
                    sparseOffsets left emit
                    sparseOffsets right emit
                collect
                    left.Count
                    sparseEligible
                    iterOffsets
                    (fun offset -> applyBin op (rangeRead left offset) (rangeRead right offset))
                    (applyBin op CellValue.Empty CellValue.Empty)
            | Some source, None when source.Count > 0 && not (isArrayExpr rightExpr) ->
                let scalar = eval rightExpr ctx
                let area = int64 source.Count
                let sparseEligible = area >= 1024L && int64 source.M.Cells.Count <= area * 4L
                match op, coerceNumber scalar with
                | "/", Ok denominator when denominator <> 0.0 ->
                    collect source.Count sparseEligible (sparseOffsets source) (fun offset -> applyBin op (rangeRead source offset) scalar) (applyBin op CellValue.Empty scalar)
                | ("+" | "-" | "*"), _ ->
                    collect source.Count sparseEligible (sparseOffsets source) (fun offset -> applyBin op (rangeRead source offset) scalar) (applyBin op CellValue.Empty scalar)
                | _ -> ValueNone
            | None, Some source when op <> "/" && source.Count > 0 && not (isArrayExpr leftExpr) ->
                let scalar = eval leftExpr ctx
                let area = int64 source.Count
                let sparseEligible = area >= 1024L && int64 source.M.Cells.Count <= area * 4L
                collect
                    source.Count
                    sparseEligible
                    (sparseOffsets source)
                    (fun offset -> applyBin op scalar (rangeRead source offset))
                    (applyBin op scalar CellValue.Empty)
            | _ -> ValueNone
        | UnaryOp (op, innerExpr) when op = "+" || op = "-" || op = "%" ->
            match rangeAccessOf model workbook innerExpr with
            | Some source when source.Count > 0 ->
                let area = int64 source.Count
                let sparseEligible = area >= 1024L && int64 source.M.Cells.Count <= area * 4L
                collect
                    source.Count
                    sparseEligible
                    (sparseOffsets source)
                    (fun offset -> applyUnary op (rangeRead source offset))
                    (applyUnary op CellValue.Empty)
            | _ -> ValueNone
        | _ -> ValueNone
    let trySparseNpvArithmetic rate sourceExpr =
        let discountedRepeated startIndex repeats value =
            if repeats <= 0L || value = 0.0 then
                0.0
            elif abs rate < 1e-12 then
                value * float repeats
            else
                let q = 1.0 / (1.0 + rate)
                value * (q ** float startIndex) * (1.0 - (q ** float repeats)) / (1.0 - q)
        let sparseOffsets (source: RangeAccess) (emit: int -> unit) =
            source.M.Cells
            |> Map.iter (fun addr _ ->
                match tryRangeOffsetForCount source source.Count addr with
                | ValueSome offset -> emit offset
                | ValueNone -> ())
        let collect count sparseEligible iterOffsets readAt gapValue =
            if not sparseEligible then
                ValueNone
            else
                let offsets = HashSet<int>()
                iterOffsets (fun offset -> offsets.Add offset |> ignore)
                let sortedOffsets = hashSetToArray offsets
                Array.sortInPlace sortedOffsets
                let mutable currentOffset = 0
                let mutable npvIndex = 1L
                let mutable total = 0.0
                let mutable firstErr = ValueNone
                let consume repeats value =
                    if repeats > 0L then
                        match unwrapValue value with
                        | CellValue.Error e ->
                            match firstErr with
                            | ValueNone -> firstErr <- ValueSome e
                            | ValueSome _ -> ()
                        | CellValue.Number n ->
                            total <- total + discountedRepeated npvIndex repeats n
                            npvIndex <- npvIndex + repeats
                        | CellValue.Boolean b ->
                            total <- total + discountedRepeated npvIndex repeats (if b then 1.0 else 0.0)
                            npvIndex <- npvIndex + repeats
                        | _ -> ()
                let mutable i = 0
                while ValueOption.isNone firstErr && i < sortedOffsets.Length do
                    let offset = sortedOffsets.[i]
                    if offset > currentOffset then
                        consume (int64 (offset - currentOffset)) gapValue
                    if ValueOption.isNone firstErr then
                        consume 1L (readAt offset)
                    currentOffset <- offset + 1
                    i <- i + 1
                if ValueOption.isNone firstErr && currentOffset < count then
                    consume (int64 (count - currentOffset)) gapValue
                match firstErr with
                | ValueSome e -> ValueSome (err e)
                | ValueNone -> ValueSome (CellValue.Number total)
        match sourceExpr with
        | BinOp (leftExpr, op, rightExpr) when op = "+" || op = "-" || op = "*" || op = "/" ->
            match rangeAccessOf model workbook leftExpr, rangeAccessOf model workbook rightExpr with
            | Some left, Some right when op <> "/" && left.Count = right.Count && left.W = right.W && left.Count > 0 ->
                let area = int64 left.Count
                let sparseEligible = area >= 1024L && int64 left.M.Cells.Count + int64 right.M.Cells.Count <= area * 4L
                let iterOffsets emit =
                    sparseOffsets left emit
                    sparseOffsets right emit
                collect
                    left.Count
                    sparseEligible
                    iterOffsets
                    (fun offset -> applyBin op (rangeRead left offset) (rangeRead right offset))
                    (applyBin op CellValue.Empty CellValue.Empty)
            | Some source, None when source.Count > 0 && not (isArrayExpr rightExpr) ->
                let scalar = eval rightExpr ctx
                let area = int64 source.Count
                let sparseEligible = area >= 1024L && int64 source.M.Cells.Count <= area * 4L
                match op, coerceNumber scalar with
                | "/", Ok denominator when denominator <> 0.0 ->
                    collect source.Count sparseEligible (sparseOffsets source) (fun offset -> applyBin op (rangeRead source offset) scalar) (applyBin op CellValue.Empty scalar)
                | ("+" | "-" | "*"), _ ->
                    collect source.Count sparseEligible (sparseOffsets source) (fun offset -> applyBin op (rangeRead source offset) scalar) (applyBin op CellValue.Empty scalar)
                | _ -> ValueNone
            | None, Some source when op <> "/" && source.Count > 0 && not (isArrayExpr leftExpr) ->
                let scalar = eval leftExpr ctx
                let area = int64 source.Count
                let sparseEligible = area >= 1024L && int64 source.M.Cells.Count <= area * 4L
                collect
                    source.Count
                    sparseEligible
                    (sparseOffsets source)
                    (fun offset -> applyBin op scalar (rangeRead source offset))
                    (applyBin op scalar CellValue.Empty)
            | _ -> ValueNone
        | UnaryOp (op, innerExpr) when op = "+" || op = "-" || op = "%" ->
            match rangeAccessOf model workbook innerExpr with
            | Some source when source.Count > 0 ->
                let area = int64 source.Count
                let sparseEligible = area >= 1024L && int64 source.M.Cells.Count <= area * 4L
                collect
                    source.Count
                    sparseEligible
                    (sparseOffsets source)
                    (fun offset -> applyUnary op (rangeRead source offset))
                    (applyUnary op CellValue.Empty)
            | _ -> ValueNone
        | _ -> ValueNone
    let trySparseIrrArithmetic guess sourceExpr =
        let discountedRepeated rate startIndex repeats value =
            if repeats <= 0L || value = 0.0 then
                0.0
            elif abs rate < 1e-12 then
                value * float repeats
            else
                let q = 1.0 / (1.0 + rate)
                value * (q ** float startIndex) * (1.0 - (q ** float repeats)) / (1.0 - q)
        let sparseOffsets (source: RangeAccess) (emit: int -> unit) =
            source.M.Cells
            |> Map.iter (fun addr _ ->
                match tryRangeOffsetForCount source source.Count addr with
                | ValueSome offset -> emit offset
                | ValueNone -> ())
        let collect count sparseEligible iterOffsets readAt gapValue =
            if not sparseEligible then
                ValueNone
            else
                let offsets = HashSet<int>()
                iterOffsets (fun offset -> offsets.Add offset |> ignore)
                let sortedOffsets = hashSetToArray offsets
                Array.sortInPlace sortedOffsets
                let segments = ResizeArray<struct (int64 * int64 * float)>()
                let mutable currentOffset = 0
                let mutable cashFlowCount = 0L
                let mutable firstErr = ValueNone
                let consume startOffset repeats value =
                    if repeats > 0L then
                        match unwrapValue value with
                        | CellValue.Error e ->
                            match firstErr with
                            | ValueNone -> firstErr <- ValueSome e
                            | ValueSome _ -> ()
                        | CellValue.Number n ->
                            segments.Add(struct (int64 startOffset, repeats, n))
                            cashFlowCount <- cashFlowCount + repeats
                        | CellValue.Boolean b ->
                            segments.Add(struct (int64 startOffset, repeats, if b then 1.0 else 0.0))
                            cashFlowCount <- cashFlowCount + repeats
                        | _ -> ()
                let mutable i = 0
                while ValueOption.isNone firstErr && i < sortedOffsets.Length do
                    let offset = sortedOffsets.[i]
                    if offset > currentOffset then
                        consume currentOffset (int64 (offset - currentOffset)) gapValue
                    if ValueOption.isNone firstErr then
                        consume offset 1L (readAt offset)
                    currentOffset <- offset + 1
                    i <- i + 1
                if ValueOption.isNone firstErr && currentOffset < count then
                    consume currentOffset (int64 (count - currentOffset)) gapValue
                match firstErr with
                | ValueSome _ -> ValueSome (err Err.value)
                | ValueNone when cashFlowCount < 2L -> ValueSome (err Err.num)
                | ValueNone ->
                    let npv rate =
                        let mutable total = 0.0
                        for struct (startIndex, repeats, value) in segments do
                            total <- total + discountedRepeated rate startIndex repeats value
                        total
                    let mutable r = guess
                    let mutable result = None
                    let mutable i = 0
                    while result.IsNone && i < 80 do
                        let f = npv r
                        if abs f < 1e-8 then result <- Some (polishRoot npv r)
                        else
                            let d = (npv (r + 1e-6) - f) / 1e-6
                            if abs d < 1e-14 then i <- 80 else r <- r - f / d
                        i <- i + 1
                    match result with
                    | Some r when not (Double.IsNaN r || Double.IsInfinity r) -> ValueSome (CellValue.Number r)
                    | _ -> ValueSome (err Err.num)
        match sourceExpr with
        | BinOp (leftExpr, op, rightExpr) when op = "+" || op = "-" || op = "*" || op = "/" ->
            match rangeAccessOf model workbook leftExpr, rangeAccessOf model workbook rightExpr with
            | Some left, Some right when op <> "/" && left.Count = right.Count && left.W = right.W && left.Count > 0 ->
                let area = int64 left.Count
                let sparseEligible = area >= 1024L && int64 left.M.Cells.Count + int64 right.M.Cells.Count <= area * 4L
                let iterOffsets emit =
                    sparseOffsets left emit
                    sparseOffsets right emit
                collect
                    left.Count
                    sparseEligible
                    iterOffsets
                    (fun offset -> applyBin op (rangeRead left offset) (rangeRead right offset))
                    (applyBin op CellValue.Empty CellValue.Empty)
            | Some source, None when source.Count > 0 && not (isArrayExpr rightExpr) ->
                let scalar = eval rightExpr ctx
                let area = int64 source.Count
                let sparseEligible = area >= 1024L && int64 source.M.Cells.Count <= area * 4L
                match op, coerceNumber scalar with
                | "/", Ok denominator when denominator <> 0.0 ->
                    collect source.Count sparseEligible (sparseOffsets source) (fun offset -> applyBin op (rangeRead source offset) scalar) (applyBin op CellValue.Empty scalar)
                | ("+" | "-" | "*"), _ ->
                    collect source.Count sparseEligible (sparseOffsets source) (fun offset -> applyBin op (rangeRead source offset) scalar) (applyBin op CellValue.Empty scalar)
                | _ -> ValueNone
            | None, Some source when op <> "/" && source.Count > 0 && not (isArrayExpr leftExpr) ->
                let scalar = eval leftExpr ctx
                let area = int64 source.Count
                let sparseEligible = area >= 1024L && int64 source.M.Cells.Count <= area * 4L
                collect
                    source.Count
                    sparseEligible
                    (sparseOffsets source)
                    (fun offset -> applyBin op scalar (rangeRead source offset))
                    (applyBin op scalar CellValue.Empty)
            | _ -> ValueNone
        | UnaryOp (op, innerExpr) when op = "+" || op = "-" || op = "%" ->
            match rangeAccessOf model workbook innerExpr with
            | Some source when source.Count > 0 ->
                let area = int64 source.Count
                let sparseEligible = area >= 1024L && int64 source.M.Cells.Count <= area * 4L
                collect
                    source.Count
                    sparseEligible
                    (sparseOffsets source)
                    (fun offset -> applyUnary op (rangeRead source offset))
                    (applyUnary op CellValue.Empty)
            | _ -> ValueNone
        | _ -> ValueNone
    // Shared index resolver for XLOOKUP / XMATCH. Streams the lookup range by
    // index (no list), returning the 0-based offset of the match. match_mode:
    // 0 exact, -1 exact-or-next-smaller, 1 exact-or-next-larger, 2 wildcard.
    // search_mode: 1 first→last, -1 last→first, 2/-2 binary asc/desc.
    let findIndexBySearchMode count searchMode predicate =
        let mutable found = ValueNone
        if searchMode = -1 || searchMode = -2 then
            let mutable k = count - 1
            while k >= 0 && found.IsNone do
                if predicate k then found <- ValueSome k
                else k <- k - 1
        else
            let mutable k = 0
            while k < count && found.IsNone do
                if predicate k then found <- ValueSome k
                else k <- k + 1
        match found with
        | ValueSome k -> Some k
        | ValueNone -> None

    let tryFindLastLessOrEqualSorted count readValue target =
        if count <= 0 then
            ValueSome None
        else
            let mutable lo = 0
            let mutable hi = count - 1
            let mutable hit = None
            let mutable supported = true
            while supported && lo <= hi do
                let mid = lo + (hi - lo) / 2
                match compareValuesNoError (readValue mid) target with
                | ValueSome c when c <= 0 ->
                    hit <- Some mid
                    lo <- mid + 1
                | ValueSome _ ->
                    hi <- mid - 1
                | ValueNone ->
                    supported <- false
            if supported then ValueSome hit else ValueNone

    let tryFindFirstGreaterOrEqualSorted count readValue target =
        if count <= 0 then
            ValueSome None
        else
            let mutable lo = 0
            let mutable hi = count - 1
            let mutable hit = None
            let mutable supported = true
            while supported && lo <= hi do
                let mid = lo + (hi - lo) / 2
                match compareValuesNoError (readValue mid) target with
                | ValueSome c when c >= 0 ->
                    hit <- Some mid
                    hi <- mid - 1
                | ValueSome _ ->
                    lo <- mid + 1
                | ValueNone ->
                    supported <- false
            if supported then ValueSome hit else ValueNone

    let tryFindFirstLessOrEqualDescending count readValue target =
        if count <= 0 then
            ValueSome None
        else
            let mutable lo = 0
            let mutable hi = count - 1
            let mutable hit = None
            let mutable supported = true
            while supported && lo <= hi do
                let mid = lo + (hi - lo) / 2
                match compareValuesNoError (readValue mid) target with
                | ValueSome c when c <= 0 ->
                    hit <- Some mid
                    hi <- mid - 1
                | ValueSome _ ->
                    lo <- mid + 1
                | ValueNone ->
                    supported <- false
            if supported then ValueSome hit else ValueNone

    let tryFindLastGreaterOrEqualDescending count readValue target =
        if count <= 0 then
            ValueSome None
        else
            let mutable lo = 0
            let mutable hi = count - 1
            let mutable hit = None
            let mutable supported = true
            while supported && lo <= hi do
                let mid = lo + (hi - lo) / 2
                match compareValuesNoError (readValue mid) target with
                | ValueSome c when c >= 0 ->
                    hit <- Some mid
                    lo <- mid + 1
                | ValueSome _ ->
                    hi <- mid - 1
                | ValueNone ->
                    supported <- false
            if supported then ValueSome hit else ValueNone

    let tryFindFirstEqualBinary count readValue target ascending =
        if count <= 0 then
            ValueSome None
        else
            let mutable lo = 0
            let mutable hi = count - 1
            let mutable hit = None
            let mutable supported = true
            while supported && lo <= hi do
                let mid = lo + (hi - lo) / 2
                match compareValuesNoError (readValue mid) target with
                | ValueSome 0 ->
                    hit <- Some mid
                    hi <- mid - 1
                | ValueSome c ->
                    if ascending then
                        if c < 0 then lo <- mid + 1 else hi <- mid - 1
                    else
                        if c > 0 then lo <- mid + 1 else hi <- mid - 1
                | ValueNone ->
                    supported <- false
            if supported then ValueSome hit else ValueNone

    let tryBinaryXMatchIndex (target: CellValue) (la: RangeAccess) (matchMode: int) (searchMode: int) =
        let read k = rangeRead la k
        match searchMode, matchMode with
        | 2, 0 -> tryFindFirstEqualBinary la.Count read target true
        | -2, 0 -> tryFindFirstEqualBinary la.Count read target false
        | 2, -1 -> tryFindLastLessOrEqualSorted la.Count read target
        | 2, 1 -> tryFindFirstGreaterOrEqualSorted la.Count read target
        | -2, -1 -> tryFindFirstLessOrEqualDescending la.Count read target
        | -2, 1 -> tryFindLastGreaterOrEqualDescending la.Count read target
        | _ -> ValueNone

    let xMatchIndex (target: CellValue) (la: RangeAccess) (matchMode: int) (searchMode: int) : int option =
        let cmp a b =
            match compareValuesNoError a b with
            | ValueSome c -> Some c
            | ValueNone -> None
        match tryBinaryXMatchIndex target la matchMode searchMode with
        | ValueSome hit -> hit
        | ValueNone ->
            let exactMatch () =
                match trySparseExactMatchResult target la searchMode with
                | ValueSome (Some k) -> Some k
                | ValueSome None -> None
                | ValueNone -> findIndexBySearchMode la.Count searchMode (fun k -> cmp target (rangeRead la k) = Some 0)
            match matchMode with
            | 2 ->
                match coerceText target with
                | Ok pat ->
                    let rx = wildcardRegex pat
                    findIndexBySearchMode la.Count searchMode (fun k -> match coerceText (rangeRead la k) with Ok t -> rx.IsMatch t | _ -> false)
                | _ -> None
            | -1 | 1 ->
                match exactMatch () with
                | Some k -> Some k
                | None ->
                    // No exact hit: closest value on the requested side
                    // (-1 → largest value < target; 1 → smallest value > target).
                    match trySparseClosestMatch target la matchMode searchMode with
                    | ValueSome k -> Some k
                    | ValueNone ->
                        let mutable best : (int * CellValue) option = None
                        let visit k =
                            let v = rangeRead la k
                            let onSide = match cmp v target with Some c -> (if matchMode = -1 then c < 0 else c > 0) | None -> false
                            if onSide then
                                match best with
                                | None -> best <- Some (k, v)
                                | Some (_, bv) -> if (match cmp v bv with Some c -> (if matchMode = -1 then c > 0 else c < 0) | None -> false) then best <- Some (k, v)
                        if searchMode = -1 || searchMode = -2 then
                            for k = la.Count - 1 downto 0 do visit k
                        else
                            for k in 0 .. la.Count - 1 do visit k
                        best |> Option.map fst
            | _ ->
                exactMatch ()
    let holidaySetFromArg arg =
        let result = HashSet<float>()
        let addHoliday holiday =
            result.Add(floor holiday) |> ignore
        match trySparseOrderedArithmeticValues arg with
        | ValueSome (Ok values) ->
            for struct (holiday, _, _) in values do
                addHoliday holiday
        | _ ->
            match numbersOfArgs [ arg ] ctx with
            | Ok holidays ->
                for holiday in holidays do
                    addHoliday holiday
            | _ -> ()
        result
    match fnName with
    // ----- aggregates (streaming one-pass folds; see foldArgs / numAggFold) --
    | "SUM" ->
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone ->
                match trySparseMapAggregate () with
                | ValueSome result -> result
                | ValueNone ->
                    match trySequenceAggregate () with
                    | ValueSome result -> result
                    | ValueNone ->
                        match trySequenceStreamAggregate () with
                        | ValueSome result -> result
                        | ValueNone ->
                            match tryRandArrayAggregate () with
                            | ValueSome result -> result
                            | ValueNone ->
                                match tryMakeArrayAggregate () with
                                | ValueSome result -> result
                                | ValueNone ->
                                    match trySparseSortPassthroughAggregate () with
                                    | ValueSome result -> result
                                    | ValueNone ->
                                        match trySparseTransposePassthroughAggregate () with
                                        | ValueSome result -> result
                                        | ValueNone ->
                                            match trySparseFilterAggregate () with
                                            | ValueSome result -> result
                                            | ValueNone ->
                                                match trySparseByRowAggregate () with
                                                | ValueSome result -> result
                                                | ValueNone ->
                                                    match trySparseScanAggregate () with
                                                    | ValueSome result -> result
                                                    | ValueNone -> foldNumericArgs (numAggFold 0.0 (+) (fun _ a -> CellValue.Number a)) args ctx
    | "PRODUCT" ->
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone ->
                match trySequenceStreamAggregate () with
                | ValueSome result -> result
                | ValueNone ->
                    match tryRandArrayAggregate () with
                    | ValueSome result -> result
                    | ValueNone ->
                        match tryMakeArrayAggregate () with
                        | ValueSome result -> result
                        | ValueNone ->
                            match trySparseSortPassthroughAggregate () with
                            | ValueSome result -> result
                            | ValueNone ->
                                match trySparseTransposePassthroughAggregate () with
                                | ValueSome result -> result
                                | ValueNone ->
                                    match trySparseFilterAggregate () with
                                    | ValueSome result -> result
                                    | ValueNone -> foldNumericArgs (numAggFold 1.0 (fun acc x -> flushProductUnderflow (acc * x)) (fun n a -> CellValue.Number (if n = 0 then 0.0 else a))) args ctx
    | "AVERAGE" ->
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone ->
                match trySparseMapAggregate () with
                | ValueSome result -> result
                | ValueNone ->
                    match trySequenceAggregate () with
                    | ValueSome result -> result
                    | ValueNone ->
                        match trySequenceStreamAggregate () with
                        | ValueSome result -> result
                        | ValueNone ->
                            match tryRandArrayAggregate () with
                            | ValueSome result -> result
                            | ValueNone ->
                                match tryMakeArrayAggregate () with
                                | ValueSome result -> result
                                | ValueNone ->
                                    match trySparseSortPassthroughAggregate () with
                                    | ValueSome result -> result
                                    | ValueNone ->
                                        match trySparseTransposePassthroughAggregate () with
                                        | ValueSome result -> result
                                        | ValueNone ->
                                            match trySparseFilterAggregate () with
                                            | ValueSome result -> result
                                            | ValueNone ->
                                                match trySparseScanAggregate () with
                                                | ValueSome result -> result
                                                | ValueNone -> foldNumericArgs (numAggFold 0.0 (+) (fun n a -> if n = 0 then err Err.div0 else CellValue.Number (a / float n))) args ctx
    | "MIN" ->
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone ->
                match trySparseMapAggregate () with
                | ValueSome result -> result
                | ValueNone ->
                    match trySequenceAggregate () with
                    | ValueSome result -> result
                    | ValueNone ->
                        match trySequenceStreamAggregate () with
                        | ValueSome result -> result
                        | ValueNone ->
                            match tryRandArrayAggregate () with
                            | ValueSome result -> result
                            | ValueNone ->
                                match tryMakeArrayAggregate () with
                                | ValueSome result -> result
                                | ValueNone ->
                                    match trySparseSortPassthroughAggregate () with
                                    | ValueSome result -> result
                                    | ValueNone ->
                                        match trySparseTransposePassthroughAggregate () with
                                        | ValueSome result -> result
                                        | ValueNone ->
                                            match trySparseFilterAggregate () with
                                            | ValueSome result -> result
                                            | ValueNone ->
                                                match trySparseScanAggregate () with
                                                | ValueSome result -> result
                                                | ValueNone -> foldNumericArgs (numAggFold infinity min (fun n a -> CellValue.Number (if n = 0 then 0.0 else a))) args ctx
    | "MAX" ->
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone ->
                match trySparseMapAggregate () with
                | ValueSome result -> result
                | ValueNone ->
                    match trySequenceAggregate () with
                    | ValueSome result -> result
                    | ValueNone ->
                        match trySequenceStreamAggregate () with
                        | ValueSome result -> result
                        | ValueNone ->
                            match tryRandArrayAggregate () with
                            | ValueSome result -> result
                            | ValueNone ->
                                match tryMakeArrayAggregate () with
                                | ValueSome result -> result
                                | ValueNone ->
                                    match trySparseSortPassthroughAggregate () with
                                    | ValueSome result -> result
                                    | ValueNone ->
                                        match trySparseTransposePassthroughAggregate () with
                                        | ValueSome result -> result
                                        | ValueNone ->
                                            match trySparseFilterAggregate () with
                                            | ValueSome result -> result
                                            | ValueNone ->
                                                match trySparseScanAggregate () with
                                                | ValueSome result -> result
                                                | ValueNone -> foldNumericArgs (numAggFold (-infinity) max (fun n a -> CellValue.Number (if n = 0 then 0.0 else a))) args ctx
    | "MEDIAN" ->
        match args with
        | [ sourceExpr ] ->
            match numericBufferOfArithmeticValue sourceExpr with
            | ValueSome (Result.Error e) -> err e
            | ValueSome (Ok (struct (values, count))) -> medianSelected values count
            | ValueNone ->
                match trySparseOrderedArithmeticValues sourceExpr with
                | ValueSome (Ok values) -> medianCompressed values
                | ValueSome (Result.Error e) -> err e
                | ValueNone ->
                    match rangeBounds sourceExpr with
                    | Some _ ->
                        match numericBufferOfRangeValue sourceExpr with
                        | Result.Error e -> err e
                        | Ok (struct (values, count)) -> medianSelected values count
                    | None ->
                        match numbersArrayOfArgs args ctx with
                        | Result.Error e -> err e
                        | Ok values -> medianSelected values values.Length
        | _ ->
            match numbersArrayOfArgs args ctx with
            | Result.Error e -> err e
            | Ok values -> medianSelected values values.Length
    | "COUNT" ->
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone ->
                match trySparseMapAggregate () with
                | ValueSome result -> result
                | ValueNone ->
                    match trySequenceAggregate () with
                    | ValueSome result -> result
                    | ValueNone ->
                        match trySequenceStreamAggregate () with
                        | ValueSome result -> result
                        | ValueNone ->
                            match tryRandArrayAggregate () with
                            | ValueSome result -> result
                            | ValueNone ->
                                match tryMakeArrayAggregate () with
                                | ValueSome result -> result
                                | ValueNone ->
                                    match trySparseSortPassthroughAggregate () with
                                    | ValueSome result -> result
                                    | ValueNone ->
                                        match trySparseTransposePassthroughAggregate () with
                                        | ValueSome result -> result
                                        | ValueNone ->
                                            match trySparseFilterAggregate () with
                                            | ValueSome result -> result
                                            | ValueNone ->
                                                match trySparseByRowAggregate () with
                                                | ValueSome result -> result
                                                | ValueNone ->
                                                    match trySparseScanAggregate () with
                                                    | ValueSome result -> result
                                                    | ValueNone -> foldArgsIgnoringMissingCellsWith directCountArgument (countFold true (fun v -> match v with CellValue.Number _ -> true | _ -> false)) args ctx
    | "COUNTA" ->
        let mutable fastCount = 0
        let mutable slowArgs = []
        for arg in args do
            match arithmeticArrayNonBlankCount model workbook arg with
            | ValueSome count -> fastCount <- fastCount + count
            | ValueNone -> slowArgs <- arg :: slowArgs
        let slowCount =
            match slowArgs with
            | [] -> 0.0
            | _ ->
                match foldArgsIgnoringMissingCells (countFold false (fun v -> match v with CellValue.Empty -> false | _ -> true)) (List.rev slowArgs) ctx with
                | CellValue.Number n -> n
                | _ -> 0.0
        CellValue.Number (float fastCount + slowCount)
    | "COUNTBLANK" -> countBlankArgs args ctx
    | "STDEV" | "STDEV.S" | "STDEVP" | "STDEV.P" ->
        let sample = fnName = "STDEV" || fnName = "STDEV.S"
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone -> foldNumericArgs (varianceFold sample true) args ctx
    | "VAR" | "VAR.S" | "VARP" | "VAR.P" ->
        let sample = fnName = "VAR" || fnName = "VAR.S"
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone -> foldNumericArgs (varianceFold sample false) args ctx
    | "SUBTOTAL" | "AGGREGATE" ->
        // SUBTOTAL(fnum, ref…) / AGGREGATE(fnum, options, ref…). Aggregate
        // over the refs while (a) skipping rows hidden by a filter — and, for
        // SUBTOTAL 101–111 or AGGREGATE options 1/3/5/7, manually-hidden rows —
        // and (b) skipping cells that are themselves SUBTOTAL/AGGREGATE so
        // nested subtotals aren't double-counted. AGGREGATE additionally drops
        // error cells under options 2/3/6/7.
        let isAggregate = fnName = "AGGREGATE"
        let refStart = if isAggregate then 2 else 1
        match argNum 0 with
        | Result.Error e -> err e
        | Ok fnf ->
            let fnRaw = int fnf
            let baseFn = if fnRaw >= 101 then fnRaw - 100 else fnRaw
            let opts = if isAggregate then (match argNum 1 with Ok o -> int o | _ -> 0) else 0
            let ignoreHidden =
                if isAggregate then (opts = 1 || opts = 3 || opts = 5 || opts = 7)
                else fnRaw >= 101
            let ignoreErrors = isAggregate && (opts = 2 || opts = 3 || opts = 6 || opts = 7)
            let skipNested = not isAggregate || opts <= 3
            let hasFilters = not (Map.isEmpty model.Filters)
            let mutable filterColsCache: int array option = None
            let mutable filterValuesCache: Set<string> array option = None
            // (The rows a filter hides are kept in the model: they are what is skipped.)
            let hasFilterProjection = hasFilters
            let filterCriteria () =
                match filterColsCache, filterValuesCache with
                | Some filterCols, Some filterValues -> filterCols, filterValues
                | _ ->
                    let filterEntries = model.Filters |> Map.toArray
                    let filterCols = filterEntries |> Array.map fst
                    let filterValues = filterEntries |> Array.map (fun _ -> Set.empty<string>)
                    filterColsCache <- Some filterCols
                    filterValuesCache <- Some filterValues
                    filterCols, filterValues
            let filterIndexFor (filterCols: int array) col =
                match filterCols.Length with
                | 1 ->
                    if col = filterCols.[0] then 0 else -1
                | 2 ->
                    if col = filterCols.[0] then 0
                    elif col = filterCols.[1] then 1
                    else -1
                | 3 ->
                    if col = filterCols.[0] then 0
                    elif col = filterCols.[1] then 1
                    elif col = filterCols.[2] then 2
                    else -1
                | _ ->
                    let mutable lo = 0
                    let mutable hi = filterCols.Length
                    while lo < hi do
                        let mid = lo + (hi - lo) / 2
                        if filterCols.[mid] < col then lo <- mid + 1 else hi <- mid
                    if lo < filterCols.Length && filterCols.[lo] = col then lo else -1
            let rowValuesPassActiveFilters (rowValues: CellValue array) (filterValues: Set<string> array) =
                let mutable passes = true
                let mutable i = 0
                while passes && i < rowValues.Length do
                    passes <- Set.contains (displayValue rowValues.[i]) filterValues.[i]
                    i <- i + 1
                passes
            let filterEmptyAllowances (filterValues: Set<string> array) =
                let allowsEmpty = Array.zeroCreate<bool> filterValues.Length
                let mutable rejectedCount = 0
                for i = 0 to filterValues.Length - 1 do
                    let allowed = Set.contains "" filterValues.[i]
                    allowsEmpty.[i] <- allowed
                    if not allowed then
                        rejectedCount <- rejectedCount + 1
                allowsEmpty, rejectedCount
            let touchedRowValuesPassActiveFilters
                (rowValues: CellValue array)
                (filterValues: Set<string> array)
                (touchedIndexes: ResizeArray<int>)
                missingRejectedCount =
                let mutable passes = missingRejectedCount = 0
                let mutable i = 0
                while passes && i < touchedIndexes.Count do
                    let filterIndex = touchedIndexes.[i]
                    passes <- Set.contains (displayValue rowValues.[filterIndex]) filterValues.[filterIndex]
                    i <- i + 1
                passes
            let filterColumnsAreContiguous (filterCols: int array) =
                let mutable contiguous = filterCols.Length > 0
                let mutable i = 1
                while contiguous && i < filterCols.Length do
                    contiguous <- filterCols.[i] = filterCols.[i - 1] + 1
                    i <- i + 1
                contiguous
            let filterIndexForScan contiguous firstCol lastCol (filterCols: int array) col =
                if contiguous then
                    if col >= firstCol && col <= lastCol then col - firstCol else -1
                else
                    filterIndexFor filterCols col
            let rowPassesActiveFilters (r: int) =
                let filterCols, filterValues = filterCriteria ()
                let mutable passes = true
                let mutable i = 0
                while passes && i < filterCols.Length do
                    let value = getCellValue { Row = r; Col = filterCols.[i] } model
                    passes <- Set.contains (displayValue value) filterValues.[i]
                    i <- i + 1
                passes
            let rowOk useFilterProjection (r: int) =
                (not hasFilters
                 || (if useFilterProjection then not (isFilterHiddenRow r model) else rowPassesActiveFilters r))
                && (not ignoreHidden || getRowHeight r model > 0.0)
            let nested (r: int) (c: int) =
                skipNested &&
                (match getCell { Row = r; Col = c } model with
                 | Some cd ->
                     match cd.Value with
                     | CellValue.Formula (ftext, _) ->
                         formulaTextMayUseSubtotal ftext
                         && (match parseFormula ftext with Ok e -> exprUsesSubtotal e | _ -> false)
                     | _ -> false
                 | None -> false)
            // Cell coordinates covered by one reference argument.
            let argArea (a: Expr) : (int * int * int * int) option =
                match rangeBounds a with
                | Some b -> Some b
                | None ->
                    match a with
                    | TableRef sel ->
                        match tableSelArea sel workbook ctx.Sheet ctx.CurrentCell with
                        | Result.Ok (r0, c0, r1, c1) -> Some (r0, c0, r1, c1)
                        | Result.Error _ -> None
                    | _ -> None
            let iterReferenceCells (r0, c0, r1, c1) (addValue: CellValue -> unit) =
                let area = int64 (r1 - r0 + 1) * int64 (c1 - c0 + 1)
                let useFilterProjection = hasFilterProjection
                if area >= 1024L && int64 model.Cells.Count <= area * 4L && hasFilters && not useFilterProjection then
                    let filterCols, filterValues = filterCriteria ()
                    let singleValueColumn = c0 = c1
                    let rowCells =
                        if singleValueColumn then
                            Unchecked.defaultof<ResizeArray<int * CellValue>>
                        else
                            ResizeArray<int * CellValue>()
                    let mutable hasSingleRowCell = false
                    let mutable singleRowCellValue = CellValue.Empty
                    let mutable currentRow = System.Int32.MinValue
                    let mutable hasRow = false
                    let clearRowCells () =
                        if singleValueColumn then
                            hasSingleRowCell <- false
                        else
                            rowCells.Clear()
                    let addTargetCell col value =
                        if col >= c0 && col <= c1 then
                            if singleValueColumn then
                                hasSingleRowCell <- true
                                singleRowCellValue <- value
                            else
                                rowCells.Add(col, value)
                    let addPassedRowValues () =
                        if singleValueColumn then
                            if hasSingleRowCell && not (nested currentRow c0) then
                                addValue (unwrapValue singleRowCellValue)
                        else
                            for col, value in rowCells do
                                if not (nested currentRow col) then
                                    addValue (unwrapValue value)
                    if filterCols.Length <= 3 then
                        let rowValues = Array.create filterCols.Length CellValue.Empty
                        let finishRow () =
                            if hasRow
                               && rowValuesPassActiveFilters rowValues filterValues
                               && (not ignoreHidden || getRowHeight currentRow model > 0.0) then
                                addPassedRowValues ()
                        model.Cells
                        |> Map.iter (fun addr cell ->
                            if addr.Row >= r0 && addr.Row <= r1 then
                                if addr.Row <> currentRow then
                                    finishRow()
                                    currentRow <- addr.Row
                                    hasRow <- true
                                    Array.fill rowValues 0 rowValues.Length CellValue.Empty
                                    clearRowCells ()
                                let filterIndex = filterIndexFor filterCols addr.Col
                                if filterIndex >= 0 then
                                    rowValues.[filterIndex] <- cell.Value
                                addTargetCell addr.Col cell.Value)
                        finishRow()
                    else
                        let rowValues = Array.create filterCols.Length CellValue.Empty
                        let touchedIndexes = ResizeArray<int>()
                        let emptyAllowed, emptyRejectedCount = filterEmptyAllowances filterValues
                        let mutable missingRejectedCount = emptyRejectedCount
                        let contiguousFilters = filterColumnsAreContiguous filterCols
                        let firstFilterCol = filterCols.[0]
                        let lastFilterCol = filterCols.[filterCols.Length - 1]
                        let finishRow () =
                            if hasRow
                               && touchedRowValuesPassActiveFilters rowValues filterValues touchedIndexes missingRejectedCount
                               && (not ignoreHidden || getRowHeight currentRow model > 0.0) then
                                addPassedRowValues ()
                        model.Cells
                        |> Map.iter (fun addr cell ->
                            if addr.Row >= r0 && addr.Row <= r1 then
                                if addr.Row <> currentRow then
                                    finishRow()
                                    currentRow <- addr.Row
                                    hasRow <- true
                                    touchedIndexes.Clear()
                                    missingRejectedCount <- emptyRejectedCount
                                    clearRowCells ()
                                let filterIndex = filterIndexForScan contiguousFilters firstFilterCol lastFilterCol filterCols addr.Col
                                if filterIndex >= 0 then
                                    touchedIndexes.Add filterIndex
                                    if not emptyAllowed.[filterIndex] then
                                        missingRejectedCount <- missingRejectedCount - 1
                                    rowValues.[filterIndex] <- cell.Value
                                addTargetCell addr.Col cell.Value)
                        finishRow()
                elif area >= 1024L && int64 model.Cells.Count <= area * 4L then
                    if useFilterProjection then
                        let hiddenRows = model.FilterHiddenRows
                        let mutable hiddenIndex = 0
                        let rowVisibleByProjection row =
                            while hiddenIndex < hiddenRows.Length && hiddenRows.[hiddenIndex] < row do
                                hiddenIndex <- hiddenIndex + 1
                            hiddenIndex >= hiddenRows.Length || hiddenRows.[hiddenIndex] <> row

                        model.Cells
                        |> Map.iter (fun addr cell ->
                            if addr.Row >= r0 && addr.Row <= r1 && addr.Col >= c0 && addr.Col <= c1
                               && rowVisibleByProjection addr.Row
                               && (not ignoreHidden || getRowHeight addr.Row model > 0.0)
                               && not (nested addr.Row addr.Col) then
                                addValue (unwrapValue cell.Value))
                    else
                        model.Cells
                        |> Map.iter (fun addr cell ->
                            if addr.Row >= r0 && addr.Row <= r1 && addr.Col >= c0 && addr.Col <= c1
                               && rowOk useFilterProjection addr.Row && not (nested addr.Row addr.Col) then
                                addValue (unwrapValue cell.Value))
                else
                    for r in r0 .. r1 do
                        if rowOk useFilterProjection r then
                            for c in c0 .. c1 do
                                if not (nested r c) then
                                    addValue (gridCell model r c)
            let addReferenceArg (addValue: CellValue -> unit) a =
                match argArea a with
                | Some area -> iterReferenceCells area addValue
                | None -> addValue (eval a ctx)
            if baseFn = 9 then
                let mutable total = 0.0
                let mutable firstErr = ValueNone
                let addSumValue value =
                    if ValueOption.isNone firstErr then
                        match unwrapValue value with
                        | CellValue.Error e when ignoreErrors -> ()
                        | CellValue.Error e -> firstErr <- ValueSome e
                        | CellValue.Number n -> total <- total + n
                        | CellValue.Boolean b -> total <- total + (if b then 1.0 else 0.0)
                        | _ -> ()
                let mutable argIndex = 0
                for a in args do
                    if argIndex >= refStart then
                        addReferenceArg addSumValue a
                    argIndex <- argIndex + 1
                match firstErr with
                | ValueSome e -> err e
                | ValueNone -> CellValue.Number total
            elif baseFn = 2 then
                let mutable count = 0
                let addCountValue value =
                    match unwrapValue value with
                    | CellValue.Number _ -> count <- count + 1
                    | _ -> ()
                let mutable argIndex = 0
                for a in args do
                    if argIndex >= refStart then
                        addReferenceArg addCountValue a
                    argIndex <- argIndex + 1
                CellValue.Number (float count)
            elif baseFn = 3 then
                let mutable count = 0
                let addCountAValue value =
                    match unwrapValue value with
                    | CellValue.Error _ when ignoreErrors -> ()
                    | CellValue.Empty -> ()
                    | _ -> count <- count + 1
                let mutable argIndex = 0
                for a in args do
                    if argIndex >= refStart then
                        addReferenceArg addCountAValue a
                    argIndex <- argIndex + 1
                CellValue.Number (float count)
            elif baseFn = 1 || baseFn = 4 || baseFn = 5 || baseFn = 6 then
                let mutable total = 0.0
                let mutable count = 0
                let mutable maxValue = -infinity
                let mutable minValue = infinity
                let mutable product = 1.0
                let mutable firstErr = ValueNone
                let addNumericValue n =
                    count <- count + 1
                    total <- total + n
                    if n > maxValue then maxValue <- n
                    if n < minValue then minValue <- n
                    product <- product * n
                let addAggregateValue value =
                    if ValueOption.isNone firstErr then
                        match unwrapValue value with
                        | CellValue.Error e when ignoreErrors -> ()
                        | CellValue.Error e -> firstErr <- ValueSome e
                        | CellValue.Number n -> addNumericValue n
                        | CellValue.Boolean b -> addNumericValue (if b then 1.0 else 0.0)
                        | _ -> ()
                let mutable argIndex = 0
                for a in args do
                    if argIndex >= refStart then
                        addReferenceArg addAggregateValue a
                    argIndex <- argIndex + 1
                match firstErr with
                | ValueSome e -> err e
                | ValueNone ->
                    match baseFn with
                    | 1 -> if count = 0 then err Err.div0 else CellValue.Number (total / float count)
                    | 4 -> CellValue.Number (if count = 0 then 0.0 else maxValue)
                    | 5 -> CellValue.Number (if count = 0 then 0.0 else minValue)
                    | 6 -> CellValue.Number (if count = 0 then 0.0 else product)
                    | _ -> err Err.value
            else
                let values = ResizeArray<CellValue>()
                let addValue value =
                    if ignoreErrors then
                        match unwrapValue value with
                        | CellValue.Error _ -> ()
                        | _ -> values.Add value
                    else
                        values.Add value
                let mutable argIndex = 0
                for a in args do
                    if argIndex >= refStart then
                        addReferenceArg addValue a
                    argIndex <- argIndex + 1
                let numbersFromValues () =
                    let nums = ResizeArray<float>()
                    let mutable firstErr = ValueNone
                    let mutable index = 0
                    while ValueOption.isNone firstErr && index < values.Count do
                        match unwrapValue values.[index] with
                        | CellValue.Error e -> firstErr <- ValueSome e
                        | CellValue.Number n -> nums.Add n
                        | CellValue.Boolean b -> nums.Add(if b then 1.0 else 0.0)
                        | _ -> ()
                        index <- index + 1
                    match firstErr with
                    | ValueSome e -> Result.Error e
                    | ValueNone -> Ok nums
                let withNums empty f =
                    match numbersFromValues () with
                    | Result.Error e -> err e
                    | Ok nums when nums.Count = 0 -> empty
                    | Ok nums -> f nums
                let variance sample (nums: ResizeArray<float>) =
                    let n = nums.Count
                    if (sample && n < 2) || n < 1 then err Err.div0
                    else
                        let mutable sum = 0.0
                        for value in nums do
                            sum <- sum + value
                        let mean = sum / float n
                        let mutable ss = 0.0
                        for value in nums do
                            ss <- ss + (value - mean) ** 2.0
                        CellValue.Number (ss / (if sample then float (n - 1) else float n))
                match baseFn with
                | 1 -> withNums (err Err.div0) (fun nums ->
                    let mutable total = 0.0
                    for value in nums do
                        total <- total + value
                    CellValue.Number (total / float nums.Count))
                | 2 ->
                    let mutable count = 0
                    for v in values do
                        match unwrapValue v with
                        | CellValue.Number _ -> count <- count + 1
                        | _ -> ()
                    CellValue.Number (float count)
                | 3 ->
                    let mutable count = 0
                    for v in values do
                        match unwrapValue v with
                        | CellValue.Empty -> ()
                        | _ -> count <- count + 1
                    CellValue.Number (float count)
                | 4 -> withNums (CellValue.Number 0.0) (fun nums ->
                    let mutable maxValue = nums.[0]
                    for i = 1 to nums.Count - 1 do
                        if nums.[i] > maxValue then maxValue <- nums.[i]
                    CellValue.Number maxValue)
                | 5 -> withNums (CellValue.Number 0.0) (fun nums ->
                    let mutable minValue = nums.[0]
                    for i = 1 to nums.Count - 1 do
                        if nums.[i] < minValue then minValue <- nums.[i]
                    CellValue.Number minValue)
                | 6 -> withNums (CellValue.Number 0.0) (fun nums ->
                    let mutable product = 1.0
                    for value in nums do
                        product <- product * value
                    CellValue.Number product)
                | 7 -> withNums (err Err.div0) (fun ns -> match variance true ns with CellValue.Number v -> CellValue.Number (sqrt v) | e -> e)
                | 8 -> withNums (err Err.div0) (fun ns -> match variance false ns with CellValue.Number v -> CellValue.Number (sqrt v) | e -> e)
                | 10 -> withNums (err Err.div0) (variance true)
                | 11 -> withNums (err Err.div0) (variance false)
                | 12 -> withNums (err Err.num) (fun nums ->
                    CellValue.Number (medianArray (nums.ToArray())))
                | _ -> err Err.value   // AGGREGATE 13–19 (MODE/LARGE/SMALL/PERCENTILE/QUARTILE) — later
    | "LARGE" | "SMALL" ->
        if nargs <> 2 then err Err.value else
        match argNum 1 with
        | Result.Error e -> err e
        | Ok k ->
            let finish values count =
                let ki = int k - 1
                if ki < 0 || ki >= count then err Err.num
                else
                    let index = if fnName = "LARGE" then count - 1 - ki else ki
                    CellValue.Number (selectNumberAtCount count index values)
            let sourceExpr = argExpr 0
            match numericBufferOfArithmeticValue sourceExpr with
            | ValueSome (Result.Error e) -> err e
            | ValueSome (Ok (struct (values, count))) -> finish values count
            | ValueNone ->
                match trySparseOrderedArithmeticValues sourceExpr with
                | ValueSome (Result.Error e) -> err e
                | ValueSome (Ok values) -> rankedCompressed (fnName = "LARGE") (int k) values
                | ValueNone ->
                    if int k = 1 then
                        match rangeAccessOf model workbook sourceExpr with
                        | Some source ->
                            match rankOneRangeValue (fnName = "LARGE") source with
                            | Ok value -> CellValue.Number value
                            | Result.Error e -> err e
                        | None ->
                            match rangeBounds sourceExpr with
                            | Some _ ->
                                match numericBufferOfRangeValue sourceExpr with
                                | Result.Error e -> err e
                                | Ok (struct (values, count)) -> finish values count
                            | None ->
                                match numbersArrayOfArgs [ sourceExpr ] ctx with
                                | Result.Error e -> err e
                                | Ok values -> finish values values.Length
                    else
                        match rangeBounds sourceExpr with
                        | Some _ ->
                            match numericBufferOfRangeValue sourceExpr with
                            | Result.Error e -> err e
                            | Ok (struct (values, count)) -> finish values count
                        | None ->
                            match numbersArrayOfArgs [ sourceExpr ] ctx with
                            | Result.Error e -> err e
                            | Ok values -> finish values values.Length
    // ----- criteria aggregates (streamed by index; see rangeAccessOf) -------
    | "COUNTIF" ->
        if nargs <> 2 then err Err.value else
        match rangeAccessOf model workbook (argExpr 0) with
        | None -> err Err.value
        | Some a ->
            let pred = criteriaPredicate (argRaw 1)
            CellValue.Number (float (countIfRangeAccess a pred))
    | "SUMIF" ->
        // SUMIF(range, criteria, [sum_range])
        if nargs < 2 then err Err.value else
        match rangeAccessOf model workbook (argExpr 0) with
        | None -> err Err.value
        | Some crit ->
            let sum =
                if nargs >= 3 then (match rangeAccessOf model workbook (argExpr 2) with Some s -> s | None -> crit)
                else crit
            let pred = criteriaPredicate (argRaw 1)
            sumIfRangeAccess crit sum pred
    | "AVERAGEIF" ->
        if nargs < 2 then err Err.value else
        match rangeAccessOf model workbook (argExpr 0) with
        | None -> err Err.value
        | Some crit ->
            let avg =
                if nargs >= 3 then (match rangeAccessOf model workbook (argExpr 2) with Some s -> s | None -> crit)
                else crit
            let pred = criteriaPredicate (argRaw 1)
            averageIfRangeAccess crit avg pred
    | "SUMIFS" | "COUNTIFS" ->
        // SUMIFS(sum_range, crit_range1, crit1, crit_range2, crit2, …)
        // COUNTIFS(crit_range1, crit1, …)  — no sum range. Streamed by index:
        // each (crit range, predicate) pair is read in lockstep, no lists.
        let isSum = fnName = "SUMIFS"
        let pairsStart = if isSum then 1 else 0
        let pairCount = nargs - pairsStart
        if pairCount <= 0 || pairCount % 2 <> 0 then err Err.value else
        let crits = ResizeArray<RangeAccess * (CellValue -> bool)>(pairCount / 2)
        let mutable invalidCriteria = false
        let mutable count = System.Int32.MaxValue
        let mutable argIndex = pairsStart
        while argIndex < nargs do
            match rangeAccessOf model workbook (argExpr argIndex) with
            | Some access ->
                crits.Add(access, criteriaPredicate (eval (argExpr (argIndex + 1)) ctx))
                if access.Count < count then count <- access.Count
            | None -> invalidCriteria <- true
            argIndex <- argIndex + 2
        if invalidCriteria then err Err.value else
        let sumAcc = if isSum then rangeAccessOf model workbook (argExpr 0) else None
        let sparseResult =
            match trySparseIfs crits count sumAcc with
            | ValueSome result -> ValueSome result
            | ValueNone -> trySparseBlankIfs crits count sumAcc
        match sparseResult with
        | ValueSome (total, hits) ->
            if isSum then CellValue.Number total else CellValue.Number (float hits)
        | ValueNone ->
            let mutable total = 0.0
            let mutable hits = 0
            let mutable firstError = None
            let critReaders =
                Array.init crits.Count (fun i ->
                    let access, pred = crits.[i]
                    if pred CellValue.Empty then null else rangeReadCached access count)
            let sumReader =
                match sumAcc with
                | Some s -> rangeReadCached s count
                | None -> null
            for k in 0 .. count - 1 do
                let mutable matched = true
                let mutable critIndex = 0
                while matched && critIndex < crits.Count do
                    let access, pred = crits.[critIndex]
                    let reader = critReaders.[critIndex]
                    let value = if isNull reader then rangeRead access k else reader.Read k
                    if not (pred value) then
                        matched <- false
                    critIndex <- critIndex + 1
                if matched then
                    hits <- hits + 1
                    match sumAcc with
                    | Some s when k < s.Count ->
                        match unwrapValue (sumReader.Read k) with
                        | CellValue.Error e -> if firstError.IsNone then firstError <- Some e
                        | value -> (match tryAggregateNumber value with ValueSome n -> total <- total + n | ValueNone -> ())
                    | _ -> ()
            // (An error among the cells added is the result, as SUMIF's.)
            match firstError with
            | Some e when isSum -> err e
            | _ -> if isSum then CellValue.Number total else CellValue.Number (float hits)
    | "MAXIFS" | "MINIFS" | "AVERAGEIFS" ->
        // MAXIFS(value_range, crit_range1, crit1, …); same shape as SUMIFS.
        // One streamed pass tracks the matched count + running min/max/sum.
        let fn = fnName
        let pairCount = nargs - 1
        if pairCount <= 0 || pairCount % 2 <> 0 then err Err.value else
        match rangeAccessOf model workbook (argExpr 0) with
        | Some valAcc ->
            let crits = ResizeArray<RangeAccess * (CellValue -> bool)>(pairCount / 2)
            let mutable invalidCriteria = false
            let mutable count = valAcc.Count
            let mutable argIndex = 1
            while argIndex < nargs do
                match rangeAccessOf model workbook (argExpr argIndex) with
                | Some access ->
                    crits.Add(access, criteriaPredicate (eval (argExpr (argIndex + 1)) ctx))
                    if access.Count < count then count <- access.Count
                | None -> invalidCriteria <- true
                argIndex <- argIndex + 2
            if invalidCriteria then err Err.value else
            let n, acc =
                let sparseResult =
                    match trySparseValueIfs fn valAcc crits count with
                    | ValueSome result -> ValueSome result
                    | ValueNone -> trySparseBlankValueIfs fn valAcc crits count
                match sparseResult with
                | ValueSome result -> result
                | ValueNone ->
                    let mutable n = 0
                    let mutable acc = if fn = "MINIFS" then infinity elif fn = "MAXIFS" then -infinity else 0.0
                    let valueReader = rangeReadCached valAcc count
                    let critReaders =
                        Array.init crits.Count (fun i ->
                            let access, pred = crits.[i]
                            if pred CellValue.Empty then null else rangeReadCached access count)
                    for k in 0 .. count - 1 do
                        let mutable matched = true
                        let mutable critIndex = 0
                        while matched && critIndex < crits.Count do
                            let access, pred = crits.[critIndex]
                            let reader = critReaders.[critIndex]
                            let value = if isNull reader then rangeRead access k else reader.Read k
                            if not (pred value) then
                                matched <- false
                            critIndex <- critIndex + 1
                        if matched then
                            match tryAggregateNumber (valueReader.Read k) with
                            | ValueSome x ->
                                n <- n + 1
                                acc <- match fn with "MINIFS" -> min acc x | "MAXIFS" -> max acc x | _ -> acc + x
                            | ValueNone -> ()
                    n, acc
            match fn with
            | "AVERAGEIFS" -> if n = 0 then err Err.div0 else CellValue.Number (acc / float n)
            | _ -> CellValue.Number (if n = 0 then 0.0 else acc)
        | _ -> err Err.value
    | "SUMPRODUCT" ->
        // Element-wise product of equally shaped arrays, then sum. Each
        // argument is evaluated as an array, so expressions such as
        // `(C1:C5>3)*C1:C5` or `--(D1:D5="fig")` work; entries that are not
        // numbers (text, blanks, booleans) count as 0 and keep their position,
        // which keeps parallel arrays aligned.
        if List.isEmpty args then err Err.value else
        let arrays = args |> List.map (fun arg -> evalArray arg ctx) |> List.toArray
        let rows = arrays.[0].Length
        let cols = if rows > 0 then arrays.[0].[0].Length else 0
        if arrays |> Array.exists (fun arr -> arr.Length <> rows || (rows > 0 && arr.[0].Length <> cols)) then
            // A 1×1 error argument reports itself; any other mismatch is #VALUE!.
            let firstError =
                arrays |> Array.tryPick (fun arr ->
                    if arr.Length = 1 && arr.[0].Length = 1 then
                        (match unwrapValue arr.[0].[0] with CellValue.Error e -> Some e | _ -> None)
                    else None)
            err (defaultArg firstError Err.value)
        else
            let mutable total = 0.0
            let mutable firstErr = ValueNone
            for r in 0 .. rows - 1 do
                for c in 0 .. cols - 1 do
                    let mutable product = 1.0
                    for arr in arrays do
                        match unwrapValue arr.[r].[c] with
                        | CellValue.Number n -> product <- flushProductUnderflow (product * n)
                        | CellValue.Error e -> (if ValueOption.isNone firstErr then firstErr <- ValueSome e)
                        | _ -> product <- 0.0
                    total <- total + product
            match firstErr with
            | ValueSome e -> err e
            | ValueNone -> CellValue.Number total
    // ----- math -------------------------------------------------------------
    | "ABS" -> unaryNum (abs >> CellValue.Number)
    | "SQRT" -> unaryNum (fun x -> if x < 0.0 then err Err.num else CellValue.Number (sqrt x))
    | "SIGN" -> unaryNum (fun x -> CellValue.Number (float (sign x)))
    | "INT" -> unaryNum (fun x -> CellValue.Number (floor x))
    | "TRUNC" ->
        match argNum 0 with
        | Result.Error e -> err e
        | Ok x ->
            let digits = if nargs >= 2 then (match argNum 1 with Ok d -> int d | _ -> 0) else 0
            let factor = 10.0 ** float digits
            CellValue.Number (Math.Truncate(x * factor) / factor)
    | "ROUND" | "ROUNDUP" | "ROUNDDOWN" ->
        if nargs < 1 then err Err.value else
        match argNum 0, (if nargs >= 2 then argNum 1 else Ok 0.0) with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok x, Ok d ->
            let factor = 10.0 ** d
            let scaled = x * factor
            let rounded =
                match fnName with
                | "ROUNDUP" -> (if scaled < 0.0 then floor scaled else ceil scaled)
                | "ROUNDDOWN" -> Math.Truncate scaled
                | _ -> roundHalfAway scaled
            CellValue.Number (rounded / factor)
    | "MOD" ->
        if nargs <> 2 then err Err.value else
        match argNum 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok _, Ok d when d = 0.0 -> err Err.div0
        | Ok a, Ok d ->
            // Excel MOD takes the sign of the divisor.
            let r = a - d * floor (a / d)
            CellValue.Number r
    | "POWER" ->
        if nargs <> 2 then err Err.value else
        match argNum 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok b, Ok p -> powerProduct b p
    // ----- extended math ----------------------------------------------------
    | "PI" -> CellValue.Number Math.PI
    | "EXP" -> unaryNum (exp >> CellValue.Number)
    | "LN" -> unaryNum (fun x -> if x <= 0.0 then err Err.num else CellValue.Number (log x))
    | "LOG10" -> unaryNum (fun x -> if x <= 0.0 then err Err.num else CellValue.Number (log10 x))
    | "LOG" ->
        match argNum 0, (if nargs >= 2 then argNum 1 else Ok 10.0) with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok x, Ok b -> if x <= 0.0 || b <= 0.0 || b = 1.0 then err Err.num else CellValue.Number (log x / log b)
    | "SIN" -> unaryNum (sin >> CellValue.Number)
    | "COS" -> unaryNum (cos >> CellValue.Number)
    | "TAN" -> unaryNum (tan >> CellValue.Number)
    | "RADIANS" -> unaryNum (fun d -> CellValue.Number (d * Math.PI / 180.0))
    | "DEGREES" -> unaryNum (fun r -> CellValue.Number (r * 180.0 / Math.PI))
    | "QUOTIENT" ->
        match argNum 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok a, Ok b -> if b = 0.0 then err Err.div0 else CellValue.Number (Math.Truncate(a / b))
    | "CEILING" | "FLOOR" ->
        let up = fnName = "CEILING"
        match argNum 0, (if nargs >= 2 then argNum 1 else Ok 1.0) with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok x, Ok sg -> if sg = 0.0 then CellValue.Number 0.0 else CellValue.Number ((if up then ceil (x / sg) else floor (x / sg)) * sg)
    | "MROUND" ->
        match argNum 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok x, Ok m -> if m = 0.0 then CellValue.Number 0.0 else CellValue.Number (roundHalfAway (x / m) * m)
    | "EVEN" -> unaryNum (fun x -> let n = ceil (abs x / 2.0) * 2.0 in CellValue.Number (if x < 0.0 then -n else n))
    | "ODD" -> unaryNum (fun x -> let n = 2.0 * ceil ((abs x - 1.0) / 2.0) + 1.0 in CellValue.Number (if x < 0.0 then -n else n))
    | "SUMSQ" ->
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone -> foldNumericArgs (numAggFold 0.0 (fun acc x -> acc + x * x) (fun _ sum -> CellValue.Number sum)) args ctx
    | "FACT" ->
        match argNum 0 with
        | Result.Error e -> err e
        | Ok x ->
            let n = int (floor x)
            if n < 0 then err Err.num
            else (let mutable p = 1.0 in (for i in 2 .. n do p <- p * float i); CellValue.Number p)
    | "COMBIN" ->
        match argNum 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok nf, Ok kf ->
            let n, k = int nf, int kf
            if n < 0 || k < 0 || k > n then err Err.num
            else (let mutable r = 1.0 in (for i in 0 .. k - 1 do r <- r * float (n - i) / float (i + 1)); CellValue.Number (Math.Round r))
    | "GCD" | "LCM" ->
        if fnName = "GCD" then
            foldArgsIgnoringMissingCells (integerAggFold 0L gcdInt64) args ctx
        else
            foldArgsIgnoringMissingCells (integerAggFold 1L (fun acc x -> if acc = 0L || x = 0L then 0L else acc / gcdInt64 acc x * x)) args ctx
    // ----- extended stats ---------------------------------------------------
    | "AVERAGEA" | "MAXA" | "MINA" ->
        // A-variants count text as 0 and booleans as 1/0; blanks are ignored.
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone -> foldArgsIgnoringMissingCells (aVariantFold fnName) args ctx
    | "MODE" | "MODE.SNGL" ->
        match args with
        | [ sourceExpr ] ->
            match numericBufferOfArithmeticValue sourceExpr with
            | ValueSome (Result.Error e) -> err e
            | ValueSome (Ok (struct (values, count))) -> modeSelected values count
            | ValueNone ->
                match trySparseOrderedArithmeticValues sourceExpr with
                | ValueSome (Ok values) -> modeCompressed values
                | ValueSome (Result.Error e) -> err e
                | ValueNone -> foldNumericArgs (modeFold ()) args ctx
        | _ -> foldNumericArgs (modeFold ()) args ctx
    | "DEVSQ" ->
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone -> foldNumericArgs devSqFold args ctx
    | "AVEDEV" ->
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone -> foldNumericArgs (aveDevFold ()) args ctx
    | "GEOMEAN" ->
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone -> foldNumericArgs geoMeanFold args ctx
    | "HARMEAN" ->
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone -> foldNumericArgs harmonicMeanFold args ctx
    | "RANK" | "RANK.EQ" ->
        if nargs < 2 then err Err.value else
        match argNum 0 with
        | Result.Error e -> err e
        | Ok x ->
            let order = if nargs >= 3 then (match argNum 2 with Ok o -> int o | _ -> 0) else 0
            let sourceExpr = argExpr 1
            match numericBufferOfArithmeticValue sourceExpr with
            | ValueSome (Result.Error e) -> err e
            | ValueSome (Ok (struct (values, count))) -> rankSelected x (order = 0) values count
            | ValueNone ->
                match trySparseOrderedArithmeticValues sourceExpr with
                | ValueSome (Ok values) -> rankCompressed x (order = 0) values
                | ValueSome (Result.Error e) -> err e
                | ValueNone -> foldArgsIgnoringMissingCells (rankFold x (order = 0)) [ sourceExpr ] ctx
    | "PERCENTILE" | "PERCENTILE.INC" | "QUARTILE" | "QUARTILE.INC" ->
        let isQ = fnName.StartsWith "QUARTILE"
        if nargs < 2 then err Err.value else
        match argNum 1 with
        | Result.Error e -> err e
        | Ok p ->
            let k = if isQ then p / 4.0 else p
            let sourceExpr = argExpr 0
            match numericBufferOfArithmeticValue sourceExpr with
            | ValueSome (Result.Error e) -> err e
            | ValueSome (Ok (struct (values, count))) -> percentileSelected values count k
            | ValueNone ->
                match trySparseOrderedArithmeticValues sourceExpr with
                | ValueSome (Result.Error e) -> err e
                | ValueSome (Ok values) -> percentileCompressed k values
                | ValueNone ->
                    match rangeAccessOf model workbook sourceExpr with
                    | Some source when k = 0.0 || k = 1.0 ->
                        match rankOneRangeValue (k = 1.0) source with
                        | Ok value -> CellValue.Number value
                        | Result.Error e -> err e
                    | _ ->
                        match rangeBounds sourceExpr with
                        | Some _ ->
                            match numericBufferOfRangeValue sourceExpr with
                            | Result.Error e -> err e
                            | Ok (struct (values, count)) -> percentileSelected values count k
                        | None ->
                            match numbersArrayOfArgs [ sourceExpr ] ctx with
                            | Result.Error e -> err e
                            | Ok values -> percentileSelected values values.Length k
    // ----- extended text ----------------------------------------------------
    | "VALUE" ->
        match argText 0 with
        | Result.Error e -> err e
        | Ok s ->
            let t = s.Trim()
            match Double.TryParse(NumberText.cutToSignificantDigits t, NumberStyles.Any, CultureInfo.InvariantCulture) with
            | true, n -> CellValue.Number n
            | _ ->
                if t.EndsWith "%" then
                    match Double.TryParse(NumberText.cutToSignificantDigits (t.TrimEnd('%')), NumberStyles.Any, CultureInfo.InvariantCulture) with
                    | true, n -> CellValue.Number (n / 100.0)
                    | _ -> err Err.value
                else
                    match DateTime.TryParse(t, CultureInfo.InvariantCulture, DateTimeStyles.None) with
                    | true, d -> CellValue.Number (dateToSerial d)
                    | _ -> err Err.value
    | "NUMBERVALUE" ->
        // NUMBERVALUE(text, [decimal_sep], [group_sep]) — parse with custom
        // separators (e.g. "1.234,56" with dec="," grp=".").
        match argText 0 with
        | Result.Error e -> err e
        | Ok s ->
            let dec = if nargs >= 2 then (match argText 1 with Ok d when d.Length > 0 -> d.[0] | _ -> '.') else '.'
            let grp = if nargs >= 3 then (match argText 2 with Ok g when g.Length > 0 -> g.[0] | _ -> ',') else ','
            let cleaned = (s.Replace(string grp, "").Replace(dec, '.')).Replace(" ", "")
            let isPct = cleaned.EndsWith "%"
            match Double.TryParse(NumberText.cutToSignificantDigits (cleaned.TrimEnd('%')), NumberStyles.Any, CultureInfo.InvariantCulture) with
            | true, n -> CellValue.Number (if isPct then n / 100.0 else n)
            | _ -> err Err.value
    | "EXACT" ->
        match argText 0, argText 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok a, Ok b -> CellValue.Boolean (String.Equals(a, b, StringComparison.Ordinal))
    | "CHAR" -> (match argNum 0 with Result.Error e -> err e | Ok n -> (let i = int n in if i < 1 || i > 255 then err Err.value else CellValue.String (string (char i))))
    | "UNICHAR" ->
        match argNum 0 with
        | Result.Error e -> err e
        | Ok n ->
            let i = int n
            if i < 1 || i > 0x10FFFF || (i >= 0xD800 && i <= 0xDFFF) then err Err.value
            else CellValue.String (stringFromCodePoint i)
    | "CODE" | "UNICODE" -> (match argText 0 with Result.Error e -> err e | Ok s -> if s.Length = 0 then err Err.value else CellValue.Number (float (int s.[0])))
    | "CLEAN" ->
        match argText 0 with
        | Result.Error e -> err e
        | Ok s ->
            let cleaned = Text.StringBuilder(s.Length)
            for ch in s do
                if int ch >= 32 then
                    cleaned.Append ch |> ignore
            CellValue.String (cleaned.ToString())
    | "TEXTBEFORE" | "TEXTAFTER" ->
        if nargs < 2 then err Err.value else
        match argText 0, argText 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok s, Ok delim ->
            if delim = "" then err Err.value else
            let instance = if nargs >= 3 then (match argNum 2 with Ok n -> int n | _ -> 1) else 1
            let rec findNth start count =
                let idx = s.IndexOf(delim, start, StringComparison.OrdinalIgnoreCase)
                if idx < 0 then -1 elif count <= 1 then idx else findNth (idx + delim.Length) (count - 1)
            let idx = findNth 0 (max 1 instance)
            if idx < 0 then err Err.na
            elif fnName = "TEXTBEFORE" then CellValue.String (s.Substring(0, idx))
            else CellValue.String (s.Substring(idx + delim.Length))
    | "FIXED" ->
        match argNum 0 with
        | Result.Error e -> err e
        | Ok x ->
            let dec = if nargs >= 2 then (match argNum 1 with Ok d -> int d | _ -> 2) else 2
            let noCommas = if nargs >= 3 then (match argBool 2 with Ok b -> b | _ -> false) else false
            let d = max 0 dec
            let factor = 10.0 ** float d
            let rounded = roundHalfAway (x * factor) / factor
            CellValue.String (NumberText.numberToString ((if noCommas then "F" else "N") + string d) rounded)
    | "T" -> (match unwrapValue (argRaw 0) with CellValue.String s -> CellValue.String s | _ -> CellValue.String "")
    // ----- extended date/time -----------------------------------------------
    | "TIME" ->
        match argNum 0, argNum 1, argNum 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok h, Ok m, Ok s -> CellValue.Number ((((h * 3600.0 + m * 60.0 + s) % 86400.0) + 86400.0) % 86400.0 / 86400.0)
    | "HOUR" -> unaryNum (fun s -> CellValue.Number (float (serialToDate s).Hour))
    | "MINUTE" -> unaryNum (fun s -> CellValue.Number (float (serialToDate s).Minute))
    | "SECOND" -> unaryNum (fun s -> CellValue.Number (float (serialToDate s).Second))
    | "EDATE" ->
        match argNum 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok s, Ok months ->
            match monthsFromSerial date1904 s (int months) false with
            | ValueSome moved -> CellValue.Number moved
            | ValueNone -> CellValue.Number (dateToSerial ((serialToDate s).AddMonths(int months)))
    | "DAYS" ->
        match argNum 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok endD, Ok startD -> CellValue.Number (floor endD - floor startD)
    | "DATEVALUE" ->
        match argText 0 with
        | Result.Error e -> err e
        | Ok s ->
            match DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None) with
            | true, d -> CellValue.Number (float (int (dateToSerial d)))
            | _ -> err Err.value
    | "ISOWEEKNUM" ->
        // ISO-8601 week = week number of the Thursday in the same week.
        unaryNum (fun s ->
            let d = serialToDate s
            let dow = if d.DayOfWeek = DayOfWeek.Sunday then 7 else int d.DayOfWeek
            let thursday = d.AddDays(float (4 - dow))
            CellValue.Number (float ((thursday.DayOfYear - 1) / 7 + 1)))
    | "WEEKNUM" ->
        match argNum 0 with
        | Result.Error e -> err e
        | Ok s ->
            let d = serialToDate s
            let jan1 = utcDate d.Year 1 1
            let returnType = if nargs >= 2 then (match argNum 1 with Ok t -> int t | _ -> 1) else 1
            let startDow = if returnType = 2 then int DayOfWeek.Monday else int DayOfWeek.Sunday
            let jan1Dow = (int jan1.DayOfWeek - startDow + 7) % 7
            CellValue.Number (float (((d - jan1).Days + jan1Dow) / 7 + 1))
    | "DATEDIF" ->
        if nargs <> 3 then err Err.value else
        match argNum 0, argNum 1, argText 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok s1, Ok s2, Ok unit ->
            let d1, d2 = serialToDate s1, serialToDate s2
            if d2 < d1 then err Err.num else
            let monthsDiff = (d2.Year - d1.Year) * 12 + d2.Month - d1.Month - (if d2.Day < d1.Day then 1 else 0)
            let res =
                match unit.ToUpperInvariant() with
                | "Y" -> float ((d2.Year - d1.Year) - (if (d2.Month, d2.Day) < (d1.Month, d1.Day) then 1 else 0))
                | "M" -> float monthsDiff
                | "D" -> floor s2 - floor s1
                | "YM" -> float (((monthsDiff % 12) + 12) % 12)
                | "YD" ->
                    let anniv = d1.AddYears(d2.Year - d1.Year - (if (d2.Month, d2.Day) < (d1.Month, d1.Day) then 1 else 0))
                    float (d2 - anniv).Days
                | "MD" ->
                    if d2.Day >= d1.Day then float (d2.Day - d1.Day)
                    else (let pm = d2.AddMonths(-1) in float (DateTime.DaysInMonth(pm.Year, pm.Month) - d1.Day + d2.Day))
                | _ -> nan
            if Double.IsNaN res then err Err.num else CellValue.Number res
    | "NETWORKDAYS" ->
        if nargs < 2 then err Err.value else
        match argNum 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok s1, Ok s2 ->
            let holidays =
                if nargs >= 3 then holidaySetFromArg (argExpr 2)
                else HashSet<float>()
            let a, b = int (floor (min s1 s2)), int (floor (max s1 s2))
            let mutable count = 0
            for day in a .. b do
                let dow = (serialToDate (float day)).DayOfWeek
                if dow <> DayOfWeek.Saturday && dow <> DayOfWeek.Sunday && not (holidays.Contains(float day)) then count <- count + 1
            CellValue.Number ((if s2 < s1 then -1.0 else 1.0) * float count)
    | "WORKDAY" ->
        if nargs < 2 then err Err.value else
        match argNum 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok s, Ok daysN ->
            let holidays =
                if nargs >= 3 then holidaySetFromArg (argExpr 2)
                else HashSet<float>()
            let step = if daysN < 0.0 then -1 else 1
            let mutable remaining = int (abs daysN)
            let mutable cur = int (floor s)
            while remaining > 0 do
                cur <- cur + step
                let dow = (serialToDate (float cur)).DayOfWeek
                if dow <> DayOfWeek.Saturday && dow <> DayOfWeek.Sunday && not (holidays.Contains(float cur)) then remaining <- remaining - 1
            CellValue.Number (float cur)
    // ----- financial (time value of money) ----------------------------------
    // Sign convention: cash you receive is +, cash you pay is −.
    | "PMT" ->
        if nargs < 3 then err Err.value else
        match argNum 0, argNum 1, argNum 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok r, Ok np, Ok pv ->
            match optNum 3 0.0, optNum 4 0.0 with
            | Result.Error e, _ | _, Result.Error e -> err e
            | Ok fv, Ok ty ->
                if np = 0.0 then err Err.num
                else CellValue.Number (loanPayment r np pv fv ty)
    | "FV" ->
        if nargs < 3 then err Err.value else
        match argNum 0, argNum 1, argNum 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok r, Ok np, Ok pmt ->
            match optNum 3 0.0, optNum 4 0.0 with
            | Result.Error e, _ | _, Result.Error e -> err e
            | Ok pv, Ok ty ->
                if r = 0.0 then CellValue.Number (-(pv + pmt * np))
                else (let p = compound r np in CellValue.Number (-(pv * p + pmt * (1.0 + r * ty) * ((p - 1.0) / r))))
    | "PV" ->
        if nargs < 3 then err Err.value else
        match argNum 0, argNum 1, argNum 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok r, Ok np, Ok pmt ->
            match optNum 3 0.0, optNum 4 0.0 with
            | Result.Error e, _ | _, Result.Error e -> err e
            | Ok fv, Ok ty ->
                if r = 0.0 then CellValue.Number (-(fv + pmt * np))
                else (let p = (1.0 + r) ** np in CellValue.Number (-(fv + pmt * (1.0 + r * ty) * (p - 1.0) / r) / p))
    | "NPER" ->
        if nargs < 3 then err Err.value else
        match argNum 0, argNum 1, argNum 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok r, Ok pmt, Ok pv ->
            match optNum 3 0.0, optNum 4 0.0 with
            | Result.Error e, _ | _, Result.Error e -> err e
            | Ok fv, Ok ty ->
                if r = 0.0 then (if pmt = 0.0 then err Err.num else CellValue.Number (-(pv + fv) / pmt))
                else
                    let a = pmt * (1.0 + r * ty)
                    let num1 = a - fv * r
                    let den1 = a + pv * r
                    if num1 / den1 <= 0.0 then err Err.num else CellValue.Number (log (num1 / den1) / log (1.0 + r))
    | "RATE" ->
        if nargs < 3 then err Err.value else
        match argNum 0, argNum 1, argNum 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok np, Ok pmt, Ok pv ->
            match optNum 3 0.0, optNum 4 0.0, optNum 5 0.1 with
            | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
            | Ok fv, Ok ty, Ok guess ->
                let tvm r = if abs r < 1e-12 then pv + pmt * np + fv else (let p = (1.0 + r) ** np in pv * p + pmt * (1.0 + r * ty) * (p - 1.0) / r + fv)
                let mutable r = guess
                let mutable result = None
                let mutable i = 0
                while result.IsNone && i < 80 do
                    let f = tvm r
                    if abs f < 1e-8 then result <- Some r
                    else
                        let d = (tvm (r + 1e-6) - f) / 1e-6
                        if abs d < 1e-14 then i <- 80 else r <- r - f / d
                    i <- i + 1
                match result with Some r -> CellValue.Number r | None -> err Err.num
    // GETPIVOTDATA reads the pivot table as it stands on its sheet (Model.PivotLayout): the row
    // whose items are those named, in the column of the data field (and of the column items named).
    | "GETPIVOTDATA" ->
        let pivotCell =
            match if nargs >= 2 then Some (argExpr 1) else None with
            | Some (CellRef r) -> Some (r.Sheet, r.Row, r.Col)
            | Some (RangeRef (a, _)) -> Some (a.Sheet, a.Row, a.Col)
            | _ -> None
        match argText 0, pivotCell with
        | Result.Error e, _ -> err e
        | _, None -> err Err.ref
        | Ok dataName, Some (sheet, row, col) ->
            // (A measure of a data model is named "[Measures].[name]": the data field is "name".)
            let dataName =
                let t = dataName.Trim()
                if t.StartsWith("[Measures].[", StringComparison.OrdinalIgnoreCase) && t.EndsWith "]" then t.Substring(12, t.Length - 13) else t
            match resolveModel sheet model workbook |> Option.bind (fun m -> m.Pivots |> List.tryFind (fun p -> row >= p.Top && row <= p.Bottom && col >= p.Left && col <= p.Right) |> Option.map (fun p -> m, p)) with
            | None -> err Err.ref
            | Some _ when nargs % 2 = 1 -> err Err.ref
            | Some (sheetModel, pivot) ->
                let same (a: string) (b: string) = String.Equals(a.Trim(), b.Trim(), StringComparison.OrdinalIgnoreCase)
                let valueAt r c = getCellValue { Row = r; Col = c } sheetModel |> unwrapValue
                let sameItem (label: CellValue) (item: CellValue) =
                    match label, unwrapValue item with
                    | CellValue.Number a, CellValue.Number b -> a = b
                    | CellValue.String a, CellValue.String b -> same a b || same (a.Replace(" Total", "")) b
                    | CellValue.Number a, CellValue.String b -> same (string a) b
                    | CellValue.String a, CellValue.Number b -> same a (string b)
                    | CellValue.Boolean a, CellValue.Boolean b -> a = b
                    | _ -> false
                let dataIndex =
                    pivot.DataFields |> List.tryFindIndex (fun name ->
                        same name dataName || same name ("Sum of " + dataName) || (name.EndsWith(" of " + dataName.Trim(), StringComparison.OrdinalIgnoreCase)))
                // The items named, by field.
                let pairs = [ for i in 2 .. 2 .. nargs - 1 -> (match argText i with Ok field -> field | _ -> ""), argRaw (i + 1) ]
                let lineRows = pivot.RowLines |> List.mapi (fun k line -> pivot.Top + pivot.FirstDataRow + k, line)
                let lineCols = pivot.ColumnLines |> List.mapi (fun k line -> pivot.Left + pivot.FirstDataCol + k, line)
                // The items of each row (column): those it repeats from the one before, and its own, read off its label.
                let paths (lines: (int * PivotLine) list) (labelAt: int -> int -> CellValue) (fields: string list) =
                    let mutable before : CellValue list = []
                    [ for at, line in lines ->
                        if line.Kind = "blank" then at, line, None
                        elif line.Kind = "grand" then at, line, None
                        else
                            // A deleted parent line cannot lend its labels to a surviving
                            // child that repeats them. Such a child has no complete path.
                            let validRepeat = line.Repeat >= 0 && line.Repeat <= before.Length
                            let items = if line.Items.IsEmpty then [ 0 ] else line.Items
                            let own =
                                items |> List.mapi (fun offset item ->
                                    let level = line.Repeat + offset
                                    if level < fields.Length && fields.[level] = "" then
                                        CellValue.Number (float item)
                                    else labelAt at level)
                            let path = (List.truncate line.Repeat before) @ own
                            if line.Kind = "data" && validRepeat then before <- path
                            at, line, (if validRepeat then Some path else None) ]
                let rowLabelColumn level = pivot.Left + (if pivot.FirstDataCol > 1 then min level (pivot.FirstDataCol - 1) else 0)
                let rowPaths = paths lineRows (fun r level -> valueAt r (rowLabelColumn level)) pivot.RowFields
                let colPaths = paths lineCols (fun c level -> valueAt (pivot.Top + pivot.FirstHeaderRow + level) c) pivot.ColumnFields
                // What each axis must be: the item of each field named, the data field where the data fields are a field of it.
                let wants (fields: string list) =
                    fields
                    |> List.mapi (fun level field -> level, field)
                    |> List.choose (fun (level, field) ->
                        if field = "" then dataIndex |> Option.map (fun index -> level, CellValue.Number (float index))
                        else pairs |> List.tryFind (fun (name, _) -> same name field) |> Option.map (fun (_, item) -> level, item))
                let known = pairs |> List.forall (fun (name, _) -> List.exists (same name) pivot.RowFields || List.exists (same name) pivot.ColumnFields)
                let pick (fields: string list) (lines: (int * PivotLine * CellValue list option) list) (want: (int * CellValue) list) =
                    if want.IsEmpty then
                        match lines |> List.tryFind (fun (_, line, _) -> line.Kind = "grand") with
                        | Some (at, _, _) -> Some at
                        // A fieldless axis has one value column/row, not a grand-total
                        // item. A remaining member of a real field is not its deleted total.
                        | None ->
                            match fields, lines with
                            | [], [ (at, line, _) ] when line.Kind = "data" -> Some at
                            | _ -> None
                    else
                        let depth = (want |> List.map fst |> List.max) + 1
                        lines
                        |> List.tryFind (fun (_, line, path) ->
                            match path with
                            | Some path when path.Length = depth && line.Kind <> "grand" ->
                                want |> List.forall (fun (level, item) -> sameItem path.[level] item)
                            | _ -> false)
                        |> Option.map (fun (at, _, _) -> at)
                if dataIndex.IsNone || not known then err Err.ref
                else
                    match pick pivot.RowFields rowPaths (wants pivot.RowFields), pick pivot.ColumnFields colPaths (wants pivot.ColumnFields) with
                    | Some r, Some c ->
                        match valueAt r c with
                        | CellValue.Empty -> err Err.ref
                        | value -> value
                    | _ -> err Err.ref
    // (IPMT, PPMT, CUMIPMT and CUMPRINC as LibreOffice computes them - Excel's results: the balance
    // after a period is the future value of the loan so far.)
    | "IPMT" | "PPMT" ->
        if nargs < 4 then err Err.value else
        match argNum 0, argNum 1, argNum 2, argNum 3 with
        | Result.Error e, _, _, _ | _, Result.Error e, _, _ | _, _, Result.Error e, _ | _, _, _, Result.Error e -> err e
        | Ok r, Ok per, Ok np, Ok pv ->
            match optNum 4 0.0, optNum 5 0.0 with
            | Result.Error e, _ | _, Result.Error e -> err e
            | Ok fv, Ok ty ->
                if per < 1.0 || per >= np + 1.0 || np <= 0.0 then err Err.num
                else
                    let pmt = loanPayment r np pv fv ty
                    let ipmt = loanInterest r per pmt pv ty
                    CellValue.Number (if name = "IPMT" then ipmt else pmt - ipmt)
    | "CUMIPMT" | "CUMPRINC" ->
        if nargs < 6 then err Err.value else
        match argNum 0, argNum 1, argNum 2, argNum 3, argNum 4, argNum 5 with
        | Ok r, Ok np, Ok pv, Ok first, Ok last, Ok ty ->
            let first, last, ty = Math.Truncate first, Math.Truncate last, Math.Truncate ty
            if r <= 0.0 || np <= 0.0 || pv <= 0.0 || first < 1.0 || last < first || last > np || (ty <> 0.0 && ty <> 1.0) then err Err.num
            else
                let pmt = loanPayment r np pv 0.0 ty
                let mutable total = 0.0
                for per in first .. last do
                    let ipmt = loanInterest r per pmt pv ty
                    total <- total + (if name = "CUMIPMT" then ipmt else pmt - ipmt)
                CellValue.Number total
        | results ->
            let (a, b, c, d, e, f) = results
            match [ a; b; c; d; e; f ] |> List.tryPick (function Result.Error code -> Some code | _ -> None) with
            | Some code -> err code
            | None -> err Err.value
    | "XNPV" ->
        if nargs < 3 then err Err.value else
        match argNum 0, numbersArrayOfArgs [ argExpr 1 ] ctx, numbersArrayOfArgs [ argExpr 2 ] ctx with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok r, Ok values, Ok dates ->
            if values.Length <> dates.Length || values.Length = 0 then err Err.num
            elif r <= -1.0 then err Err.num
            else CellValue.Number (datedPresentValue r values dates)
    | "XIRR" ->
        if nargs < 2 then err Err.value else
        match numbersArrayOfArgs [ argExpr 0 ] ctx, numbersArrayOfArgs [ argExpr 1 ] ctx with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok values, Ok dates ->
            match optNum 2 0.1 with
            | Result.Error e -> err e
            | Ok guess ->
            if values.Length <> dates.Length || values.Length < 2 || not (Array.exists (fun v -> v > 0.0) values) || not (Array.exists (fun v -> v < 0.0) values) then err Err.num
            // Excel rejects an out-of-domain guess even when the flows have another valid root.
            elif guess <= -1.0 then err Err.num
            else
                let f r = datedPresentValue r values dates
                match rootOf f guess with
                | Some r -> CellValue.Number r
                | None ->
                    // A rate nearer -1 than the rates looked at: a large outlay with little back
                    // soon. Excel 2019 finds one down to about 1e-9 above -1 (-0.9999999978 for
                    // -1,000,000 and 6,000 a month for three months), and none for two months.
                    let near = [| for k in 9 .. -1 .. 2 -> -1.0 + 10.0 ** float -k |]
                    let values' = near |> Array.map f
                    let finite (v: float) = not (Double.IsNaN v || Double.IsInfinity v)
                    match Seq.init (near.Length - 1) id |> Seq.tryFind (fun k -> finite values'.[k] && finite values'.[k + 1] && values'.[k] * values'.[k + 1] <= 0.0) with
                    | None -> err Err.num
                    | Some k ->
                        let mutable lo = near.[k]
                        let mutable hi = near.[k + 1]
                        let mutable flo = values'.[k]
                        for _ in 1 .. 200 do
                            let mid = (lo + hi) / 2.0
                            let fm = f mid
                            if flo * fm <= 0.0 then hi <- mid
                            else
                                lo <- mid
                                flo <- fm
                        CellValue.Number ((lo + hi) / 2.0)
    | "MIRR" ->
        if nargs < 3 then err Err.value else
        match numbersArrayOfArgs [ argExpr 0 ] ctx, argNum 1, argNum 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok values, Ok financeRate, Ok reinvestRate ->
            let n = float values.Length
            let presentOf (rate: float) (keep: float -> bool) =
                values |> Array.mapi (fun i v -> if keep v then v / ((1.0 + rate) ** float (i + 1)) else 0.0) |> Array.sum
            let positive = presentOf reinvestRate (fun v -> v > 0.0)
            let negative = presentOf financeRate (fun v -> v < 0.0)
            if negative = 0.0 || positive = 0.0 || n < 2.0 then err Err.div0
            else CellValue.Number ((-positive * ((1.0 + reinvestRate) ** n) / (negative * (1.0 + financeRate))) ** (1.0 / (n - 1.0)) - 1.0)
    | "NPV" ->
        if nargs < 2 then err Err.value else
        match argNum 0 with
        | Result.Error e -> err e
        | Ok r ->
            match args with
            | [_; cashFlowExpr] ->
                match trySparseNpvArithmetic r cashFlowExpr with
                | ValueSome result -> result
                | ValueNone -> foldArgsIgnoringMissingCells (npvFold r) (List.tail args) ctx
            | _ -> foldArgsIgnoringMissingCells (npvFold r) (List.tail args) ctx
    | "IRR" ->
        if nargs < 1 then err Err.value else
        // (An error in the guess is the result; one among the cash flows is #VALUE!, as Excel's.)
        match optNum 1 0.1 with
        | Result.Error e -> err e
        | Ok guess ->
        match trySparseIrrArithmetic guess (argExpr 0) with
        | ValueSome result -> result
        | ValueNone ->
            match numbersArrayOfArgs [ argExpr 0 ] ctx with
            | Result.Error _ -> err Err.value
            | Ok cashFlows when cashFlows.Length >= 2 ->
                let npv r =
                    let mutable total = 0.0
                    for i = 0 to cashFlows.Length - 1 do
                        total <- total + cashFlows.[i] / ((1.0 + r) ** float i)
                    total
                match rootOf npv guess with
                | Some r -> CellValue.Number r
                | None -> err Err.num
            | Ok _ -> err Err.num
    | "SLN" ->
        match argNum 0, argNum 1, argNum 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok cost, Ok salvage, Ok life -> if life = 0.0 then err Err.div0 else CellValue.Number ((cost - salvage) / life)
    // ----- extended info ----------------------------------------------------
    | "ISEVEN" -> (match argNum 0 with Result.Error e -> err e | Ok x -> CellValue.Boolean (int64 (Math.Truncate x) % 2L = 0L))
    | "ISODD" -> (match argNum 0 with Result.Error e -> err e | Ok x -> CellValue.Boolean (int64 (Math.Truncate x) % 2L <> 0L))
    | "ISNONTEXT" -> (if nargs <> 1 then err Err.value else CellValue.Boolean (match unwrapValue (argRaw 0) with CellValue.String _ -> false | _ -> true))
    | "N" ->
        match unwrapValue (argRaw 0) with
        | CellValue.Number n -> CellValue.Number n
        | CellValue.Boolean b -> CellValue.Number (if b then 1.0 else 0.0)
        | CellValue.Error e -> err e
        | _ -> CellValue.Number 0.0
    | "ISFORMULA" ->
        match List.tryHead args with
        | Some (CellRef r) ->
            match resolveModel r.Sheet model workbook with
            | Some m -> CellValue.Boolean (match getCellValue { Row = r.Row; Col = r.Col } m with CellValue.Formula _ -> true | _ -> false)
            | None -> err Err.ref
        | _ -> CellValue.Boolean false
    // ----- logical ----------------------------------------------------------
    | "IF" ->
        match args with
        | [ c; t ] ->
            match coerceBool (eval c ctx) with
            | Ok true -> eval t ctx
            | Ok false -> CellValue.Boolean false
            | Result.Error e -> err e
        | [ c; t; f ] ->
            match coerceBool (eval c ctx) with
            | Ok true -> eval t ctx
            | Ok false -> eval f ctx
            | Result.Error e -> err e
        | _ -> err Err.value
    | "AND" | "OR" | "XOR" ->
        foldArgsIgnoringMissingCells (logicalFold fnName) args ctx
    | "NOT" ->
        if nargs <> 1 then err Err.value else
        match argBool 0 with Ok b -> CellValue.Boolean (not b) | Result.Error e -> err e
    | "IFERROR" ->
        if nargs <> 2 then err Err.value else
        let v = argRaw 0
        if isErr v then argRaw 1 else v
    | "IFNA" ->
        if nargs <> 2 then err Err.value else
        match unwrapValue (argRaw 0) with
        | CellValue.Error e when e = Err.na -> argRaw 1
        | v -> v
    | "IFS" ->
        // IFS(cond1, val1, cond2, val2, …)
        if nargs < 2 || nargs % 2 <> 0 then err Err.value else
        let rec loop i =
            if i >= nargs then err Err.na
            else
                match coerceBool (argRaw i) with
                | Result.Error e -> err e
                | Ok true -> argRaw (i + 1)
                | Ok false -> loop (i + 2)
        loop 0
    | "SWITCH" ->
        // SWITCH(expr, case1, val1, …, [default])
        if nargs < 3 then err Err.value else
        let target = argRaw 0
        let rec loop i =
            if i + 1 >= nargs then
                // odd trailing arg = default
                if i < nargs then argRaw i else err Err.na
            else
                match compareOperands target (argRaw i) with
                | Ok 0 -> argRaw (i + 1)
                | _ -> loop (i + 2)
        loop 1
    | "LET" ->
        // LET(name1, value1, name2, value2, …, calculation). Names parse as
        // NameRef. Each value is inlined (as an expression) into the later
        // values and the final calculation, left to right, so earlier names
        // are visible to later ones. Inlining (rather than pre-evaluating)
        // keeps range-valued names usable by the structural functions.
        if nargs < 3 || nargs % 2 = 0 then err Err.value else
        let rec build i (subst: Map<string, Expr>) =
            if i = nargs - 1 then eval (substituteNames subst (argExpr i)) ctx
            else
                match argExpr i with
                | NameRef nm ->
                    let valExpr = substituteNames subst (argExpr (i + 1))
                    build (i + 2) (Map.add (nm.ToUpperInvariant()) valExpr subst)
                | _ -> err Err.value  // name position must be a bare identifier
        build 0 Map.empty
    | "LAMBDA" ->
        // A bare (uncalled) LAMBDA has no scalar value. It becomes useful when
        // bound by LET and called — `LET(f, LAMBDA(x,x*2), f(5))` — or passed
        // to MAP/REDUCE/SCAN/MAKEARRAY, both handled by beta-reduction.
        err Err.calc
    | "TRUE" -> CellValue.Boolean true
    | "FALSE" -> CellValue.Boolean false
    | "NA" -> err Err.na
    // ----- info -------------------------------------------------------------
    | "ISBLANK" ->
        if nargs <> 1 then err Err.value
        else CellValue.Boolean (match unwrapValue (argRaw 0) with CellValue.Empty -> true | _ -> false)
    | "ISERROR" ->
        if nargs <> 1 then err Err.value else CellValue.Boolean (isErr (argRaw 0))
    | "ISNA" ->
        if nargs <> 1 then err Err.value
        else CellValue.Boolean (match unwrapValue (argRaw 0) with CellValue.Error e when e = Err.na -> true | _ -> false)
    | "ISNUMBER" ->
        if nargs <> 1 then err Err.value
        else CellValue.Boolean (match unwrapValue (argRaw 0) with CellValue.Number _ -> true | _ -> false)
    | "ISTEXT" ->
        if nargs <> 1 then err Err.value
        else CellValue.Boolean (match unwrapValue (argRaw 0) with CellValue.String _ -> true | _ -> false)
    | "ISLOGICAL" ->
        if nargs <> 1 then err Err.value
        else CellValue.Boolean (match unwrapValue (argRaw 0) with CellValue.Boolean _ -> true | _ -> false)
    // ----- text -------------------------------------------------------------
    | "CONCAT" | "CONCATENATE" ->
        let sb = Text.StringBuilder()
        let mutable firstError: string option = None
        iterFlattenedArgsIgnoringMissingCells args ctx (fun value ->
            match firstError with
            | Some _ -> ()
            | None ->
                match coerceText value with
                | Ok text -> sb.Append text |> ignore
                | Result.Error e -> firstError <- Some e)
        match firstError with
        | Some e -> err e
        | None -> CellValue.String (sb.ToString())
    | "TEXTJOIN" ->
        // TEXTJOIN(delimiter, ignore_empty, text…)
        if nargs < 3 then err Err.value else
        match argText 0, argBool 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok delim, Ok ignoreEmpty ->
            let sb = Text.StringBuilder()
            let mutable needsDelimiter = false
            let appendPart (text: string) =
                if needsDelimiter then
                    sb.Append delim |> ignore
                sb.Append text |> ignore
                needsDelimiter <- true
            let rec emitTextArgs index remaining =
                match remaining with
                | [] -> ()
                | _ :: tail when index < 2 -> emitTextArgs (index + 1) tail
                | arg :: tail ->
                    let iterArg =
                        if ignoreEmpty then iterFlattenedArgIgnoringMissingCells else iterFlattenedArg
                    iterArg arg ctx (fun value ->
                        match unwrapValue value with
                        | CellValue.Empty when ignoreEmpty -> ()
                        | CellValue.String "" when ignoreEmpty -> ()
                        | other ->
                            match coerceText other with
                            | Ok text -> appendPart text
                            | Result.Error _ -> appendPart "")
                    emitTextArgs (index + 1) tail
            emitTextArgs 0 args
            CellValue.String (sb.ToString())
    | "LEN" ->
        if nargs <> 1 then err Err.value
        else match argText 0 with Ok s -> CellValue.Number (float s.Length) | Result.Error e -> err e
    | "TRIM" ->
        if nargs <> 1 then err Err.value
        else match argText 0 with
             | Ok s ->
                // Excel TRIM collapses runs of interior spaces to one.
                CellValue.String (trimExcelText s)
             | Result.Error e -> err e
    | "UPPER" ->
        if nargs <> 1 then err Err.value
        else match argText 0 with Ok s -> CellValue.String (s.ToUpperInvariant()) | Result.Error e -> err e
    | "LOWER" ->
        if nargs <> 1 then err Err.value
        else match argText 0 with Ok s -> CellValue.String (s.ToLowerInvariant()) | Result.Error e -> err e
    | "PROPER" ->
        if nargs <> 1 then err Err.value
        else match argText 0 with
             | Ok s -> CellValue.String (properCase s)
             | Result.Error e -> err e
    | "LEFT" ->
        if nargs < 1 then err Err.value else
        match argText 0 with
        | Result.Error e -> err e
        | Ok s ->
            let cnt = if nargs >= 2 then (match argNum 1 with Ok n -> int n | _ -> 1) else 1
            if cnt < 0 then err Err.value
            else CellValue.String (s.Substring(0, min s.Length cnt))
    | "RIGHT" ->
        if nargs < 1 then err Err.value else
        match argText 0 with
        | Result.Error e -> err e
        | Ok s ->
            let cnt = if nargs >= 2 then (match argNum 1 with Ok n -> int n | _ -> 1) else 1
            if cnt < 0 then err Err.value
            else
                let c = min s.Length cnt
                CellValue.String (s.Substring(s.Length - c, c))
    | "MID" ->
        if nargs <> 3 then err Err.value else
        match argText 0, argNum 1, argNum 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok s, Ok startN, Ok lenN ->
            let start0 = int startN - 1   // Excel MID is 1-indexed
            if start0 < 0 || lenN < 0.0 then err Err.value
            elif start0 >= s.Length then CellValue.String ""
            else
                let take = max 0 (min (s.Length - start0) (int lenN))
                CellValue.String (s.Substring(start0, take))
    | "REPT" ->
        if nargs <> 2 then err Err.value else
        match argText 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok s, Ok cnt ->
            if cnt < 0.0 then err Err.value
            else CellValue.String (String.replicate (int cnt) s)
    | "FIND" | "SEARCH" ->
        // FIND is case-sensitive, SEARCH is not. Both 1-indexed.
        if nargs < 2 then err Err.value else
        match argText 0, argText 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok needle, Ok hay ->
            let startAt = if nargs >= 3 then (match argNum 2 with Ok n -> int n - 1 | _ -> 0) else 0
            if startAt < 0 || startAt > hay.Length then err Err.value
            else
                let idx =
                    if fnName = "FIND" then
                        hay.IndexOf(needle, startAt, StringComparison.Ordinal)
                    elif hasCriteriaWildcard needle then
                        // SEARCH understands `*`, `?` and `~`.
                        let m = (wildcardSearchRegex needle).Match(hay.Substring startAt)
                        if m.Success then startAt + m.Index else -1
                    else
                        indexOfOrdinalIgnoreCase needle hay startAt
                if idx < 0 then err Err.value else CellValue.Number (float (idx + 1))
    | "SUBSTITUTE" ->
        // SUBSTITUTE(text, old, new, [instance])
        if nargs < 3 then err Err.value else
        match argText 0, argText 1, argText 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok s, Ok oldT, Ok newT ->
            if oldT = "" then CellValue.String s
            elif nargs >= 4 then
                match argNum 3 with
                | Result.Error e -> err e
                | Ok inst ->
                    // Replace only the inst-th occurrence.
                    let target = int inst
                    let mutable count = 0
                    let mutable idx = 0
                    let sb = System.Text.StringBuilder(s.Length)
                    while idx < s.Length do
                        if sliceEqualsOrdinal s idx oldT then
                            count <- count + 1
                            if count = target then sb.Append(newT) |> ignore
                            else sb.Append(oldT) |> ignore
                            idx <- idx + oldT.Length
                        else
                            sb.Append(s.[idx]) |> ignore
                            idx <- idx + 1
                    CellValue.String (sb.ToString())
            else
                CellValue.String (s.Replace(oldT, newT))
    | "REPLACE" ->
        // REPLACE(text, start, num_chars, new_text) — 1-indexed start.
        if nargs <> 4 then err Err.value else
        match argText 0, argNum 1, argNum 2, argText 3 with
        | Result.Error e, _, _, _ | _, Result.Error e, _, _
        | _, _, Result.Error e, _ | _, _, _, Result.Error e -> err e
        | Ok s, Ok startN, Ok lenN, Ok newT ->
            let start0 = max 0 (int startN - 1)
            let cnt = max 0 (int lenN)
            if start0 > s.Length then CellValue.String (s + newT)
            else
                let before = s.Substring(0, start0)
                let after = if start0 + cnt < s.Length then s.Substring(start0 + cnt) else ""
                CellValue.String (before + newT + after)
    | "TEXT" ->
        if nargs <> 2 then err Err.value else
        match argNum 0, argText 1 with
        | Result.Error e, _ -> err e
        | _, Result.Error e -> err e
        | Ok n, Ok fmt ->
            try CellValue.String (CellText.formatNumberWithCodeIn date1904 fmt n)
            with _ -> CellValue.String (numberToText n)
    // ----- lookup -----------------------------------------------------------
    | "CHOOSE" ->
        if nargs < 2 then err Err.value else
        match argNum 0 with
        | Result.Error e -> err e
        | Ok idx ->
            let i = int idx
            if i < 1 || i >= nargs then err Err.value else argRaw i
    | "INDEX" ->
        // INDEX(array, row_num, [col_num]) — 1-indexed; 0 = whole row/col,
        // collapsed here to the first cell of that line. The array may be a
        // range, an array constant or any expression that yields an array.
        if nargs < 2 then err Err.value else
        let source =
            match rangeAccessOf model workbook (argExpr 0) with
            | Some access when access.W > 0 ->
                Ok (access.Count / access.W, access.W, (fun r c -> rangeRead access (r * access.W + c)))
            | _ ->
                let arr = evalArray (argExpr 0) ctx
                let rows = arr.Length
                let cols = if rows > 0 then arr.[0].Length else 0
                if rows = 1 && cols = 1 then
                    match unwrapValue arr.[0].[0] with
                    | CellValue.Error e -> Result.Error e
                    | _ -> Ok (rows, cols, (fun r c -> unwrapValue arr.[r].[c]))
                elif rows = 0 || cols = 0 then Result.Error Err.ref
                else Ok (rows, cols, (fun r c -> unwrapValue arr.[r].[c]))
        match source with
        | Result.Error e -> err e
        | Ok (rows, cols, read) ->
            let indexAt i =
                if i >= nargs then Ok 1.0
                else
                    match argExpr i with
                    | MissingArg -> Ok 1.0
                    | _ -> argNum i
            match indexAt 1, indexAt 2 with
            | Result.Error e, _ | _, Result.Error e -> err e
            | Ok rowN, Ok colN ->
                let first = max 1 (int rowN) - 1
                let second = max 1 (int colN) - 1
                // Single-row or single-col array: the lone index walks it.
                let r, c =
                    if nargs < 3 && rows = 1 then 0, first
                    elif nargs < 3 && cols = 1 then first, 0
                    else first, second
                if r < 0 || r >= rows || c < 0 || c >= cols then err Err.ref
                else read r c
    | "MATCH" ->
        // MATCH(lookup, range, [match_type]) — 1=≤(sorted asc), 0=exact,
        // -1=≥(sorted desc). Defaults to 1. Streamed by index (no cell list).
        if nargs < 2 then err Err.value else
        match rangeAccessOf model workbook (argExpr 1) with
        | None -> err Err.ref
        | Some la ->
            let target = argRaw 0
            let matchType = if nargs >= 3 then (match argNum 2 with Ok m -> int m | _ -> 1) else 1
            let result =
                if matchType = 0 then
                    match wildcardLookupPredicate target with
                    | Some matches -> findIndexBySearchMode la.Count 1 (fun k -> matches (rangeRead la k))
                    | None ->
                    match trySparseExactMatchResult target la 1 with
                    | ValueSome (Some k) -> Some k
                    | ValueSome None -> None
                    | ValueNone ->
                        findIndexBySearchMode la.Count 1 (fun k -> valuesEqualNoError target (rangeRead la k))
                else
                    // type 1: largest value ≤ target (asc); -1: smallest ≥ target (desc).
                    // Both pick the LAST index satisfying the side condition (old tryLast).
                    let linear () =
                        let ok v =
                            match compareValuesNoError v target with
                            | ValueSome c -> if matchType = 1 then c <= 0 else c >= 0
                            | ValueNone -> false
                        let mutable last = None
                        for k in 0 .. la.Count - 1 do if ok (rangeRead la k) then last <- Some k
                        last
                    if matchType = 1 then
                        match tryFindLastLessOrEqualSorted la.Count (fun k -> rangeRead la k) target with
                        | ValueSome hit -> hit
                        | ValueNone -> linear ()
                    else
                        linear ()
            match result with
            | Some k -> CellValue.Number (float (k + 1))
            | None -> err Err.na
    | "VLOOKUP" | "HLOOKUP" ->
        // VLOOKUP(lookup, table, col_index, [range_lookup])
        // HLOOKUP(lookup, table, row_index, [range_lookup]). Streamed: scan the
        // first column/row by index, then read the result cell directly.
        if nargs < 3 then err Err.value else
        match rangeAccessOf model workbook (argExpr 1) with
        | None -> err Err.ref
        | Some t ->
            let isV = fnName = "VLOOKUP"
            let rows = t.Count / t.W
            let cols = t.W
            let target = argRaw 0
            match argNum 2 with
            | Result.Error e -> err e
            | Ok idxN ->
                let idx = int idxN
                let approx = if nargs >= 4 then (match argBool 3 with Ok b -> b | _ -> true) else true
                let lineLen = if isV then rows else cols
                let lineCell i = if isV then rangeRead t (i * t.W) else rangeRead t i
                let lookupLine =
                    if isV then
                        { M = t.M; SR = t.SR; SC = t.SC; W = 1; Count = rows }
                    else
                        { M = t.M; SR = t.SR; SC = t.SC; W = cols; Count = cols }
                let hit =
                    if approx then
                        // largest value ≤ target (assumes sorted ascending)
                        match tryFindLastLessOrEqualSorted lineLen lineCell target with
                        | ValueSome hit -> hit
                        | ValueNone ->
                            let mutable last = None
                            for i in 0 .. lineLen - 1 do
                                match compareValuesNoError (lineCell i) target with
                                | ValueSome c when c <= 0 -> last <- Some i
                                | _ -> ()
                            last
                    else
                        match wildcardLookupPredicate target with
                        | Some matches -> findIndexBySearchMode lineLen 1 (fun i -> matches (lineCell i))
                        | None ->
                        match trySparseExactMatchResult target lookupLine 1 with
                        | ValueSome (Some k) -> Some k
                        | ValueSome None -> None
                        | ValueNone ->
                            findIndexBySearchMode lineLen 1 (fun i -> valuesEqualNoError target (lineCell i))
                match hit with
                | None -> err Err.na
                | Some li ->
                    if isV then
                        if idx < 1 || idx > cols then err Err.ref else rangeRead t (li * t.W + (idx - 1))
                    else
                        if idx < 1 || idx > rows then err Err.ref else rangeRead t ((idx - 1) * t.W + li)
    | "XLOOKUP" ->
        // XLOOKUP(lookup, lookup_array, return_array, [if_not_found],
        //         [match_mode], [search_mode]). The matched index into
        //         lookup_array reads return_array directly (no lists).
        if nargs < 3 then err Err.value else
        let target = argRaw 0
        match rangeAccessOf model workbook (argExpr 1), rangeAccessOf model workbook (argExpr 2) with
        | Some la, Some ra ->
            let matchMode = if nargs >= 5 then (match argNum 4 with Ok m -> int m | _ -> 0) else 0
            let searchMode = if nargs >= 6 then (match argNum 5 with Ok s -> int s | _ -> 1) else 1
            match xMatchIndex target la matchMode searchMode with
            | Some k -> if k < ra.Count then rangeRead ra k else err Err.value
            | None -> if nargs >= 4 && argExpr 3 <> MissingArg then argRaw 3 else err Err.na
        | _ -> err Err.value
    | "XMATCH" ->
        // XMATCH(lookup, lookup_array, [match_mode], [search_mode]) → 1-based
        // position of the match within lookup_array (streamed).
        if nargs < 2 then err Err.value else
        let target = argRaw 0
        match rangeAccessOf model workbook (argExpr 1) with
        | Some la ->
            let matchMode = if nargs >= 3 then (match argNum 2 with Ok m -> int m | _ -> 0) else 0
            let searchMode = if nargs >= 4 then (match argNum 3 with Ok s -> int s | _ -> 1) else 1
            match xMatchIndex target la matchMode searchMode with
            | Some k -> CellValue.Number (float (k + 1))
            | None -> err Err.na
        | None -> err Err.value
    | "LOOKUP" ->
        // Vector form: LOOKUP(value, lookup_vector, [result_vector]). Finds the
        // largest lookup value ≤ target (assumes ascending) and returns the
        // aligned result-vector element.
        if nargs < 2 then err Err.value else
        match rangeAccessOf model workbook (argExpr 1) with
        | None -> err Err.na
        | Some lv ->
            let target = argRaw 0
            let resv = if nargs >= 3 then rangeAccessOf model workbook (argExpr 2) else Some lv
            let hit =
                match tryFindLastLessOrEqualSorted lv.Count (fun k -> rangeRead lv k) target with
                | ValueSome hit -> hit
                | ValueNone ->
                    let mutable hit = None
                    for k in 0 .. lv.Count - 1 do
                        match compareValuesNoError (rangeRead lv k) target with
                        | ValueSome c when c <= 0 -> hit <- Some k
                        | _ -> ()
                    hit
            match hit, resv with
            | Some k, Some rv -> if k < rv.Count then rangeRead rv k else err Err.na
            | _ -> err Err.na
    // ----- date -------------------------------------------------------------
    | "TODAY" -> CellValue.Number (dateToSerial DateTime.Today)
    | "NOW" -> CellValue.Number (dateToSerial DateTime.Now)
    | "DATE" ->
        if nargs <> 3 then err Err.value else
        match argNum 0, argNum 1, argNum 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok y, Ok m, Ok d ->
            try
                let baseDate = utcDate (int y) 1 1
                let dt = baseDate.AddMonths(int m - 1).AddDays(d - 1.0)
                CellValue.Number (dateToSerial dt)
            with _ -> err Err.num
    | "YEAR" -> unaryNum (fun s -> CellValue.Number (float (serialToDate s).Year))
    | "MONTH" -> unaryNum (fun s -> CellValue.Number (float (serialToDate s).Month))
    | "DAY" -> unaryNum (fun s -> CellValue.Number (float (serialToDate s).Day))
    | "WEEKDAY" ->
        // WEEKDAY(serial, [return_type]): 1 = Sunday..Saturday as 1..7
        // (default), 2 = Monday..Sunday as 1..7, 3 = Monday..Sunday as 0..6,
        // 11..17 = the week starting Monday..Sunday, numbered 1..7.
        if nargs < 1 then err Err.value else
        let returnType =
            if nargs < 2 then Ok 1.0
            else (match argExpr 1 with MissingArg -> Ok 1.0 | _ -> argNum 1)
        match argNum 0, returnType with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok s, Ok kind ->
            let dow = int (serialToDate s).DayOfWeek  // Sunday = 0
            let fromMonday = (dow + 6) % 7            // Monday = 0
            match int kind with
            | 1 -> CellValue.Number (float (dow + 1))
            | 2 -> CellValue.Number (float (fromMonday + 1))
            | 3 -> CellValue.Number (float fromMonday)
            | k when k >= 11 && k <= 17 -> CellValue.Number (float ((fromMonday - (k - 11) + 7) % 7 + 1))
            | _ -> err Err.num
    | "EOMONTH" ->
        if nargs <> 2 then err Err.value else
        match argNum 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok s, Ok months ->
            match monthsFromSerial date1904 s (int months) true with
            | ValueSome moved -> CellValue.Number moved
            | ValueNone ->
            let d = (serialToDate s).AddMonths(int months)
            let eom = utcDate d.Year d.Month (DateTime.DaysInMonth(d.Year, d.Month))
            CellValue.Number (dateToSerial eom)
    // ----- volatile --------------------------------------------------------
    | "RAND" -> CellValue.Number (rng.NextDouble())
    | "RANDBETWEEN" ->
        if nargs <> 2 then err Err.value else
        match argNum 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok lo, Ok hi ->
            let l, h = int (ceil lo), int (floor hi)
            if l > h then err Err.num
            else CellValue.Number (float (rng.Next(l, h + 1)))
    | _ -> evalExtendedFunction fnName args ctx

/// Functions added after the main dispatcher grew large: reference
/// information, inverse/hyperbolic trigonometry, day-count and working-day
/// calendars, regression and normal-distribution statistics, matrix scalars,
/// database functions, base conversion and bit operations.
and evalExtendedFunction (fnName: string) (args: Expr list) (ctx: EvalContext) : CellValue =
    let date1904 = date1904Of ctx
    let serialToDate (s: float) = serialToDateIn date1904 s
    let dateToSerial (d: DateTime) = dateToSerialIn date1904 d
    let argsArray = List.toArray args
    let nargs = argsArray.Length
    let supplied i = i < nargs && argsArray.[i] <> MissingArg
    let argRaw i = eval argsArray.[i] ctx
    let argNum i = coerceNumber (argRaw i)
    let argText i = coerceText (argRaw i)
    let numOr i (fallback: float) = if supplied i then argNum i else Ok fallback
    let finite (x: float) = if Double.IsNaN x || Double.IsInfinity x then err Err.num else CellValue.Number x
    let unary (f: float -> CellValue) =
        if nargs <> 1 then err Err.value
        else match argNum 0 with Ok x -> f x | Result.Error e -> err e
    /// The rectangle an argument refers to, when it is a reference.
    let rec referenceBounds (e: Expr) : (string option * int * int * int * int) option =
        match e with
        | CellRef r -> Some (r.Sheet, r.Row, r.Col, r.Row, r.Col)
        | RangeRef (a, b) -> let sr, sc, er, ec = normalizeRange a b in Some (a.Sheet, sr, sc, er, ec)
        | SpillRef anchor -> let a, b = spillRangeIn anchor ctx.Model ctx.Workbook in referenceBounds (RangeRef (a, b))
        | TableRef sel ->
            match tableSelArea sel ctx.Workbook ctx.Sheet ctx.CurrentCell with
            | Ok (r0, c0, r1, c1) when r0 <= r1 -> Some (None, r0, c0, r1, c1)
            | _ -> None
        | NameRef n ->
            match ctx.Workbook |> Option.bind (resolveNameExpr n ctx.Sheet) with
            | Some (NameRef _) | None -> None
            | Some resolved -> referenceBounds resolved
        | _ -> None
    let arrayOf i : ArrVal = evalArray argsArray.[i] ctx
    let firstErrorIn (arr: ArrVal) =
        arr |> Array.tryPick (Array.tryPick (fun v -> match unwrapValue v with CellValue.Error e -> Some e | _ -> None))
    /// Numbers of an array argument, row-major; text, booleans and blanks are skipped.
    let numbersOf i : Result<float[], string> =
        let arr = arrayOf i
        match firstErrorIn arr with
        | Some e -> Result.Error e
        | None ->
            Ok [| for row in arr do
                    for v in row do
                        match unwrapValue v with
                        | CellValue.Number x -> yield x
                        | _ -> () |]
    /// Position-aligned (y, x) pairs of two equally sized arrays where both are numbers.
    let pairsOf iy ix : Result<(float * float)[], string> =
        let ys = arrayOf iy |> Array.collect id
        let xs = arrayOf ix |> Array.collect id
        let errorOf (vs: CellValue[]) = vs |> Array.tryPick (fun v -> match unwrapValue v with CellValue.Error e -> Some e | _ -> None)
        match errorOf ys, errorOf xs with
        | Some e, _ | _, Some e -> Result.Error e
        | None, None ->
            if ys.Length <> xs.Length then Result.Error Err.na
            else
                Ok [| for i in 0 .. ys.Length - 1 do
                        match unwrapValue ys.[i], unwrapValue xs.[i] with
                        | CellValue.Number y, CellValue.Number x -> yield y, x
                        | _ -> () |]
    /// (n, mean y, mean x, Syy, Sxx, Sxy) of the pairs.
    let regression (pairs: (float * float)[]) =
        let count = float pairs.Length
        let meanY = (pairs |> Array.sumBy fst) / count
        let meanX = (pairs |> Array.sumBy snd) / count
        let mutable syy = 0.0
        let mutable sxx = 0.0
        let mutable sxy = 0.0
        for (y, x) in pairs do
            syy <- syy + (y - meanY) * (y - meanY)
            sxx <- sxx + (x - meanX) * (x - meanX)
            sxy <- sxy + (y - meanY) * (x - meanX)
        count, meanY, meanX, syy, sxx, sxy
    // Standard normal CDF (Marsaglia's series; exact to double precision).
    let normCdf (z: float) =
        if z < -8.0 then 0.0
        elif z > 8.0 then 1.0
        else
            let q = z * z
            let mutable sum = z
            let mutable previous = 0.0
            let mutable term = z
            let mutable i = 1.0
            while sum <> previous do
                previous <- sum
                i <- i + 2.0
                term <- term * q / i
                sum <- sum + term
            0.5 + sum * exp (-0.5 * q - 0.91893853320467274178)
    let normPdf (z: float) = exp (-0.5 * z * z) / sqrt (2.0 * Math.PI)
    // Inverse standard normal: Acklam's rational approximation, then one
    // Halley step against the exact CDF.
    let normInv (p: float) =
        let a = [| -3.969683028665376e+01; 2.209460984245205e+02; -2.759285104469687e+02; 1.383577518672690e+02; -3.066479806614716e+01; 2.506628277459239e+00 |]
        let b = [| -5.447609879822406e+01; 1.615858368580409e+02; -1.556989798598866e+02; 6.680131188771972e+01; -1.328068155288572e+01 |]
        let c = [| -7.784894002430293e-03; -3.223964580411365e-01; -2.400758277161838e+00; -2.549732539343734e+00; 4.374664141464968e+00; 2.938163982698783e+00 |]
        let d = [| 7.784695709041462e-03; 3.224671290700398e-01; 2.445134137142996e+00; 3.754408661907416e+00 |]
        let low = 0.02425
        let x =
            if p < low then
                let q = sqrt (-2.0 * log p)
                (((((c.[0] * q + c.[1]) * q + c.[2]) * q + c.[3]) * q + c.[4]) * q + c.[5]) / ((((d.[0] * q + d.[1]) * q + d.[2]) * q + d.[3]) * q + 1.0)
            elif p <= 1.0 - low then
                let q = p - 0.5
                let r = q * q
                (((((a.[0] * r + a.[1]) * r + a.[2]) * r + a.[3]) * r + a.[4]) * r + a.[5]) * q / (((((b.[0] * r + b.[1]) * r + b.[2]) * r + b.[3]) * r + b.[4]) * r + 1.0)
            else
                let q = sqrt (-2.0 * log (1.0 - p))
                -(((((c.[0] * q + c.[1]) * q + c.[2]) * q + c.[3]) * q + c.[4]) * q + c.[5]) / ((((d.[0] * q + d.[1]) * q + d.[2]) * q + d.[3]) * q + 1.0)
        let e = normCdf x - p
        let u = e * sqrt (2.0 * Math.PI) * exp (x * x / 2.0)
        x - u / (1.0 + x * u / 2.0)
    let isLeapYear year = (year % 4 = 0 && year % 100 <> 0) || year % 400 = 0
    let lastDayOfMonth (d: DateTime) = d.Day = DateTime.DaysInMonth(d.Year, d.Month)
    let days360 (startSerial: float) (endSerial: float) (european: bool) =
        let d1 = serialToDate (floor startSerial)
        let d2 = serialToDate (floor endSerial)
        let mutable day1 = d1.Day
        let mutable day2 = d2.Day
        if european then
            if day1 = 31 then day1 <- 30
            if day2 = 31 then day2 <- 30
        else
            let startIsFebEnd = d1.Month = 2 && lastDayOfMonth d1
            if startIsFebEnd && d2.Month = 2 && lastDayOfMonth d2 then day2 <- 30
            if startIsFebEnd then day1 <- 30
            if day2 = 31 && day1 >= 30 then day2 <- 30
            if day1 = 31 then day1 <- 30
        float ((d2.Year - d1.Year) * 360 + (d2.Month - d1.Month) * 30 + (day2 - day1))
    /// Monday = 0 … Sunday = 6 for a date serial.
    let mondayIndex (serial: int) = ((serial + int (dateSystemDays date1904) + 5) % 7 + 7) % 7
    /// Weekend mask (Monday..Sunday) from a WORKDAY.INTL / NETWORKDAYS.INTL weekend argument.
    let weekendMask i : Result<bool[], string> =
        if not (supplied i) then Ok [| false; false; false; false; false; true; true |]
        else
            match unwrapValue (argRaw i) with
            | CellValue.String pattern ->
                if pattern.Length = 7 && pattern |> Seq.forall (fun ch -> ch = '0' || ch = '1') && pattern <> "1111111" then
                    Ok (pattern.ToCharArray() |> Array.map (fun ch -> ch = '1'))
                else Result.Error Err.value
            | CellValue.Error e -> Result.Error e
            | other ->
                match coerceNumber other with
                | Result.Error e -> Result.Error e
                | Ok code ->
                    let mask = Array.create 7 false
                    match int code with
                    | k when k >= 1 && k <= 7 ->
                        // 1 = Sat+Sun, 2 = Sun+Mon, … 7 = Fri+Sat
                        let first = (k + 4) % 7
                        mask.[first] <- true
                        mask.[(first + 1) % 7] <- true
                        Ok mask
                    | k when k >= 11 && k <= 17 ->
                        // 11 = Sunday only, 12 = Monday only, … 17 = Saturday only
                        mask.[(k - 11 + 6) % 7] <- true
                        Ok mask
                    | _ -> Result.Error Err.num
    let holidaysOf i : Result<HashSet<int>, string> =
        let set = HashSet<int>()
        if not (supplied i) then Ok set
        else
            match numbersOf i with
            | Result.Error e -> Result.Error e
            | Ok days ->
                for day in days do set.Add(int (floor day)) |> ignore
                Ok set
    let parseDigits (text: string) (radix: int) (maxDigits: int) : float option =
        let t = text.Trim().ToUpperInvariant()
        if t.Length = 0 || t.Length > maxDigits then None
        else
            let mutable value = 0.0
            let mutable valid = true
            for ch in t do
                let digit =
                    if ch >= '0' && ch <= '9' then int ch - int '0'
                    elif ch >= 'A' && ch <= 'F' then int ch - int 'A' + 10
                    else 99
                if digit >= radix then valid <- false
                else value <- value * float radix + float digit
            if not valid then None
            else
                // A full-width number is two's complement.
                let limit = Math.Pow(float radix, float maxDigits)
                if t.Length = maxDigits && value >= limit / 2.0 then Some (value - limit) else Some value
    let formatDigits (value: float) (radix: int) (maxDigits: int) (places: float option) : CellValue =
        let limit = Math.Pow(float radix, float maxDigits)
        let whole = Math.Truncate value
        if whole < -(limit / 2.0) || whole > limit / 2.0 - 1.0 then err Err.num
        else
            let mutable remaining = if whole < 0.0 then whole + limit else whole
            let mutable digits = if remaining = 0.0 then "0" else ""
            while remaining > 0.0 do
                let digit = int (remaining - Math.Floor(remaining / float radix) * float radix)
                digits <- "0123456789ABCDEF".Substring(digit, 1) + digits
                remaining <- Math.Floor(remaining / float radix)
            match places with
            | Some p when whole >= 0.0 ->
                let width = int (Math.Truncate p)
                if width < digits.Length || width > maxDigits then err Err.num
                else CellValue.String (String.replicate (width - digits.Length) "0" + digits)
            | _ -> CellValue.String digits
    let fromBase radix maxDigits =
        if nargs <> 1 then err Err.value
        else
            match argText 0 with
            | Result.Error e -> err e
            | Ok text ->
                match parseDigits text radix maxDigits with
                | Some value -> CellValue.Number value
                | None -> err Err.num
    let toBase radix maxDigits =
        if nargs < 1 || nargs > 2 then err Err.value
        else
            match argNum 0, (if supplied 1 then argNum 1 |> Result.map Some else Ok None) with
            | Result.Error e, _ | _, Result.Error e -> err e
            | Ok value, Ok places -> formatDigits value radix maxDigits places
    let bitOperation (op: int64 -> int64 -> int64) =
        if nargs <> 2 then err Err.value
        else
            match argNum 0, argNum 1 with
            | Result.Error e, _ | _, Result.Error e -> err e
            | Ok a, Ok b ->
                let limit = 281474976710656.0   // 2^48
                if a < 0.0 || b < 0.0 || a >= limit || b >= limit || a <> Math.Floor a || b <> Math.Floor b then err Err.num
                else CellValue.Number (float (op (int64 a) (int64 b)))
    /// Square numeric matrix of an argument, or the error to report.
    let squareMatrix i : Result<float[][], string> =
        let arr = arrayOf i
        let size = arr.Length
        if size = 0 || arr |> Array.exists (fun row -> row.Length <> size) then Result.Error Err.value
        else
            match firstErrorIn arr with
            | Some e -> Result.Error e
            | None ->
                let mutable numeric = true
                let m =
                    arr |> Array.map (Array.map (fun v ->
                        match unwrapValue v with
                        | CellValue.Number x -> x
                        | _ -> numeric <- false; 0.0))
                if numeric then Ok m else Result.Error Err.value
    /// Database functions: the numeric/raw values of `field` in the records that match `criteria`.
    let databaseValues () : Result<CellValue[], string> =
        if nargs <> 3 then Result.Error Err.value
        else
            let database = arrayOf 0
            let criteria = arrayOf 2
            if database.Length < 1 || criteria.Length < 2 then Result.Error Err.value
            else
                let header (v: CellValue) = match coerceText v with Ok t -> t.Trim() | _ -> ""
                let columns = database.[0] |> Array.map header
                let columnIndex (name: string) = columns |> Array.tryFindIndex (fun c -> equalsOrdinalIgnoreCase c (name.Trim()))
                let field =
                    match unwrapValue (argRaw 1) with
                    | CellValue.Number index when index >= 1.0 && int index <= columns.Length -> Ok (int index - 1)
                    | CellValue.String name ->
                        (match columnIndex name with Some i -> Ok i | None -> Result.Error Err.value)
                    | CellValue.Error e -> Result.Error e
                    | _ -> Result.Error Err.value
                match field with
                | Result.Error e -> Result.Error e
                | Ok fieldIndex ->
                    let criteriaColumns = criteria.[0] |> Array.map (fun v -> columnIndex (header v))
                    // Each criteria row is an AND of its non-blank cells; rows are OR-ed.
                    let tests =
                        [| for r in 1 .. criteria.Length - 1 ->
                            [| for c in 0 .. criteria.[r].Length - 1 do
                                match unwrapValue criteria.[r].[c], criteriaColumns.[c] with
                                | CellValue.Empty, _ -> ()
                                | CellValue.String "", _ -> ()
                                | condition, Some column -> yield column, compileCriteria condition
                                | _, None -> yield -1, (fun _ -> false) |] |]
                    let matches (record: CellValue[]) =
                        tests |> Array.exists (fun row ->
                            row |> Array.forall (fun (column, test) -> column >= 0 && column < record.Length && test record.[column]))
                    Ok [| for r in 1 .. database.Length - 1 do
                            if matches database.[r] && fieldIndex < database.[r].Length then
                                yield unwrapValue database.[r].[fieldIndex] |]
    let databaseNumbers (finish: float[] -> CellValue) =
        match databaseValues () with
        | Result.Error e -> err e
        | Ok values ->
            match values |> Array.tryPick (fun v -> match v with CellValue.Error e -> Some e | _ -> None) with
            | Some e -> err e
            | None -> finish (values |> Array.choose (fun v -> match v with CellValue.Number x -> Some x | _ -> None))
    match fnName with
    // ----- reference information --------------------------------------------
    | "ROW" | "COLUMN" ->
        let pick (row: int) (col: int) = CellValue.Number (float ((if fnName = "ROW" then row else col) + 1))
        if nargs = 0 then
            match ctx.CurrentCell with
            | Some cell -> pick cell.Row cell.Col
            | None -> err Err.value
        else
            match referenceBounds argsArray.[0] with
            | Some (_, sr, sc, _, _) -> pick sr sc
            | None -> (match argsArray.[0] with ErrorLit e -> err e | _ -> err Err.value)
    | "ROWS" | "COLUMNS" ->
        if nargs <> 1 then err Err.value
        else
            match referenceBounds argsArray.[0] with
            | Some (_, sr, sc, er, ec) -> CellValue.Number (float (if fnName = "ROWS" then er - sr + 1 else ec - sc + 1))
            | None ->
                let arr = arrayOf 0
                if arr.Length = 1 && arr.[0].Length = 1 then
                    match unwrapValue arr.[0].[0] with
                    | CellValue.Error e -> err e
                    | _ -> CellValue.Number 1.0
                else CellValue.Number (float (if fnName = "ROWS" then arr.Length else (if arr.Length > 0 then arr.[0].Length else 0)))
    | "ADDRESS" ->
        // ADDRESS(row, column, [abs_num], [a1], [sheet_text])
        if nargs < 2 then err Err.value
        else
            match argNum 0, argNum 1, numOr 2 1.0 with
            | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
            | Ok rowN, Ok colN, Ok absN ->
                let row = int rowN
                let col = int colN
                let kind = int absN
                if row < 1 || row > lastSheetRow + 1 || col < 1 || col > lastSheetCol + 1 || kind < 1 || kind > 4 then err Err.value
                else
                    let a1 = if supplied 3 then (match coerceBool (argRaw 3) with Ok flag -> flag | _ -> true) else true
                    let rowAbs = kind = 1 || kind = 2
                    let colAbs = kind = 1 || kind = 3
                    let cellText =
                        if a1 then
                            (if colAbs then "$" else "") + columnToLetter (col - 1) + (if rowAbs then "$" else "") + string row
                        else
                            (if rowAbs then "R" + string row else "R[" + string row + "]")
                            + (if colAbs then "C" + string col else "C[" + string col + "]")
                    let prefix =
                        if supplied 4 then (match argText 4 with Ok sheet -> sheetPrefixText (Some sheet) | _ -> "")
                        else ""
                    CellValue.String (prefix + cellText)
    | "ISERR" ->
        // Any error except #N/A.
        if nargs <> 1 then err Err.value
        else
            match unwrapValue (argRaw 0) with
            | CellValue.Error code -> CellValue.Boolean (code <> Err.na)
            | _ -> CellValue.Boolean false
    | "SINGLE" ->
        // `@reference`: the one cell of a row/column range that shares the
        // formula's row/column; the top-left value of anything else.
        if nargs <> 1 then err Err.value
        else
            match argsArray.[0], ctx.CurrentCell with
            | RangeRef (a, b), Some cell ->
                let sr, sc, er, ec = normalizeRange a b
                if sr = er && sc = ec then eval (CellRef a) ctx
                elif sc = ec && cell.Row >= sr && cell.Row <= er then eval (CellRef { a with Row = cell.Row; Col = sc }) ctx
                elif sr = er && cell.Col >= sc && cell.Col <= ec then eval (CellRef { a with Row = sr; Col = cell.Col }) ctx
                else err Err.value
            | RangeRef (a, b), None -> firstRangeRefValue a b ctx.Model ctx.Workbook
            | other, _ ->
                if isArrayExpr other then
                    let arr = evalArray other ctx
                    if arr.Length > 0 && arr.[0].Length > 0 then arr.[0].[0] else err Err.value
                else eval other ctx
    | "ISREF" ->
        if nargs <> 1 then err Err.value
        else CellValue.Boolean (referenceBounds argsArray.[0]).IsSome
    | "ERROR.TYPE" ->
        if nargs <> 1 then err Err.value
        else
            match unwrapValue (argRaw 0) with
            | CellValue.Error code ->
                (match Err.typeNumber code with
                 | 0 -> err Err.na
                 | number -> CellValue.Number (float number))
            | _ -> err Err.na
    | "TYPE" ->
        if nargs <> 1 then err Err.value
        else
            let isMultiCell =
                match referenceBounds argsArray.[0] with
                | Some (_, sr, sc, er, ec) -> er > sr || ec > sc
                | None ->
                    match argsArray.[0] with
                    | ArrayLit rows -> rows.Length > 1 || (rows.Length = 1 && rows.[0].Length > 1)
                    | Function (name, _) when isArrayFunction name ->
                        let arr = arrayOf 0
                        arr.Length > 1 || (arr.Length = 1 && arr.[0].Length > 1)
                    | _ -> false
            if isMultiCell then CellValue.Number 64.0
            else
                match unwrapValue (argRaw 0) with
                | CellValue.Number _ | CellValue.Empty -> CellValue.Number 1.0
                | CellValue.String _ -> CellValue.Number 2.0
                | CellValue.Boolean _ -> CellValue.Number 4.0
                | _ -> CellValue.Number 16.0
    | "HYPERLINK" ->
        // The cell shows the friendly name (or the link when none is given).
        if nargs < 1 || nargs > 2 then err Err.value
        elif supplied 1 then argRaw 1
        else argRaw 0
    // ----- trigonometry -----------------------------------------------------
    | "ASIN" -> unary (fun x -> if x < -1.0 || x > 1.0 then err Err.num else CellValue.Number (Math.Asin x))
    | "ACOS" -> unary (fun x -> if x < -1.0 || x > 1.0 then err Err.num else CellValue.Number (Math.Acos x))
    | "ATAN" -> unary (fun x -> CellValue.Number (Math.Atan x))
    | "ATAN2" ->
        if nargs <> 2 then err Err.value
        else
            match argNum 0, argNum 1 with
            | Result.Error e, _ | _, Result.Error e -> err e
            | Ok x, Ok y -> if x = 0.0 && y = 0.0 then err Err.div0 else CellValue.Number (Math.Atan2(y, x))
    | "SINH" -> unary (fun x -> finite ((exp x - exp (-x)) / 2.0))
    | "COSH" -> unary (fun x -> finite ((exp x + exp (-x)) / 2.0))
    | "TANH" ->
        unary (fun x ->
            if x > 20.0 then CellValue.Number 1.0
            elif x < -20.0 then CellValue.Number -1.0
            else (let a, b = exp x, exp (-x) in CellValue.Number ((a - b) / (a + b))))
    | "ASINH" -> unary (fun x -> CellValue.Number (if x < 0.0 then -(log (-x + sqrt (x * x + 1.0))) else log (x + sqrt (x * x + 1.0))))
    | "ACOSH" -> unary (fun x -> if x < 1.0 then err Err.num else CellValue.Number (log (x + sqrt (x * x - 1.0))))
    | "ATANH" -> unary (fun x -> if x <= -1.0 || x >= 1.0 then err Err.num else CellValue.Number (0.5 * log ((1.0 + x) / (1.0 - x))))
    | "COT" -> unary (fun x -> if x = 0.0 then err Err.div0 else finite (1.0 / Math.Tan x))
    // ----- dates ------------------------------------------------------------
    | "DAYS360" ->
        if nargs < 2 then err Err.value
        else
            match argNum 0, argNum 1 with
            | Result.Error e, _ | _, Result.Error e -> err e
            | Ok startSerial, Ok endSerial ->
                let european = supplied 2 && (match coerceBool (argRaw 2) with Ok flag -> flag | _ -> false)
                CellValue.Number (days360 startSerial endSerial european)
    | "YEARFRAC" ->
        // YEARFRAC(start, end, [basis]): 0 = US 30/360, 1 = actual/actual,
        // 2 = actual/360, 3 = actual/365, 4 = European 30/360.
        if nargs < 2 then err Err.value
        else
            match argNum 0, argNum 1, numOr 2 0.0 with
            | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
            | Ok first, Ok second, Ok basisN ->
                let startSerial = floor (min first second)
                let endSerial = floor (max first second)
                let actual = endSerial - startSerial
                match int basisN with
                | 0 -> CellValue.Number (days360 startSerial endSerial false / 360.0)
                | 4 -> CellValue.Number (days360 startSerial endSerial true / 360.0)
                | 2 -> CellValue.Number (actual / 360.0)
                | 3 -> CellValue.Number (actual / 365.0)
                | 1 ->
                    let d1 = serialToDate startSerial
                    let d2 = serialToDate endSerial
                    let withinOneYear =
                        d1.Year = d2.Year
                        || (d2.Year = d1.Year + 1 && (d2.Month < d1.Month || (d2.Month = d1.Month && d2.Day <= d1.Day)))
                    let denominator =
                        if withinOneYear then
                            if d1.Year = d2.Year then (if isLeapYear d1.Year then 366.0 else 365.0)
                            else
                                // The span crosses New Year: it is a 366-day year only if it contains a 29 February.
                                let leapDayInSpan =
                                    (isLeapYear d1.Year && (d1.Month < 3))
                                    || (isLeapYear d2.Year && (d2.Month > 2 || (d2.Month = 2 && d2.Day = 29)))
                                if leapDayInSpan then 366.0 else 365.0
                        else
                            let mutable total = 0.0
                            for year in d1.Year .. d2.Year do
                                total <- total + (if isLeapYear year then 366.0 else 365.0)
                            total / float (d2.Year - d1.Year + 1)
                    CellValue.Number (actual / denominator)
                | _ -> err Err.num
    | "TIMEVALUE" ->
        if nargs <> 1 then err Err.value
        else
            match argText 0 with
            | Result.Error e -> err e
            | Ok text ->
                let clock = Text.RegularExpressions.Regex.Match(text, "(\\d{1,2}):(\\d{1,2})(?::(\\d{1,2}(?:\\.\\d+)?))?\\s*([AaPp])?")
                let hourOnly = Text.RegularExpressions.Regex.Match(text, "(\\d{1,2})\\s*([AaPp])[Mm]?\\s*$")
                let build (hourText: string) (minuteText: string) (secondText: string) (meridiem: string) =
                    let hour = float (Int32.Parse hourText)
                    let minute = if minuteText = "" then 0.0 else float (Int32.Parse minuteText)
                    let second = if secondText = "" then 0.0 else Double.Parse(secondText, CultureInfo.InvariantCulture)
                    let hour =
                        match meridiem.ToUpperInvariant() with
                        | "P" -> if hour < 12.0 then hour + 12.0 else hour
                        | "A" -> if hour = 12.0 then 0.0 else hour
                        | _ -> hour
                    if (meridiem <> "" && Int32.Parse hourText > 12) || minute > 59.0 || second >= 60.0 then err Err.value
                    else
                        let fraction = (hour * 3600.0 + minute * 60.0 + second) / 86400.0
                        CellValue.Number (fraction - Math.Floor fraction)
                if clock.Success then build clock.Groups.[1].Value clock.Groups.[2].Value clock.Groups.[3].Value clock.Groups.[4].Value
                elif hourOnly.Success then build hourOnly.Groups.[1].Value "" "" hourOnly.Groups.[2].Value
                else err Err.value
    | "NETWORKDAYS.INTL" ->
        if nargs < 2 then err Err.value
        else
            match argNum 0, argNum 1, weekendMask 2, holidaysOf 3 with
            | Result.Error e, _, _, _ | _, Result.Error e, _, _ | _, _, Result.Error e, _ | _, _, _, Result.Error e -> err e
            | Ok first, Ok second, Ok weekend, Ok holidays ->
                let lo = int (floor (min first second))
                let hi = int (floor (max first second))
                let mutable count = 0
                for day in lo .. hi do
                    if not weekend.[mondayIndex day] && not (holidays.Contains day) then count <- count + 1
                CellValue.Number (float (if second < first then -count else count))
    | "WORKDAY.INTL" ->
        if nargs < 2 then err Err.value
        else
            match argNum 0, argNum 1, weekendMask 2, holidaysOf 3 with
            | Result.Error e, _, _, _ | _, Result.Error e, _, _ | _, _, Result.Error e, _ | _, _, _, Result.Error e -> err e
            | Ok start, Ok days, Ok weekend, Ok holidays ->
                let step = if days < 0.0 then -1 else 1
                let mutable remaining = int (abs (Math.Truncate days))
                let mutable current = int (floor start)
                while remaining > 0 do
                    current <- current + step
                    if not weekend.[mondayIndex current] && not (holidays.Contains current) then remaining <- remaining - 1
                CellValue.Number (float current)
    // ----- regression and correlation ---------------------------------------
    | "CORREL" | "PEARSON" | "RSQ" | "SLOPE" | "INTERCEPT" | "STEYX" | "COVAR" | "COVARIANCE.P" | "COVARIANCE.S" ->
        if nargs <> 2 then err Err.value
        else
            // CORREL/COVAR take (array1, array2); the regression functions take (known_y's, known_x's).
            match pairsOf 0 1 with
            | Result.Error e -> err e
            | Ok pairs when pairs.Length = 0 -> err Err.div0
            | Ok pairs ->
                let count, meanY, meanX, syy, sxx, sxy = regression pairs
                match fnName with
                | "COVAR" | "COVARIANCE.P" -> CellValue.Number (sxy / count)
                | "COVARIANCE.S" -> if count < 2.0 then err Err.div0 else CellValue.Number (sxy / (count - 1.0))
                | "CORREL" | "PEARSON" -> if sxx = 0.0 || syy = 0.0 then err Err.div0 else CellValue.Number (sxy / sqrt (sxx * syy))
                | "RSQ" -> if sxx = 0.0 || syy = 0.0 then err Err.div0 else CellValue.Number (sxy * sxy / (sxx * syy))
                | "SLOPE" -> if sxx = 0.0 then err Err.div0 else CellValue.Number (sxy / sxx)
                | "INTERCEPT" -> if sxx = 0.0 then err Err.div0 else CellValue.Number (meanY - sxy / sxx * meanX)
                | _ ->
                    if count < 3.0 || sxx = 0.0 then err Err.div0
                    else CellValue.Number (sqrt (max 0.0 (syy - sxy * sxy / sxx) / (count - 2.0)))
    | "FORECAST" | "FORECAST.LINEAR" ->
        if nargs <> 3 then err Err.value
        else
            match argNum 0, pairsOf 1 2 with
            | Result.Error e, _ | _, Result.Error e -> err e
            | Ok _, Ok pairs when pairs.Length = 0 -> err Err.na
            | Ok x, Ok pairs ->
                let _, meanY, meanX, _, sxx, sxy = regression pairs
                if sxx = 0.0 then err Err.div0 else CellValue.Number (meanY + sxy / sxx * (x - meanX))
    | "RANK.AVG" ->
        if nargs < 2 then err Err.value
        else
            match argNum 0, numbersOf 1, numOr 2 0.0 with
            | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
            | Ok x, Ok values, Ok order ->
                let equal = values |> Array.filter (fun v -> v = x) |> Array.length
                if equal = 0 then err Err.na
                else
                    let ahead = values |> Array.filter (fun v -> if order = 0.0 then v > x else v < x) |> Array.length
                    CellValue.Number (float ahead + (float equal + 1.0) / 2.0)
    | "PERCENTRANK" | "PERCENTRANK.INC" ->
        if nargs < 2 then err Err.value
        else
            match numbersOf 0, argNum 1, numOr 2 3.0 with
            | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
            | Ok values, Ok x, Ok significance ->
                let sorted = Array.sort values
                let count = sorted.Length
                if count = 0 || significance < 1.0 then err Err.num
                elif x < sorted.[0] || x > sorted.[count - 1] then err Err.na
                elif count = 1 then CellValue.Number 1.0
                else
                    let below = sorted |> Array.filter (fun v -> v < x) |> Array.length
                    let rank =
                        if sorted |> Array.exists (fun v -> v = x) then float below / float (count - 1)
                        else
                            let lower = sorted.[below - 1]
                            let upper = sorted.[below]
                            (float (below - 1) + (x - lower) / (upper - lower)) / float (count - 1)
                    // Truncated (not rounded) to the requested number of digits.
                    let scale = Math.Pow(10.0, Math.Truncate significance)
                    CellValue.Number (Math.Floor(rank * scale + 1e-9) / scale)
    // ----- normal distribution ----------------------------------------------
    | "NORM.S.DIST" | "NORMSDIST" ->
        if nargs < 1 then err Err.value
        else
            match argNum 0 with
            | Result.Error e -> err e
            | Ok z ->
                let cumulative = fnName = "NORMSDIST" || not (supplied 1) || (match coerceBool (argRaw 1) with Ok flag -> flag | _ -> true)
                CellValue.Number (if cumulative then normCdf z else normPdf z)
    | "NORM.DIST" | "NORMDIST" ->
        if nargs < 4 then err Err.value
        else
            match argNum 0, argNum 1, argNum 2, coerceBool (argRaw 3) with
            | Result.Error e, _, _, _ | _, Result.Error e, _, _ | _, _, Result.Error e, _ | _, _, _, Result.Error e -> err e
            | Ok x, Ok mean, Ok sd, Ok cumulative ->
                if sd <= 0.0 then err Err.num
                else
                    let z = (x - mean) / sd
                    CellValue.Number (if cumulative then normCdf z else normPdf z / sd)
    | "NORM.S.INV" | "NORMSINV" ->
        unary (fun p -> if p <= 0.0 || p >= 1.0 then err Err.num else CellValue.Number (normInv p))
    | "NORM.INV" | "NORMINV" ->
        if nargs <> 3 then err Err.value
        else
            match argNum 0, argNum 1, argNum 2 with
            | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
            | Ok p, Ok mean, Ok sd ->
                if p <= 0.0 || p >= 1.0 || sd <= 0.0 then err Err.num
                else CellValue.Number (mean + sd * normInv p)
    // ----- matrices ---------------------------------------------------------
    | "MDETERM" ->
        if nargs <> 1 then err Err.value
        else
            match squareMatrix 0 with
            | Result.Error e -> err e
            | Ok m ->
                let size = m.Length
                let a = m |> Array.map Array.copy
                let mutable det = 1.0
                let mutable col = 0
                while det <> 0.0 && col < size do
                    let mutable pivot = col
                    for r in col + 1 .. size - 1 do
                        if abs a.[r].[col] > abs a.[pivot].[col] then pivot <- r
                    if a.[pivot].[col] = 0.0 then det <- 0.0
                    else
                        if pivot <> col then
                            let swap = a.[pivot]
                            a.[pivot] <- a.[col]
                            a.[col] <- swap
                            det <- -det
                        det <- det * a.[col].[col]
                        for r in col + 1 .. size - 1 do
                            let factor = a.[r].[col] / a.[col].[col]
                            for c in col .. size - 1 do
                                a.[r].[c] <- a.[r].[c] - factor * a.[col].[c]
                    col <- col + 1
                CellValue.Number det
    // ----- database functions -----------------------------------------------
    | "DSUM" -> databaseNumbers (fun xs -> CellValue.Number (Array.sum xs))
    | "DPRODUCT" -> databaseNumbers (fun xs -> CellValue.Number (if xs.Length = 0 then 0.0 else Array.fold (*) 1.0 xs))
    | "DCOUNT" -> databaseNumbers (fun xs -> CellValue.Number (float xs.Length))
    | "DAVERAGE" -> databaseNumbers (fun xs -> if xs.Length = 0 then err Err.div0 else CellValue.Number (Array.average xs))
    | "DMAX" -> databaseNumbers (fun xs -> CellValue.Number (if xs.Length = 0 then 0.0 else Array.max xs))
    | "DMIN" -> databaseNumbers (fun xs -> CellValue.Number (if xs.Length = 0 then 0.0 else Array.min xs))
    | "DCOUNTA" ->
        match databaseValues () with
        | Result.Error e -> err e
        | Ok values -> CellValue.Number (float (values |> Array.filter (fun v -> v <> CellValue.Empty) |> Array.length))
    | "DGET" ->
        match databaseValues () with
        | Result.Error e -> err e
        | Ok values ->
            if values.Length = 0 then err Err.value
            elif values.Length > 1 then err Err.num
            else values.[0]
    // ----- base conversion and bit operations ---------------------------------
    | "BIN2DEC" -> fromBase 2 10
    | "OCT2DEC" -> fromBase 8 10
    | "HEX2DEC" -> fromBase 16 10
    | "DEC2BIN" -> toBase 2 10
    | "DEC2OCT" -> toBase 8 10
    | "DEC2HEX" -> toBase 16 10
    | "BITAND" -> bitOperation (fun a b -> a &&& b)
    | "BITOR" -> bitOperation (fun a b -> a ||| b)
    | "BITXOR" -> bitOperation (fun a b -> a ^^^ b)
    | _ ->
        // A function of the table, one not loaded yet, a provider's - or unknown.
        callTableFunction fnName args ctx

/// Array-returning functions added alongside `evalExtendedFunction`: stacking
/// and slicing of arrays, matrix product/inverse, FREQUENCY, and the array
/// forms of ROW / COLUMN.
and evalExtendedArrayFunction (name: string) (args: Expr list) (ctx: EvalContext) : ArrVal =
    let argsArray = List.toArray args
    let nargs = argsArray.Length
    let supplied i = i < nargs && argsArray.[i] <> MissingArg
    let errArr e : ArrVal = [| [| err e |] |]
    let arrayOf i : ArrVal = evalArray argsArray.[i] ctx
    let widthOf (a: ArrVal) = if a.Length > 0 then a.[0].Length else 0
    let isErrorArray (a: ArrVal) =
        if a.Length = 1 && a.[0].Length = 1 then
            match unwrapValue a.[0].[0] with
            | CellValue.Error e -> Some e
            | _ -> None
        else None
    let intArg i : Result<int, string> =
        match coerceNumber (eval argsArray.[i] ctx) with
        | Ok n -> Ok (int (Math.Truncate n))
        | Result.Error e -> Result.Error e
    /// Indices given as scalars and/or arrays, flattened (CHOOSECOLS / CHOOSEROWS).
    let indicesFrom first : Result<int list, string> =
        let mutable result = Ok []
        for i in first .. nargs - 1 do
            match result with
            | Result.Error _ -> ()
            | Ok acc ->
                let mutable collected = acc
                let mutable failure = None
                for row in arrayOf i do
                    for v in row do
                        match coerceNumber v with
                        | Ok n -> collected <- int (Math.Truncate n) :: collected
                        | Result.Error e -> failure <- Some e
                result <- (match failure with Some e -> Result.Error e | None -> Ok collected)
        result |> Result.map List.rev
    /// 1-based index (negative counts from the end) to a 0-based position.
    let position (index: int) (size: int) =
        if index > 0 && index <= size then Some (index - 1)
        elif index < 0 && -index <= size then Some (size + index)
        else None
    let numericMatrix i : Result<float[][], string> =
        let arr = arrayOf i
        match isErrorArray arr with
        | Some e -> Result.Error e
        | None ->
            let mutable numeric = arr.Length > 0
            let m =
                arr |> Array.map (Array.map (fun v ->
                    match unwrapValue v with
                    | CellValue.Number x -> x
                    | _ -> numeric <- false; 0.0))
            if numeric then Ok m else Result.Error Err.value
    match name with
    | "ROW" | "COLUMN" ->
        match (if nargs = 1 then argsArray.[0] else MissingArg) with
        | RangeRef (a, b) ->
            let sr, sc, er, ec = normalizeRange a b
            if name = "ROW" then [| for r in sr .. er -> [| CellValue.Number (float (r + 1)) |] |]
            else [| [| for c in sc .. ec -> CellValue.Number (float (c + 1)) |] |]
        | _ -> [| [| evalExtendedFunction name args ctx |] |]
    | "VSTACK" | "HSTACK" ->
        if nargs = 0 then errArr Err.value
        else
            let parts = [| for i in 0 .. nargs - 1 -> arrayOf i |]
            let cell (a: ArrVal) r c = if r < a.Length && c < a.[r].Length then a.[r].[c] else err Err.na
            if name = "VSTACK" then
                let width = parts |> Array.map widthOf |> Array.max
                [| for part in parts do
                    for r in 0 .. part.Length - 1 ->
                        [| for c in 0 .. width - 1 -> cell part r c |] |]
            else
                let height = parts |> Array.map (fun part -> part.Length) |> Array.max
                [| for r in 0 .. height - 1 ->
                    [| for part in parts do
                        for c in 0 .. widthOf part - 1 -> cell part r c |] |]
    | "TAKE" | "DROP" ->
        if nargs < 2 then errArr Err.value
        else
            let source = arrayOf 0
            let rows = source.Length
            let cols = widthOf source
            let count i = if supplied i then intArg i |> Result.map Some else Ok None
            match isErrorArray source, count 1, count 2 with
            | Some e, _, _ -> errArr e
            | _, Result.Error e, _ | _, _, Result.Error e -> errArr e
            | None, Ok rowCount, Ok colCount ->
                // (first index, length) of the slice kept along one axis.
                let slice (size: int) (n: int option) =
                    match n with
                    | None -> Some (0, size)
                    | Some 0 when name = "TAKE" -> None
                    | Some k when name = "TAKE" ->
                        let keep = min size (abs k)
                        Some ((if k > 0 then 0 else size - keep), keep)
                    | Some k ->
                        let dropped = min size (abs k)
                        if dropped = size then None
                        else Some ((if k >= 0 then dropped else 0), size - dropped)
                match slice rows rowCount, slice cols colCount with
                | Some (r0, rn), Some (c0, cn) when rn > 0 && cn > 0 ->
                    [| for r in r0 .. r0 + rn - 1 -> [| for c in c0 .. c0 + cn - 1 -> source.[r].[c] |] |]
                | _ -> errArr Err.calc
    | "CHOOSECOLS" | "CHOOSEROWS" ->
        if nargs < 2 then errArr Err.value
        else
            let source = arrayOf 0
            let byCols = name = "CHOOSECOLS"
            let size = if byCols then widthOf source else source.Length
            match isErrorArray source, indicesFrom 1 with
            | Some e, _ -> errArr e
            | None, Result.Error e -> errArr e
            | None, Ok indices ->
                let positions = indices |> List.map (fun index -> position index size)
                if positions |> List.exists Option.isNone then errArr Err.value
                else
                    let picked = positions |> List.choose id |> List.toArray
                    if byCols then source |> Array.map (fun row -> picked |> Array.map (fun c -> row.[c]))
                    else picked |> Array.map (fun r -> source.[r])
    | "TOCOL" | "TOROW" ->
        if nargs < 1 then errArr Err.value
        else
            let source = arrayOf 0
            let ignore = if supplied 1 then (match intArg 1 with Ok n -> n | _ -> 0) else 0
            let byColumn = supplied 2 && (match coerceBool (eval argsArray.[2] ctx) with Ok flag -> flag | _ -> false)
            let ordered =
                if byColumn then [| for c in 0 .. widthOf source - 1 do for r in 0 .. source.Length - 1 -> source.[r].[c] |]
                else Array.collect id source
            let kept =
                ordered |> Array.filter (fun v ->
                    match unwrapValue v with
                    | CellValue.Empty -> ignore <> 1 && ignore <> 3
                    | CellValue.Error _ -> ignore <> 2 && ignore <> 3
                    | _ -> true)
            if kept.Length = 0 then errArr Err.calc
            elif name = "TOCOL" then kept |> Array.map (fun v -> [| v |])
            else [| kept |]
    | "MMULT" ->
        if nargs <> 2 then errArr Err.value
        else
            match numericMatrix 0, numericMatrix 1 with
            | Result.Error e, _ | _, Result.Error e -> errArr e
            | Ok a, Ok b ->
                let inner = if a.Length > 0 then a.[0].Length else 0
                if inner <> b.Length then errArr Err.value
                else
                    let cols = b.[0].Length
                    [| for r in 0 .. a.Length - 1 ->
                        [| for c in 0 .. cols - 1 ->
                            let mutable sum = 0.0
                            for k in 0 .. inner - 1 do sum <- sum + a.[r].[k] * b.[k].[c]
                            CellValue.Number sum |] |]
    | "MINVERSE" ->
        if nargs <> 1 then errArr Err.value
        else
            match numericMatrix 0 with
            | Result.Error e -> errArr e
            | Ok m when m |> Array.exists (fun row -> row.Length <> m.Length) -> errArr Err.value
            | Ok m ->
                // Gauss-Jordan elimination with partial pivoting on [m | I].
                let size = m.Length
                let a = m |> Array.mapi (fun r row -> Array.append row (Array.init size (fun c -> if c = r then 1.0 else 0.0)))
                let mutable singular = false
                for col in 0 .. size - 1 do
                    if not singular then
                        let mutable pivot = col
                        for r in col + 1 .. size - 1 do
                            if abs a.[r].[col] > abs a.[pivot].[col] then pivot <- r
                        if abs a.[pivot].[col] < 1e-14 then singular <- true
                        else
                            let swap = a.[pivot]
                            a.[pivot] <- a.[col]
                            a.[col] <- swap
                            let divisor = a.[col].[col]
                            for c in 0 .. 2 * size - 1 do a.[col].[c] <- a.[col].[c] / divisor
                            for r in 0 .. size - 1 do
                                if r <> col then
                                    let factor = a.[r].[col]
                                    if factor <> 0.0 then
                                        for c in 0 .. 2 * size - 1 do a.[r].[c] <- a.[r].[c] - factor * a.[col].[c]
                if singular then errArr Err.num
                else [| for r in 0 .. size - 1 -> [| for c in 0 .. size - 1 -> CellValue.Number a.[r].[size + c] |] |]
    | "FREQUENCY" ->
        if nargs <> 2 then errArr Err.value
        else
            let numbers (arr: ArrVal) = [| for row in arr do for v in row do match unwrapValue v with CellValue.Number x -> yield x | _ -> () |]
            let data = numbers (arrayOf 0)
            let bins = numbers (arrayOf 1) |> Array.sort
            let counts = Array.zeroCreate<int> (bins.Length + 1)
            for x in data do
                // First bin whose upper bound is >= x; past the last bin otherwise.
                let mutable slot = bins.Length
                let mutable i = 0
                while slot = bins.Length && i < bins.Length do
                    if x <= bins.[i] then slot <- i
                    i <- i + 1
                counts.[slot] <- counts.[slot] + 1
            counts |> Array.map (fun n -> [| CellValue.Number (float n) |])
    | _ -> callTableArrayFunction name args ctx

/// The evaluator a function of the table is handed (made once).
and functionEvaluator () : FunctionEvaluator =
    match sharedEvaluator with
    | Some evaluator -> evaluator
    | None ->
        let evaluator =
            { Eval = eval
              EvalArray = evalArray
              Reference = argumentReference
              IsArray = isArrayExpr
              Criteria = criteriaPredicate
              Call = fun name args ctx -> eval (Function (name, args)) ctx }
        sharedEvaluator <- Some evaluator
        evaluator

/// See `FunctionEvaluator.Reference`.
and argumentReference (e: Expr) (ctx: EvalContext) : Expr option =
    match e with
    | CellRef _ | RangeRef _ -> Some e
    | SpillRef anchor -> Some (RangeRef (spillRangeIn anchor ctx.Model ctx.Workbook))
    | NameRef name ->
        match ctx.Workbook |> Option.bind (resolveNameForEval name ctx.Sheet) with
        | Some resolved when resolved <> e -> argumentReference resolved ctx
        | _ -> None
    | Function (name, args) when isReferenceFunction name ->
        match resolveReferenceFunction name args ctx with
        | (CellRef _ | RangeRef _) as r -> Some r
        | _ -> None
    | Function (name, args) when isIndexFunction name ->
        match indexAsReference args ctx with
        | Some ((CellRef _ | RangeRef _) as r) -> Some r
        | _ -> None
    | _ -> None

/// A call whose name the built-in catalogue does not know, for one value.
and callTableFunction (name: string) (args: Expr list) (ctx: EvalContext) : CellValue =
    let topLeft (a: ArrVal) = if a.Length > 0 && a.[0].Length > 0 then a.[0].[0] else err Err.value
    match tryFunctionDef name with
    | Some def ->
        let count = List.length args
        if count < def.MinArgs || (def.MaxArgs >= 0 && count > def.MaxArgs) then err Err.value
        else
            match def.Impl with
            | ScalarFunction f -> f (functionEvaluator ()) args ctx
            | ArrayFunction f -> topLeft (f (functionEvaluator ()) args ctx)
    | None ->
        if isProviderFunction name then topLeft (callProviderFunction name args ctx)
        else
            match catalogEntry name with
            | Some entry ->
                requestFunctionWhileEvaluating entry.Module
                err Err.busy
            | None -> err Err.name

/// A call whose name the built-in catalogue does not know, for an array.
and callTableArrayFunction (name: string) (args: Expr list) (ctx: EvalContext) : ArrVal =
    match tryFunctionDef name with
    | Some def ->
        let count = List.length args
        if count < def.MinArgs || (def.MaxArgs >= 0 && count > def.MaxArgs) then [| [| err Err.value |] |]
        else
            match def.Impl with
            | ArrayFunction f -> f (functionEvaluator ()) args ctx
            | ScalarFunction f -> [| [| f (functionEvaluator ()) args ctx |] |]
    | None ->
        if isProviderFunction name then callProviderFunction name args ctx
        else [| [| callTableFunction name args ctx |] |]

/// A provider function's call: its arguments evaluated, handed to `callProvider`.
and callProviderFunction (name: string) (args: Expr list) (ctx: EvalContext) : ArrVal =
    let arguments =
        [| for arg in args ->
            match arg with
            | MissingArg -> [| [| CellValue.Empty |] |]
            | _ -> evalArray arg ctx |> Array.map (Array.map unwrapValue) |]
    let row, col = match ctx.CurrentCell with Some cell -> cell.Row, cell.Col | None -> -1, -1
    callProvider { Name = upperName name; Arguments = arguments; Sheet = ctx.Sheet; Row = row; Col = col }

/// Parse + evaluate a formula string in one shot. Defined names resolve at
/// the workbook's active sheet — the right scope for an interactive
/// edit-buffer preview, which is what this entry point serves.
/// The value of a cell's formula: `eval`, with what Excel does at the end of a formula whose last
/// operation is an addition or a subtraction - a result within 8 units in the last place of the larger
/// operand (`cancelsWithin`) is 0. (Since Excel 97; not inside a function or a deeper operation.
/// `=1-0.9-0.1` is 0, `=1*(1-0.9-0.1)` is not, nor is `=(1-0.9-0.1)`: in Excel the parentheses
/// around the whole formula are its last operation (`wrappedForms`); `=(1-0.9)-0.1` is 0.)
let evalCell (expr: Expr) (ctx: EvalContext) : CellValue =
    let value = eval expr ctx
    match expr, value with
    | BinOp (left, ("+" | "-" as op), right), CellValue.Number x when x <> 0.0 && abs x < 1e-3 && isNull (formTableGet wrappedForms (box expr)) ->
        match unwrapValue (eval left ctx), unwrapValue (eval right ctx) with
        | CellValue.Number a, CellValue.Number b ->
            let b = if op = "-" then -b else b
            if (a > 0.0) <> (b > 0.0) && cancelsWithin x a b then CellValue.Number 0.0 else value
        | _ -> value
    // A formula whose value is an empty cell (`=B5`, `=IF(c,"",B5)`) is 0: what reads the cell
    // sees a number (`=A1=""` is FALSE), as Excel's.
    | _, v when unwrapValue v = CellValue.Empty -> CellValue.Number 0.0
    // (Excel has no negative zero: `=-MIN(0,5)` is 0.)
    | _, CellValue.Number x when x = 0.0 && 1.0 / x < 0.0 -> CellValue.Number 0.0
    | _ -> value

/// A value of a dynamic array as its cell shows it: an empty cell of the array is 0 (`=B1:B3` with B2
/// empty shows 0 in its second cell), as a formula's value is (`evalCell`).
/// Of formulas 0 .. n-1, each with the formulas it reads (`graph`): which are on a cycle (a component
/// of more than one, or one that reads itself), and an order of the others in which each comes after
/// what it reads that is not on a cycle. (Tarjan's algorithm without recursion, then Kahn's.)
let liveCycleOrder (graph: int[][]) : int[] * bool[] =
    let n = graph.Length
    let index = Array.create n -1
    let low = Array.create n 0
    let onStack = Array.create n false
    let onCycle = Array.create n false
    let stack = Stack<int>()
    let work = Stack<struct (int * int)>()
    let mutable counter = 0
    for root in 0 .. n - 1 do
        if index.[root] < 0 then
            work.Push(struct (root, 0))
            while work.Count > 0 do
                let struct (v, start) = work.Pop()
                if start = 0 then
                    index.[v] <- counter
                    low.[v] <- counter
                    counter <- counter + 1
                    stack.Push v
                    onStack.[v] <- true
                else
                    let child = graph.[v].[start - 1]
                    if low.[child] < low.[v] then low.[v] <- low.[child]
                let mutable next = start
                let mutable descended = false
                while not descended && next < graph.[v].Length do
                    let w = graph.[v].[next]
                    next <- next + 1
                    if index.[w] < 0 then
                        work.Push(struct (v, next))
                        work.Push(struct (w, 0))
                        descended <- true
                    elif onStack.[w] && index.[w] < low.[v] then low.[v] <- index.[w]
                if not descended && low.[v] = index.[v] then
                    let members = ResizeArray<int>()
                    let mutable popping = true
                    while popping do
                        let w = stack.Pop()
                        onStack.[w] <- false
                        members.Add w
                        if w = v then popping <- false
                    if members.Count > 1 || Array.contains v graph.[v] then
                        for w in members do onCycle.[w] <- true
    let waiting = graph |> Array.map (fun deps -> deps |> Array.filter (fun j -> not onCycle.[j]) |> Array.length)
    let readers = Array.init n (fun _ -> ResizeArray<int>())
    graph |> Array.iteri (fun i deps -> for j in deps do if not onCycle.[j] then readers.[j].Add i)
    let ready = Queue<int>()
    for i in 0 .. n - 1 do if not onCycle.[i] && waiting.[i] = 0 then ready.Enqueue i
    let order = ResizeArray<int>(n)
    while ready.Count > 0 do
        let i = ready.Dequeue()
        order.Add i
        for reader in readers.[i] do
            waiting.[reader] <- waiting.[reader] - 1
            if waiting.[reader] = 0 && not onCycle.[reader] then ready.Enqueue reader
    order.ToArray(), onCycle

let spilledValue (value: CellValue) : CellValue =
    match value with
    | CellValue.Empty -> CellValue.Number 0.0
    | CellValue.Formula (_, CellValue.Empty) -> CellValue.Number 0.0
    | _ -> value

/// The areas a formula reads when it is evaluated now, as (sheet, first row, first column, last row,
/// last column); the sheet None is the formula's own. As Excel's calculation goes: the branch of an
/// IF, IFS, IFERROR, IFNA or CHOOSE that the values now do not take is not read. (A circular
/// reference only through such branches is none: Excel calculates `=IF(B1="",5,A2)` in A1 with
/// `=A1` in A2.) Inside the arguments of any other function everything is taken as read. A reference
/// to a sheet that is not there is left out.
let liveReferences (expr: Expr) (ctx: EvalContext) : (Rbt.Spreadsheet.Workbook.SheetId option * int * int * int * int) list =
    let found = ResizeArray<Rbt.Spreadsheet.Workbook.SheetId option * int * int * int * int>()
    let sheetOf (sheet: string option) =
        match sheet with
        | None -> Some ctx.Sheet
        | Some name ->
            match ctx.Workbook |> Option.bind (Rbt.Spreadsheet.Workbook.resolveSheetName name) with
            | Some sid -> Some (Some sid)
            | None -> None
    let add (sheet: string option) r0 c0 r1 c1 =
        match sheetOf sheet with
        | Some sid -> found.Add((sid, min r0 r1, min c0 c1, max r0 r1, max c0 c1))
        | None -> ()
    let truth (condition: Expr) =
        match unwrapValue (eval condition ctx) with
        | CellValue.Boolean b -> Some b
        | CellValue.Number n -> Some (n <> 0.0)
        | CellValue.Error _ -> None
        | other -> match coerceBool other with Ok b -> Some b | Result.Error _ -> None
    let rec walk (pruning: bool) (depth: int) (e: Expr) =
        match e with
        | Number _ | Str _ | Bool _ | ErrorLit _ | ArrayLit _ | MissingArg -> ()
        | CellRef r when isThreeDSheet r.Sheet -> walk false depth (expandThreeD e ctx)
        | RangeRef (a, _) when isThreeDSheet a.Sheet -> walk false depth (expandThreeD e ctx)
        | CellRef r -> add r.Sheet r.Row r.Col r.Row r.Col
        | RangeRef (a, b) -> add a.Sheet a.Row a.Col b.Row b.Col
        | SpillRef anchor ->
            let a, b = spillRangeIn anchor ctx.Model ctx.Workbook
            add a.Sheet a.Row a.Col b.Row b.Col
        | TableRef sel ->
            match tableRefArea ctx.Workbook ctx.Sheet ctx.CurrentCell sel with
            | Some (sid, (r0, c0, r1, c1)) -> found.Add((Some sid, r0, c0, r1, c1))
            | None -> ()
        | NameRef name ->
            if depth < 16 then
                match ctx.Workbook |> Option.bind (resolveNameForEval name ctx.Sheet) with
                | Some resolved -> walk pruning (depth + 1) resolved
                | None -> ()
        | UnaryOp (_, inner) -> walk pruning depth inner
        | BinOp (l, _, r) -> walk pruning depth l; walk pruning depth r
        | Function (name, args) when isReferenceFunction name ->
            // OFFSET, INDIRECT ...: what they read is where they point now.
            for arg in args do walk false depth arg
            match resolveReferenceFunction name args ctx with
            | (CellRef _ | RangeRef _) as target -> walk false depth target
            | _ -> ()
        | Function (name, args) ->
            let all () = for arg in args do walk false depth arg
            if not pruning then all ()
            else
                match name.ToUpperInvariant(), args with
                | "IF", [ c; t ] ->
                    walk false depth c
                    match truth c with
                    | Some true -> walk true depth t
                    | Some false -> ()
                    | None -> walk false depth t
                | "IF", [ c; t; f ] ->
                    walk false depth c
                    match truth c with
                    | Some true -> walk true depth t
                    | Some false -> walk true depth f
                    | None -> walk false depth t; walk false depth f
                | ("IFERROR" | "IFNA" as fn), [ v; alt ] ->
                    walk false depth v
                    match unwrapValue (eval v ctx) with
                    | CellValue.Error code when fn = "IFERROR" || code = Err.na -> walk true depth alt
                    | _ -> ()
                | "IFS", _ when args.Length >= 2 && args.Length % 2 = 0 ->
                    let pairs = List.chunkBySize 2 args
                    let rec go (rest: Expr list list) =
                        match rest with
                        | [ c; v ] :: more ->
                            walk false depth c
                            match truth c with
                            | Some true -> walk true depth v
                            | Some false -> go more
                            | None -> walk false depth v; go more
                        | _ -> ()
                    go pairs
                | "CHOOSE", index :: choices when not choices.IsEmpty ->
                    walk false depth index
                    match unwrapValue (eval index ctx) with
                    | CellValue.Number n when int n >= 1 && int n <= choices.Length -> walk true depth (List.item (int n - 1) choices)
                    | _ -> for choice in choices do walk false depth choice
                | _ -> all ()
    walk true 0 expr
    List.ofSeq found

/// The value of a formula cell whose value now is `cached`: `evalCell`, except
/// that a formula calling a provider function nobody provides keeps the value
/// it has (`keepsCachedValue`) - the value the file came with. One with no
/// value yet (or #BUSY!) is calculated, and shows the function's error.
let evalFormulaCell (expr: Expr) (ctx: EvalContext) (cached: CellValue) : CellValue =
    let cached = if isNull (box cached) then CellValue.Empty else unwrapValue cached
    match cached with
    | CellValue.Empty -> evalCell expr ctx
    | CellValue.Error e when e = Err.busy -> evalCell expr ctx
    | _ -> if keepsCachedValue expr then cached else evalCell expr ctx

/// What a formula on a circular reference, or one that reads one, shows without iterative
/// calculation: Excel does not calculate it, and it keeps the value it has (`value`, the cell's; one
/// that has none, 0 - as a formula just typed into such a place).
let keptValue (value: CellValue) : CellValue =
    match (match value with CellValue.Formula (_, cached) -> unwrapValue cached | other -> unwrapValue other) with
    | CellValue.Empty -> CellValue.Number 0.0
    | CellValue.Error e when e = Err.cycle -> CellValue.Number 0.0
    | other -> other

let evaluateFormula (formula: string) (model: SpreadsheetModel)
                    (workbook: Rbt.Spreadsheet.Workbook.Workbook option) : CellValue =
    match parseFormulaShared formula with
    | Ok expr ->
        let ctx =
            { Model = model
              Workbook = workbook
              Sheet = workbook |> Option.map (fun w -> w.ActiveSheet)
              CurrentCell = None }
        evalCell expr ctx
    | Result.Error _ -> CellValue.Error Err.syntax

// ---------------------------------------------------------------------------
// Whole-model recalculation (dependency-ordered). Calc.fs owns the
// incremental engine; this is the simple full pass the shared model uses.
// ---------------------------------------------------------------------------

type private FormulaCell = {
    Address: CellAddress
    Cell: CellData
    Formula: string
    Expr: Expr
    ParseError: string option
}

let private collectFormulaCells (model: SpreadsheetModel) =
    let formulaCells = ResizeArray<FormulaCell>()
    model.Cells
    |> Map.iter (fun addr cell ->
        match cell.Value with
        | CellValue.Formula (formula, _) ->
            match parseFormulaShared formula with
            | Ok expr ->
                formulaCells.Add {
                    Address = addr
                    Cell = cell
                    Formula = formula
                    Expr = expr
                    ParseError = None
                }
            | Result.Error error ->
                formulaCells.Add {
                    Address = addr
                    Cell = cell
                    Formula = formula
                    Expr = Number 0.0
                    ParseError = Some error
                }
        | _ -> ())
    formulaCells

/// Recalculate every formula cell in a model, in dependency order.
/// `currentSheet` scopes defined-name resolution to the sheet being
/// recalculated (the caller knows which sheet `model` belongs to).
let rec recalculateAllFormulas (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                               (currentSheet: Rbt.Spreadsheet.Workbook.SheetId option)
                               (model: SpreadsheetModel) : SpreadsheetModel =
    recalculateFormulasKeepingBlocked Set.empty workbook currentSheet model

/// `blockedAnchors` are array formulas whose spill is obstructed: they keep
/// showing #SPILL! instead of collapsing to their top-left element.
and private recalculateFormulasKeepingBlocked (blockedAnchors: Set<CellAddress>)
                                              (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                                              (currentSheet: Rbt.Spreadsheet.Workbook.SheetId option)
                                              (model: SpreadsheetModel) : SpreadsheetModel =
    let formulaCells = collectFormulaCells model
    if formulaCells.Count = 0 then model else
        let formulaByAddress = Dictionary<CellAddress, FormulaCell>(formulaCells.Count)
        let addIndexed (index: Dictionary<int, ResizeArray<CellAddress>>) key addr =
            match Portable.Dict.tryFind key index with
            | ValueSome cells -> cells.Add addr
            | ValueNone ->
                let cells = ResizeArray<CellAddress>()
                cells.Add addr
                index.[key] <- cells
        for fc in formulaCells do
            formulaByAddress.[fc.Address] <- fc

        let mutable formulaRows = ValueNone
        let mutable formulaCols = ValueNone
        let formulaIndexBy keyOf =
            let index = Dictionary<int, ResizeArray<CellAddress>>()
            for fc in formulaCells do
                addIndexed index (keyOf fc.Address) fc.Address
            index
        let formulaRowsIndex () =
            match formulaRows with
            | ValueSome index -> index
            | ValueNone ->
                let index = formulaIndexBy (fun addr -> addr.Row)
                formulaRows <- ValueSome index
                index
        let formulaColsIndex () =
            match formulaCols with
            | ValueSome index -> index
            | ValueNone ->
                let index = formulaIndexBy (fun addr -> addr.Col)
                formulaCols <- ValueSome index
                index

        let referencedFormulaCells (owner: CellAddress) (expr: Expr) =
            let refs = ResizeArray<CellAddress>()
            let addCell addr =
                if formulaByAddress.ContainsKey addr then
                    refs.Add addr
            let addFormulaCellsInRange a b =
                let sr, sc, er, ec = normalizeRange a b
                let rowCount = er - sr + 1
                let colCount = ec - sc + 1
                if rowCount <= colCount then
                    let formulaRows = formulaRowsIndex ()
                    for row in sr .. er do
                        match Portable.Dict.tryFind row formulaRows with
                        | ValueSome candidates ->
                            for addr in candidates do
                                if addr.Col >= sc && addr.Col <= ec then
                                    refs.Add addr
                        | ValueNone -> ()
                else
                    let formulaCols = formulaColsIndex ()
                    for col in sc .. ec do
                        match Portable.Dict.tryFind col formulaCols with
                        | ValueSome candidates ->
                            for addr in candidates do
                                if addr.Row >= sr && addr.Row <= er then
                                    refs.Add addr
                        | ValueNone -> ()
            let stack = ResizeArray<struct (Set<string> * Expr)>()
            stack.Add(struct (Set.empty, expr))
            while stack.Count > 0 do
                let last = stack.Count - 1
                let struct (seen, e) = stack.[last]
                Portable.Stack.dropLast stack
                match e with
                | Number _ | Str _ | Bool _ | ErrorLit _ | ArrayLit _ | MissingArg -> ()
                | TableRef sel ->
                    // A structured reference reads the table's cells: formula
                    // cells among them (a calculated column) must come first.
                    match tableRefArea workbook currentSheet (Some owner) sel with
                    | Some (sid, (r0, c0, r1, c1)) when currentSheet.IsNone || currentSheet = Some sid ->
                        let corner row col = { Sheet = None; Row = row; Col = col; RowAbs = false; ColAbs = false }
                        addFormulaCellsInRange (corner r0 c0) (corner r1 c1)
                    | _ -> ()
                | NameRef name ->
                    let key = name.ToUpperInvariant()
                    if not (Set.contains key seen) then
                        match workbook |> Option.bind (resolveNameExpr name None) with
                        | Some resolved -> stack.Add(struct (Set.add key seen, resolved))
                        | None -> ()
                | CellRef r when r.Sheet.IsSome -> ()
                | CellRef r | SpillRef r -> addCell { Row = r.Row; Col = r.Col }
                | RangeRef (a, _) when a.Sheet.IsSome -> ()
                | RangeRef (a, b) -> addFormulaCellsInRange a b
                | UnaryOp (_, inner) -> stack.Add(struct (seen, inner))
                | BinOp (l, _, r) ->
                    stack.Add(struct (seen, r))
                    stack.Add(struct (seen, l))
                | Function (_, args) ->
                    for arg in args do
                        stack.Add(struct (seen, arg))
            refs.ToArray()

        let referencedFormulaCellsFor (fc: FormulaCell) =
            referencedFormulaCells fc.Address fc.Expr

        let dependencies = Dictionary<CellAddress, CellAddress array>(formulaCells.Count)
        for fc in formulaCells do
            match fc.ParseError with
            | Some _ -> ()
            | None ->
                let refs = referencedFormulaCellsFor fc
                if refs.Length > 0 then
                    dependencies.[fc.Address] <- refs

        let visited = HashSet<CellAddress>()
        let activeStack = ResizeArray<CellAddress>()
        let activeIndex = Dictionary<CellAddress, int>()
        let order = ResizeArray<CellAddress>(formulaCells.Count)
        let cyclic = HashSet<CellAddress>()

        let markCycleFrom cycleStart =
            for i = cycleStart to activeStack.Count - 1 do
                cyclic.Add activeStack.[i] |> ignore

        let depsFor addr =
            match Portable.Dict.tryFind addr dependencies with
            | ValueSome deps -> deps
            | ValueNone -> Array.empty

        let enter (stack: Stack<struct (CellAddress * CellAddress array * int)>) addr =
            let mutable cycleStart = 0
            if activeIndex.TryGetValue(addr, &cycleStart) then
                markCycleFrom cycleStart
            elif visited.Add addr then
                activeIndex.[addr] <- activeStack.Count
                activeStack.Add addr
                stack.Push(struct (addr, depsFor addr, 0))

        for fc in formulaCells do
            let stack = Stack<struct (CellAddress * CellAddress array * int)>()
            enter stack fc.Address
            while stack.Count > 0 do
                let struct (addr, deps, index) = stack.Pop()
                if index >= deps.Length then
                    activeIndex.Remove addr |> ignore
                    Portable.Stack.dropLast activeStack
                    order.Add addr
                else
                    stack.Push(struct (addr, deps, index + 1))
                    enter stack deps.[index]

        let mutable updatedCells = model.Cells
        let mutable changed = false
        let applyFormulaResult addr (fc: FormulaCell) result =
            let nextValue = CellValue.Formula (fc.Formula, result)
            if fc.Cell.Value <> nextValue then
                updatedCells <- Map.add addr { fc.Cell with Value = nextValue } updatedCells
                changed <- true
        for addr in order do
            match Portable.Dict.tryFind addr formulaByAddress with
            | ValueSome fc ->
                match fc.ParseError with
                | Some parseError ->
                    ignore parseError
                    applyFormulaResult addr fc (CellValue.Error Err.syntax)
                | None when cyclic.Contains addr ->
                    applyFormulaResult addr fc (keptValue fc.Cell.Value)
                | None when Set.contains addr blockedAnchors ->
                    applyFormulaResult addr fc (CellValue.Error Err.spill)
                | None ->
                    let updated = { model with Cells = updatedCells }
                    let ctx = { Model = updated; Workbook = workbook; Sheet = currentSheet; CurrentCell = Some addr }
                    let result = evalFormulaCell fc.Expr ctx (match fc.Cell.Value with CellValue.Formula (_, cached) -> cached | _ -> CellValue.Empty)
                    applyFormulaResult addr fc result
            | ValueNone -> ()
        // A circular reference of the references written that the calculation does not follow
        // (`liveReferences`: the branch of an IF the values do not take): its formulas, and what
        // reads them, calculated in the order of what they read now, that order found again from
        // their new values until it stays; those still on a cycle of it keep their values, and what reads them.
        let cyclicFormulas = cyclic |> Seq.filter (fun addr -> match Portable.Dict.tryFind addr formulaByAddress with ValueSome fc -> fc.ParseError.IsNone && not (Set.contains addr blockedAnchors) | ValueNone -> false) |> Seq.toArray
        if cyclicFormulas.Length > 0 then
            let readersOf = Dictionary<CellAddress, ResizeArray<CellAddress>>()
            for KeyValue(addr, deps) in dependencies do
                for dep in deps do
                    match Portable.Dict.tryFind dep readersOf with
                    | ValueSome list -> list.Add addr
                    | ValueNone ->
                        let list = ResizeArray<CellAddress>()
                        list.Add addr
                        readersOf.[dep] <- list
            let inSet = HashSet<CellAddress>()
            let pending = Stack<CellAddress>(cyclicFormulas)
            while pending.Count > 0 do
                let addr = pending.Pop()
                if inSet.Add addr then
                    match Portable.Dict.tryFind addr readersOf with
                    | ValueSome list -> for reader in list do if not (Set.contains reader blockedAnchors) then pending.Push reader
                    | ValueNone -> ()
            let members = order |> Seq.filter (fun addr -> inSet.Contains addr) |> Seq.toArray
            let n = members.Length
            let place = Dictionary<CellAddress, int>()
            members |> Array.iteri (fun i addr -> place.[addr] <- i)
            let liveDeps (i: int) : int[] =
                let addr = members.[i]
                match Portable.Dict.tryFind addr formulaByAddress with
                | ValueSome fc when fc.ParseError.IsNone ->
                    let ctx = { Model = { model with Cells = updatedCells }; Workbook = workbook; Sheet = currentSheet; CurrentCell = Some addr }
                    let deps = HashSet<int>()
                    for (sheet, r0, c0, r1, c1) in liveReferences fc.Expr ctx do
                        if sheet.IsNone || sheet = currentSheet then
                            if int64 (r1 - r0 + 1) * int64 (c1 - c0 + 1) <= int64 n then
                                for r in r0 .. r1 do
                                    for c in c0 .. c1 do
                                        match place.TryGetValue { Row = r; Col = c } with
                                        | true, j -> deps.Add j |> ignore
                                        | _ -> ()
                            else
                                for j in 0 .. n - 1 do
                                    let other = members.[j]
                                    if other.Row >= r0 && other.Row <= r1 && other.Col >= c0 && other.Col <= c1 then deps.Add j |> ignore
                    Seq.toArray deps |> Array.sort
                | _ -> [||]
            let mutable graph = Array.init n liveDeps
            let mutable sweeps = 0
            let mutable going = true
            while going do
                sweeps <- sweeps + 1
                let evaluate (i: int) (forced: CellValue option) =
                    let addr = members.[i]
                    match Portable.Dict.tryFind addr formulaByAddress with
                    | ValueSome fc when fc.ParseError.IsNone ->
                        let result =
                            match forced with
                            | Some value -> value
                            | None ->
                                let ctx = { Model = { model with Cells = updatedCells }; Workbook = workbook; Sheet = currentSheet; CurrentCell = Some addr }
                                evalFormulaCell fc.Expr ctx (match fc.Cell.Value with CellValue.Formula (_, cached) -> cached | _ -> CellValue.Empty)
                        // (Against what the cell has now: the pass above may have given it the cycle error.)
                        let nextValue = CellValue.Formula (fc.Formula, result)
                        match Map.tryFind addr updatedCells with
                        | Some current when current.Value = nextValue -> ()
                        | _ ->
                            updatedCells <- Map.add addr { fc.Cell with Value = nextValue } updatedCells
                            changed <- true
                    | _ -> ()
                let calculation, onCycle = liveCycleOrder graph
                // (On a cycle or reading one: not calculated - the value the formula has.)
                let kept = Array.copy onCycle
                for i in calculation do
                    if graph.[i] |> Array.exists (fun j -> kept.[j]) then kept.[i] <- true
                let held (i: int) =
                    match Portable.Dict.tryFind members.[i] formulaByAddress with
                    | ValueSome fc -> Some (keptValue fc.Cell.Value)
                    | ValueNone -> None
                for i in 0 .. n - 1 do
                    if onCycle.[i] then
                        cyclic.Add members.[i] |> ignore
                        evaluate i (held i)
                    else cyclic.Remove members.[i] |> ignore
                for i in calculation do evaluate i (if kept.[i] then held i else None)
                let next = Array.init n liveDeps
                if next = graph || sweeps >= 8 then going <- false
                else graph <- next
        // OFFSET / INDIRECT pick their precedents while evaluating, so the
        // ordering above cannot place them after the cells they read: one of
        // them may have run before its target was recalculated. Re-evaluate
        // those formulas and everything downstream until they settle.
        let dynamicCells =
            [ for fc in formulaCells do
                if fc.ParseError.IsNone
                   && (indexOfOrdinalIgnoreCase "OFFSET" fc.Formula 0 >= 0 || indexOfOrdinalIgnoreCase "INDIRECT" fc.Formula 0 >= 0)
                   && containsDynamicReference fc.Expr
                   && not (cyclic.Contains fc.Address) && not (Set.contains fc.Address blockedAnchors) then
                    yield fc.Address ]
        if not dynamicCells.IsEmpty then
            let dependents = Dictionary<CellAddress, ResizeArray<CellAddress>>()
            for KeyValue(addr, deps) in dependencies do
                for dep in deps do
                    match Portable.Dict.tryFind dep dependents with
                    | ValueSome list -> list.Add addr
                    | ValueNone ->
                        let list = ResizeArray<CellAddress>()
                        list.Add addr
                        dependents.[dep] <- list
            let affected = HashSet<CellAddress>()
            let pending = Stack<CellAddress>(dynamicCells)
            while pending.Count > 0 do
                let addr = pending.Pop()
                if affected.Add addr then
                    match Portable.Dict.tryFind addr dependents with
                    | ValueSome list -> for dependent in list do pending.Push dependent
                    | ValueNone -> ()
            let mutable settling = true
            let mutable rounds = 0
            while settling && rounds < 3 do
                settling <- false
                rounds <- rounds + 1
                for addr in order do
                    if affected.Contains addr && not (cyclic.Contains addr) && not (Set.contains addr blockedAnchors) then
                        match Portable.Dict.tryFind addr formulaByAddress with
                        | ValueSome fc when fc.ParseError.IsNone ->
                            let updated = { model with Cells = updatedCells }
                            let ctx = { Model = updated; Workbook = workbook; Sheet = currentSheet; CurrentCell = Some addr }
                            let nextValue = CellValue.Formula (fc.Formula, evalFormulaCell fc.Expr ctx (match fc.Cell.Value with CellValue.Formula (_, cached) -> cached | _ -> CellValue.Empty))
                            let current = Map.tryFind addr updatedCells |> Option.map (fun cell -> cell.Value)
                            if current <> Some nextValue then
                                updatedCells <- Map.add addr { fc.Cell with Value = nextValue } updatedCells
                                changed <- true
                                settling <- true
                        | _ -> ()
        if changed then { model with Cells = updatedCells } |> invalidateFilterProjection else model

/// True if the expression tree uses a dynamic-array (spilling) function.
let containsArrayFunction (e: Expr) : bool =
    let stack = ResizeArray<Expr>()
    stack.Add e
    let mutable found = false
    while not found && stack.Count > 0 do
        let last = stack.Count - 1
        let expr = stack.[last]
        Portable.Stack.dropLast stack
        match expr with
        | Function (name, _) when isArrayFunction name -> found <- true
        | Function (_, args) ->
            for arg in args do
                stack.Add arg
        | UnaryOp (_, inner) -> stack.Add inner
        | BinOp (left, _, right) ->
            stack.Add left
            stack.Add right
        | _ -> ()
    found

/// `containsArrayFunction e || isArrayExprCore false dynamicReferences e` in
/// one walk of the formula (it is asked of every formula when the dependency
/// graph is built). Every node is visited once, with whether it stands where
/// `isArrayExprCore` would look at it (`deep`): there the whole test applies,
/// elsewhere only "is it a call of an array function".
let private spillsInOneWalk (dynamicReferences: bool) (e: Expr) : bool =
    let stack = ResizeArray<Expr>()
    let deepStack = ResizeArray<bool>()
    stack.Add e
    deepStack.Add true
    let mutable found = false
    while not found && stack.Count > 0 do
        let last = stack.Count - 1
        let expr = stack.[last]
        let deep = deepStack.[last]
        Portable.Stack.dropLast stack
        deepStack.RemoveAt last
        match expr with
        | Function (name, args) ->
            if isArrayFunction name then found <- true
            elif not deep then
                for arg in args do
                    stack.Add arg
                    deepStack.Add false
            elif name = unionOperatorName then found <- true
            elif isReferenceFunction name then
                if dynamicReferences then found <- true
                else
                    let operands = name = rangeOperatorName || name = intersectionOperatorName
                    for arg in args do
                        stack.Add arg
                        deepStack.Add operands
            elif isIndexFunction name then
                if indexMaySelectArea args then found <- true
                else
                    let mutable first = true
                    for arg in args do
                        stack.Add arg
                        deepStack.Add (not first)
                        first <- false
            else
                let numbersLines =
                    match args with
                    | [ RangeRef _ ] -> equalsOrdinalIgnoreCase name "ROW" || equalsOrdinalIgnoreCase name "COLUMN"
                    | _ -> false
                if numbersLines then found <- true
                else
                    match functionSignature name with
                    | ValueSome signature ->
                        let mutable index = 0
                        for arg in args do
                            stack.Add arg
                            deepStack.Add (paramKindAt signature index = ScalarParam)
                            index <- index + 1
                    | ValueNone ->
                        for arg in args do
                            stack.Add arg
                            deepStack.Add false
        | UnaryOp (_, inner) ->
            stack.Add inner
            deepStack.Add deep
        | BinOp (left, _, right) ->
            stack.Add left
            deepStack.Add deep
            stack.Add right
            deepStack.Add deep
        | _ when not deep -> ()
        | RangeRef (a, b) -> if not (isWholeColumnRange a b || isWholeRowRange a b) then found <- true
        | CellRef r -> if isThreeDSheet r.Sheet then found <- true
        | TableRef sel when sel.Item = TblThisRow && sel.ColEnd.IsNone -> ()
        | TableRef _ | ArrayLit _ | SpillRef _ -> found <- true
        | _ -> ()
    found

/// Does a formula written in a cell produce an array that spills into its
/// neighbours? It does when it uses a dynamic-array function, or when its
/// value is otherwise an array: a multi-cell reference (`=A1:A3`,
/// `=OFFSET(A1,0,0,3,1)`, `=INDEX(A1:C3,0,2)`), an operator or scalar
/// function applied to one (`=A1:A3*2`, `=LEN(A1:A3)`), an array constant.
/// A whole-column / whole-row reference in a scalar position does not count —
/// that is implicit intersection, see `eval`.
let exprSpills (e: Expr) : bool = spillsInOneWalk true e

/// `exprSpills` without the benefit of the doubt for OFFSET / INDIRECT: true
/// only when the formula's text alone shows it yields an array. The
/// incremental engine uses it — a single-cell OFFSET is ordinary there, and
/// one that did spill has left its mark in the sheet's spill map.
let exprSpillsStatically (e: Expr) : bool = spillsInOneWalk false e

/// `exprSpills` / `exprSpillsStatically` as they are defined - two walks of
/// the formula: the reference the tests check the one walk against.
let exprSpillsReference (dynamicReferences: bool) (e: Expr) : bool =
    containsArrayFunction e || isArrayExprCore false dynamicReferences e

/// True if the expression holds a 3-D reference (`Sheet1:Sheet3!A1`). Its
/// precedents span several sheets and are only known with the workbook's tab
/// order, so the incremental engine leaves such workbooks to the full pass.
let containsThreeDReference (e: Expr) : bool =
    let stack = ResizeArray<Expr>()
    stack.Add e
    let mutable found = false
    while not found && stack.Count > 0 do
        let last = stack.Count - 1
        let expr = stack.[last]
        Portable.Stack.dropLast stack
        match expr with
        | CellRef r | SpillRef r -> if isThreeDSheet r.Sheet then found <- true
        | RangeRef (a, _) -> if isThreeDSheet a.Sheet then found <- true
        | Function (_, args) -> for arg in args do stack.Add arg
        | UnaryOp (_, inner) -> stack.Add inner
        | BinOp (left, _, right) ->
            stack.Add left
            stack.Add right
        | _ -> ()
    found

/// Cheap text test: can this formula possibly spill? Saves parsing the ones
/// that plainly cannot.
/// What `formulaTextMaySpill` found for a formula text: every recalculation
/// asks for every formula of the workbook again.
let private maySpillByText = Dictionary<string, bool>()

let private formulaTextMaySpill (formula: string) =
    if isNull formula then false
    else
        match Portable.Dict.tryFind formula maySpillByText with
        | ValueSome known -> known
        | ValueNone ->
            let found =
                formula.IndexOf ':' >= 0 || formula.IndexOf '{' >= 0 || formula.IndexOf '#' >= 0 || formula.IndexOf '[' >= 0
                || (let upper = formula.ToUpperInvariant()
                    upper.IndexOf("OFFSET", StringComparison.Ordinal) >= 0 || upper.IndexOf("INDIRECT", StringComparison.Ordinal) >= 0
                    || upper.IndexOf("INDEX", StringComparison.Ordinal) >= 0)
                || formulaTextMayContainArrayFunctionName formula
            if maySpillByText.Count > 100000 then maySpillByText.Clear()
            maySpillByText.[formula] <- found
            found

/// Materialize dynamic-array spills for rendering. For each formula whose
/// formula uses a spilling function (SEQUENCE/UNIQUE/SORT/FILTER/…), evaluate
/// the array and write its values into the anchor + neighbouring cells
/// (row-major). A non-empty target cell (other than the anchor) blocks the
/// spill → the anchor shows #SPILL!. Pure; returns a new model. Only formulas
/// that actually use an array function are touched — every other cell is
/// untouched, so this is safe to run after `recalculateAllFormulas` on any
/// workbook. (Interactive clear-on-recalc + spill-cell dependencies are a
/// later step; this is the render-side materialization.)
let private spillArraysCore (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                            (currentSheet: Rbt.Spreadsheet.Workbook.SheetId option)
                            (model: SpreadsheetModel) : SpreadsheetModel * bool * Set<CellAddress> =
    let previousSpill = model.Spill
    let mutable blocked = Set.empty
    let anchors = ResizeArray<CellAddress * CellData * string * Expr>()
    let arrayFormulaExpr formula =
        if not (formulaTextMaySpill formula) then
            None
        else
            match parseFormulaShared formula with
            | Ok expr when exprSpills expr -> Some expr
            | _ -> None
    model.Cells
    |> Map.iter (fun addr cell ->
        match cell.Value with
        | CellValue.Formula (formula, _) ->
            match arrayFormulaExpr formula with
            | Some expr -> anchors.Add(addr, cell, formula, expr)
            | _ -> ()
        | _ -> ())
    if anchors.Count = 0 then
        if Map.isEmpty previousSpill then
            model, false, Set.empty
        else
            let mutable updatedCells = model.Cells
            let mutable changed = false
            previousSpill
            |> Map.iter (fun addr _ ->
                match Map.tryFind addr updatedCells with
                | Some cd ->
                    match cd.Value with
                    | CellValue.Formula _ -> ()
                    | _ ->
                        updatedCells <- Map.remove addr updatedCells
                        changed <- true
                | None -> ())
            if changed then
                ({ model with Cells = updatedCells; Spill = Map.empty } |> invalidateFilterProjection), true, Set.empty
            else
                { model with Spill = Map.empty }, true, Set.empty
    else
    let mutable updatedCells = model.Cells
    let mutable spillMap = Map.empty
    let mutable changed = false
    let setCellIfChanged addr nextCell =
        match Map.tryFind addr updatedCells with
        | Some existing when existing = nextCell -> ()
        | _ ->
            updatedCells <- Map.add addr nextCell updatedCells
            changed <- true
    let removeCellIfPresent addr =
        if Map.containsKey addr updatedCells then
            updatedCells <- Map.remove addr updatedCells
            changed <- true
    // What a spill-range reference (`A1#`) reads while the arrays are laid out:
    // the arrays laid out so far as they are now, the others as they were -
    // a formula above the array it refers to sees the whole array too.
    let readsSpillRanges = anchors |> Seq.exists (fun (_, _, formula, _) -> formula.IndexOf '#' >= 0)
    let mutable viewSpill = previousSpill
    let previousByAnchor = Dictionary<CellAddress, ResizeArray<CellAddress>>()
    if readsSpillRanges then
        previousSpill
        |> Map.iter (fun spilled anchor ->
            match Portable.Dict.tryFind anchor previousByAnchor with
            | ValueSome cells -> cells.Add spilled
            | ValueNone ->
                let cells = ResizeArray<CellAddress>()
                cells.Add spilled
                previousByAnchor.[anchor] <- cells)
    for (addr, cell, formula, expr) in anchors do
        // (Its own array is not there yet: a formula that refers to its own spill range reads its cell alone.)
        if readsSpillRanges then
            match Portable.Dict.tryFind addr previousByAnchor with
            | ValueSome cells -> for spilled in cells do viewSpill <- Map.remove spilled viewSpill
            | ValueNone -> ()
        let updated = { model with Cells = updatedCells; Spill = (if readsSpillRanges then viewSpill else spillMap) }
        let arr = evalArray expr { Model = updated; Workbook = workbook; Sheet = currentSheet; CurrentCell = Some addr }
        let nr = arr.Length
        let nc = if nr > 0 then arr.[0].Length else 0
        if nr * nc <= 1 then
            let v = if nr * nc = 1 then spilledValue arr.[0].[0] else CellValue.Error Err.value
            setCellIfChanged addr { cell with Value = CellValue.Formula (formula, v) }
        else
            let occupied r c =
                if r = 0 && c = 0 then false
                else
                    let target = { Row = addr.Row + r; Col = addr.Col + c }
                    match Map.tryFind target updatedCells with
                    | Some cd ->
                        let occupiedByCell =
                            match cd.Value with
                            | CellValue.Empty -> false
                            | _ -> true
                        let reusablePreviousSpill = Map.containsKey target previousSpill
                        let occupiedByCurrentSpill = Map.containsKey target spillMap
                        (occupiedByCell && not reusablePreviousSpill) || occupiedByCurrentSpill
                    | None -> false
            let mutable collides = false
            let mutable r = 0
            while not collides && r < nr do
                let mutable c = 0
                while not collides && c < nc do
                    if occupied r c then
                        collides <- true
                    c <- c + 1
                r <- r + 1
            if collides then
                blocked <- Set.add addr blocked
                setCellIfChanged addr { cell with Value = CellValue.Formula (formula, CellValue.Error Err.spill) }
            else
                setCellIfChanged addr { cell with Value = CellValue.Formula (formula, spilledValue arr.[0].[0]) }
                for r in 0 .. nr - 1 do
                    for c in 0 .. nc - 1 do
                        if not (r = 0 && c = 0) then
                            let sa = { Row = addr.Row + r; Col = addr.Col + c }
                            setCellIfChanged sa { Value = spilledValue arr.[r].[c]; Format = cell.Format; Comment = None }
                            spillMap <- Map.add sa addr spillMap
                            if readsSpillRanges then viewSpill <- Map.add sa addr viewSpill
    previousSpill
    |> Map.iter (fun addr _ ->
        if not (Map.containsKey addr spillMap) then
            match Map.tryFind addr updatedCells with
            | Some cd ->
                match cd.Value with
                | CellValue.Formula _ -> ()
                | _ -> removeCellIfPresent addr
            | None -> ())
    let spillChanged = spillMap <> previousSpill
    if not changed && not spillChanged then
        model, false, blocked
    else
        ({ model with Cells = updatedCells; Spill = spillMap } |> invalidateFilterProjection), true, blocked

let spillArrays (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                (currentSheet: Rbt.Spreadsheet.Workbook.SheetId option)
                (model: SpreadsheetModel) : SpreadsheetModel =
    let spilled, _, _ = spillArraysCore workbook currentSheet model
    spilled

/// Recalculate every formula and (re)materialize dynamic-array spills,
/// clearing the previous spill's ghost cells. This is the interactive
/// edit→recompute step: a spilling formula (FILTER/SORT/GROUPBY/…) updates and
/// re-spills automatically when its source changes. Two recalc passes bracket
/// the spill so formulas that *read* spilled cells also settle.
let recalcAndSpill (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                   (currentSheet: Rbt.Spreadsheet.Workbook.SheetId option)
                   (model: SpreadsheetModel) : SpreadsheetModel =
    if Map.isEmpty model.Spill && not (cellsHaveFormulaCells model.Cells) then model else
    let recalculated = recalculateAllFormulas workbook currentSheet model
    let spilled, spillChanged, blocked = spillArraysCore workbook currentSheet recalculated
    if spillChanged || not (Set.isEmpty blocked) then
        recalculateFormulasKeepingBlocked blocked workbook currentSheet spilled
    else
        spilled

/// The full recalculation as it was before the workbook had one dependency
/// order (see `recalculateWorkbook`): every sheet on its own (formulas +
/// spills), in tab order, the pass repeated until nothing changes — bounded,
/// because volatile functions (RAND, NOW) never settle. One run can leave a
/// long chain through several sheets short of its values, and a circular
/// reference that crosses sheets is not marked. Kept as the reference the
/// tests check the ordered pass against, and for a host that never loaded
/// the Recalc module.
let recalculateWorkbookSettling (workbook: Rbt.Spreadsheet.Workbook.Workbook) : Rbt.Spreadsheet.Workbook.Workbook =
    let needsCalc (sheet: Rbt.Spreadsheet.Workbook.Sheet) =
        not (Map.isEmpty sheet.Model.Spill) || cellsHaveFormulaCells sheet.Model.Cells
    let formulaSheets =
        workbook.SheetOrder
        |> List.filter (fun id ->
            match Map.tryFind id workbook.Sheets with
            | Some sheet -> needsCalc sheet
            | None -> false)
    // (One sheet settles in one pass. Several: a value can travel from sheet to
    // sheet and back - a name on one sheet read by a schedule on another whose
    // totals the first sheet shows - once per pass at worst, so the passes go
    // on until one changes nothing, within a bound that volatile functions
    // and circular references across sheets, which never settle, stop at.)
    let sheetCount = List.length formulaSheets
    let maxPasses = if sheetCount <= 1 then sheetCount else min 8 (2 * sheetCount + 2)
    let mutable wb = workbook
    let mutable pass = 0
    let mutable changed = true
    while changed && pass < maxPasses do
        changed <- false
        for id in formulaSheets do
            match Map.tryFind id wb.Sheets with
            | Some sheet ->
                let recalculated = recalcAndSpill (Some wb) (Some id) sheet.Model
                if not (obj.ReferenceEquals(recalculated, sheet.Model)) then
                    wb <- Rbt.Spreadsheet.Workbook.updateSheetModel id recalculated wb
                    changed <- true
            | None -> ()
        pass <- pass + 1
    wb

/// The workbook-wide recalculation, installed by Recalc (which compiles after
/// this file): held in a record, as mutable function values are not reliable
/// in Fable.
type WorkbookPass = { Run: Rbt.Spreadsheet.Workbook.Workbook -> Rbt.Spreadsheet.Workbook.Workbook }
let mutable private workbookPass : WorkbookPass option = None
let installWorkbookPass (pass: WorkbookPass) : unit = workbookPass <- Some pass

/// Recalculate every sheet of a workbook: every formula of every sheet in
/// one dependency order that follows references across sheets, defined names,
/// structured references and the cells dynamic arrays spill into, so that one
/// run reaches the values (see Recalc.recalculateAll). Formulas on a circular
/// reference - within a sheet or across sheets - keep their values.
let recalculateWorkbook (workbook: Rbt.Spreadsheet.Workbook.Workbook) : Rbt.Spreadsheet.Workbook.Workbook =
    if workbookPass.IsNone then
#if !FABLE_COMPILER
        // (A .NET library initialises a file when one of its values is first
        // read: reading Recalc.workbookPassInstalled installs the pass.)
        try
            let recalcModule = typeof<Expr>.Assembly.GetType "Rbt.Spreadsheet.Recalc"
            if not (isNull recalcModule) then
                let installed = recalcModule.GetProperty "workbookPassInstalled"
                if not (isNull installed) then installed.GetValue(null, null) |> ignore
        with _ -> ()
#else
        ()
#endif
    match workbookPass with
    | Some pass -> pass.Run workbook
    | None -> recalculateWorkbookSettling workbook

/// Evaluate a formula written for one cell as if it stood `rows` / `cols`
/// away, at `cell`: relative references move, `$`-locked ones stay. This is
/// how one conditional-format or validation formula serves a whole range.
/// The result is unwrapped; an array result yields its top-left element.
let evaluateFormulaShifted (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                           (sheet: Rbt.Spreadsheet.Workbook.SheetId option)
                           (formula: string) (rows: int) (cols: int)
                           (model: SpreadsheetModel) (cell: CellAddress) : CellValue =
    match parseFormulaShared formula with
    | Result.Error _ -> CellValue.Error Err.syntax
    | Ok expr ->
        let expr = if rows = 0 && cols = 0 then expr else shiftExpr rows cols expr
        let ctx = { Model = model; Workbook = workbook; Sheet = sheet; CurrentCell = Some cell }
        if exprSpillsStatically expr then
            let arr = evalArray expr ctx
            if arr.Length > 0 && arr.[0].Length > 0 then unwrapValue arr.[0].[0] else CellValue.Error Err.value
        else unwrapValue (evalCell expr ctx)

/// Every value a formula gives, row by row, evaluated as `evaluateFormulaShifted`
/// evaluates it: the cells of a reference (a range, OFFSET, INDEX:INDEX, a
/// defined name that is one of these), the items of an array, one value for
/// anything else. None for a formula that cannot be read.
let evaluateFormulaValuesShifted (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                                 (sheet: Rbt.Spreadsheet.Workbook.SheetId option)
                                 (formula: string) (rows: int) (cols: int)
                                 (model: SpreadsheetModel) (cell: CellAddress) : CellValue list option =
    match parseFormulaShared formula with
    | Result.Error _ -> None
    | Ok expr ->
        let expr = if rows = 0 && cols = 0 then expr else shiftExpr rows cols expr
        let ctx = { Model = model; Workbook = workbook; Sheet = sheet; CurrentCell = Some cell }
        try
            let arr = evalArray expr ctx
            Some [ for row in arr do
                    for value in row -> unwrapValue value ]
        with _ -> None

/// The values a formula gives, as rows (a range or an array keeps its shape; anything else is one row
/// of one value), evaluated as `evaluateFormulaShifted` evaluates it. None for a formula that cannot be read.
let evaluateFormulaRows (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                        (sheet: Rbt.Spreadsheet.Workbook.SheetId option)
                        (formula: string) (model: SpreadsheetModel) (cell: CellAddress) : CellValue[][] option =
    match parseFormulaShared formula with
    | Result.Error _ -> None
    | Ok expr ->
        let ctx = { Model = model; Workbook = workbook; Sheet = sheet; CurrentCell = Some cell }
        try Some (evalArray expr ctx |> Array.map (Array.map unwrapValue))
        with _ -> None

/// The area a formula refers to - a range, or a function that gives a reference (OFFSET, INDEX, INDIRECT) - as
/// (first row, first column, last row, last column), evaluated on `sheet` of the workbook; None for anything else.
/// (A print area written as a formula: vertex42's `OFFSET(MortgageCalculator!$A$1,0,0,ROW(...)+1+$K$25,...)`.)
let evaluateFormulaArea (workbook: Rbt.Spreadsheet.Workbook.Workbook option)
                        (sheet: Rbt.Spreadsheet.Workbook.SheetId option)
                        (formula: string) (model: SpreadsheetModel) : (int * int * int * int) option =
    match parseFormulaShared formula with
    | Result.Error _ -> None
    | Ok expr ->
        let ctx = { Model = model; Workbook = workbook; Sheet = sheet; CurrentCell = None }
        try
            // (`asReference` of a call that is no function: names, OFFSET / INDIRECT / INDEX as references.)
            match referenceRect (resolveReferenceFunction "" [ expr ] ctx) ctx with
            | Ok (_, r0, c0, r1, c1) -> Some (min r0 r1, min c0 c1, max r0 r1, max c0 c1)
            | Result.Error _ -> None
        with _ -> None

/// Installs what Model.fs needs from this file, which compiles after it: the
/// formula rewrites for paste, fill and row/column edits, and the evaluator
/// for validation and conditional-format formulas. Model.fs reads this value
/// to force the installation on hosts that have not touched this module yet.
let formulaHooksInstalled : bool =
    installFormulaHooks
        { Shift = shiftFormulaText
          Adjust = adjustFormulaTextForEdit
          Move = moveFormulaText
          Evaluate = fun formula rows cols model cell -> evaluateFormulaShifted None None formula rows cols model cell
          Sheets = rewriteFormulaTextForSheets
          Display = fun date1904 format value -> Rbt.Spreadsheet.CellText.cellDisplayTextIn date1904 None 1.0e6 format value }
    true
