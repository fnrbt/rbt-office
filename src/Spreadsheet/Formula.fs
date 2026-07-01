module RedBat.Spreadsheet.Formula

open System
open System.Collections.Generic
open System.Globalization
open RedBat.Spreadsheet.Model

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

let private addBoundedCacheEntry maxEntries (cache: Dictionary<'K, 'V>) (order: Queue<'K>) key value =
    while cache.Count >= maxEntries && order.Count > 0 do
        cache.Remove(order.Dequeue()) |> ignore
    if cache.Count >= maxEntries then
        cache.Clear()
        order.Clear()
    cache.[key] <- value
    order.Enqueue key
    value

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
    /// Every recognized error literal, for the lexer.
    let all = [ "#DIV/0!"; "#VALUE!"; "#REF!"; "#NAME?"; "#NUM!"; "#N/A"; "#NULL!" ]

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
    /// A bare `#` that isn't part of an error literal — the spill operator
    /// suffix in `A1#`.
    | THash
    | TComma | TColon | TBang
    | TEnd

let private tokenizeFrom (input: string) startIndex : Token array =
    let n = input.Length
    let tokens = System.Collections.Generic.List<Token>(max 1 (n - startIndex + 1))
    let mutable i = startIndex
    while i < n do
        let c = input.[i]
        if Char.IsWhiteSpace c then
            i <- i + 1
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
            tokens.Add(TNumber (Double.Parse(text, CultureInfo.InvariantCulture)))
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
            tokens.Add(TIdent (input.Substring(start, i - start)))
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
                    let inner = input.Substring(start, (min j n) - start)
                    tokens.Add(TTableSpec inner)
                    i <- (if j < n then j + 1 else n)
                | ']' -> tokens.Add(TRBracket); i <- i + 1
                | ',' -> tokens.Add(TComma); i <- i + 1
                | ':' -> tokens.Add(TColon); i <- i + 1
                | '!' -> tokens.Add(TBang); i <- i + 1
                | _ -> i <- i + 1  // skip unknown
    tokens.Add(TEnd)
    tokens.ToArray()

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

type private Parser = {
    Tokens: Token[]
    mutable Pos: int
    mutable PrimaryDepth: int
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
        | TLParen ->
            p.Pos <- p.Pos + 1
            let e = parseComparison p
            expectP p TRParen
            e
        | TSheet sheetName ->
            p.Pos <- p.Pos + 1
            expectP p TBang
            parseSheetQualified p (Some sheetName)
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
                if equalsOrdinalIgnoreCase ident "LAMBDA" && peekP p = TLParen then
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
    match nextP p with
    | TIdent ident ->
        match parseRefToken ident with
        | Some refAst ->
            let withSheet = { refAst with Sheet = sheet }
            match peekP p with
            | TColon ->
                p.Pos <- p.Pos + 1
                match nextP p with
                | TIdent endIdent ->
                    match parseRefToken endIdent with
                    | Some endRef -> RangeRef(withSheet, { endRef with Sheet = sheet })
                    | None -> failwithf "Invalid cell ref after ':': %s" endIdent
                | other -> failwithf "Expected cell ref after ':', got %A" other
            | _ -> CellRef withSheet
        | None ->
            // Sheet-qualified defined name (rare). Treat as NameRef.
            NameRef ident
    | other -> failwithf "Expected reference after sheet '!', got %A" other

and private parseArgs (p: Parser) : Expr list =
    if peekP p = TRParen then []
    else
        let rec loop acc =
            let e = parseComparison p
            match peekP p with
            | TComma -> p.Pos <- p.Pos + 1; loop (e :: acc)
            | _ -> List.rev (e :: acc)
        loop []

and private parseRange (p: Parser) : Expr =
    let left = parsePrimary p
    match peekP p, left with
    | TColon, CellRef a ->
        p.Pos <- p.Pos + 1
        match parsePrimary p with
        | CellRef b -> RangeRef(a, b)
        | _ -> failwith "Expected cell reference after ':'"
    | _ -> left

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
let parseFormula (input: string) : Result<Expr, string> =
    try
        let start = if input.Length > 0 && input.[0] = '=' then 1 else 0
        let tokens = tokenizeFrom input start
        let p = { Tokens = tokens; Pos = 0; PrimaryDepth = 0 }
        let e = parseComparison p
        match peekP p with
        | TEnd -> Ok e
        | t -> Result.Error (sprintf "Unexpected trailing token %A" t)
    with ex -> Result.Error ex.Message

let internal parsedFormulaResultRetainedTextCharCount maxChars (parsed: Result<Expr, string>) =
    let addCapped (amount: int) (total: int) =
        if amount <= 0 then total
        elif total > maxChars || amount > maxChars - total then maxChars + 1
        else total + amount

    let addOptionString (value: string option) total =
        match value with
        | Some value -> addCapped value.Length total
        | None -> total

    let addRefText (ref: RefAst) total =
        addOptionString ref.Sheet total

    let addTableText (table: TableSel) total =
        total
        |> addCapped table.Table.Length
        |> addOptionString table.ColStart
        |> addOptionString table.ColEnd

    match parsed with
    | Result.Error error -> min (maxChars + 1) error.Length
    | Ok expr ->
        let stack = ResizeArray<Expr>()
        stack.Add expr
        let mutable total = 0
        while total <= maxChars && stack.Count > 0 do
            let last = stack.Count - 1
            let expr = stack.[last]
            stack.RemoveAt last
            match expr with
            | Number _
            | Bool _
            | ArrayLit _ -> ()
            | Str value
            | ErrorLit value
            | NameRef value ->
                total <- addCapped value.Length total
            | CellRef ref
            | SpillRef ref ->
                total <- addRefText ref total
            | RangeRef (left, right) ->
                total <- addRefText left total
                total <- addRefText right total
            | TableRef table ->
                total <- addTableText table total
            | UnaryOp (op, operand) ->
                total <- addCapped op.Length total
                stack.Add operand
            | BinOp (left, op, right) ->
                total <- addCapped op.Length total
                stack.Add left
                stack.Add right
            | Function (name, args) ->
                total <- addCapped name.Length total
                let mutable remaining = args
                while total <= maxChars && not remaining.IsEmpty do
                    stack.Add remaining.Head
                    remaining <- remaining.Tail
        total

let private parsedFormulaCacheMaxEntries = 32768
let private parsedFormulaCacheMaxTotalChars = 5000000
let private parsedFormulaCache = Dictionary<string, Result<Expr, string>>()
let private parsedFormulaCacheOrder = Queue<string>()
let mutable private parsedFormulaCacheTotalChars = 0
type private ParsedFormulaFastCacheEntry = {
    Formula: string
    Parsed: Result<Expr, string>
}
let private parsedFormulaFastSlots = Array.zeroCreate<ParsedFormulaFastCacheEntry> 256

let private clearParsedFormulaFastSlotsIn (slots: ParsedFormulaFastCacheEntry array) =
    for index = 0 to slots.Length - 1 do
        slots.[index] <- Unchecked.defaultof<ParsedFormulaFastCacheEntry>

let private clearParsedFormulaFastSlots () =
    clearParsedFormulaFastSlotsIn parsedFormulaFastSlots

let private parsedFormulaFastSlotIn (slots: ParsedFormulaFastCacheEntry array) (input: string) =
    let hashCode =
        if isNull input then 0
        else System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(input)
    hashCode &&& (slots.Length - 1)

let private setParsedFormulaFastSlotIn (slots: ParsedFormulaFastCacheEntry array) input parsed =
    slots.[parsedFormulaFastSlotIn slots input] <- {
        Formula = input
        Parsed = parsed
    }

let private setParsedFormulaFastSlot input parsed =
    setParsedFormulaFastSlotIn parsedFormulaFastSlots input parsed

let private tryGetParsedFormulaFastSlotIn (slots: ParsedFormulaFastCacheEntry array) input =
    let fast = slots.[parsedFormulaFastSlotIn slots input]
    if not (obj.ReferenceEquals(fast, null))
       && obj.ReferenceEquals(fast.Formula, input) then
        ValueSome fast.Parsed
    else
        ValueNone

let private tryGetParsedFormulaFastSlot input =
    tryGetParsedFormulaFastSlotIn parsedFormulaFastSlots input

let private parsedFormulaAddCapped (amount: int) (total: int) =
    if amount <= 0 then total
    elif total > parsedFormulaCacheMaxTotalChars || amount > parsedFormulaCacheMaxTotalChars - total then
        parsedFormulaCacheMaxTotalChars + 1
    else
        total + amount

let private parsedFormulaCacheEntryChars (input: string) (parsed: Result<Expr, string>) =
    0
    |> parsedFormulaAddCapped input.Length
    |> parsedFormulaAddCapped (parsedFormulaResultRetainedTextCharCount parsedFormulaCacheMaxTotalChars parsed)

let private removeParsedFormulaCacheEntry input =
    let mutable previous = Unchecked.defaultof<Result<Expr, string>>
    if parsedFormulaCache.TryGetValue(input, &previous) then
        parsedFormulaCache.Remove input |> ignore
        parsedFormulaCacheTotalChars <- max 0 (parsedFormulaCacheTotalChars - parsedFormulaCacheEntryChars input previous)
        clearParsedFormulaFastSlots ()

let private addParsedFormulaCacheEntry (input: string) (parsed: Result<Expr, string>) =
    let entryChars = parsedFormulaCacheEntryChars input parsed
    if entryChars > parsedFormulaCacheMaxTotalChars then
        removeParsedFormulaCacheEntry input
        parsed
    else
        let mutable previous = Unchecked.defaultof<Result<Expr, string>>
        if parsedFormulaCache.TryGetValue(input, &previous) then
            parsedFormulaCacheTotalChars <- max 0 (parsedFormulaCacheTotalChars - parsedFormulaCacheEntryChars input previous)
            clearParsedFormulaFastSlots ()
        else
            parsedFormulaCacheOrder.Enqueue input
        parsedFormulaCache.[input] <- parsed
        parsedFormulaCacheTotalChars <- parsedFormulaCacheTotalChars + entryChars
        while (parsedFormulaCache.Count > parsedFormulaCacheMaxEntries || parsedFormulaCacheTotalChars > parsedFormulaCacheMaxTotalChars)
              && parsedFormulaCacheOrder.Count > 0 do
            let oldest = parsedFormulaCacheOrder.Dequeue()
            if oldest <> input then
                removeParsedFormulaCacheEntry oldest
        if parsedFormulaCache.Count > parsedFormulaCacheMaxEntries || parsedFormulaCacheTotalChars > parsedFormulaCacheMaxTotalChars then
            let current = parsedFormulaCache.[input]
            parsedFormulaCache.Clear()
            parsedFormulaCacheOrder.Clear()
            clearParsedFormulaFastSlots ()
            parsedFormulaCacheOrder.Enqueue input
            parsedFormulaCache.[input] <- current
            parsedFormulaCacheTotalChars <- entryChars
        if parsedFormulaCache.ContainsKey input then
            setParsedFormulaFastSlot input parsed
        parsed

let private parseFormulaShared (input: string) : Result<Expr, string> =
    match tryGetParsedFormulaFastSlot input with
    | ValueSome cached -> cached
    | ValueNone ->
        let mutable cached = Unchecked.defaultof<Result<Expr, string>>
        if parsedFormulaCache.TryGetValue(input, &cached) then
            setParsedFormulaFastSlot input cached
            cached
        else
            cached <- parseFormula input
            addParsedFormulaCacheEntry input cached

/// Resolve a defined name to its final parsed expression, scoped to the
/// current sheet (a sheet-local name shadows a workbook-global one). A
/// name's refers-to text is itself a formula — usually a range, sometimes
/// a constant or another name; name→name hops are followed with a cycle
/// guard. Returns None when the name is unknown, won't parse, or cycles.
let resolveNameExpr (name: string) (currentSheet: RedBat.Spreadsheet.Workbook.SheetId option)
                    (workbook: RedBat.Spreadsheet.Workbook.Workbook) : Expr option =
    let rec resolve (n: string) (seen: Set<string>) (hops: int) =
        let key = n.ToUpperInvariant()
        if hops > 16 || Set.contains key seen then None
        else
            match RedBat.Spreadsheet.Workbook.getName currentSheet n workbook with
            | None -> None
            | Some refersTo ->
                let body = if refersTo.StartsWith "=" then refersTo else "=" + refersTo
                match parseFormula body with
                | Result.Error _ -> None
                | Ok (NameRef inner) -> resolve inner (Set.add key seen) (hops + 1)
                | Ok expr -> Some expr
    resolve name Set.empty 0

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
        n.ToString("G15", CultureInfo.InvariantCulture)

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
        match Double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture) with
        | true, n -> Ok n
        | _ ->
            // Excel also coerces "TRUE"/"FALSE" text in arithmetic? No — it
            // does not; only the boolean *value* coerces. Text TRUE → #VALUE!.
            Result.Error Err.value
    | CellValue.Formula _ -> Result.Error Err.value  // unreachable after unwrap

let inline private tryAggregateNumber (value: CellValue) : float voption =
    match unwrapValue value with
    | CellValue.Number n -> ValueSome n
    | CellValue.Boolean b -> ValueSome (if b then 1.0 else 0.0)
    | CellValue.Empty -> ValueSome 0.0
    | CellValue.String s ->
        let struct (start, length) = trimBounds s
        match Double.TryParse(trimmedNumberText s start length, NumberStyles.Float, CultureInfo.InvariantCulture) with
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
            match Double.TryParse(trimmedNumberText s start length, NumberStyles.Float, CultureInfo.InvariantCulture) with
            | true, n -> Ok (n <> 0.0)
            | _ -> Result.Error Err.value
    | CellValue.Formula _ -> Result.Error Err.value

// ---------------------------------------------------------------------------
// Reference resolution + evaluation
// ---------------------------------------------------------------------------

let private normalizeRange (a: RefAst) (b: RefAst) =
    (min a.Row b.Row, min a.Col b.Col, max a.Row b.Row, max a.Col b.Col)

/// Set of cell addresses an expression references *within the current
/// model*. Cross-sheet refs (`r.Sheet = Some _`) are deliberately excluded
/// — single-model recalc ordering can't span sheets, so they're treated as
/// external constants. A `NameRef` is resolved through the workbook (when
/// one is supplied) and its target cells contribute to the set, so
/// `=SUM(MyRange)` registers a dependency on the range's members. Defined
/// names resolve at workbook-global scope here: a missed sheet-local name
/// only affects recalc *ordering*, never the computed values, since `eval`
/// resolves names with the correct scope. Used by the dependency tracker.
let referencesInExpr (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                     (expr: Expr) : Set<CellAddress> =
    let refs = HashSet<CellAddress>()
    let stack = ResizeArray<struct (Set<string> * Expr)>()
    stack.Add(struct (Set.empty, expr))
    while stack.Count > 0 do
        let last = stack.Count - 1
        let struct (seen, e) = stack.[last]
        stack.RemoveAt last
        match e with
        | Number _ | Str _ | Bool _ | ErrorLit _ | TableRef _ | ArrayLit _ -> ()
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
            let sr, sc, er, ec = normalizeRange a b
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
let referencesQualified (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                        (homeSheet: RedBat.Spreadsheet.Workbook.SheetId)
                        (expr: Expr) : Set<RedBat.Spreadsheet.Workbook.SheetCell> =
    let sheetIdOf (name: string option) =
        match name with
        | None -> Some homeSheet
        | Some n -> workbook |> Option.bind (RedBat.Spreadsheet.Workbook.resolveSheetName n)
    let refs = HashSet<RedBat.Spreadsheet.Workbook.SheetCell>()
    let stack = ResizeArray<struct (Set<string> * Expr)>()
    stack.Add(struct (Set.empty, expr))
    while stack.Count > 0 do
        let last = stack.Count - 1
        let struct (seen, e) = stack.[last]
        stack.RemoveAt last
        match e with
        | Number _ | Str _ | Bool _ | ErrorLit _ | TableRef _ | ArrayLit _ -> ()
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
                let sr, sc, er, ec = normalizeRange a b
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
    { Sheet: RedBat.Spreadsheet.Workbook.SheetId; R0: int; C0: int; R1: int; C1: int }

let rectContainsCell (rect: DepRect) (sc: RedBat.Spreadsheet.Workbook.SheetCell) =
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
let referencesRanged (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                     (homeSheet: RedBat.Spreadsheet.Workbook.SheetId)
                     (expr: Expr) : Set<RedBat.Spreadsheet.Workbook.SheetCell> * DepRect list =
    let sheetIdOf (name: string option) =
        match name with
        | None -> Some homeSheet
        | Some n -> workbook |> Option.bind (RedBat.Spreadsheet.Workbook.resolveSheetName n)
    let cells = System.Collections.Generic.HashSet<RedBat.Spreadsheet.Workbook.SheetCell>()
    let rects = ResizeArray<DepRect>()
    let stack = ResizeArray<struct (Set<string> * Expr)>()
    let push seen expr = stack.Add(struct (seen, expr))
    push Set.empty expr
    while stack.Count > 0 do
        let last = stack.Count - 1
        let struct (seen, e) = stack.[last]
        stack.RemoveAt last
        match e with
        | Number _ | Str _ | Bool _ | ErrorLit _ | TableRef _ | ArrayLit _ -> ()
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
                if (er - sr + 1) * (ec - sc + 1) > rangeDepThreshold then
                    rects.Add { Sheet = sid; R0 = sr; C0 = sc; R1 = er; C1 = ec }
                else
                    for row in sr .. er do
                        for col in sc .. ec do
                            cells.Add { Sheet = sid; Cell = { Row = row; Col = col } } |> ignore
            | None -> ()
        | UnaryOp (_, inner) -> push seen inner
        | BinOp (l, _, r) ->
            push seen l
            push seen r
        | Function (_, args) ->
            for arg in args do
                push seen arg
    let mutable cellSet = Set.empty
    for cell in cells do
        cellSet <- Set.add cell cellSet
    (cellSet, resizeArrayToList rects)

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
    | Number _ | Str _ | Bool _ | ErrorLit _ | NameRef _ | TableRef _ | ArrayLit _ -> expr
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
    | Number _ | Str _ | Bool _ | ErrorLit _ | CellRef _ | RangeRef _ | TableRef _ | ArrayLit _ | SpillRef _ -> expr
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
                bind subst restParts []
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
        stack.RemoveAt last

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
        stack.RemoveAt last
        match expr with
        | Function (name, args) ->
            if equalsOrdinalIgnoreCase name "RAND"
               || equalsOrdinalIgnoreCase name "RANDBETWEEN"
               || equalsOrdinalIgnoreCase name "NOW"
               || equalsOrdinalIgnoreCase name "TODAY"
               || equalsOrdinalIgnoreCase name "OFFSET"
               || equalsOrdinalIgnoreCase name "INDIRECT" then
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
        stack.RemoveAt last
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
                         (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                         : SpreadsheetModel option =
    match sheet with
    | None -> Some model
    | Some name ->
        match workbook with
        | None -> None
        | Some wb ->
            RedBat.Spreadsheet.Workbook.resolveSheetName name wb
            |> Option.bind (fun sid -> Map.tryFind sid wb.Sheets)
            |> Option.map (fun s -> s.Model)

/// Read a single cell's value, resolving cross-sheet references.
let private readCell (r: RefAst) (model: SpreadsheetModel)
                     (workbook: RedBat.Spreadsheet.Workbook.Workbook option) : CellValue =
    match resolveModel r.Sheet model workbook with
    | Some m -> getCellValue { Row = r.Row; Col = r.Col } m |> unwrapValue
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

/// Expand a RangeRef to the flat list of unwrapped cell values, row-major.
/// The sheet qualifier on the start ref selects the target sheet.
let iterRangeRef (a: RefAst) (b: RefAst) (model: SpreadsheetModel)
                 (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                 (emit: CellValue -> unit) : unit =
    match resolveModel a.Sheet model workbook with
    | Some m ->
        let sr, sc, er, ec = normalizeRange a b
        for row in sr .. er do
            for col in sc .. ec do
                emit (getCellValue { Row = row; Col = col } m |> unwrapValue)
    | None -> emit (err Err.ref)

let expandRangeRef (a: RefAst) (b: RefAst) (model: SpreadsheetModel)
                   (workbook: RedBat.Spreadsheet.Workbook.Workbook option) : CellValue list =
    let values = ResizeArray<CellValue>()
    iterRangeRef a b model workbook (fun value -> values.Add value)
    resizeArrayToList values

let firstRangeRefValue (a: RefAst) (b: RefAst) (model: SpreadsheetModel)
                       (workbook: RedBat.Spreadsheet.Workbook.Workbook option) : CellValue =
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
                 (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                 (sheet: RedBat.Spreadsheet.Workbook.SheetId option)
                 (currentCell: CellAddress option) : Result<int * int * int * int, string> =
    match workbook with
    | None -> Result.Error Err.ref
    | Some wb ->
        let tableOpt =
            if sel.Table <> "" then RedBat.Spreadsheet.Workbook.getTable sel.Table wb
            else currentCell |> Option.bind (fun c -> RedBat.Spreadsheet.Workbook.findTableContaining sheet c wb)
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
                    match RedBat.Spreadsheet.Workbook.tableColumnIndex t cs with
                    | None -> None
                    | Some ci ->
                        match sel.ColEnd with
                        | None -> Some (t.StartCol + ci, t.StartCol + ci)
                        | Some ce ->
                            RedBat.Spreadsheet.Workbook.tableColumnIndex t ce
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

/// Resolve a structured reference (`Table[[#Totals],[Amount]]` etc.) to the
/// flat list of cell values it covers (row-major over the selected rows ×
/// columns). `sheet`/`currentCell` resolve this-row and unqualified refs.
let iterTableSel (sel: TableSel)
                 (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                 (sheet: RedBat.Spreadsheet.Workbook.SheetId option)
                 (currentCell: CellAddress option)
                 (model: SpreadsheetModel)
                 (emit: CellValue -> unit) : unit =
    match tableSelArea sel workbook sheet currentCell with
    | Result.Error e -> emit (err e)
    | Result.Ok (r0, c0, r1, c1) ->
        for row in r0 .. r1 do
            for col in c0 .. c1 do
                emit (getCellValue { Row = row; Col = col } model |> unwrapValue)

let expandTableSel (sel: TableSel)
                   (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                   (sheet: RedBat.Spreadsheet.Workbook.SheetId option)
                   (currentCell: CellAddress option)
                   (model: SpreadsheetModel) : CellValue list =
    let values = ResizeArray<CellValue>()
    iterTableSel sel workbook sheet currentCell model (fun value -> values.Add value)
    resizeArrayToList values

let firstTableSelValue (sel: TableSel)
                       (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                       (sheet: RedBat.Spreadsheet.Workbook.SheetId option)
                       (currentCell: CellAddress option)
                       (model: SpreadsheetModel) : CellValue =
    match tableSelArea sel workbook sheet currentCell with
    | Result.Error e -> err e
    | Result.Ok (r0, c0, r1, _) ->
        if r0 > r1 then err Err.ref
        else getCellValue { Row = r0; Col = c0 } model |> unwrapValue

// --- Date helpers -----------------------------------------------------------
// Excel serial dates: serial 1 = 1900-01-01. Excel's phantom 1900-02-29 (the
// 1900-leap-year bug) makes the simple offset exact for every date from
// 1900-03-01 (serial 61) onward, which covers every realistic spreadsheet.
let private excelEpoch = DateTime(1899, 12, 30)
let private serialToDate (s: float) : DateTime = excelEpoch.AddDays(s)
let private dateToSerial (d: DateTime) : float = (d - excelEpoch).TotalDays

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

// --- Criteria parsing (SUMIF / COUNTIF / …) ---------------------------------
let private appendRegexLiteralChar (sb: Text.StringBuilder) (ch: char) =
    match ch with
    | '\\' | '^' | '$' | '.' | '|' | '?' | '*' | '+' | '(' | ')' | '[' | ']' | '{' | '}' ->
        sb.Append('\\').Append(ch) |> ignore
    | _ ->
        sb.Append(ch) |> ignore

/// Translate Excel's `*` / `?` wildcards into a case-insensitive regex.
type private WildcardRegexCacheEntry = {
    Regex: System.Text.RegularExpressions.Regex
    PatternChars: int
}

let private wildcardRegexCache = Dictionary<string, WildcardRegexCacheEntry>()
let private wildcardRegexCacheOrder = Queue<string>()
let private wildcardRegexCacheMaxEntries = 1024
let private wildcardRegexCacheMaxTotalPatternChars = 1000000
let mutable private wildcardRegexCacheTotalPatternChars = 0

let private wildcardRegexEntryPatternChars (pattern: string) regexPatternLength =
    let total = pattern.Length + regexPatternLength
    if total <= 0 then 1 else total

let private wildcardRegexPatternCharCount (pattern: string) =
    let mutable total = 2 // ^ and $
    let mutable i = 0
    while i < pattern.Length do
        match pattern.[i] with
        | '~' when i + 1 < pattern.Length ->
            let ch = pattern.[i + 1]
            total <-
                total
                + match ch with
                  | '\\' | '^' | '$' | '.' | '|' | '?' | '*' | '+' | '(' | ')' | '[' | ']' | '{' | '}' -> 2
                  | _ -> 1
            i <- i + 2
        | '*' ->
            total <- total + 2
            i <- i + 1
        | '?' ->
            total <- total + 1
            i <- i + 1
        | ch ->
            total <-
                total
                + match ch with
                  | '\\' | '^' | '$' | '.' | '|' | '?' | '*' | '+' | '(' | ')' | '[' | ']' | '{' | '}' -> 2
                  | _ -> 1
            i <- i + 1
    total

let private addWildcardRegexCacheEntry (pattern: string) regexPatternLength regex =
    let entry = {
        Regex = regex
        PatternChars = wildcardRegexEntryPatternChars pattern regexPatternLength
    }
    if entry.PatternChars > wildcardRegexCacheMaxTotalPatternChars then
        let mutable previous = Unchecked.defaultof<WildcardRegexCacheEntry>
        if wildcardRegexCache.TryGetValue(pattern, &previous) then
            wildcardRegexCache.Remove pattern |> ignore
            wildcardRegexCacheTotalPatternChars <- max 0 (wildcardRegexCacheTotalPatternChars - previous.PatternChars)
        regex
    else
        let mutable previous = Unchecked.defaultof<WildcardRegexCacheEntry>
        if wildcardRegexCache.TryGetValue(pattern, &previous) then
            wildcardRegexCacheTotalPatternChars <- max 0 (wildcardRegexCacheTotalPatternChars - previous.PatternChars)
        else
            wildcardRegexCacheOrder.Enqueue pattern
        wildcardRegexCache.[pattern] <- entry
        wildcardRegexCacheTotalPatternChars <- wildcardRegexCacheTotalPatternChars + entry.PatternChars
        while (wildcardRegexCache.Count > wildcardRegexCacheMaxEntries || wildcardRegexCacheTotalPatternChars > wildcardRegexCacheMaxTotalPatternChars)
              && wildcardRegexCacheOrder.Count > 0 do
            let oldest = wildcardRegexCacheOrder.Dequeue()
            if oldest <> pattern then
                let mutable oldestEntry = Unchecked.defaultof<WildcardRegexCacheEntry>
                if wildcardRegexCache.TryGetValue(oldest, &oldestEntry) then
                    wildcardRegexCache.Remove oldest |> ignore
                    wildcardRegexCacheTotalPatternChars <- max 0 (wildcardRegexCacheTotalPatternChars - oldestEntry.PatternChars)
        if wildcardRegexCache.Count > wildcardRegexCacheMaxEntries || wildcardRegexCacheTotalPatternChars > wildcardRegexCacheMaxTotalPatternChars then
            let current = wildcardRegexCache.[pattern]
            wildcardRegexCache.Clear()
            wildcardRegexCacheOrder.Clear()
            wildcardRegexCacheOrder.Enqueue pattern
            wildcardRegexCache.[pattern] <- current
            wildcardRegexCacheTotalPatternChars <- current.PatternChars
        regex

let private wildcardRegex (pattern: string) =
    let mutable cached = Unchecked.defaultof<WildcardRegexCacheEntry>
    if wildcardRegexCache.TryGetValue(pattern, &cached) then
        cached.Regex
    else
        let sb = System.Text.StringBuilder()
        sb.Append('^') |> ignore
        let mutable i = 0
        while i < pattern.Length do
            match pattern.[i] with
            | '~' when i + 1 < pattern.Length ->
                // ~* / ~? escape the wildcard.
                appendRegexLiteralChar sb pattern.[i + 1]
                i <- i + 2
            | '*' -> sb.Append(".*") |> ignore; i <- i + 1
            | '?' -> sb.Append('.') |> ignore; i <- i + 1
            | c -> appendRegexLiteralChar sb c; i <- i + 1
        sb.Append('$') |> ignore
        let regexPattern = sb.ToString()
        let regex =
            System.Text.RegularExpressions.Regex(
                regexPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        addWildcardRegexCacheEntry pattern regexPattern.Length regex

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

let private tryCriteriaNumber value =
    match unwrapValue value with
    | CellValue.Number n -> ValueSome n
    | CellValue.Boolean b -> ValueSome (if b then 1.0 else 0.0)
    | CellValue.Empty -> ValueSome 0.0
    | CellValue.String s ->
        let struct (start, length) = trimBounds s
        match Double.TryParse(trimmedNumberText s start length, NumberStyles.Float, CultureInfo.InvariantCulture) with
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
            match Double.TryParse(trimmedNumberText s start length, NumberStyles.Float, CultureInfo.InvariantCulture) with
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
    match unwrapValue criteria with
    | CellValue.Number cn ->
        fun v -> match tryCriteriaNumber v with ValueSome x -> x = cn | ValueNone -> false
    | CellValue.Boolean cb ->
        fun v -> match tryCriteriaBool v with ValueSome x -> x = cb | ValueNone -> false
    | CellValue.String raw ->
        let op, restStart = parseCriteriaOperator raw
        let rest = if restStart = 0 then raw else raw.Substring(restStart)
        match Double.TryParse(rest, NumberStyles.Float, CultureInfo.InvariantCulture) with
        | true, num ->
            fun v ->
                match tryCriteriaNumber v with
                | ValueSome x ->
                    match op with
                    | CriteriaGe -> x >= num
                    | CriteriaLe -> x <= num
                    | CriteriaGt -> x > num
                    | CriteriaLt -> x < num
                    | CriteriaNe -> x <> num
                    | CriteriaEq -> x = num
                | ValueNone -> false
        | _ ->
            match op with
            | CriteriaGt | CriteriaLt | CriteriaGe | CriteriaLe ->
                fun v ->
                    match criteriaTextCompare rest v with
                    | ValueSome c ->
                        match op with
                        | CriteriaGt -> c > 0
                        | CriteriaLt -> c < 0
                        | CriteriaGe -> c >= 0
                        | _ -> c <= 0
                    | ValueNone -> false
            | CriteriaNe ->
                if hasCriteriaWildcard rest then
                    let rx = wildcardRegex rest
                    fun v -> match coerceText v with Ok t -> not (rx.IsMatch t) | _ -> true
                else
                    fun v -> match criteriaTextCompare rest v with ValueSome c -> c <> 0 | ValueNone -> true
            | CriteriaEq ->
                if hasCriteriaWildcard rest then
                    let rx = wildcardRegex rest
                    fun v -> match coerceText v with Ok t -> rx.IsMatch t | _ -> false
                else
                    fun v -> match criteriaTextCompare rest v with ValueSome c -> c = 0 | ValueNone -> false
    | CellValue.Empty ->
        fun v -> match unwrapValue v with CellValue.Empty -> true | _ -> false
    | _ -> fun _ -> false

let private criteriaPredicateCacheMaxEntries = 2048
let private criteriaPredicateCacheMaxTotalTextChars = 1000000

type CriteriaPredicateCache() =
    let values = Dictionary<CellValue, CellValue -> bool>()
    let textChars = Dictionary<CellValue, int>()
    let order = Queue<CellValue>()
    let mutable totalTextChars = 0

    member internal _.Values = values
    member internal _.TextChars = textChars
    member internal _.Order = order
    member internal _.TotalTextChars
        with get () = totalTextChars
        and set value = totalTextChars <- value

let private criteriaPredicateCacheAddCapped amount total =
    if amount <= 0 then total
    elif total > criteriaPredicateCacheMaxTotalTextChars
         || amount > criteriaPredicateCacheMaxTotalTextChars - total then
        criteriaPredicateCacheMaxTotalTextChars + 1
    else
        total + amount

let private criteriaValueTextCharCount value =
    let mutable total = 0
    let mutable current = value
    let mutable doneCounting = false
    while total <= criteriaPredicateCacheMaxTotalTextChars && not doneCounting do
        match current with
        | CellValue.String value
        | CellValue.Error value ->
            total <- criteriaPredicateCacheAddCapped value.Length total
            doneCounting <- true
        | CellValue.Formula (formula, cachedValue) ->
            total <- criteriaPredicateCacheAddCapped formula.Length total
            current <- cachedValue
        | CellValue.Empty
        | CellValue.Number _
        | CellValue.Boolean _ ->
            doneCounting <- true
    total

let private criteriaPredicateRetainedTextCharCount criteria =
    let baseTextChars = max 1 (criteriaValueTextCharCount criteria)
    match unwrapValue criteria with
    | CellValue.String raw ->
        let op, restStart = parseCriteriaOperator raw
        let rest = if restStart = 0 then raw else raw.Substring(restStart)
        match Double.TryParse(rest, NumberStyles.Float, CultureInfo.InvariantCulture) with
        | true, _ -> baseTextChars
        | _ ->
            match op with
            | CriteriaEq | CriteriaNe when hasCriteriaWildcard rest ->
                baseTextChars
                |> criteriaPredicateCacheAddCapped (wildcardRegexEntryPatternChars rest (wildcardRegexPatternCharCount rest))
            | _ -> baseTextChars
    | _ -> baseTextChars

let private removeCriteriaPredicateCacheEntry criteria (cache: CriteriaPredicateCache) =
    let mutable entryTextChars = 0
    if cache.TextChars.TryGetValue(criteria, &entryTextChars) then
        cache.TotalTextChars <- max 0 (cache.TotalTextChars - entryTextChars)
        cache.TextChars.Remove criteria |> ignore
    cache.Values.Remove criteria |> ignore

let private clearCriteriaPredicateCache (cache: CriteriaPredicateCache) =
    cache.Values.Clear()
    cache.TextChars.Clear()
    cache.Order.Clear()
    cache.TotalTextChars <- 0

let private addCriteriaPredicateCacheEntry (cache: CriteriaPredicateCache) criteria pred entryTextChars =
    if entryTextChars > criteriaPredicateCacheMaxTotalTextChars then
        removeCriteriaPredicateCacheEntry criteria cache
    else
        let existed = cache.Values.ContainsKey criteria
        if existed then
            removeCriteriaPredicateCacheEntry criteria cache
        else
            cache.Order.Enqueue criteria
        cache.Values.[criteria] <- pred
        cache.TextChars.[criteria] <- entryTextChars
        cache.TotalTextChars <- cache.TotalTextChars + entryTextChars

        let mutable skippedCurrent = false
        while (cache.Values.Count > criteriaPredicateCacheMaxEntries
               || cache.TotalTextChars > criteriaPredicateCacheMaxTotalTextChars)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if oldest = criteria then
                skippedCurrent <- true
            else
                removeCriteriaPredicateCacheEntry oldest cache
        if skippedCurrent && cache.Values.ContainsKey criteria then
            cache.Order.Enqueue criteria

        if cache.Values.Count > criteriaPredicateCacheMaxEntries
           || cache.TotalTextChars > criteriaPredicateCacheMaxTotalTextChars then
            clearCriteriaPredicateCache cache
            cache.Order.Enqueue criteria
            cache.Values.[criteria] <- pred
            cache.TextChars.[criteria] <- entryTextChars
            cache.TotalTextChars <- entryTextChars

let private cachedCriteriaPredicate (cache: CriteriaPredicateCache option) criteria =
    match cache with
    | None -> compileCriteria criteria
    | Some cache ->
        let mutable pred = Unchecked.defaultof<CellValue -> bool>
        if cache.Values.TryGetValue(criteria, &pred) then
            pred
        else
            pred <- compileCriteria criteria
            let entryTextChars = criteriaPredicateRetainedTextCharCount criteria
            addCriteriaPredicateCacheEntry cache criteria pred entryTextChars
            pred

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
        stack.RemoveAt last
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

let private compareValuesNoError (l: CellValue) (r: CellValue) : int voption =
    let left = unwrapValue l
    match left with
    | CellValue.Error _ -> ValueNone
    | CellValue.Number x ->
        let right = unwrapValue r
        match right with
        | CellValue.Error _ -> ValueNone
        | CellValue.Number y -> ValueSome (compareFloats x y)
        | CellValue.Empty -> ValueSome (compareFloats x 0.0)
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
        | _ -> ValueSome (compareInts (cellValueRank left) (cellValueRank right))
    | CellValue.Empty ->
        let right = unwrapValue r
        match right with
        | CellValue.Error _ -> ValueNone
        | CellValue.Empty -> ValueSome 0
        | CellValue.Number y -> ValueSome (compareFloats 0.0 y)
        | CellValue.String y -> ValueSome (String.Compare("", y, StringComparison.OrdinalIgnoreCase))
        | _ -> ValueSome (compareInts (cellValueRank left) (cellValueRank right))
    | _ ->
        let right = unwrapValue r
        match right with
        | CellValue.Error _ -> ValueNone
        | _ -> ValueSome (compareInts (cellValueRank left) (cellValueRank right))

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

/// Compare two scalar values for the comparison operators. Excel orders
/// number < text < FALSE < TRUE; numbers compare numerically, text
/// case-insensitively. Returns -1 / 0 / 1, or an Error value to propagate.
let private compareValues (l: CellValue) (r: CellValue) : Result<int, string> =
    match compareValuesNoError l r with
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

/// Resolve a range-shaped expression (RangeRef or single CellRef) to a
/// `RangeAccess`, honoring a sheet qualifier. None for non-range expressions.
let rangeAccessOf (model: SpreadsheetModel) (workbook: RedBat.Spreadsheet.Workbook.Workbook option) (e: Expr) : RangeAccess option =
    let mk sheet sr sc er ec =
        match resolveModel sheet model workbook with
        | Some m -> Some { M = m; SR = sr; SC = sc; W = ec - sc + 1; Count = (er - sr + 1) * (ec - sc + 1) }
        | None -> None
    match e with
    | RangeRef (a, b) -> let sr, sc, er, ec = normalizeRange a b in mk a.Sheet sr sc er ec
    | CellRef r -> mk r.Sheet r.Row r.Col r.Row r.Col
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

let private sumIfRangeAccess (crit: RangeAccess) (sum: RangeAccess) (pred: CellValue -> bool) =
    let count = min crit.Count sum.Count
    let area = int64 count
    let mutable total = 0.0
    let addNumber value =
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
                | ValueSome _ -> addNumber (unwrapValue cell.Value)
                | ValueNone -> ()
                match rangeOffsetWithinCount crit count critEr critEc addr with
                | ValueSome k when not (pred (unwrapValue cell.Value)) -> excluded.Add k |> ignore
                | _ -> ())
        else
            sum.M.Cells
            |> Map.iter (fun addr cell ->
                match rangeOffsetWithinCount sum count sumEr sumEc addr with
                | ValueSome _ -> addNumber (unwrapValue cell.Value)
                | ValueNone -> ())
            crit.M.Cells
            |> Map.iter (fun addr cell ->
                match rangeOffsetWithinCount crit count critEr critEc addr with
                | ValueSome k when not (pred (unwrapValue cell.Value)) -> excluded.Add k |> ignore
                | _ -> ())
        for k in excluded do
            if k < sum.Count then
                subtractNumber (sumReader.Read k)
    elif area >= 1024L && int64 crit.M.Cells.Count <= area * 4L then
        crit.M.Cells
        |> Map.iter (fun addr cell ->
            match tryRangeOffsetForCount crit count addr with
            | ValueSome k when pred (unwrapValue cell.Value) -> addNumber (sumReader.Read k)
            | _ -> ())
    else
        for k in 0 .. count - 1 do
            if pred (critReader.Read k) then
                addNumber (sumReader.Read k)
    total

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
                        match tryAggregateNumber (sumReader.Read k) with
                        | ValueSome n -> total <- total + n
                        | ValueNone -> ()
                    | _ -> ()
            | _ -> ())
        ValueSome (total, hits)

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
            let sumReader = rangeReadCached s count
            let subtractNumber value =
                match tryAggregateNumber value with
                | ValueSome n -> total <- total - n
                | ValueNone -> ()
            let struct (er, ec) = rangeEndForCount s count
            s.M.Cells
            |> Map.iter (fun addr cell ->
                match rangeOffsetWithinCount s count er ec addr with
                | ValueSome _ ->
                    match tryAggregateNumber (unwrapValue cell.Value) with
                    | ValueSome n -> total <- total + n
                    | ValueNone -> ()
                | _ -> ())
            for k in excluded do
                if k < s.Count then
                    subtractNumber (sumReader.Read k)
            ValueSome (total, hits)
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
                  let found = found || abs (n - target) < 1e-12
                  let rank = if (if descending then n > target else n < target) then rank + 1 else rank
                  struct (ValueNone, found, rank)
              | CellValue.Boolean b ->
                  let n = if b then 1.0 else 0.0
                  let found = found || abs (n - target) < 1e-12
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
       "BYROW"; "BYCOL" |]

#if FABLE_COMPILER
let isArrayFunction (name: string) =
    arrayFunctionNames |> Array.exists (equalsOrdinalIgnoreCase name)
#else
let private arrayFunctionSet = HashSet<string>(arrayFunctionNames, StringComparer.OrdinalIgnoreCase)
let isArrayFunction (name: string) = arrayFunctionSet.Contains name
#endif

/// Does this expression evaluate to an array (a spill)? An array function, or
/// an operator whose operand is itself an array (`SEQUENCE(5)*2`, `A1:A8>2`).
/// Aggregates consume such an argument over its whole array, not its top-left.
let isArrayExpr (e: Expr) : bool =
    let stack = ResizeArray<Expr>()
    stack.Add e
    let mutable found = false
    while not found && stack.Count > 0 do
        let last = stack.Count - 1
        let expr = stack.[last]
        stack.RemoveAt last
        match expr with
        | Function (name, _) when isArrayFunction name -> found <- true
        // A range/table inside an operator broadcasts element-wise (`A1:A8+B1:B8`).
        // A *bare* range arg never reaches this test — aggregates match RangeRef
        // first and stream it — so this only fires inside Unary/BinOp/array funcs.
        | RangeRef _ | TableRef _ | ArrayLit _ | SpillRef _ -> found <- true
        | UnaryOp (_, inner) -> stack.Add inner
        | BinOp (left, _, right) ->
            stack.Add left
            stack.Add right
        | _ -> ()
    found

let arithmeticArrayNonBlankCount model workbook e =
    let rangeShape expr =
        match rangeAccessOf model workbook expr with
        | Some access when access.W > 0 -> ValueSome (access.Count / access.W, access.W)
        | _ -> ValueNone
    let scalarShape expr =
        if isArrayExpr expr then ValueNone else ValueSome (1, 1)
    let broadcastShape left right =
        match left, right with
        | ValueSome (lr, lc), ValueSome (rr, rc)
            when (lr = 1 || rr = 1 || lr = rr) && (lc = 1 || rc = 1 || lc = rc) ->
            ValueSome (max lr rr * max lc rc)
        | ValueSome _, ValueSome _ -> ValueSome 1
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

/// Apply a prefix/postfix unary operator to one value (shared by scalar eval
/// and element-wise array broadcasting).
let applyUnary (op: string) (v: CellValue) : CellValue =
    match op with
    | "-" -> (match coerceNumber v with Ok n -> CellValue.Number (-n) | Result.Error e -> err e)
    | "+" -> (match coerceNumber v with Ok n -> CellValue.Number n | Result.Error e -> err e)
    | "%" -> (match coerceNumber v with Ok n -> CellValue.Number (n / 100.0) | Result.Error e -> err e)
    | _ -> err Err.value

/// Apply a binary operator to two values (shared by scalar eval and array
/// broadcasting). Mirrors the scalar BinOp semantics exactly.
let applyBin (op: string) (lv: CellValue) (rv: CellValue) : CellValue =
    match firstError [ lv; rv ] with
    | Some e -> e
    | None ->
        match op with
        | "&" ->
            (match coerceText lv, coerceText rv with
             | Ok a, Ok b -> CellValue.String (a + b)
             | Result.Error e, _ | _, Result.Error e -> err e)
        | "=" | "<>" | "<" | ">" | "<=" | ">=" ->
            (match compareValues lv rv with
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
                match op with
                | "+" -> CellValue.Number (a + b)
                | "-" -> CellValue.Number (a - b)
                | "*" -> CellValue.Number (a * b)
                | "/" -> if b = 0.0 then err Err.div0 else CellValue.Number (a / b)
                | "^" -> (let p = a ** b in if Double.IsNaN p || Double.IsInfinity p then err Err.num else CellValue.Number p)
                | _ -> err Err.value)

/// Everything an evaluation needs beyond the expression itself: the sheet
/// model being evaluated, the workbook (for cross-sheet refs and defined
/// names), and which sheet is "current" (for scoping sheet-local names).
type EvalContext = {
    Model: SpreadsheetModel
    Workbook: RedBat.Spreadsheet.Workbook.Workbook option
    Sheet: RedBat.Spreadsheet.Workbook.SheetId option
    CriteriaCache: CriteriaPredicateCache option
    /// Address of the cell being evaluated, when known. Needed to resolve
    /// this-row structured refs (`Table[@Col]` / `[@Col]`) and to find the
    /// containing table for unqualified refs. None for ad-hoc evaluation.
    CurrentCell: CellAddress option
}

/// Build a context with no workbook — same-sheet formulas only, no defined
/// names. Convenient for callers that just have a bare model.
let contextOf (model: SpreadsheetModel) : EvalContext =
    { Model = model; Workbook = None; Sheet = None; CriteriaCache = None; CurrentCell = None }

/// Evaluate an expression to a scalar CellValue.
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
            match resolveNameExpr name ctx.Sheet wb with
            | Some e -> eval e ctx
            | None -> err Err.name
    | CellRef r -> readCell r model workbook
    | RangeRef (a, b) ->
        // A bare range in scalar context yields its first cell (Excel's
        // implicit-intersection collapse is close enough for our use).
        firstRangeRefValue a b model workbook
    | TableRef sel -> firstTableSelValue sel workbook ctx.Sheet ctx.CurrentCell model
    | ArrayLit a -> if a.Length > 0 && a.[0].Length > 0 then a.[0].[0] else err Err.value
    | SpillRef anchor -> let a, b = spillRangeOf anchor model in eval (RangeRef (a, b)) ctx
    | UnaryOp (op, operand) -> applyUnary op (eval operand ctx)
    | BinOp _ -> evalBinOpChain expr ctx
    | Function (name, args) ->
        // An array-producing function in scalar context collapses to its
        // top-left element (implicit intersection; spill placement is Step 2).
        if isArrayFunction name then
            match tryEvalArrayTopLeft name args ctx with
            | ValueSome value -> value
            | ValueNone ->
                let arr = evalArray expr ctx
                if arr.Length > 0 && arr.[0].Length > 0 then arr.[0].[0] else err Err.value
        else evalFunction name args ctx

and evalBinOpChain (expr: Expr) (ctx: EvalContext) : CellValue =
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
            let colsN =
                if nargs >= 2 then tryNumber 1 else Ok 1.0
            let startN =
                if nargs >= 3 then tryNumber 2 else Ok 1.0
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
                match compareValuesNoError left right with
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
        let a, b = spillRangeOf anchor model
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
        match workbook |> Option.bind (resolveNameExpr name ctx.Sheet) with
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
    | SpillRef anchor -> let a, b = spillRangeOf anchor model in iterRange a b
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
        match workbook |> Option.bind (resolveNameExpr name ctx.Sheet) with
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
    | Function (name, args) when isArrayFunction name -> evalArrayFunction (name.ToUpperInvariant()) args ctx
    | RangeRef (a, b) ->
        match resolveModel a.Sheet ctx.Model ctx.Workbook with
        | Some m ->
            let sr, sc, er, ec = normalizeRange a b
            [| for r in sr .. er -> [| for c in sc .. ec -> getCellValue { Row = r; Col = c } m |> unwrapValue |] |]
        | None -> scalar (err Err.ref)
    | TableRef sel -> [| for v in expandTableSel sel ctx.Workbook ctx.Sheet ctx.CurrentCell ctx.Model -> [| v |] |]
    | ArrayLit a -> a
    | SpillRef anchor -> let a, b = spillRangeOf anchor ctx.Model in evalArray (RangeRef (a, b)) ctx
    | NameRef name ->
        match ctx.Workbook |> Option.bind (resolveNameExpr name ctx.Sheet) with
        | Some (RangeRef (a, b)) -> evalArray (RangeRef (a, b)) ctx
        | Some (TableRef sel) -> evalArray (TableRef sel) ctx
        | Some e -> scalar (eval e ctx)
        | None -> scalar (err Err.name)
    | UnaryOp (op, inner) -> evalArray inner ctx |> Array.map (Array.map (applyUnary op))
    | BinOp (l, op, r) -> broadcastArr op (evalArray l ctx) (evalArray r ctx)
    | _ -> scalar (eval expr ctx)

/// Element-wise binary op with Excel-style broadcasting: a 1×1 operand is
/// reused for every element; equal shapes pair up; otherwise → #N/A.
and broadcastArr (op: string) (la: ArrVal) (ra: ArrVal) : ArrVal =
    let lr = la.Length
    let lc = if lr > 0 then la.[0].Length else 0
    let rr = ra.Length
    let rc = if rr > 0 then ra.[0].Length else 0
    if (lr <> 1 && rr <> 1 && lr <> rr) || (lc <> 1 && rc <> 1 && lc <> rc) then [| [| err Err.na |] |]
    else
        let nr = max lr rr
        let nc = max lc rc
        let getL r c = la.[(if lr = 1 then 0 else r)].[(if lc = 1 then 0 else c)]
        let getR r c = ra.[(if rr = 1 then 0 else r)].[(if rc = 1 then 0 else c)]
        [| for r in 0 .. nr - 1 -> [| for c in 0 .. nc - 1 -> applyBin op (getL r c) (getR r c) |] |]

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
        let rows = int (num 0)
        let cols = if nargs >= 2 then int (num 1) else 1
        let start = if nargs >= 3 then num 2 else 1.0
        let step = if nargs >= 4 then num 3 else 1.0
        if rows < 1 || cols < 1 then errArr Err.value
        else [| for r in 0 .. rows - 1 -> [| for c in 0 .. cols - 1 -> CellValue.Number (start + float (r * cols + c) * step) |] |]
    | "RANDARRAY" ->
        let rows = if nargs >= 1 then int (num 0) else 1
        let cols = if nargs >= 2 then int (num 1) else 1
        let lo = if nargs >= 3 then num 2 else 0.0
        let hi = if nargs >= 4 then num 3 else 1.0
        let whole = nargs >= 5 && boolAt 4
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
                match compareValuesNoError left right with
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
        let parseAggArg i = match coerceText (eval (argExpr i) ctx) with Ok s -> RedBat.Spreadsheet.Pivot.parseAgg s |> Option.defaultValue RedBat.Spreadsheet.Pivot.Sum | _ -> RedBat.Spreadsheet.Pivot.Sum
        let agg = if nargs >= 3 then parseAggArg 2 else RedBat.Spreadsheet.Pivot.Sum
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
                        | RedBat.Spreadsheet.Pivot.Count -> CellValue.Number(float count)
                        | RedBat.Spreadsheet.Pivot.Sum -> CellValue.Number sum
                        | RedBat.Spreadsheet.Pivot.Average -> if numericCount = 0 then CellValue.Empty else CellValue.Number(sum / float numericCount)
                        | RedBat.Spreadsheet.Pivot.Max -> if numericCount = 0 then CellValue.Empty else CellValue.Number maxValue
                        | RedBat.Spreadsheet.Pivot.Min -> if numericCount = 0 then CellValue.Empty else CellValue.Number minValue
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
        | ValueNone -> RedBat.Spreadsheet.Pivot.groupBy (evalArray (argExpr 0) ctx) (evalArray (argExpr 1) ctx) agg
    | "PIVOTBY" ->
        // PIVOTBY(row_range, col_range, value_range, [agg]) — live spilling
        // cross-tab.
        if nargs < 3 then errArr Err.value else
        let parseAggArg i = match coerceText (eval (argExpr i) ctx) with Ok s -> RedBat.Spreadsheet.Pivot.parseAgg s |> Option.defaultValue RedBat.Spreadsheet.Pivot.Sum | _ -> RedBat.Spreadsheet.Pivot.Sum
        let agg = if nargs >= 4 then parseAggArg 3 else RedBat.Spreadsheet.Pivot.Sum
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
                        | RedBat.Spreadsheet.Pivot.Count -> CellValue.Number(float count)
                        | RedBat.Spreadsheet.Pivot.Sum -> CellValue.Number sum
                        | RedBat.Spreadsheet.Pivot.Average -> if numericCount = 0 then CellValue.Empty else CellValue.Number(sum / float numericCount)
                        | RedBat.Spreadsheet.Pivot.Max -> if numericCount = 0 then CellValue.Empty else CellValue.Number maxValue
                        | RedBat.Spreadsheet.Pivot.Min -> if numericCount = 0 then CellValue.Empty else CellValue.Number minValue
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
        | ValueNone -> RedBat.Spreadsheet.Pivot.pivotBy (evalArray (argExpr 0) ctx) (evalArray (argExpr 1) ctx) (evalArray (argExpr 2) ctx) agg
    | _ -> errArr Err.name

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
        | SpillRef anchor -> let a, b = spillRangeOf anchor model in iterRange a b
        | TableRef sel -> iterTableSel sel workbook ctx.Sheet ctx.CurrentCell model (fun v -> s <- f.Step s v)
        | _ when isArrayExpr arg ->
            for row in evalArray arg ctx do
                for v in row do s <- f.Step s (unwrapValue v)
        | NameRef n ->
            match workbook |> Option.bind (resolveNameExpr n ctx.Sheet) with
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
        | SpillRef anchor -> let a, b = spillRangeOf anchor model in iterRange a b
        | TableRef sel -> iterTableSel sel workbook ctx.Sheet ctx.CurrentCell model (fun v -> s <- f.Step s v)
        | _ when isArrayExpr arg ->
            for row in evalArray arg ctx do
                for v in row do s <- f.Step s (unwrapValue v)
        | NameRef n ->
            match workbook |> Option.bind (resolveNameExpr n ctx.Sheet) with
            | Some (RangeRef (a, b)) -> iterRange a b
            | Some (TableRef sel) -> iterTableSel sel workbook ctx.Sheet ctx.CurrentCell model (fun v -> s <- f.Step s v)
            | Some inner -> s <- f.Step s (unwrapValue (eval inner ctx))
            | None -> s <- f.Step s (err Err.name)
        | _ -> s <- f.Step s (unwrapValue (eval arg ctx))
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
        | SpillRef anchor -> let a, b = spillRangeOf anchor model in iterRange a b
        | TableRef sel -> iterTableSel sel workbook ctx.Sheet ctx.CurrentCell model addValue
        | _ when ValueOption.isSome (arithmeticArrayNonBlankCount model workbook arg) -> ()
        | _ when isArrayExpr arg ->
            for row in evalArray arg ctx do
                for value in row do addValue value
        | NameRef n ->
            match workbook |> Option.bind (resolveNameExpr n ctx.Sheet) with
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
    foldArgsIgnoringMissingCells
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
    foldArgsIgnoringMissingCells
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
    let model = ctx.Model
    let workbook = ctx.Workbook
    let fnName = name.ToUpperInvariant()
    let argsArray =
        match fnName with
        | "SUM" | "PRODUCT" | "AVERAGE" | "MIN" | "MAX" | "MEDIAN"
        | "COUNT" | "COUNTA" | "COUNTBLANK"
        | "STDEV" | "STDEV.S" | "STDEVP" | "STDEV.P"
        | "VAR" | "VAR.S" | "VARP" | "VAR.P" -> Array.empty
        | _ -> List.toArray args
    let nargs = argsArray.Length
    let argExpr i = argsArray.[i]
    let argRaw i = eval (argExpr i) ctx
    let argNum i = coerceNumber (argRaw i)
    let argText i = coerceText (argRaw i)
    let argBool i = coerceBool (argRaw i)
    // Apply a unary numeric function to arg 0 with error propagation.
    let unaryNum f =
        match argNum 0 with Ok x -> f x | Result.Error e -> err e
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
                | CellValue.Boolean b ->
                    values.[count] <- if b then 1.0 else 0.0
                    count <- count + 1
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
        | None -> addValue (eval e ctx)
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
                | "%" -> Ok (n / 100.0)
                | _ -> Result.Error Err.value
        let arithmeticNumber op a b =
            match op with
            | "+" -> Ok (a + b)
            | "-" -> Ok (a - b)
            | "*" -> Ok (a * b)
            | "/" -> if b = 0.0 then Result.Error Err.div0 else Ok (a / b)
            | "^" ->
                let p = a ** b
                if Double.IsNaN p || Double.IsInfinity p then Result.Error Err.num else Ok p
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
            found <- found || abs (value - target) < 1e-12
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
                productValue <- productValue * n
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
            found <- found || abs (value - target) < 1e-12
            if (if descending then value > target else value < target) then
                rank <- rank + repeats
        if found then CellValue.Number (float rank) else err Err.na
    let isErr v = match unwrapValue v with CellValue.Error _ -> true | _ -> false
    let repeatedProduct factor repeats =
        if repeats <= 0L then 1.0
        elif factor = 0.0 then 0.0
        elif factor = 1.0 then 1.0
        elif factor = -1.0 then if repeats % 2L = 0L then 1.0 else -1.0
        else Math.Pow(factor, float repeats)
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
                                    productValue <- productValue * repeatedProduct n repeats
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
                                let visit offset =
                                    if visited.Add offset then
                                        consume offset 1L (readAt offset)
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
                                productValue <- productValue * repeatedProduct n repeats
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
                            let visit offset =
                                if visited.Add offset then
                                    consume offset 1L (readAt offset)
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
                            results.RemoveAt last
                            result

                        while frames.Count > 0 do
                            let last = frames.Count - 1
                            let struct (kind, expr, op) = frames.[last]
                            frames.RemoveAt last

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
                            match Double.TryParse(trimmedNumberText s start length, NumberStyles.Float, CultureInfo.InvariantCulture) with
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
                            results.RemoveAt last
                            result

                        while frames.Count > 0 do
                            let last = frames.Count - 1
                            let struct (kind, expr, op) = frames.[last]
                            frames.RemoveAt last

                            match kind with
                            | 1 ->
                                let struct (innerKind, n, e) = popResult ()
                                if innerKind = 0 then
                                    match op with
                                    | "+" -> results.Add(numberOk n)
                                    | "-" -> results.Add(numberOk -n)
                                    | "%" -> results.Add(numberOk (n / 100.0))
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
                                    | "*" -> results.Add(numberOk (leftNumber * rightNumber))
                                    | "/" when rightNumber = 0.0 -> results.Add(numberError Err.div0)
                                    | "/" -> results.Add(numberOk (leftNumber / rightNumber))
                                    | "^" -> results.Add(numberOk (Math.Pow(leftNumber, rightNumber)))
                                    | _ -> results.Add numberUnsupported
                            | _ ->
                                match expr with
                                | Number n -> results.Add(numberOk n)
                                | Bool b -> results.Add(numberOk (if b then 1.0 else 0.0))
                                | ErrorLit e -> results.Add(numberError e)
                                | Str s ->
                                    let struct (start, length) = trimBounds s
                                    match Double.TryParse(trimmedNumberText s start length, NumberStyles.Float, CultureInfo.InvariantCulture) with
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
                                let c = compareFloats leftNumber rightNumber
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
                results.RemoveAt last
                result

            while frames.Count > 0 do
                let last = frames.Count - 1
                let struct (kind, expr, op) = frames.[last]
                frames.RemoveAt last

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
                results.RemoveAt last
                result

            while frames.Count > 0 do
                let last = frames.Count - 1
                let struct (kind, expr, op) = frames.[last]
                frames.RemoveAt last

                match kind with
                | 1 ->
                    let struct (innerKind, n, e) = popResult ()
                    if innerKind = 0 then
                        match op with
                        | "+" -> results.Add(numericOk n)
                        | "-" -> results.Add(numericOk -n)
                        | "%" -> results.Add(numericOk (n / 100.0))
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
                        | "*" -> results.Add(numericOk (leftValue * rightValue))
                        | "/" when rightValue = 0.0 -> results.Add(numericError Err.div0)
                        | "/" -> results.Add(numericOk (leftValue / rightValue))
                        | "^" -> results.Add(numericOk (Math.Pow(leftValue, rightValue)))
                        | "=" | "<>" | "<" | ">" | "<=" | ">=" ->
                            let c = compareFloats leftValue rightValue
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
                    | ValueSome (Ok (struct (lr, lc, lk))), ValueSome (Ok (struct (rr, rc, rk))) when lr = rr && lc = rc ->
                        ValueSome (Ok (comparisonValue op (compareFloats lk rk)))
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
                        stack.RemoveAt last

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
                        when lrc = rrc && lr = rr && lc = rc ->
                        ValueSome (Ok (comparisonValue op (compareFloats lk rk)))
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
                        stack.RemoveAt last

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
                            | Ok n -> Ok (factor * n)
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
                        stack.RemoveAt last

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

                if not (isMulDiv root) then
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
                        stack.RemoveAt last

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
                                        elif value = 0.0 then
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
                                    | Ok n -> Ok (scalar * n)
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
                                        | "%" -> Ok (n / 100.0)
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
                                            | "*" -> Ok (n * scalar)
                                            | "/" when scalar <> 0.0 -> Ok (n / scalar)
                                            | "/" -> Result.Error Err.div0
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
                                                | "*" -> Ok (scalar * n)
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
                                        | Ok n -> Ok (comparisonValue (compare n scalar))
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
                                            | Ok n -> Ok (comparisonValue (compare scalar n))
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
                            productValue <- productValue * n
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
                        productValue <- productValue * n
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
                    | "PRODUCT" -> CellValue.Number (Math.Pow(constant, total))
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
                        struct (true, ValueSome (Ok (rows, cols, scaleBilinear factor value)))
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
                        stack.RemoveAt last

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

                if not (isMulDiv root) then
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
                        stack.RemoveAt last

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
                                        elif value = 0.0 then
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
                                struct (true, ValueSome (Ok (rows, cols, scaleBilinear scalar arrayValue)))
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
                                        match tryBilinear rowName colName body with
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
                                    | "*" -> ValueSome (Ok (rows, cols, scaleBilinear scalar value))
                                    | "/" when scalar <> 0.0 -> ValueSome (Ok (rows, cols, scaleBilinear (1.0 / scalar) value))
                                    | "/" -> ValueSome (Result.Error Err.div0)
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
            let allTrue, allFalse =
                match op with
                | "=" -> minValue = maxValue && minValue = scalar, scalar < minValue || scalar > maxValue
                | "<>" -> scalar < minValue || scalar > maxValue, minValue = maxValue && minValue = scalar
                | "<" -> maxValue < scalar, minValue >= scalar
                | ">" -> minValue > scalar, maxValue <= scalar
                | "<=" -> maxValue <= scalar, minValue > scalar
                | _ -> minValue >= scalar, maxValue < scalar
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
                    | Ok scalar -> ValueSome (makeStream rows cols value scalar compare)
                | ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                | _ ->
                    match tryBilinearMakeArrayExpr right with
                    | ValueSome (Ok (rows, cols, value)) when not (isArrayExpr left) ->
                        match coerceNumber (eval left ctx) with
                        | Result.Error e -> ValueSome (Result.Error e)
                        | Ok scalar -> ValueSome (makeStream rows cols value scalar (fun value scalar -> compare scalar value))
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
                            match tryBilinear rowName colName body with
                            | ValueSome (Result.Error e) -> ValueSome (err e)
                            | ValueSome (Ok (struct (rowColCoeff, rowCoeff, colCoeff, constant))) ->
                                if fnName = "PRODUCT" then
                                    if rowColCoeff = 0.0 && rowCoeff = 0.0 && colCoeff = 0.0 then
                                        ValueSome (CellValue.Number (Math.Pow(constant, float totalCells)))
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
                                    productValue <- productValue * n
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
            | Some arg -> eval arg ctx |> coerceNumber
            | None -> Ok defaultValue
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
            | Some arg -> eval arg ctx |> coerceNumber
            | None -> Ok defaultValue
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
                        ValueSome (Ok (count, (fun () -> factor * nextNumber ()), countAsNumber))
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
                        stack.RemoveAt last

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

                if not (isMulDiv root) then
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
                        stack.RemoveAt last

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
                                        elif value = 0.0 then
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
                                ValueSome (Ok (streamCount, (fun () -> scalar * streamNext ()), streamCountAsNumber))
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
                                    | "%" -> fun () -> nextNumber () / 100.0
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
                                    | "*" -> ValueSome (Ok (count, (fun () -> nextNumber () * scalar), countAsNumber))
                                    | "/" when scalar <> 0.0 -> ValueSome (Ok (count, (fun () -> nextNumber () / scalar), countAsNumber))
                                    | "/" -> ValueSome (Result.Error Err.div0)
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
                                        | "*" -> ValueSome (Ok (count, (fun () -> scalar * nextNumber ()), countAsNumber))
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
                                    ValueSome (Ok (count, (fun () -> comparisonValue (compare (nextNumber ()) scalar)), false))
                            | ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                            | _ ->
                                match sequenceStreamOfExpr right with
                                | ValueSome (Ok (count, nextNumber, _)) when not (isArrayExpr left) ->
                                    match coerceNumber (eval left ctx) with
                                    | Result.Error e -> ValueSome (Result.Error e)
                                    | Ok scalar ->
                                        ValueSome (Ok (count, (fun () -> comparisonValue (compare scalar (nextNumber ()))), false))
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
                            productValue <- productValue * n
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
            | Some arg -> eval arg ctx |> coerceNumber
            | None -> Ok defaultValue
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
                        ValueSome (Ok (count, (fun () -> factor * nextNumber ()), countAsNumber))
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
                        stack.RemoveAt last

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

                if not (isMulDiv root) then
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
                        stack.RemoveAt last

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
                                        elif value = 0.0 then
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
                                ValueSome (Ok (streamCount, (fun () -> scalar * streamNext ()), streamCountAsNumber))
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
                                    | "%" -> fun () -> nextNumber () / 100.0
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
                                    | "*" -> ValueSome (Ok (count, (fun () -> nextNumber () * scalar), countAsNumber))
                                    | "/" when scalar <> 0.0 -> ValueSome (Ok (count, (fun () -> nextNumber () / scalar), countAsNumber))
                                    | "/" -> ValueSome (Result.Error Err.div0)
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
                                        | "*" -> ValueSome (Ok (count, (fun () -> scalar * nextNumber ()), countAsNumber))
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
                                    ValueSome (Ok (count, (fun () -> comparisonValue (compare (nextNumber ()) scalar)), false))
                            | ValueSome (Result.Error e) -> ValueSome (Result.Error e)
                            | _ ->
                                match randStreamOfExpr right with
                                | ValueSome (Ok (count, nextNumber, _)) when not (isArrayExpr left) ->
                                    match coerceNumber (eval left ctx) with
                                    | Result.Error e -> ValueSome (Result.Error e)
                                    | Ok scalar ->
                                        ValueSome (Ok (count, (fun () -> comparisonValue (compare scalar (nextNumber ()))), false))
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
                            productValue <- productValue * n
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
                match compareValuesNoError left right with
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
                                productValue <- productValue * n
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
                | ValueSome e -> ValueSome (err e)
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
                        if abs f < 1e-8 then result <- Some r
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
                            for k in la.Count - 1 .. -1 .. 0 do visit k
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
                                                    | ValueNone -> foldArgsIgnoringMissingCells (numAggFold 0.0 (+) (fun _ a -> CellValue.Number a)) args ctx
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
                                    | ValueNone -> foldArgsIgnoringMissingCells (numAggFold 1.0 ( * ) (fun n a -> CellValue.Number (if n = 0 then 0.0 else a))) args ctx
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
                                                | ValueNone -> foldArgsIgnoringMissingCells (numAggFold 0.0 (+) (fun n a -> if n = 0 then err Err.div0 else CellValue.Number (a / float n))) args ctx
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
                                                | ValueNone -> foldArgsIgnoringMissingCells (numAggFold infinity min (fun n a -> CellValue.Number (if n = 0 then 0.0 else a))) args ctx
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
                                                | ValueNone -> foldArgsIgnoringMissingCells (numAggFold (-infinity) max (fun n a -> CellValue.Number (if n = 0 then 0.0 else a))) args ctx
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
                                                    | ValueNone -> foldArgsIgnoringMissingCells (countFold true (fun v -> match v with CellValue.Number _ -> true | _ -> false)) args ctx
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
            | ValueNone -> foldArgsIgnoringMissingCells (varianceFold sample true) args ctx
    | "VAR" | "VAR.S" | "VARP" | "VAR.P" ->
        let sample = fnName = "VAR" || fnName = "VAR.S"
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone -> foldArgsIgnoringMissingCells (varianceFold sample false) args ctx
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
            let hasFilterProjection = hasFilters && model.FilterHiddenRows.Length > 0
            let filterCriteria () =
                match filterColsCache, filterValuesCache with
                | Some filterCols, Some filterValues -> filterCols, filterValues
                | _ ->
                    let filterEntries = model.Filters |> Map.toArray
                    let filterCols = filterEntries |> Array.map fst
                    let filterValues = filterEntries |> Array.map (fun (_, filter) -> filter.VisibleValues)
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
            let pred = cachedCriteriaPredicate ctx.CriteriaCache (argRaw 1)
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
            let pred = cachedCriteriaPredicate ctx.CriteriaCache (argRaw 1)
            CellValue.Number (sumIfRangeAccess crit sum pred)
    | "AVERAGEIF" ->
        if nargs < 2 then err Err.value else
        match rangeAccessOf model workbook (argExpr 0) with
        | None -> err Err.value
        | Some crit ->
            let avg =
                if nargs >= 3 then (match rangeAccessOf model workbook (argExpr 2) with Some s -> s | None -> crit)
                else crit
            let pred = cachedCriteriaPredicate ctx.CriteriaCache (argRaw 1)
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
                crits.Add(access, cachedCriteriaPredicate ctx.CriteriaCache (eval (argExpr (argIndex + 1)) ctx))
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
                    | Some s when k < s.Count -> (match tryAggregateNumber (sumReader.Read k) with ValueSome n -> total <- total + n | ValueNone -> ())
                    | _ -> ()
            if isSum then CellValue.Number total else CellValue.Number (float hits)
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
                    crits.Add(access, cachedCriteriaPredicate ctx.CriteriaCache (eval (argExpr (argIndex + 1)) ctx))
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
        // Element-wise product of parallel ranges, then sum.
        let buffers = ResizeArray<struct (float[] * int)>()
        let mutable firstErr = ValueNone
        for arg in args do
            match firstErr with
            | ValueSome _ -> ()
            | ValueNone ->
                match numericBufferOfRangeValue arg with
                | Result.Error e -> firstErr <- ValueSome e
                | Ok buffer -> buffers.Add buffer
        match firstErr with
        | ValueSome e -> err e
        | ValueNone ->
            if buffers.Count = 0 then CellValue.Number 0.0
            else
                let mutable len =
                    let struct (_, count) = buffers.[0]
                    count
                for struct (_, count) in buffers do
                    if count < len then len <- count
                let mutable total = 0.0
                for i in 0 .. len - 1 do
                    let mutable product = 1.0
                    for struct (array, _) in buffers do
                        product <- product * array.[i]
                    total <- total + product
                CellValue.Number total
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
                | _ -> Math.Round(scaled, MidpointRounding.AwayFromZero)
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
        | Ok b, Ok p ->
            let r = b ** p
            if Double.IsNaN r || Double.IsInfinity r then err Err.num else CellValue.Number r
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
        | Ok x, Ok m -> if m = 0.0 then CellValue.Number 0.0 else CellValue.Number (Math.Round(x / m, MidpointRounding.AwayFromZero) * m)
    | "EVEN" -> unaryNum (fun x -> let n = ceil (abs x / 2.0) * 2.0 in CellValue.Number (if x < 0.0 then -n else n))
    | "ODD" -> unaryNum (fun x -> let n = 2.0 * ceil ((abs x - 1.0) / 2.0) + 1.0 in CellValue.Number (if x < 0.0 then -n else n))
    | "SUMSQ" ->
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone -> foldArgsIgnoringMissingCells (numAggFold 0.0 (fun acc x -> acc + x * x) (fun _ sum -> CellValue.Number sum)) args ctx
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
                | ValueNone -> foldArgsIgnoringMissingCells (modeFold ()) args ctx
        | _ -> foldArgsIgnoringMissingCells (modeFold ()) args ctx
    | "DEVSQ" ->
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone -> foldArgsIgnoringMissingCells devSqFold args ctx
    | "AVEDEV" ->
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone -> foldArgsIgnoringMissingCells (aveDevFold ()) args ctx
    | "GEOMEAN" ->
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone -> foldArgsIgnoringMissingCells geoMeanFold args ctx
    | "HARMEAN" ->
        match tryRangeUnaryAggregate () with
        | ValueSome result -> result
        | ValueNone ->
            match tryRangeBinOpAggregate () with
            | ValueSome result -> result
            | ValueNone -> foldArgsIgnoringMissingCells harmonicMeanFold args ctx
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
            match Double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture) with
            | true, n -> CellValue.Number n
            | _ ->
                if t.EndsWith "%" then
                    match Double.TryParse(t.TrimEnd('%'), NumberStyles.Any, CultureInfo.InvariantCulture) with
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
            match Double.TryParse(cleaned.TrimEnd('%'), NumberStyles.Any, CultureInfo.InvariantCulture) with
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
            let rounded = Math.Round(x, d, MidpointRounding.AwayFromZero)
            CellValue.String (rounded.ToString((if noCommas then "F" else "N") + string d, CultureInfo.InvariantCulture))
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
        | Ok s, Ok months -> CellValue.Number (dateToSerial ((serialToDate s).AddMonths(int months)))
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
            let jan1 = DateTime(d.Year, 1, 1)
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
            let fv = if nargs >= 4 then (match argNum 3 with Ok v -> v | _ -> 0.0) else 0.0
            let ty = if nargs >= 5 then (match argNum 4 with Ok v -> v | _ -> 0.0) else 0.0
            if np = 0.0 then err Err.num
            elif r = 0.0 then CellValue.Number (-(pv + fv) / np)
            else (let p = (1.0 + r) ** np in CellValue.Number (-(pv * p + fv) * r / ((1.0 + r * ty) * (p - 1.0))))
    | "FV" ->
        if nargs < 3 then err Err.value else
        match argNum 0, argNum 1, argNum 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok r, Ok np, Ok pmt ->
            let pv = if nargs >= 4 then (match argNum 3 with Ok v -> v | _ -> 0.0) else 0.0
            let ty = if nargs >= 5 then (match argNum 4 with Ok v -> v | _ -> 0.0) else 0.0
            if r = 0.0 then CellValue.Number (-(pv + pmt * np))
            else (let p = (1.0 + r) ** np in CellValue.Number (-(pv * p + pmt * (1.0 + r * ty) * (p - 1.0) / r)))
    | "PV" ->
        if nargs < 3 then err Err.value else
        match argNum 0, argNum 1, argNum 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok r, Ok np, Ok pmt ->
            let fv = if nargs >= 4 then (match argNum 3 with Ok v -> v | _ -> 0.0) else 0.0
            let ty = if nargs >= 5 then (match argNum 4 with Ok v -> v | _ -> 0.0) else 0.0
            if r = 0.0 then CellValue.Number (-(fv + pmt * np))
            else (let p = (1.0 + r) ** np in CellValue.Number (-(fv + pmt * (1.0 + r * ty) * (p - 1.0) / r) / p))
    | "NPER" ->
        if nargs < 3 then err Err.value else
        match argNum 0, argNum 1, argNum 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok r, Ok pmt, Ok pv ->
            let fv = if nargs >= 4 then (match argNum 3 with Ok v -> v | _ -> 0.0) else 0.0
            let ty = if nargs >= 5 then (match argNum 4 with Ok v -> v | _ -> 0.0) else 0.0
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
            let fv = if nargs >= 4 then (match argNum 3 with Ok v -> v | _ -> 0.0) else 0.0
            let ty = if nargs >= 5 then (match argNum 4 with Ok v -> v | _ -> 0.0) else 0.0
            let guess = if nargs >= 6 then (match argNum 5 with Ok v -> v | _ -> 0.1) else 0.1
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
        let guess = if nargs >= 2 then (match argNum 1 with Ok v -> v | _ -> 0.1) else 0.1
        match trySparseIrrArithmetic guess (argExpr 0) with
        | ValueSome result -> result
        | ValueNone ->
            match numbersArrayOfArgs [ argExpr 0 ] ctx with
            | Result.Error e -> err e
            | Ok cashFlows when cashFlows.Length >= 2 ->
                let npv r =
                    let mutable total = 0.0
                    for i = 0 to cashFlows.Length - 1 do
                        total <- total + cashFlows.[i] / ((1.0 + r) ** float i)
                    total
                let mutable r = guess
                let mutable result = None
                let mutable i = 0
                while result.IsNone && i < 80 do
                    let f = npv r
                    if abs f < 1e-8 then result <- Some r
                    else
                        let d = (npv (r + 1e-6) - f) / 1e-6
                        if abs d < 1e-14 then i <- 80 else r <- r - f / d
                    i <- i + 1
                match result with Some r when not (Double.IsNaN r || Double.IsInfinity r) -> CellValue.Number r | _ -> err Err.num
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
                match compareValues target (argRaw i) with
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
            CellValue.String (s.Substring(0, max 0 (min s.Length cnt)))
    | "RIGHT" ->
        if nargs < 1 then err Err.value else
        match argText 0 with
        | Result.Error e -> err e
        | Ok s ->
            let cnt = if nargs >= 2 then (match argNum 1 with Ok n -> int n | _ -> 1) else 1
            let c = max 0 (min s.Length cnt)
            CellValue.String (s.Substring(s.Length - c, c))
    | "MID" ->
        if nargs <> 3 then err Err.value else
        match argText 0, argNum 1, argNum 2 with
        | Result.Error e, _, _ | _, Result.Error e, _ | _, _, Result.Error e -> err e
        | Ok s, Ok startN, Ok lenN ->
            let start0 = int startN - 1   // Excel MID is 1-indexed
            if start0 < 0 then err Err.value
            elif start0 >= s.Length then CellValue.String ""
            else
                let take = max 0 (min (s.Length - start0) (int lenN))
                CellValue.String (s.Substring(start0, take))
    | "REPT" ->
        if nargs <> 2 then err Err.value else
        match argText 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok s, Ok cnt -> CellValue.String (String.replicate (max 0 (int cnt)) s)
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
                        hay.IndexOf(needle, startAt)
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
        // Minimal TEXT — passes the format through .NET. Common Excel codes
        // (0, 0.00, #,##0, 0%) are .NET-compatible enough for the usual case.
        if nargs <> 2 then err Err.value else
        match argNum 0, argText 1 with
        | Result.Error e, _ -> err e
        | _, Result.Error e -> err e
        | Ok n, Ok fmt ->
            try CellValue.String (n.ToString(fmt, CultureInfo.InvariantCulture))
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
        // INDEX(range, row_num, [col_num]) — 1-indexed; 0 = whole row/col,
        // collapsed here to the first cell of that line.
        if nargs < 2 then err Err.value else
        match rangeBounds (argExpr 0) with
        | None -> err Err.ref
        | Some (sr, sc, er, ec) ->
            match argNum 1 with
            | Result.Error e -> err e
            | Ok rowN ->
                let colN = if nargs >= 3 then (match argNum 2 with Ok c -> int c | _ -> 1) else 1
                let rows = er - sr + 1
                let csz = ec - sc + 1
                // Single-row or single-col range: the lone index walks it.
                let r, c =
                    if rows = 1 then sr, sc + (max 1 (int rowN)) - 1
                    elif csz = 1 then sr + (max 1 (int rowN)) - 1, sc
                    else sr + (max 1 (int rowN)) - 1, sc + (max 1 colN) - 1
                if r < sr || r > er || c < sc || c > ec then err Err.ref
                else gridCell model r c
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
            | None -> if nargs >= 4 then argRaw 3 else err Err.na
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
                let baseDate = DateTime(int y, 1, 1)
                let dt = baseDate.AddMonths(int m - 1).AddDays(d - 1.0)
                CellValue.Number (dateToSerial dt)
            with _ -> err Err.num
    | "YEAR" -> unaryNum (fun s -> CellValue.Number (float (serialToDate s).Year))
    | "MONTH" -> unaryNum (fun s -> CellValue.Number (float (serialToDate s).Month))
    | "DAY" -> unaryNum (fun s -> CellValue.Number (float (serialToDate s).Day))
    | "WEEKDAY" ->
        match argNum 0 with
        | Result.Error e -> err e
        | Ok s ->
            // Default type 1: Sunday = 1 … Saturday = 7.
            let dow = int (serialToDate s).DayOfWeek  // Sunday = 0
            CellValue.Number (float (dow + 1))
    | "EOMONTH" ->
        if nargs <> 2 then err Err.value else
        match argNum 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok s, Ok months ->
            let d = (serialToDate s).AddMonths(int months)
            let eom = DateTime(d.Year, d.Month, DateTime.DaysInMonth(d.Year, d.Month))
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
    | "OFFSET" ->
        // OFFSET(reference, rows, cols, [height], [width]) — scalar context
        // collapses to the top-left cell of the resulting block.
        if nargs < 3 then err Err.value else
        let anchor =
            match argExpr 0 with
            | CellRef r -> Some (r.Sheet, r.Row, r.Col)
            | RangeRef (a, _) -> Some (a.Sheet, a.Row, a.Col)
            | _ -> None
        match anchor with
        | None -> err Err.ref
        | Some (sheet, br, bc) ->
            match argNum 1, argNum 2 with
            | Result.Error e, _ | _, Result.Error e -> err e
            | Ok dr, Ok dc ->
                let r = br + int dr
                let c = bc + int dc
                if r < 0 || c < 0 then err Err.ref
                else
                    match resolveModel sheet model workbook with
                    | Some m -> getCellValue { Row = r; Col = c } m |> unwrapValue
                    | None -> err Err.ref
    | "INDIRECT" ->
        // INDIRECT(ref_text) — parse the text as an A1 reference and read it.
        if nargs < 1 then err Err.value else
        match argText 0 with
        | Result.Error e -> err e
        | Ok refText ->
            match parseFormulaShared ("=" + refText) with
            | Ok (CellRef _ as e) | Ok (RangeRef _ as e) -> eval e ctx
            | _ -> err Err.ref
    | _ ->
        // Unknown / unsupported function.
        err Err.name

/// Parse + evaluate a formula string in one shot. Defined names resolve at
/// the workbook's active sheet — the right scope for an interactive
/// edit-buffer preview, which is what this entry point serves.
let evaluateFormula (formula: string) (model: SpreadsheetModel)
                    (workbook: RedBat.Spreadsheet.Workbook.Workbook option) : CellValue =
    match parseFormulaShared formula with
    | Ok expr ->
        let ctx =
            { Model = model
              Workbook = workbook
              Sheet = workbook |> Option.map (fun w -> w.ActiveSheet)
              CriteriaCache = Some (CriteriaPredicateCache())
              CurrentCell = None }
        eval expr ctx
    | Result.Error msg -> CellValue.Error msg

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

let private recalcLocalParseCacheMaxEntries = 8192
let private recalcLocalParseCacheMaxTotalChars = 5000000

let private recalcReferenceCacheMaxEntries = 8192
let private recalcReferenceCacheMaxTotalFormulaChars = 5000000
let private recalcReferenceCacheMaxTotalRefs = 200000

let private recalcArrayFormulaCacheMaxEntries = 8192
let private recalcArrayFormulaCacheMaxTotalChars = 5000000

let private recalcCacheAddCapped (maxTotal: int) (amount: int) (total: int) =
    if amount <= 0 then total
    elif total > maxTotal || amount > maxTotal - total then maxTotal + 1
    else total + amount

let private recalcLocalParseCacheEntryChars (formula: string) (parsed: Result<Expr, string>) =
    0
    |> recalcCacheAddCapped recalcLocalParseCacheMaxTotalChars formula.Length
    |> recalcCacheAddCapped
        recalcLocalParseCacheMaxTotalChars
        (parsedFormulaResultRetainedTextCharCount recalcLocalParseCacheMaxTotalChars parsed)

let private rememberRecalcLocalParseCacheEntry
    (cache: Dictionary<string, Result<Expr, string>>)
    (order: Queue<string>)
    (totalChars: byref<int>)
    (formula: string)
    (parsed: Result<Expr, string>) =
    let entryChars = recalcLocalParseCacheEntryChars formula parsed
    if entryChars <= recalcLocalParseCacheMaxTotalChars then
        let mutable previous = Unchecked.defaultof<Result<Expr, string>>
        if cache.TryGetValue(formula, &previous) then
            totalChars <- max 0 (totalChars - recalcLocalParseCacheEntryChars formula previous)
        else
            order.Enqueue formula
        cache.[formula] <- parsed
        totalChars <- totalChars + entryChars
        while (cache.Count > recalcLocalParseCacheMaxEntries
               || totalChars > recalcLocalParseCacheMaxTotalChars)
              && order.Count > 0 do
            let oldest = order.Dequeue()
            if oldest <> formula then
                let mutable oldestParsed = Unchecked.defaultof<Result<Expr, string>>
                if cache.TryGetValue(oldest, &oldestParsed) then
                    cache.Remove oldest |> ignore
                    totalChars <- max 0 (totalChars - recalcLocalParseCacheEntryChars oldest oldestParsed)

let private recalcReferenceCacheFormulaChars (formula: string) =
    max 1 formula.Length

let private recalcReferenceCacheRefsCount (refs: CellAddress array) =
    max 1 refs.Length

type private RecalcReferenceFastCacheEntry = {
    Formula: string
    Refs: CellAddress array
}

let private clearRecalcReferenceFastSlots (slots: RecalcReferenceFastCacheEntry array) =
    for index = 0 to slots.Length - 1 do
        slots.[index] <- Unchecked.defaultof<RecalcReferenceFastCacheEntry>

let private recalcReferenceFastSlot (slots: RecalcReferenceFastCacheEntry array) (formula: string) =
    let hashCode =
        if isNull formula then 0
        else System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(formula)
    hashCode &&& (slots.Length - 1)

let private setRecalcReferenceFastSlot (slots: RecalcReferenceFastCacheEntry array) formula refs =
    slots.[recalcReferenceFastSlot slots formula] <- {
        Formula = formula
        Refs = refs
    }

let private tryGetRecalcReferenceFastSlot (slots: RecalcReferenceFastCacheEntry array) formula =
    let fast = slots.[recalcReferenceFastSlot slots formula]
    if not (obj.ReferenceEquals(fast, null))
       && obj.ReferenceEquals(fast.Formula, formula) then
        ValueSome fast.Refs
    else
        ValueNone

let private rememberRecalcReferenceCacheEntry
    (cache: Dictionary<string, CellAddress array>)
    (order: Queue<string>)
    (totalFormulaChars: byref<int>)
    (totalRefs: byref<int>)
    formula
    refs =
    let formulaChars = recalcReferenceCacheFormulaChars formula
    let refCount = recalcReferenceCacheRefsCount refs
    if formulaChars <= recalcReferenceCacheMaxTotalFormulaChars
       && refCount <= recalcReferenceCacheMaxTotalRefs then
        let mutable previous = Unchecked.defaultof<CellAddress array>
        if cache.TryGetValue(formula, &previous) then
            totalFormulaChars <- max 0 (totalFormulaChars - recalcReferenceCacheFormulaChars formula)
            totalRefs <- max 0 (totalRefs - recalcReferenceCacheRefsCount previous)
        else
            order.Enqueue formula
        cache.[formula] <- refs
        totalFormulaChars <- totalFormulaChars + formulaChars
        totalRefs <- totalRefs + refCount
        while (cache.Count > recalcReferenceCacheMaxEntries
               || totalFormulaChars > recalcReferenceCacheMaxTotalFormulaChars
               || totalRefs > recalcReferenceCacheMaxTotalRefs)
              && order.Count > 0 do
            let oldest = order.Dequeue()
            if oldest <> formula then
                let mutable oldestRefs = Unchecked.defaultof<CellAddress array>
                if cache.TryGetValue(oldest, &oldestRefs) then
                    cache.Remove oldest |> ignore
                    totalFormulaChars <- max 0 (totalFormulaChars - recalcReferenceCacheFormulaChars oldest)
                    totalRefs <- max 0 (totalRefs - recalcReferenceCacheRefsCount oldestRefs)

let private recalcArrayFormulaCacheEntryChars (formula: string) (expr: Expr option) =
    let exprTextChars =
        match expr with
        | Some expr -> parsedFormulaResultRetainedTextCharCount recalcArrayFormulaCacheMaxTotalChars (Ok expr)
        | None -> 0
    0
    |> recalcCacheAddCapped recalcArrayFormulaCacheMaxTotalChars formula.Length
    |> recalcCacheAddCapped recalcArrayFormulaCacheMaxTotalChars exprTextChars

let private rememberRecalcArrayFormulaCacheEntry
    (cache: Dictionary<string, Expr option>)
    (order: Queue<string>)
    (totalChars: byref<int>)
    (formula: string)
    (expr: Expr option) =
    let entryChars = recalcArrayFormulaCacheEntryChars formula expr
    if entryChars <= recalcArrayFormulaCacheMaxTotalChars then
        let mutable previous = Unchecked.defaultof<Expr option>
        if cache.TryGetValue(formula, &previous) then
            totalChars <- max 0 (totalChars - recalcArrayFormulaCacheEntryChars formula previous)
        else
            order.Enqueue formula
        cache.[formula] <- expr
        totalChars <- totalChars + entryChars
        while (cache.Count > recalcArrayFormulaCacheMaxEntries
               || totalChars > recalcArrayFormulaCacheMaxTotalChars)
              && order.Count > 0 do
            let oldest = order.Dequeue()
            if oldest <> formula then
                let mutable oldestExpr = Unchecked.defaultof<Expr option>
                if cache.TryGetValue(oldest, &oldestExpr) then
                    cache.Remove oldest |> ignore
                    totalChars <- max 0 (totalChars - recalcArrayFormulaCacheEntryChars oldest oldestExpr)

type private RecalcArrayFormulaFastCacheEntry = {
    Formula: string
    Expr: Expr option
}

let private clearRecalcArrayFormulaFastSlots (slots: RecalcArrayFormulaFastCacheEntry array) =
    for index = 0 to slots.Length - 1 do
        slots.[index] <- Unchecked.defaultof<RecalcArrayFormulaFastCacheEntry>

let private recalcArrayFormulaFastSlot (slots: RecalcArrayFormulaFastCacheEntry array) (formula: string) =
    let hashCode =
        if isNull formula then 0
        else System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(formula)
    hashCode &&& (slots.Length - 1)

let private setRecalcArrayFormulaFastSlot (slots: RecalcArrayFormulaFastCacheEntry array) formula expr =
    slots.[recalcArrayFormulaFastSlot slots formula] <- {
        Formula = formula
        Expr = expr
    }

let private tryGetRecalcArrayFormulaFastSlot (slots: RecalcArrayFormulaFastCacheEntry array) formula =
    let fast = slots.[recalcArrayFormulaFastSlot slots formula]
    if not (obj.ReferenceEquals(fast, null))
       && obj.ReferenceEquals(fast.Formula, formula) then
        ValueSome fast.Expr
    else
        ValueNone

let private collectFormulaCells (model: SpreadsheetModel) =
    let formulaCells = ResizeArray<FormulaCell>()
    let parseCache = Dictionary<string, Result<Expr, string>>()
    let parseCacheOrder = Queue<string>()
    let parseCacheFastSlots = Array.zeroCreate<ParsedFormulaFastCacheEntry> 256
    let mutable parseCacheTotalChars = 0
    let parseFormulaCached formula =
        match tryGetParsedFormulaFastSlotIn parseCacheFastSlots formula with
        | ValueSome parsed -> parsed
        | ValueNone ->
            let mutable parsed = Unchecked.defaultof<Result<Expr, string>>
            if parseCache.TryGetValue(formula, &parsed) then
                setParsedFormulaFastSlotIn parseCacheFastSlots formula parsed
                parsed
            else
                parsed <- parseFormulaShared formula
                rememberRecalcLocalParseCacheEntry parseCache parseCacheOrder &parseCacheTotalChars formula parsed
                clearParsedFormulaFastSlotsIn parseCacheFastSlots
                if parseCache.ContainsKey formula then
                    setParsedFormulaFastSlotIn parseCacheFastSlots formula parsed
                parsed
    model.Cells
    |> Map.iter (fun addr cell ->
        match cell.Value with
        | CellValue.Formula (formula, _) ->
            match parseFormulaCached formula with
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
let recalculateAllFormulas (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                           (currentSheet: RedBat.Spreadsheet.Workbook.SheetId option)
                           (model: SpreadsheetModel) : SpreadsheetModel =
    match cellsFormulaCountIfKnown model.Cells with
    | Some 0 -> model
    | _ ->
        let formulaCells = collectFormulaCells model
        rememberCellsFormulaCount model.Cells formulaCells.Count
        if formulaCells.Count = 0 then model else
        let criteriaCache = CriteriaPredicateCache()

        let formulaByAddress = Dictionary<CellAddress, FormulaCell>(formulaCells.Count)
        let addIndexed (index: Dictionary<int, ResizeArray<CellAddress>>) key addr =
            match index.TryGetValue key with
            | true, cells -> cells.Add addr
            | false, _ ->
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

        let referencedFormulaCells (expr: Expr) =
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
                        match formulaRows.TryGetValue row with
                        | true, candidates ->
                            for addr in candidates do
                                if addr.Col >= sc && addr.Col <= ec then
                                    refs.Add addr
                        | false, _ -> ()
                else
                    let formulaCols = formulaColsIndex ()
                    for col in sc .. ec do
                        match formulaCols.TryGetValue col with
                        | true, candidates ->
                            for addr in candidates do
                                if addr.Row >= sr && addr.Row <= er then
                                    refs.Add addr
                        | false, _ -> ()
            let stack = ResizeArray<struct (Set<string> * Expr)>()
            stack.Add(struct (Set.empty, expr))
            while stack.Count > 0 do
                let last = stack.Count - 1
                let struct (seen, e) = stack.[last]
                stack.RemoveAt last
                match e with
                | Number _ | Str _ | Bool _ | ErrorLit _ | TableRef _ | ArrayLit _ -> ()
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

        let referencedFormulaCellsCache = Dictionary<string, CellAddress array>()
        let referencedFormulaCellsCacheOrder = Queue<string>()
        let referencedFormulaCellsFastSlots = Array.zeroCreate<RecalcReferenceFastCacheEntry> 256
        let mutable referencedFormulaCellsCacheTotalFormulaChars = 0
        let mutable referencedFormulaCellsCacheTotalRefs = 0
        let referencedFormulaCellsFor (fc: FormulaCell) =
            match tryGetRecalcReferenceFastSlot referencedFormulaCellsFastSlots fc.Formula with
            | ValueSome refs -> refs
            | ValueNone ->
                let mutable refs = Unchecked.defaultof<CellAddress array>
                if referencedFormulaCellsCache.TryGetValue(fc.Formula, &refs) then
                    setRecalcReferenceFastSlot referencedFormulaCellsFastSlots fc.Formula refs
                    refs
                else
                    refs <- referencedFormulaCells fc.Expr
                    rememberRecalcReferenceCacheEntry
                        referencedFormulaCellsCache
                        referencedFormulaCellsCacheOrder
                        &referencedFormulaCellsCacheTotalFormulaChars
                        &referencedFormulaCellsCacheTotalRefs
                        fc.Formula
                        refs
                    clearRecalcReferenceFastSlots referencedFormulaCellsFastSlots
                    if referencedFormulaCellsCache.ContainsKey fc.Formula then
                        setRecalcReferenceFastSlot referencedFormulaCellsFastSlots fc.Formula refs
                    refs

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
            match dependencies.TryGetValue addr with
            | true, deps -> deps
            | false, _ -> Array.empty

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
                    activeStack.RemoveAt(activeStack.Count - 1)
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
            match formulaByAddress.TryGetValue addr with
            | true, fc ->
                match fc.ParseError with
                | Some parseError ->
                    applyFormulaResult addr fc (CellValue.Error parseError)
                | None when cyclic.Contains addr ->
                    applyFormulaResult addr fc (CellValue.Error Err.cycle)
                | None ->
                    let updated = { model with Cells = updatedCells }
                    let ctx = { Model = updated; Workbook = workbook; Sheet = currentSheet; CriteriaCache = Some criteriaCache; CurrentCell = Some addr }
                    let result = eval fc.Expr ctx
                    applyFormulaResult addr fc result
            | false, _ -> ()
        rememberCellsFormulaCount updatedCells formulaCells.Count
        if changed then { model with Cells = updatedCells } else model

/// True if the expression tree uses a dynamic-array (spilling) function.
let containsArrayFunction (e: Expr) : bool =
    let stack = ResizeArray<Expr>()
    stack.Add e
    let mutable found = false
    while not found && stack.Count > 0 do
        let last = stack.Count - 1
        let expr = stack.[last]
        stack.RemoveAt last
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

/// Materialize dynamic-array spills for rendering. For each formula whose
/// formula uses a spilling function (SEQUENCE/UNIQUE/SORT/FILTER/…), evaluate
/// the array and write its values into the anchor + neighbouring cells
/// (row-major). A non-empty target cell (other than the anchor) blocks the
/// spill → the anchor shows #SPILL!. Pure; returns a new model. Only formulas
/// that actually use an array function are touched — every other cell is
/// untouched, so this is safe to run after `recalculateAllFormulas` on any
/// workbook. (Interactive clear-on-recalc + spill-cell dependencies are a
/// later step; this is the render-side materialization.)
let private spillArraysCore (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                            (currentSheet: RedBat.Spreadsheet.Workbook.SheetId option)
                            (model: SpreadsheetModel) : SpreadsheetModel * bool =
    let previousSpill = model.Spill
    let anchors = ResizeArray<CellAddress * CellData * string * Expr>()
    let arrayFormulaCache = Dictionary<string, Expr option>()
    let arrayFormulaCacheOrder = Queue<string>()
    let arrayFormulaFastSlots = Array.zeroCreate<RecalcArrayFormulaFastCacheEntry> 256
    let mutable arrayFormulaCacheTotalChars = 0
    let arrayFormulaExpr formula =
        match tryGetRecalcArrayFormulaFastSlot arrayFormulaFastSlots formula with
        | ValueSome cached -> cached
        | ValueNone ->
            let mutable cached = Unchecked.defaultof<Expr option>
            if arrayFormulaCache.TryGetValue(formula, &cached) then
                setRecalcArrayFormulaFastSlot arrayFormulaFastSlots formula cached
                cached
            else
                cached <-
                    match parseFormulaShared formula with
                    | Ok expr when containsArrayFunction expr -> Some expr
                    | _ -> None
                rememberRecalcArrayFormulaCacheEntry
                    arrayFormulaCache
                    arrayFormulaCacheOrder
                    &arrayFormulaCacheTotalChars
                    formula
                    cached
                clearRecalcArrayFormulaFastSlots arrayFormulaFastSlots
                if arrayFormulaCache.ContainsKey formula then
                    setRecalcArrayFormulaFastSlot arrayFormulaFastSlots formula cached
                cached
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
            model, false
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
                let mutable formulaCount = 0
                updatedCells
                |> Map.iter (fun _ cell ->
                    match cell.Value with
                    | CellValue.Formula _ -> formulaCount <- formulaCount + 1
                    | _ -> ())
                rememberCellsFormulaCount updatedCells formulaCount
                { model with Cells = updatedCells; Spill = Map.empty }, true
            else
                { model with Spill = Map.empty }, true
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
    let criteriaCache = CriteriaPredicateCache()
    for (addr, cell, formula, expr) in anchors do
        let updated = { model with Cells = updatedCells; Spill = spillMap }
        let arr = evalArray expr { Model = updated; Workbook = workbook; Sheet = currentSheet; CriteriaCache = Some criteriaCache; CurrentCell = Some addr }
        let nr = arr.Length
        let nc = if nr > 0 then arr.[0].Length else 0
        if nr * nc <= 1 then
            let v = if nr * nc = 1 then arr.[0].[0] else CellValue.Error Err.value
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
                setCellIfChanged addr { cell with Value = CellValue.Formula (formula, CellValue.Error Err.spill) }
            else
                setCellIfChanged addr { cell with Value = CellValue.Formula (formula, arr.[0].[0]) }
                for r in 0 .. nr - 1 do
                    for c in 0 .. nc - 1 do
                        if not (r = 0 && c = 0) then
                            let sa = { Row = addr.Row + r; Col = addr.Col + c }
                            setCellIfChanged sa { Value = arr.[r].[c]; Format = cell.Format; Comment = None }
                            spillMap <- Map.add sa addr spillMap
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
        model, false
    else
        if changed then
            let mutable formulaCount = 0
            updatedCells
            |> Map.iter (fun _ cell ->
                match cell.Value with
                | CellValue.Formula _ -> formulaCount <- formulaCount + 1
                | _ -> ())
            rememberCellsFormulaCount updatedCells formulaCount
        { model with Cells = updatedCells; Spill = spillMap }, true

let spillArrays (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                (currentSheet: RedBat.Spreadsheet.Workbook.SheetId option)
                (model: SpreadsheetModel) : SpreadsheetModel =
    spillArraysCore workbook currentSheet model |> fst

/// Recalculate every formula and (re)materialize dynamic-array spills,
/// clearing the previous spill's ghost cells. This is the interactive
/// edit→recompute step: a spilling formula (FILTER/SORT/GROUPBY/…) updates and
/// re-spills automatically when its source changes. Two recalc passes bracket
/// the spill so formulas that *read* spilled cells also settle.
let recalcAndSpill (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                   (currentSheet: RedBat.Spreadsheet.Workbook.SheetId option)
                   (model: SpreadsheetModel) : SpreadsheetModel =
    if Map.isEmpty model.Spill && not (cellsHaveFormulaCells model.Cells) then model else
    let recalculated = recalculateAllFormulas workbook currentSheet model
    let spilled, spillChanged = spillArraysCore workbook currentSheet recalculated
    if spillChanged then
        recalculateAllFormulas workbook currentSheet spilled
    else
        spilled
