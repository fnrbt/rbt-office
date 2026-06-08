module RedBat.Spreadsheet.Formula

open System
open System.Globalization
open RedBat.Spreadsheet.Model

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

type Expr =
    | Number of float
    | Str of string
    | Bool of bool
    /// Error literal written directly in the formula (`#REF!` etc). Carries
    /// the bare code.
    | ErrorLit of string
    | CellRef of RefAst
    | RangeRef of RefAst * RefAst
    | TableRef of tableName: string * columnName: string
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
    | TPlus | TMinus | TStar | TSlash | TCaret | TPercent | TAmp
    | TEq | TNe | TLt | TGt | TLe | TGe
    | TLParen | TRParen
    | TLBracket | TRBracket
    | TComma | TColon | TBang
    | TEnd

let tokenize (input: string) : Token list =
    let chars = input.ToCharArray()
    let n = chars.Length
    let tokens = System.Collections.Generic.List<Token>()
    let mutable i = 0
    while i < n do
        let c = chars.[i]
        if Char.IsWhiteSpace c then
            i <- i + 1
        elif Char.IsDigit c || (c = '.' && i + 1 < n && Char.IsDigit chars.[i + 1]) then
            // Number, with optional fraction and scientific exponent.
            let start = i
            while i < n && Char.IsDigit chars.[i] do i <- i + 1
            if i < n && chars.[i] = '.' then
                i <- i + 1
                while i < n && Char.IsDigit chars.[i] do i <- i + 1
            if i < n && (chars.[i] = 'e' || chars.[i] = 'E') then
                let save = i
                i <- i + 1
                if i < n && (chars.[i] = '+' || chars.[i] = '-') then i <- i + 1
                if i < n && Char.IsDigit chars.[i] then
                    while i < n && Char.IsDigit chars.[i] do i <- i + 1
                else
                    // 'E' not followed by an exponent — not part of the number.
                    i <- save
            let text = System.String(chars, start, i - start)
            tokens.Add(TNumber (Double.Parse(text, CultureInfo.InvariantCulture)))
        elif c = '"' then
            // String literal. "" is an escaped quote.
            i <- i + 1
            let sb = System.Text.StringBuilder()
            let mutable closed = false
            while i < n && not closed do
                if chars.[i] = '"' then
                    if i + 1 < n && chars.[i + 1] = '"' then
                        sb.Append('"') |> ignore
                        i <- i + 2
                    else
                        closed <- true
                        i <- i + 1
                else
                    sb.Append(chars.[i]) |> ignore
                    i <- i + 1
            tokens.Add(TString (sb.ToString()))
        elif c = '\'' then
            // Quoted sheet name. '' is an escaped apostrophe.
            i <- i + 1
            let sb = System.Text.StringBuilder()
            let mutable closed = false
            while i < n && not closed do
                if chars.[i] = '\'' then
                    if i + 1 < n && chars.[i + 1] = '\'' then
                        sb.Append('\'') |> ignore
                        i <- i + 2
                    else
                        closed <- true
                        i <- i + 1
                else
                    sb.Append(chars.[i]) |> ignore
                    i <- i + 1
            tokens.Add(TSheet (sb.ToString()))
        elif c = '#' then
            // Error literal: longest match against the known set.
            let rest = System.String(chars, i, n - i)
            match Err.all |> List.tryFind (fun e -> rest.StartsWith(e, StringComparison.OrdinalIgnoreCase)) with
            | Some lit ->
                tokens.Add(TError (lit.Substring 1))
                i <- i + lit.Length
            | None ->
                i <- i + 1  // stray '#', skip
        elif c = '$' || Char.IsLetter c || c = '_' then
            // Reference / name / function name. Allow '$', letters, digits,
            // '_', '.' so "$A$1", "Sheet_2", "_xlfn.FOO" all lex as one token.
            let start = i
            while i < n && (chars.[i] = '$' || Char.IsLetterOrDigit chars.[i] || chars.[i] = '_' || chars.[i] = '.') do
                i <- i + 1
            tokens.Add(TIdent (System.String(chars, start, i - start)))
        else
            // Operators / punctuation. Two-char operators checked first.
            let two = if i + 1 < n then System.String(chars, i, 2) else ""
            match two with
            | "<=" -> tokens.Add(TLe); i <- i + 2
            | ">=" -> tokens.Add(TGe); i <- i + 2
            | "<>" -> tokens.Add(TNe); i <- i + 2
            | _ ->
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
                | '[' -> tokens.Add(TLBracket); i <- i + 1
                | ']' -> tokens.Add(TRBracket); i <- i + 1
                | ',' -> tokens.Add(TComma); i <- i + 1
                | ':' -> tokens.Add(TColon); i <- i + 1
                | '!' -> tokens.Add(TBang); i <- i + 1
                | _ -> i <- i + 1  // skip unknown
    tokens.Add(TEnd)
    List.ofSeq tokens

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
    let colStart = idx
    while idx < len && Char.IsLetter token.[idx] do idx <- idx + 1
    let colPart = token.Substring(colStart, idx - colStart)
    let rowAbs = idx < len && token.[idx] = '$'
    if rowAbs then idx <- idx + 1
    let rowStart = idx
    while idx < len && Char.IsDigit token.[idx] do idx <- idx + 1
    let rowPart = token.Substring(rowStart, idx - rowStart)
    if idx <> len || colPart = "" || rowPart = "" then None
    else
        match Int32.TryParse rowPart with
        | true, r when r >= 1 ->
            Some {
                Sheet = None
                Row = r - 1
                Col = letterToColumn colPart
                RowAbs = rowAbs
                ColAbs = colAbs
            }
        | _ -> None

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

type private Parser = { Tokens: Token[]; mutable Pos: int }

let private peekP (p: Parser) = if p.Pos < p.Tokens.Length then p.Tokens.[p.Pos] else TEnd
let private nextP (p: Parser) = let t = peekP p in p.Pos <- p.Pos + 1; t
let private expectP (p: Parser) (t: Token) =
    if peekP p = t then p.Pos <- p.Pos + 1
    else failwithf "Expected %A but got %A" t (peekP p)

let rec private parsePrimary (p: Parser) : Expr =
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
    | TIdent ident ->
        p.Pos <- p.Pos + 1
        match peekP p with
        | TLParen ->
            // Function call.
            p.Pos <- p.Pos + 1
            let args = parseArgs p
            expectP p TRParen
            Function(ident, args)
        | TLBracket ->
            // Structured table reference: Table[Column].
            p.Pos <- p.Pos + 1
            match nextP p with
            | TIdent col ->
                expectP p TRBracket
                TableRef(ident, col)
            | other -> failwithf "Expected column name in table ref, got %A" other
        | TBang ->
            // Unquoted sheet-qualified ref: Sheet1!A1.
            p.Pos <- p.Pos + 1
            parseSheetQualified p (Some ident)
        | _ ->
            // Cell ref, defined name, or boolean literal.
            match ident.ToUpperInvariant() with
            | "TRUE" -> Bool true
            | "FALSE" -> Bool false
            | _ ->
                match parseRefToken ident with
                | Some refAst -> CellRef refAst
                | None -> NameRef ident
    | other -> failwithf "Unexpected token %A" other

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
        let body = if input.StartsWith "=" then input.Substring 1 else input
        let tokens = tokenize body |> Array.ofList
        let p = { Tokens = tokens; Pos = 0 }
        let e = parseComparison p
        match peekP p with
        | TEnd -> Ok e
        | t -> Result.Error (sprintf "Unexpected trailing token %A" t)
    with ex -> Result.Error ex.Message

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
let rec unwrapValue (value: CellValue) : CellValue =
    match value with
    | CellValue.Formula (_, cached) -> unwrapValue cached
    | other -> other

let private err code : CellValue = CellValue.Error code

/// First error among the operands, if any — for error propagation.
let private firstError (vs: CellValue list) : CellValue option =
    vs |> List.tryPick (fun v -> match unwrapValue v with CellValue.Error _ as e -> Some e | _ -> None)

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
        match s.Trim().ToUpperInvariant() with
        | "TRUE" -> Ok true
        | "FALSE" -> Ok false
        | "" -> Ok false
        | other ->
            match Double.TryParse(other, NumberStyles.Float, CultureInfo.InvariantCulture) with
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
    let rec go (seen: Set<string>) (e: Expr) : Set<CellAddress> =
        match e with
        | Number _ | Str _ | Bool _ | ErrorLit _ | TableRef _ -> Set.empty
        | NameRef name ->
            let key = name.ToUpperInvariant()
            if Set.contains key seen then Set.empty
            else
                match workbook |> Option.bind (resolveNameExpr name None) with
                | Some resolved -> go (Set.add key seen) resolved
                | None -> Set.empty
        | CellRef r when r.Sheet.IsSome -> Set.empty
        | CellRef r -> Set.singleton { Row = r.Row; Col = r.Col }
        | RangeRef (a, _) when a.Sheet.IsSome -> Set.empty
        | RangeRef (a, b) ->
            let sr, sc, er, ec = normalizeRange a b
            seq { for row in sr .. er do for col in sc .. ec -> { Row = row; Col = col } }
            |> Set.ofSeq
        | UnaryOp (_, inner) -> go seen inner
        | BinOp (l, _, r) -> Set.union (go seen l) (go seen r)
        | Function (_, args) -> args |> List.fold (fun acc a -> Set.union acc (go seen a)) Set.empty
    go Set.empty expr

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
    let rec go (seen: Set<string>) (e: Expr) : Set<RedBat.Spreadsheet.Workbook.SheetCell> =
        match e with
        | Number _ | Str _ | Bool _ | ErrorLit _ | TableRef _ -> Set.empty
        | NameRef name ->
            let key = name.ToUpperInvariant()
            if Set.contains key seen then Set.empty
            else
                match workbook |> Option.bind (resolveNameExpr name (Some homeSheet)) with
                | Some resolved -> go (Set.add key seen) resolved
                | None -> Set.empty
        | CellRef r ->
            match sheetIdOf r.Sheet with
            | Some sid ->
                Set.singleton
                    ({ Sheet = sid; Cell = { Row = r.Row; Col = r.Col } }
                     : RedBat.Spreadsheet.Workbook.SheetCell)
            | None -> Set.empty
        | RangeRef (a, b) ->
            match sheetIdOf a.Sheet with
            | Some sid ->
                let sr, sc, er, ec = normalizeRange a b
                seq { for row in sr .. er do
                        for col in sc .. ec ->
                            ({ Sheet = sid; Cell = { Row = row; Col = col } }
                             : RedBat.Spreadsheet.Workbook.SheetCell) }
                |> Set.ofSeq
            | None -> Set.empty
        | UnaryOp (_, inner) -> go seen inner
        | BinOp (l, _, r) -> Set.union (go seen l) (go seen r)
        | Function (_, args) -> args |> List.fold (fun acc a -> Set.union acc (go seen a)) Set.empty
    go Set.empty expr

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
    let rec go (seen: Set<string>) (e: Expr) =
        match e with
        | Number _ | Str _ | Bool _ | ErrorLit _ | TableRef _ -> ()
        | NameRef name ->
            let key = name.ToUpperInvariant()
            if not (Set.contains key seen) then
                match workbook |> Option.bind (resolveNameExpr name (Some homeSheet)) with
                | Some resolved -> go (Set.add key seen) resolved
                | None -> ()
        | CellRef r ->
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
        | UnaryOp (_, inner) -> go seen inner
        | BinOp (l, _, r) -> go seen l; go seen r
        | Function (_, args) -> for a in args do go seen a
    go Set.empty expr
    (Set.ofSeq cells, List.ofSeq rects)

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
    match expr with
    | Number _ | Str _ | Bool _ | ErrorLit _ | NameRef _ | TableRef _ -> expr
    | CellRef r -> CellRef (shiftRef r)
    | RangeRef (a, b) -> RangeRef (shiftRef a, shiftRef b)
    | UnaryOp (op, e) -> UnaryOp (op, shiftExpr dr dc e)
    | BinOp (l, op, r) -> BinOp (shiftExpr dr dc l, op, shiftExpr dr dc r)
    | Function (name, args) -> Function (name, args |> List.map (shiftExpr dr dc))

/// Replace every `NameRef` whose (case-insensitive) name is bound in `subst`
/// with its bound expression. This is how LET inlines its names: a LET value
/// expression is substituted as-is (not pre-evaluated), so a name bound to a
/// range stays a range for the structural functions (SUM/SUMIFS/INDEX/…) and
/// so dependency extraction still sees the real cell references.
let rec substituteNames (subst: Map<string, Expr>) (expr: Expr) : Expr =
    match expr with
    | NameRef n ->
        match Map.tryFind (n.ToUpperInvariant()) subst with
        | Some e -> e
        | None -> expr
    | Number _ | Str _ | Bool _ | ErrorLit _ | CellRef _ | RangeRef _ | TableRef _ -> expr
    | UnaryOp (op, e) -> UnaryOp (op, substituteNames subst e)
    | BinOp (l, op, r) -> BinOp (substituteNames subst l, op, substituteNames subst r)
    | Function (nm, fa) ->
        let fa' = fa |> List.map (substituteNames subst)
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
    match List.rev parts with
    | body :: revParams ->
        let prms = List.rev revParams
        let subst =
            List.zip (prms |> List.truncate (List.length args)) (args |> List.truncate (List.length prms))
            |> List.choose (fun (p, a) -> match p with NameRef n -> Some (n.ToUpperInvariant(), a) | _ -> None)
            |> Map.ofList
        substituteNames subst body
    | [] -> ErrorLit Err.value

/// The body of a LAMBDA, if `e` is one (`Function("LAMBDA", parts)`).
let lambdaPartsOf (e: Expr) : Expr list option =
    match e with
    | Function (n, parts) when n.ToUpperInvariant() = "LAMBDA" && not (List.isEmpty parts) -> Some parts
    | _ -> None

/// Wrap a computed value back as a literal expression, so it can be passed as
/// a LAMBDA argument (MAP/REDUCE/SCAN/MAKEARRAY feed values into lambdas).
let rec valueToExpr (v: CellValue) : Expr =
    match v with
    | CellValue.Number n -> Number n
    | CellValue.String s -> Str s
    | CellValue.Boolean b -> Bool b
    | CellValue.Error e -> ErrorLit e
    | CellValue.Empty -> Number 0.0
    | CellValue.Formula (_, inner) -> valueToExpr inner

/// True if the expression contains a volatile function (RAND/RANDBETWEEN/
/// NOW/TODAY/OFFSET/INDIRECT) anywhere in its tree. Such cells must be
/// recalculated on every pass regardless of dependency state.
let rec containsVolatile (expr: Expr) : bool =
    match expr with
    | Function (name, args) ->
        let n = name.ToUpperInvariant()
        n = "RAND" || n = "RANDBETWEEN" || n = "NOW" || n = "TODAY"
        || n = "OFFSET" || n = "INDIRECT"
        || List.exists containsVolatile args
    | UnaryOp (_, e) -> containsVolatile e
    | BinOp (l, _, r) -> containsVolatile l || containsVolatile r
    | _ -> false

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

/// Expand a RangeRef to the flat list of unwrapped cell values, row-major.
/// The sheet qualifier on the start ref selects the target sheet.
let expandRangeRef (a: RefAst) (b: RefAst) (model: SpreadsheetModel)
                   (workbook: RedBat.Spreadsheet.Workbook.Workbook option) : CellValue list =
    match resolveModel a.Sheet model workbook with
    | Some m ->
        let sr, sc, er, ec = normalizeRange a b
        [ for row in sr .. er do
            for col in sc .. ec ->
                getCellValue { Row = row; Col = col } m |> unwrapValue ]
    | None -> [ err Err.ref ]

/// Expand a structured table reference to its column's values.
let expandTableRef (tableName: string) (columnName: string)
                   (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                   (model: SpreadsheetModel) : CellValue list =
    match workbook with
    | Some wb ->
        match RedBat.Spreadsheet.Workbook.getTableColumn tableName columnName wb with
        | Some (_, startAddr, endAddr) ->
            [ for row in startAddr.Row .. endAddr.Row ->
                getCellValue { Row = row; Col = startAddr.Col } model |> unwrapValue ]
        | None -> [ err Err.ref ]
    | None -> [ err Err.ref ]

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
    let lower = text.ToLowerInvariant()
    let sb = System.Text.StringBuilder(lower.Length)
    let mutable startWord = true
    for ch in lower do
        if Char.IsLetter ch then
            sb.Append(if startWord then Char.ToUpperInvariant ch else ch) |> ignore
            startWord <- false
        else
            sb.Append ch |> ignore
            startWord <- not (Char.IsDigit ch)
    sb.ToString()

// --- Criteria parsing (SUMIF / COUNTIF / …) ---------------------------------
/// Translate Excel's `*` / `?` wildcards into a case-insensitive regex.
let private wildcardRegex (pattern: string) =
    let sb = System.Text.StringBuilder()
    sb.Append('^') |> ignore
    let mutable i = 0
    while i < pattern.Length do
        match pattern.[i] with
        | '~' when i + 1 < pattern.Length ->
            // ~* / ~? escape the wildcard.
            sb.Append(System.Text.RegularExpressions.Regex.Escape(string pattern.[i + 1])) |> ignore
            i <- i + 2
        | '*' -> sb.Append(".*") |> ignore; i <- i + 1
        | '?' -> sb.Append('.') |> ignore; i <- i + 1
        | c -> sb.Append(System.Text.RegularExpressions.Regex.Escape(string c)) |> ignore; i <- i + 1
    sb.Append('$') |> ignore
    System.Text.RegularExpressions.Regex(
        sb.ToString(), System.Text.RegularExpressions.RegexOptions.IgnoreCase)

/// Compile a criteria value (`">5"`, `"<>0"`, `"app*"`, `42`, …) into a
/// predicate over candidate cell values. Used by the *IF / *IFS family.
let private compileCriteria (criteria: CellValue) : CellValue -> bool =
    match unwrapValue criteria with
    | CellValue.Number cn ->
        fun v -> match coerceNumber v with Ok x -> x = cn | _ -> false
    | CellValue.Boolean cb ->
        fun v -> match coerceBool v with Ok x -> x = cb | _ -> false
    | CellValue.String raw ->
        let ops = [ ">="; "<="; "<>"; ">"; "<"; "=" ]
        let op, rest =
            match ops |> List.tryFind (fun o -> raw.StartsWith o) with
            | Some o -> o, raw.Substring(o.Length)
            | None -> "=", raw
        match Double.TryParse(rest, NumberStyles.Float, CultureInfo.InvariantCulture) with
        | true, num ->
            fun v ->
                match coerceNumber v with
                | Ok x ->
                    match op with
                    | ">=" -> x >= num | "<=" -> x <= num | ">" -> x > num
                    | "<" -> x < num | "<>" -> x <> num | _ -> x = num
                | _ -> false
        | _ ->
            match op with
            | ">" | "<" | ">=" | "<=" ->
                fun v ->
                    match coerceText v with
                    | Ok t ->
                        let c = String.Compare(t, rest, StringComparison.OrdinalIgnoreCase)
                        match op with
                        | ">" -> c > 0 | "<" -> c < 0 | ">=" -> c >= 0 | _ -> c <= 0
                    | _ -> false
            | "<>" ->
                let rx = wildcardRegex rest
                fun v -> match coerceText v with Ok t -> not (rx.IsMatch t) | _ -> true
            | _ ->
                let rx = wildcardRegex rest
                fun v -> match coerceText v with Ok t -> rx.IsMatch t | _ -> false
    | CellValue.Empty ->
        fun v -> match unwrapValue v with CellValue.Empty -> true | _ -> false
    | _ -> fun _ -> false

/// Static 0-indexed bounds of a range-shaped expression (RangeRef or a
/// single CellRef). None for anything else.
let private rangeBounds (e: Expr) : (int * int * int * int) option =
    match e with
    | RangeRef (a, b) -> Some (min a.Row b.Row, min a.Col b.Col, max a.Row b.Row, max a.Col b.Col)
    | CellRef r -> Some (r.Row, r.Col, r.Row, r.Col)
    | _ -> None

/// Compare two scalar values for the comparison operators. Excel orders
/// number < text < FALSE < TRUE; numbers compare numerically, text
/// case-insensitively. Returns -1 / 0 / 1, or an Error value to propagate.
let private compareValues (l: CellValue) (r: CellValue) : Result<int, string> =
    let rank v =
        match v with
        | CellValue.Number _ -> 0
        | CellValue.String _ -> 1
        | CellValue.Boolean _ -> 2
        | CellValue.Empty -> 0  // blank behaves as 0 / "" depending on the other side
        | _ -> 3
    match unwrapValue l, unwrapValue r with
    | CellValue.Error e, _ -> Result.Error e
    | _, CellValue.Error e -> Result.Error e
    | a, b ->
        match a, b with
        | CellValue.Number x, CellValue.Number y -> Ok (compare x y)
        | CellValue.String x, CellValue.String y ->
            Ok (String.Compare(x, y, StringComparison.OrdinalIgnoreCase))
        | CellValue.Boolean x, CellValue.Boolean y -> Ok (compare x y)
        | CellValue.Empty, CellValue.Empty -> Ok 0
        // Blank vs number/text: blank acts as 0 or "".
        | CellValue.Empty, CellValue.Number y -> Ok (compare 0.0 y)
        | CellValue.Number x, CellValue.Empty -> Ok (compare x 0.0)
        | CellValue.Empty, CellValue.String y -> Ok (String.Compare("", y, StringComparison.OrdinalIgnoreCase))
        | CellValue.String x, CellValue.Empty -> Ok (String.Compare(x, "", StringComparison.OrdinalIgnoreCase))
        // Cross-type: order by rank.
        | _ -> Ok (compare (rank a) (rank b))

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
let private arrayFunctionSet =
    set [ "SEQUENCE"; "RANDARRAY"; "UNIQUE"; "SORT"; "SORTBY"; "FILTER"; "TRANSPOSE"
          "MAP"; "REDUCE"; "SCAN"; "MAKEARRAY" ]
let isArrayFunction (name: string) = arrayFunctionSet.Contains(name.ToUpperInvariant())

/// Does this expression evaluate to an array (a spill)? An array function, or
/// an operator whose operand is itself an array (`SEQUENCE(5)*2`, `A1:A8>2`).
/// Aggregates consume such an argument over its whole array, not its top-left.
let rec isArrayExpr (e: Expr) : bool =
    match e with
    | Function (n, _) when isArrayFunction n -> true
    // A range/table inside an operator broadcasts element-wise (`A1:A8+B1:B8`).
    // A *bare* range arg never reaches this test — aggregates match RangeRef
    // first and stream it — so this only fires inside Unary/BinOp/array funcs.
    | RangeRef _ | TableRef _ -> true
    | UnaryOp (_, x) -> isArrayExpr x
    | BinOp (l, _, r) -> isArrayExpr l || isArrayExpr r
    | _ -> false

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
}

/// Build a context with no workbook — same-sheet formulas only, no defined
/// names. Convenient for callers that just have a bare model.
let contextOf (model: SpreadsheetModel) : EvalContext =
    { Model = model; Workbook = None; Sheet = None }

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
        match expandRangeRef a b model workbook with
        | v :: _ -> v
        | [] -> err Err.ref
    | TableRef (t, c) ->
        match expandTableRef t c workbook model with
        | v :: _ -> v
        | [] -> err Err.ref
    | UnaryOp (op, operand) -> applyUnary op (eval operand ctx)
    | BinOp (left, op, right) -> applyBin op (eval left ctx) (eval right ctx)
    | Function (name, args) ->
        // An array-producing function in scalar context collapses to its
        // top-left element (implicit intersection; spill placement is Step 2).
        if isArrayFunction name then
            let arr = evalArray expr ctx
            if arr.Length > 0 && arr.[0].Length > 0 then arr.[0].[0] else err Err.value
        else evalFunction name args ctx

/// Evaluate every argument, flattening ranges/tables into their cell values
/// (the form most aggregate functions want).
and flattenArgs (args: Expr list) (ctx: EvalContext) : CellValue list =
    let model = ctx.Model
    let workbook = ctx.Workbook
    args |> List.collect (fun arg ->
        match arg with
        | RangeRef (a, b) -> expandRangeRef a b model workbook
        | TableRef (t, c) -> expandTableRef t c workbook model
        | _ when isArrayExpr arg ->
            [ for row in evalArray arg ctx do for v in row -> unwrapValue v ]
        | NameRef name ->
            // A name standing in for a range must flatten like the range
            // would — otherwise SUM(MyRange) would see only its first cell.
            match workbook |> Option.bind (resolveNameExpr name ctx.Sheet) with
            | Some (RangeRef (a, b)) -> expandRangeRef a b model workbook
            | Some (TableRef (t, c)) -> expandTableRef t c workbook model
            | Some e -> [ eval e ctx ]
            | None -> [ err Err.name ]
        | _ -> [ eval arg ctx ])

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
    | TableRef (t, c) -> [| for v in expandTableRef t c ctx.Workbook ctx.Model -> [| v |] |]
    | NameRef name ->
        match ctx.Workbook |> Option.bind (resolveNameExpr name ctx.Sheet) with
        | Some (RangeRef (a, b)) -> evalArray (RangeRef (a, b)) ctx
        | Some (TableRef (t, c)) -> evalArray (TableRef (t, c)) ctx
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
    let nargs = List.length args
    let num i = match coerceNumber (eval (List.item i args) ctx) with Ok n -> n | _ -> nan
    let boolAt i = match coerceBool (eval (List.item i args) ctx) with Ok b -> b | _ -> false
    let errArr e : ArrVal = [| [| err e |] |]
    let firstCol (a: ArrVal) = a |> Array.map (fun row -> if row.Length > 0 then row.[0] else CellValue.Empty)
    let cmpVal a b = match compareValues a b with Ok c -> c | _ -> 0
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
        let arr = evalArray (List.item 0 args) ctx
        // Distinct rows, first occurrence kept; case-insensitive (engine-wide).
        let keyOf (row: CellValue[]) =
            row
            |> Array.map (fun v ->
                match v with
                | CellValue.Number n -> "n:" + n.ToString("R", CultureInfo.InvariantCulture)
                | CellValue.String s -> "s:" + s.ToLowerInvariant()
                | CellValue.Boolean b -> "b:" + string b
                | CellValue.Error e -> "x:" + e
                | _ -> "e")
            |> String.concat ""
        let seen = System.Collections.Generic.HashSet<string>()
        let outRows = ResizeArray<CellValue[]>()
        for row in arr do if seen.Add(keyOf row) then outRows.Add row
        if outRows.Count = 0 then errArr Err.calc else outRows.ToArray()
    | "SORT" ->
        if nargs < 1 then errArr Err.value else
        let arr = evalArray (List.item 0 args) ctx
        let col = (if nargs >= 2 then int (num 1) else 1) - 1
        let order = if nargs >= 3 then int (num 2) else 1
        arr
        |> Array.sortWith (fun r1 r2 ->
            let a = if col >= 0 && col < r1.Length then r1.[col] else CellValue.Empty
            let b = if col >= 0 && col < r2.Length then r2.[col] else CellValue.Empty
            let c = cmpVal a b in if order < 0 then -c else c)
    | "SORTBY" ->
        if nargs < 2 then errArr Err.value else
        let arr = evalArray (List.item 0 args) ctx
        let key = firstCol (evalArray (List.item 1 args) ctx)
        let order = if nargs >= 3 then int (num 2) else 1
        Array.init arr.Length id
        |> Array.sortWith (fun i j ->
            let a = if i < key.Length then key.[i] else CellValue.Empty
            let b = if j < key.Length then key.[j] else CellValue.Empty
            let c = cmpVal a b in if order < 0 then -c else c)
        |> Array.map (fun i -> arr.[i])
    | "FILTER" ->
        if nargs < 2 then errArr Err.value else
        let arr = evalArray (List.item 0 args) ctx
        let incl = firstCol (evalArray (List.item 1 args) ctx)
        let kept =
            [| for i in 0 .. arr.Length - 1 do
                 if i < incl.Length && (match coerceBool incl.[i] with Ok true -> true | _ -> false) then yield arr.[i] |]
        if kept.Length > 0 then kept
        elif nargs >= 3 then [| [| eval (List.item 2 args) ctx |] |]
        else errArr Err.calc
    | "TRANSPOSE" ->
        if nargs < 1 then errArr Err.value else
        let arr = evalArray (List.item 0 args) ctx
        let nr = arr.Length
        let nc = if nr > 0 then arr.[0].Length else 0
        [| for c in 0 .. nc - 1 -> [| for r in 0 .. nr - 1 -> arr.[r].[c] |] |]
    | "MAP" ->
        // MAP(array, lambda) — apply the lambda to each element. (Excel allows
        // several parallel arrays; v1 supports one.)
        if nargs < 2 then errArr Err.value else
        match lambdaPartsOf (List.item (nargs - 1) args) with
        | None -> errArr Err.value
        | Some parts ->
            evalArray (List.item 0 args) ctx
            |> Array.map (Array.map (fun v -> eval (applyLambdaExpr parts [ valueToExpr v ]) ctx))
    | "REDUCE" ->
        // REDUCE(initial, array, lambda(acc, value)) → scalar accumulator.
        if nargs <> 3 then errArr Err.value else
        match lambdaPartsOf (List.item 2 args) with
        | None -> errArr Err.value
        | Some parts ->
            let mutable acc = eval (List.item 0 args) ctx
            for row in evalArray (List.item 1 args) ctx do
                for v in row do acc <- eval (applyLambdaExpr parts [ valueToExpr acc; valueToExpr v ]) ctx
            [| [| acc |] |]
    | "SCAN" ->
        // SCAN(initial, array, lambda(acc, value)) → array of running accumulators.
        if nargs <> 3 then errArr Err.value else
        match lambdaPartsOf (List.item 2 args) with
        | None -> errArr Err.value
        | Some parts ->
            let mutable acc = eval (List.item 0 args) ctx
            evalArray (List.item 1 args) ctx
            |> Array.map (Array.map (fun v ->
                acc <- eval (applyLambdaExpr parts [ valueToExpr acc; valueToExpr v ]) ctx
                acc))
    | "MAKEARRAY" ->
        // MAKEARRAY(rows, cols, lambda(r, c)) — r,c are 1-based.
        if nargs <> 3 then errArr Err.value else
        match lambdaPartsOf (List.item 2 args) with
        | None -> errArr Err.value
        | Some parts ->
            let rows = int (num 0)
            let cols = int (num 1)
            if rows < 1 || cols < 1 then errArr Err.value
            else Array.init rows (fun r -> Array.init cols (fun c ->
                    eval (applyLambdaExpr parts [ Number (float (r + 1)); Number (float (c + 1)) ]) ctx))
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
        | TableRef (t, c) -> for v in expandTableRef t c workbook model do s <- f.Step s v
        | _ when isArrayExpr arg ->
            for row in evalArray arg ctx do
                for v in row do s <- f.Step s (unwrapValue v)
        | NameRef n ->
            match workbook |> Option.bind (resolveNameExpr n ctx.Sheet) with
            | Some (RangeRef (a, b)) -> iterRange a b
            | Some (TableRef (t, c)) -> for v in expandTableRef t c workbook model do s <- f.Step s v
            | Some inner -> s <- f.Step s (unwrapValue (eval inner ctx))
            | None -> s <- f.Step s (err Err.name)
        | _ -> s <- f.Step s (unwrapValue (eval arg ctx))
    f.Extract s

/// Numeric values from a flattened argument list — text and blanks dropped,
/// matching how SUM/AVERAGE/etc. ignore non-numeric cells. The first error
/// encountered is returned via the error channel so it can propagate.
and numbersOf (values: CellValue list) : Result<float list, string> =
    match firstError values with
    | Some (CellValue.Error e) -> Result.Error e
    | _ ->
        values
        |> List.choose (fun v ->
            match unwrapValue v with
            | CellValue.Number n -> Some n
            | CellValue.Boolean b -> Some (if b then 1.0 else 0.0)
            | _ -> None)
        |> Ok

/// Read a cell straight from the model (used by lookup functions that
/// index into a range rectangle).
and private gridCell (model: SpreadsheetModel) row col : CellValue =
    getCellValue { Row = row; Col = col } model |> unwrapValue

and evalFunction (name: string) (args: Expr list) (ctx: EvalContext) : CellValue =
    let model = ctx.Model
    let workbook = ctx.Workbook
    let nargs = List.length args
    let argRaw i = eval (List.item i args) ctx
    let argNum i = coerceNumber (argRaw i)
    let argText i = coerceText (argRaw i)
    let argBool i = coerceBool (argRaw i)
    let aggregateNumbers () = numbersOf (flattenArgs args ctx)
    // Run f over the aggregate numeric values; common shape for math
    // reductions. Empty list → emptyResult.
    let reduceNums emptyResult f =
        match aggregateNumbers () with
        | Result.Error e -> err e
        | Ok [] -> emptyResult
        | Ok ns -> f ns
    // Apply a unary numeric function to arg 0 with error propagation.
    let unaryNum f =
        match argNum 0 with Ok x -> f x | Result.Error e -> err e
    // (criteriaRangeExpr, criteriaValue) → the predicate-matched subset of
    // a parallel range, as a flat row-major list. Used by *IF / *IFS.
    let rangeValues (e: Expr) =
        match rangeBounds e with
        | Some (sr, sc, er, ec) ->
            [ for r in sr .. er do for c in sc .. ec -> gridCell model r c ]
        | None -> [ eval e ctx ]
    // Median of a sorted float list.
    let median (ns: float list) =
        let a = ns |> List.sort |> Array.ofList
        let n = a.Length
        if n = 0 then nan
        elif n % 2 = 1 then a.[n / 2]
        else (a.[n / 2 - 1] + a.[n / 2]) / 2.0
    let isErr v = match unwrapValue v with CellValue.Error _ -> true | _ -> false
    // Shared index resolver for XLOOKUP / XMATCH. Streams the lookup range by
    // index (no list), returning the 0-based offset of the match. match_mode:
    // 0 exact, -1 exact-or-next-smaller, 1 exact-or-next-larger, 2 wildcard.
    // search_mode: 1 first→last, -1 last→first (±2 binary treated as linear).
    let xMatchIndex (target: CellValue) (la: RangeAccess) (matchMode: int) (searchMode: int) : int option =
        let cmp a b = match compareValues a b with Ok c -> Some c | _ -> None
        let order = if searchMode = -1 || searchMode = -2 then seq { la.Count - 1 .. -1 .. 0 } else seq { 0 .. la.Count - 1 }
        match matchMode with
        | 2 ->
            match coerceText target with
            | Ok pat ->
                let rx = wildcardRegex pat
                order |> Seq.tryFind (fun k -> match coerceText (rangeRead la k) with Ok t -> rx.IsMatch t | _ -> false)
            | _ -> None
        | -1 | 1 ->
            match order |> Seq.tryFind (fun k -> cmp target (rangeRead la k) = Some 0) with
            | Some k -> Some k
            | None ->
                // No exact hit: closest value on the requested side
                // (-1 → largest value < target; 1 → smallest value > target).
                let mutable best : (int * CellValue) option = None
                for k in order do
                    let v = rangeRead la k
                    let onSide = match cmp v target with Some c -> (if matchMode = -1 then c < 0 else c > 0) | None -> false
                    if onSide then
                        match best with
                        | None -> best <- Some (k, v)
                        | Some (_, bv) -> if (match cmp v bv with Some c -> (if matchMode = -1 then c > 0 else c < 0) | None -> false) then best <- Some (k, v)
                best |> Option.map fst
        | _ ->
            order |> Seq.tryFind (fun k -> cmp target (rangeRead la k) = Some 0)

    match name.ToUpperInvariant() with
    // ----- aggregates (streaming one-pass folds; see foldArgs / numAggFold) --
    | "SUM" -> foldArgs (numAggFold 0.0 (+) (fun _ a -> CellValue.Number a)) args ctx
    | "PRODUCT" -> foldArgs (numAggFold 1.0 ( * ) (fun n a -> CellValue.Number (if n = 0 then 0.0 else a))) args ctx
    | "AVERAGE" -> foldArgs (numAggFold 0.0 (+) (fun n a -> if n = 0 then err Err.div0 else CellValue.Number (a / float n))) args ctx
    | "MIN" -> foldArgs (numAggFold infinity min (fun n a -> CellValue.Number (if n = 0 then 0.0 else a))) args ctx
    | "MAX" -> foldArgs (numAggFold (-infinity) max (fun n a -> CellValue.Number (if n = 0 then 0.0 else a))) args ctx
    | "MEDIAN" -> reduceNums (err Err.num) (fun ns -> CellValue.Number (median ns))
    | "COUNT" -> foldArgs (countFold true (fun v -> match v with CellValue.Number _ -> true | _ -> false)) args ctx
    | "COUNTA" -> foldArgs (countFold false (fun v -> match v with CellValue.Empty -> false | _ -> true)) args ctx
    | "COUNTBLANK" ->
        foldArgs (countFold false (fun v -> match v with CellValue.Empty | CellValue.String "" -> true | _ -> false)) args ctx
    | "STDEV" | "STDEVP" ->
        reduceNums (err Err.div0) (fun ns ->
            let n = List.length ns
            if (name.ToUpperInvariant() = "STDEV" && n < 2) then err Err.div0
            else
                let mean = List.average ns
                let ss = ns |> List.sumBy (fun x -> (x - mean) ** 2.0)
                let denom = if name.ToUpperInvariant() = "STDEV" then float (n - 1) else float n
                CellValue.Number (sqrt (ss / denom)))
    | "VAR" ->
        reduceNums (err Err.div0) (fun ns ->
            let n = List.length ns
            if n < 2 then err Err.div0
            else
                let mean = List.average ns
                CellValue.Number ((ns |> List.sumBy (fun x -> (x - mean) ** 2.0)) / float (n - 1)))
    | "LARGE" | "SMALL" ->
        if nargs <> 2 then err Err.value else
        match numbersOf (flattenArgs [ List.item 0 args ] ctx), argNum 1 with
        | Result.Error e, _ -> err e
        | _, Result.Error e -> err e
        | Ok ns, Ok k ->
            let sorted =
                if name.ToUpperInvariant() = "LARGE" then List.sortDescending ns
                else List.sort ns
            let ki = int k - 1
            if ki < 0 || ki >= List.length sorted then err Err.num
            else CellValue.Number sorted.[ki]
    // ----- criteria aggregates (streamed by index; see rangeAccessOf) -------
    | "COUNTIF" ->
        if nargs <> 2 then err Err.value else
        match rangeAccessOf model workbook (List.item 0 args) with
        | None -> err Err.value
        | Some a ->
            let pred = compileCriteria (argRaw 1)
            let mutable n = 0
            for k in 0 .. a.Count - 1 do if pred (rangeRead a k) then n <- n + 1
            CellValue.Number (float n)
    | "SUMIF" ->
        // SUMIF(range, criteria, [sum_range])
        if nargs < 2 then err Err.value else
        match rangeAccessOf model workbook (List.item 0 args) with
        | None -> err Err.value
        | Some crit ->
            let sum =
                if nargs >= 3 then (match rangeAccessOf model workbook (List.item 2 args) with Some s -> s | None -> crit)
                else crit
            let pred = compileCriteria (argRaw 1)
            let count = min crit.Count sum.Count
            let mutable total = 0.0
            for k in 0 .. count - 1 do
                if pred (rangeRead crit k) then
                    match coerceNumber (rangeRead sum k) with Ok n -> total <- total + n | _ -> ()
            CellValue.Number total
    | "AVERAGEIF" ->
        if nargs < 2 then err Err.value else
        match rangeAccessOf model workbook (List.item 0 args) with
        | None -> err Err.value
        | Some crit ->
            let avg =
                if nargs >= 3 then (match rangeAccessOf model workbook (List.item 2 args) with Some s -> s | None -> crit)
                else crit
            let pred = compileCriteria (argRaw 1)
            let count = min crit.Count avg.Count
            let mutable total = 0.0
            let mutable n = 0
            for k in 0 .. count - 1 do
                if pred (rangeRead crit k) then
                    match coerceNumber (rangeRead avg k) with Ok x -> total <- total + x; n <- n + 1 | _ -> ()
            if n = 0 then err Err.div0 else CellValue.Number (total / float n)
    | "SUMIFS" | "COUNTIFS" ->
        // SUMIFS(sum_range, crit_range1, crit1, crit_range2, crit2, …)
        // COUNTIFS(crit_range1, crit1, …)  — no sum range. Streamed by index:
        // each (crit range, predicate) pair is read in lockstep, no lists.
        let isSum = name.ToUpperInvariant() = "SUMIFS"
        let pairsStart = if isSum then 1 else 0
        let pairArgs = args |> List.skip pairsStart
        if List.length pairArgs % 2 <> 0 || List.isEmpty pairArgs then err Err.value else
        let pairs =
            pairArgs |> List.chunkBySize 2
            |> List.map (fun chunk ->
                match chunk with
                | [ rangeExpr; critExpr ] -> rangeAccessOf model workbook rangeExpr, compileCriteria (eval critExpr ctx)
                | _ -> None, (fun _ -> false))
        if pairs |> List.exists (fun (a, _) -> Option.isNone a) then err Err.value else
        let crits = pairs |> List.map (fun (a, p) -> a.Value, p)
        let count = crits |> List.map (fun (a, _) -> a.Count) |> List.min
        let sumAcc = if isSum then rangeAccessOf model workbook (List.item 0 args) else None
        let mutable total = 0.0
        let mutable hits = 0
        for k in 0 .. count - 1 do
            if crits |> List.forall (fun (a, pred) -> pred (rangeRead a k)) then
                hits <- hits + 1
                match sumAcc with
                | Some s when k < s.Count -> (match coerceNumber (rangeRead s k) with Ok n -> total <- total + n | _ -> ())
                | _ -> ()
        if isSum then CellValue.Number total else CellValue.Number (float hits)
    | "MAXIFS" | "MINIFS" | "AVERAGEIFS" ->
        // MAXIFS(value_range, crit_range1, crit1, …); same shape as SUMIFS.
        // One streamed pass tracks the matched count + running min/max/sum.
        let fn = name.ToUpperInvariant()
        let pairArgs = args |> List.skip 1
        if List.length pairArgs % 2 <> 0 || List.isEmpty pairArgs then err Err.value else
        let pairs =
            pairArgs |> List.chunkBySize 2
            |> List.map (fun chunk ->
                match chunk with
                | [ rangeExpr; critExpr ] -> rangeAccessOf model workbook rangeExpr, compileCriteria (eval critExpr ctx)
                | _ -> None, (fun _ -> false))
        match rangeAccessOf model workbook (List.item 0 args) with
        | Some valAcc when not (pairs |> List.exists (fun (a, _) -> Option.isNone a)) ->
            let crits = pairs |> List.map (fun (a, p) -> a.Value, p)
            let count = (valAcc.Count :: (crits |> List.map (fun (a, _) -> a.Count))) |> List.min
            let mutable n = 0
            let mutable acc = if fn = "MINIFS" then infinity elif fn = "MAXIFS" then -infinity else 0.0
            for k in 0 .. count - 1 do
                if crits |> List.forall (fun (a, pred) -> pred (rangeRead a k)) then
                    match coerceNumber (rangeRead valAcc k) with
                    | Ok x ->
                        n <- n + 1
                        acc <- match fn with "MINIFS" -> min acc x | "MAXIFS" -> max acc x | _ -> acc + x
                    | _ -> ()
            match fn with
            | "AVERAGEIFS" -> if n = 0 then err Err.div0 else CellValue.Number (acc / float n)
            | _ -> CellValue.Number (if n = 0 then 0.0 else acc)
        | _ -> err Err.value
    | "SUMPRODUCT" ->
        // Element-wise product of parallel ranges, then sum.
        let cols = args |> List.map (fun a -> numbersOf (rangeValues a))
        match cols |> List.tryPick (function Result.Error e -> Some e | _ -> None) with
        | Some e -> err e
        | None ->
            let arrs = cols |> List.choose (function Ok ns -> Some ns | _ -> None)
            match arrs with
            | [] -> CellValue.Number 0.0
            | first :: _ ->
                let len = arrs |> List.map List.length |> List.min
                let mutable total = 0.0
                for i in 0 .. len - 1 do
                    total <- total + (arrs |> List.fold (fun acc a -> acc * List.item i a) 1.0)
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
                match name.ToUpperInvariant() with
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
        let up = name.ToUpperInvariant() = "CEILING"
        match argNum 0, (if nargs >= 2 then argNum 1 else Ok 1.0) with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok x, Ok sg -> if sg = 0.0 then CellValue.Number 0.0 else CellValue.Number ((if up then ceil (x / sg) else floor (x / sg)) * sg)
    | "MROUND" ->
        match argNum 0, argNum 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok x, Ok m -> if m = 0.0 then CellValue.Number 0.0 else CellValue.Number (Math.Round(x / m, MidpointRounding.AwayFromZero) * m)
    | "EVEN" -> unaryNum (fun x -> let n = ceil (abs x / 2.0) * 2.0 in CellValue.Number (if x < 0.0 then -n else n))
    | "ODD" -> unaryNum (fun x -> let n = 2.0 * ceil ((abs x - 1.0) / 2.0) + 1.0 in CellValue.Number (if x < 0.0 then -n else n))
    | "SUMSQ" -> (match aggregateNumbers () with Result.Error e -> err e | Ok ns -> CellValue.Number (ns |> List.sumBy (fun x -> x * x)))
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
        match aggregateNumbers () with
        | Result.Error e -> err e
        | Ok ns ->
            let ints = ns |> List.map (fun x -> int64 (floor (abs x)))
            let rec gcd2 a b = if b = 0L then a else gcd2 b (a % b)
            if name.ToUpperInvariant() = "GCD" then CellValue.Number (float (List.fold gcd2 0L ints))
            else CellValue.Number (float (ints |> List.fold (fun acc x -> if acc = 0L || x = 0L then 0L else acc / gcd2 acc x * x) 1L))
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
    | "EXACT" ->
        match argText 0, argText 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok a, Ok b -> CellValue.Boolean (String.Equals(a, b, StringComparison.Ordinal))
    | "CHAR" -> (match argNum 0 with Result.Error e -> err e | Ok n -> (let i = int n in if i < 1 || i > 255 then err Err.value else CellValue.String (string (char i))))
    | "UNICHAR" -> (match argNum 0 with Result.Error e -> err e | Ok n -> (let i = int n in if i < 1 then err Err.value else CellValue.String (Char.ConvertFromUtf32 i)))
    | "CODE" | "UNICODE" -> (match argText 0 with Result.Error e -> err e | Ok s -> if s.Length = 0 then err Err.value else CellValue.Number (float (int s.[0])))
    | "CLEAN" -> (match argText 0 with Result.Error e -> err e | Ok s -> CellValue.String (System.String(s.ToCharArray() |> Array.filter (fun c -> int c >= 32))))
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
            elif name.ToUpperInvariant() = "TEXTBEFORE" then CellValue.String (s.Substring(0, idx))
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
        let bools =
            flattenArgs args ctx
            |> List.choose (fun v ->
                match unwrapValue v with
                | CellValue.Empty -> None
                | other -> Some (coerceBool other))
        match bools |> List.tryPick (function Result.Error e -> Some e | _ -> None) with
        | Some e -> err e
        | None ->
            let bs = bools |> List.choose (function Ok b -> Some b | _ -> None)
            if List.isEmpty bs then err Err.value
            else
                match name.ToUpperInvariant() with
                | "AND" -> CellValue.Boolean (List.forall id bs)
                | "OR" -> CellValue.Boolean (List.exists id bs)
                | _ -> CellValue.Boolean (bs |> List.filter id |> List.length |> fun n -> n % 2 = 1)
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
            if i = nargs - 1 then eval (substituteNames subst (List.item i args)) ctx
            else
                match List.item i args with
                | NameRef nm ->
                    let valExpr = substituteNames subst (List.item (i + 1) args)
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
        let parts = flattenArgs args ctx |> List.map coerceText
        match parts |> List.tryPick (function Result.Error e -> Some e | _ -> None) with
        | Some e -> err e
        | None -> CellValue.String (parts |> List.choose (function Ok s -> Some s | _ -> None) |> String.concat "")
    | "TEXTJOIN" ->
        // TEXTJOIN(delimiter, ignore_empty, text…)
        if nargs < 3 then err Err.value else
        match argText 0, argBool 1 with
        | Result.Error e, _ | _, Result.Error e -> err e
        | Ok delim, Ok ignoreEmpty ->
            let parts =
                flattenArgs (args |> List.skip 2) ctx
                |> List.choose (fun v ->
                    match unwrapValue v with
                    | CellValue.Empty when ignoreEmpty -> None
                    | other -> match coerceText other with Ok s -> Some s | _ -> Some "")
            CellValue.String (String.Join(delim, parts))
    | "LEN" ->
        if nargs <> 1 then err Err.value
        else match argText 0 with Ok s -> CellValue.Number (float s.Length) | Result.Error e -> err e
    | "TRIM" ->
        if nargs <> 1 then err Err.value
        else match argText 0 with
             | Ok s ->
                // Excel TRIM collapses runs of interior spaces to one.
                let collapsed = System.Text.RegularExpressions.Regex.Replace(s.Trim(), " +", " ")
                CellValue.String collapsed
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
                // Fable's IndexOf only accepts (needle, startIndex). For SEARCH's
                // case-insensitive behaviour we lowercase both sides — invariant
                // case folding preserves codepoint indices for BMP characters,
                // which is all the spreadsheet ever sees in practice.
                let idx =
                    if name.ToUpperInvariant() = "FIND" then
                        hay.IndexOf(needle, startAt)
                    else
                        hay.ToLowerInvariant().IndexOf(needle.ToLowerInvariant(), startAt)
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
                    let sb = System.Text.StringBuilder()
                    while idx < s.Length do
                        if idx + oldT.Length <= s.Length && s.Substring(idx, oldT.Length) = oldT then
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
        match rangeBounds (List.item 0 args) with
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
        match rangeAccessOf model workbook (List.item 1 args) with
        | None -> err Err.ref
        | Some la ->
            let target = argRaw 0
            let matchType = if nargs >= 3 then (match argNum 2 with Ok m -> int m | _ -> 1) else 1
            let result =
                if matchType = 0 then
                    seq { 0 .. la.Count - 1 }
                    |> Seq.tryFind (fun k -> match compareValues target (rangeRead la k) with Ok 0 -> true | _ -> false)
                else
                    // type 1: largest value ≤ target (asc); -1: smallest ≥ target (desc).
                    // Both pick the LAST index satisfying the side condition (old tryLast).
                    let ok v = match compareValues v target with Ok c -> (if matchType = 1 then c <= 0 else c >= 0) | _ -> false
                    let mutable last = None
                    for k in 0 .. la.Count - 1 do if ok (rangeRead la k) then last <- Some k
                    last
            match result with
            | Some k -> CellValue.Number (float (k + 1))
            | None -> err Err.na
    | "VLOOKUP" | "HLOOKUP" ->
        // VLOOKUP(lookup, table, col_index, [range_lookup])
        // HLOOKUP(lookup, table, row_index, [range_lookup]). Streamed: scan the
        // first column/row by index, then read the result cell directly.
        if nargs < 3 then err Err.value else
        match rangeAccessOf model workbook (List.item 1 args) with
        | None -> err Err.ref
        | Some t ->
            let isV = name.ToUpperInvariant() = "VLOOKUP"
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
                let hit =
                    if approx then
                        // largest value ≤ target (assumes sorted ascending)
                        let mutable last = None
                        for i in 0 .. lineLen - 1 do
                            if (match compareValues (lineCell i) target with Ok c -> c <= 0 | _ -> false) then last <- Some i
                        last
                    else
                        seq { 0 .. lineLen - 1 }
                        |> Seq.tryFind (fun i -> match compareValues target (lineCell i) with Ok 0 -> true | _ -> false)
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
        match rangeAccessOf model workbook (List.item 1 args), rangeAccessOf model workbook (List.item 2 args) with
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
        match rangeAccessOf model workbook (List.item 1 args) with
        | Some la ->
            let matchMode = if nargs >= 3 then (match argNum 2 with Ok m -> int m | _ -> 0) else 0
            let searchMode = if nargs >= 4 then (match argNum 3 with Ok s -> int s | _ -> 1) else 1
            match xMatchIndex target la matchMode searchMode with
            | Some k -> CellValue.Number (float (k + 1))
            | None -> err Err.na
        | None -> err Err.value
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
            match List.item 0 args with
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
            match parseFormula ("=" + refText) with
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
    match parseFormula formula with
    | Ok expr ->
        let ctx =
            { Model = model
              Workbook = workbook
              Sheet = workbook |> Option.map (fun w -> w.ActiveSheet) }
        eval expr ctx
    | Result.Error msg -> CellValue.Error msg

// ---------------------------------------------------------------------------
// Whole-model recalculation (dependency-ordered). Calc.fs owns the
// incremental engine; this is the simple full pass the shared model uses.
// ---------------------------------------------------------------------------

/// Recalculate every formula cell in a model, in dependency order.
/// `currentSheet` scopes defined-name resolution to the sheet being
/// recalculated (the caller knows which sheet `model` belongs to).
let recalculateAllFormulas (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                           (currentSheet: RedBat.Spreadsheet.Workbook.SheetId option)
                           (model: SpreadsheetModel) : SpreadsheetModel =
    let formulaCells =
        model.Cells
        |> Map.toList
        |> List.choose (fun (addr, cell) ->
            match cell.Value with
            | CellValue.Formula (formula, _) ->
                match parseFormula formula with
                | Ok expr -> Some (addr, cell, formula, expr, None)
                | Result.Error error -> Some (addr, cell, formula, Number 0.0, Some error)
            | _ -> None)

    let formulaByAddress =
        formulaCells
        |> List.map (fun (addr, cell, formula, expr, parseError) -> addr, (cell, formula, expr, parseError))
        |> Map.ofList

    let dependencies =
        formulaCells
        |> List.map (fun (addr, _, _, expr, parseError) ->
            let refs =
                match parseError with
                | Some _ -> Set.empty
                | None ->
                    referencesInExpr workbook expr
                    |> Set.filter (fun refAddr -> Map.containsKey refAddr formulaByAddress)
            addr, refs)
        |> Map.ofList

    let mutable visited = Set.empty
    let mutable visiting = Set.empty
    let mutable order = []
    let mutable cyclic = Set.empty

    let rec visit path addr =
        if Set.contains addr visiting then
            let cycleNodes =
                path
                |> List.takeWhile ((<>) addr)
                |> fun nodes -> addr :: nodes
            cyclic <- Set.union cyclic (Set.ofList cycleNodes)
        elif not (Set.contains addr visited) then
            visiting <- Set.add addr visiting
            let deps = Map.tryFind addr dependencies |> Option.defaultValue Set.empty
            for dep in deps do
                visit (addr :: path) dep
            visiting <- Set.remove addr visiting
            visited <- Set.add addr visited
            order <- addr :: order

    formulaCells |> List.iter (fun (addr, _, _, _, _) -> visit [] addr)

    let mutable updated = model
    for addr in List.rev order do
        match Map.tryFind addr formulaByAddress with
        | Some (cell, formula, _, Some parseError) ->
            updated <- setCell addr { cell with Value = CellValue.Formula (formula, CellValue.Error parseError) } updated
        | Some (cell, formula, _, None) when Set.contains addr cyclic ->
            updated <- setCell addr { cell with Value = CellValue.Formula (formula, CellValue.Error Err.cycle) } updated
        | Some (cell, formula, expr, None) ->
            let ctx = { Model = updated; Workbook = workbook; Sheet = currentSheet }
            let result = eval expr ctx
            updated <- setCell addr { cell with Value = CellValue.Formula (formula, result) } updated
        | None -> ()
    updated

/// True if the expression tree uses a dynamic-array (spilling) function.
let rec containsArrayFunction (e: Expr) : bool =
    match e with
    | Function (n, _) when isArrayFunction n -> true
    | Function (_, args) -> List.exists containsArrayFunction args
    | UnaryOp (_, x) -> containsArrayFunction x
    | BinOp (l, _, r) -> containsArrayFunction l || containsArrayFunction r
    | _ -> false

/// Materialize dynamic-array spills for rendering. For each formula whose
/// formula uses a spilling function (SEQUENCE/UNIQUE/SORT/FILTER/…), evaluate
/// the array and write its values into the anchor + neighbouring cells
/// (row-major). A non-empty target cell (other than the anchor) blocks the
/// spill → the anchor shows #SPILL!. Pure; returns a new model. Only formulas
/// that actually use an array function are touched — every other cell is
/// untouched, so this is safe to run after `recalculateAllFormulas` on any
/// workbook. (Interactive clear-on-recalc + spill-cell dependencies are a
/// later step; this is the render-side materialization.)
let spillArrays (workbook: RedBat.Spreadsheet.Workbook.Workbook option)
                (currentSheet: RedBat.Spreadsheet.Workbook.SheetId option)
                (model: SpreadsheetModel) : SpreadsheetModel =
    let ctx = { Model = model; Workbook = workbook; Sheet = currentSheet }
    let anchors =
        model.Cells
        |> Map.toList
        |> List.choose (fun (addr, cell) ->
            match cell.Value with
            | CellValue.Formula (formula, _) ->
                match parseFormula formula with
                | Ok expr when containsArrayFunction expr -> Some (addr, cell, formula, expr)
                | _ -> None
            | _ -> None)
    let mutable updated = model
    for (addr, cell, formula, expr) in anchors do
        let arr = evalArray expr ctx
        let nr = arr.Length
        let nc = if nr > 0 then arr.[0].Length else 0
        if nr * nc <= 1 then
            let v = if nr * nc = 1 then arr.[0].[0] else CellValue.Error Err.value
            updated <- setCell addr { cell with Value = CellValue.Formula (formula, v) } updated
        else
            let occupied r c =
                if r = 0 && c = 0 then false
                else
                    match Map.tryFind { Row = addr.Row + r; Col = addr.Col + c } model.Cells with
                    | Some cd -> (match cd.Value with CellValue.Empty -> false | _ -> true)
                    | None -> false
            let collides =
                seq { for r in 0 .. nr - 1 do for c in 0 .. nc - 1 do if occupied r c then yield () }
                |> Seq.isEmpty |> not
            if collides then
                updated <- setCell addr { cell with Value = CellValue.Formula (formula, CellValue.Error Err.spill) } updated
            else
                updated <- setCell addr { cell with Value = CellValue.Formula (formula, arr.[0].[0]) } updated
                for r in 0 .. nr - 1 do
                    for c in 0 .. nc - 1 do
                        if not (r = 0 && c = 0) then
                            updated <-
                                setCell { Row = addr.Row + r; Col = addr.Col + c }
                                    { Value = arr.[r].[c]; Format = cell.Format; Comment = None } updated
    updated
