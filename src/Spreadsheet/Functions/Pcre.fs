/// A regular expression engine of the PCRE2 flavour Excel's REGEXTEST, REGEXEXTRACT and REGEXREPLACE
/// use (PCRE2 10.43+, UTF mode, without UCP unless the pattern starts with (*UCP)), written here so that
/// .NET and the browser give the same answers: System.Text.RegularExpressions and JavaScript's RegExp
/// differ from PCRE2 and from each other. Its own Unicode tables (RegexUnicode.fs) for the same reason.
///
/// A pattern is parsed into a tree, compiled into a program of instructions and run by a backtracking
/// machine with an explicit stack (so a long text does not exhaust the call stack); atomic groups,
/// lookarounds and subroutine calls run their part as a nested run. Positions are code points.
///
/// What it has: literals and every escape of PCRE2 (\x{..}, \o{..}, octal, \cX, \N{U+..}, \Q..\E),
/// classes with ranges, POSIX classes and \d \w \s \h \v \p{..}; . ^ $ \A \z \Z \G \b \B \K \R \X \N;
/// greedy, lazy and possessive quantifiers ({,n} too); capturing, named ((?<n>) (?'n') (?P<n>)),
/// non-capturing, atomic, branch-reset groups; backreferences (\1 \g{-1} \k<n> (?P=n)); lookahead and
/// lookbehind (of bounded length); conditionals ((?(1)..|..), (?(<n>)..), (?(?=..)..), DEFINE);
/// subroutine calls and recursion ((?R) (?1) (?&n) \g<n>, not backtracked into, as in PCRE1); inline
/// options (?imnsxxUJ) and (?^); (*UCP) (*UTF) (*LF) and the start-of-pattern optimisation verbs;
/// (*FAIL) and the alpha assertions ((*pla: ...)). What it rejects (as PCRE2 rejects bad patterns, the
/// function's #VALUE!): the backtracking control verbs other than (*FAIL), newline conventions other
/// than LF, callouts, (?(R) conditions.
module Rbt.Spreadsheet.Functions.Pcre

open System
open System.Collections.Generic
open System.Text

/// The pattern is not one PCRE2 compiles (or one this engine does not do).
exception PatternError of string
/// The match took more steps than PCRE2's match limit allows.
exception MatchLimitExceeded

// ---------------------------------------------------------------------------
// Unicode data
// ---------------------------------------------------------------------------

let private base36 (s: string) : int =
    let mutable n = 0
    for ch in s do
        let d = if ch >= '0' && ch <= '9' then int ch - int '0' else int ch - int 'a' + 10
        n <- n * 36 + d
    n

/// Runs "length.value,..." as their start code points and values.
let private decodeRuns (data: string) : int[] * int[] =
    let parts = data.Split(',')
    let starts = Array.zeroCreate parts.Length
    let values = Array.zeroCreate parts.Length
    let mutable at = 0
    for i in 0 .. parts.Length - 1 do
        let dot = parts.[i].IndexOf '.'
        starts.[i] <- at
        values.[i] <- base36 (parts.[i].Substring(dot + 1))
        at <- at + base36 (parts.[i].Substring(0, dot))
    starts, values

/// The last index of the sorted `starts` whose start is at or below `cp`.
let private runIndex (starts: int[]) (cp: int) =
    let mutable lo = 0
    let mutable hi = starts.Length - 1
    while lo < hi do
        let mid = (lo + hi + 1) / 2
        if starts.[mid] <= cp then lo <- mid else hi <- mid - 1
    lo

let private categoryRuns = lazy (decodeRuns RegexUnicode.categoryData)
let private scriptRuns = lazy (decodeRuns RegexUnicode.scriptData)

/// The general category of a code point, as its number in `RegexUnicode.categoryNames`.
let category (cp: int) : int =
    if cp < 0 || cp > 0x10FFFF then 2
    else
        let starts, values = categoryRuns.Force()
        values.[runIndex starts cp]

let private categoryNumber (name: string) = Array.IndexOf(RegexUnicode.categoryNames, name)
let private catLu = categoryNumber "Lu"
let private catLl = categoryNumber "Ll"
let private catLt = categoryNumber "Lt"
let private catNd = categoryNumber "Nd"
let private catMn = categoryNumber "Mn"
let private catMc = categoryNumber "Mc"
let private catMe = categoryNumber "Me"
let private catPc = categoryNumber "Pc"

let private scriptSets =
    lazy (RegexUnicode.scriptSets.Split(';') |> Array.map (fun set -> set.Split('.') |> Array.map base36))

/// The scripts of a code point (its script extensions), as numbers of `RegexUnicode.scriptNames`.
let scripts (cp: int) : int[] =
    let starts, values = scriptRuns.Force()
    scriptSets.Force().[values.[runIndex starts (max 0 (min cp 0x10FFFF))]]

/// Simple case folding runs: start, count, step, delta.
let private foldRuns =
    lazy (
        let parts = RegexUnicode.foldData.Split(',')
        let starts = Array.zeroCreate parts.Length
        let counts = Array.zeroCreate parts.Length
        let steps = Array.zeroCreate parts.Length
        let deltas = Array.zeroCreate parts.Length
        for i in 0 .. parts.Length - 1 do
            let f = parts.[i].Split('.')
            starts.[i] <- base36 f.[0]
            counts.[i] <- base36 f.[1]
            steps.[i] <- int f.[2]
            deltas.[i] <- if f.[3].StartsWith "-" then -(base36 (f.[3].Substring 1)) else base36 f.[3]
        starts, counts, steps, deltas)

/// The simple case folding of a code point (itself when it has none).
let fold (cp: int) : int =
    let starts, counts, steps, deltas = foldRuns.Force()
    if cp < starts.[0] then cp
    else
        let i = runIndex starts cp
        let offset = cp - starts.[i]
        if offset < counts.[i] * steps.[i] && offset % steps.[i] = 0 then cp + deltas.[i] else cp

/// Every code point that folds to `f` other than `f` itself.
let private foldedFrom =
    lazy (
        let table = Dictionary<int, ResizeArray<int>>()
        let starts, counts, steps, deltas = foldRuns.Force()
        for i in 0 .. starts.Length - 1 do
            for k in 0 .. counts.[i] - 1 do
                let cp = starts.[i] + k * steps.[i]
                let f = cp + deltas.[i]
                match table.TryGetValue f with
                | true, list -> list.Add cp
                | _ -> table.[f] <- ResizeArray [ cp ]
        table)

/// The code points caselessly equal to `cp` (itself among them).
let caseVariants (cp: int) : int list =
    let f = fold cp
    let others =
        match foldedFrom.Force().TryGetValue f with
        | true, list -> List.ofSeq list
        | _ -> []
    (if f = cp then [ cp ] else [ cp; f ]) @ (others |> List.filter (fun x -> x <> cp && x <> f))

/// Upper / lower case of a code point (a single code point; itself when it has none of that case).
let toUpper (cp: int) =
    let c = category cp
    if c <> catLl && c <> catLt then cp
    else caseVariants cp |> List.tryFind (fun x -> category x = catLu) |> Option.defaultValue cp

let toLower (cp: int) =
    let c = category cp
    if c <> catLu && c <> catLt then cp
    else caseVariants cp |> List.tryFind (fun x -> category x = catLl) |> Option.defaultValue cp

/// A text as its code points (a lone surrogate is its own).
let codePoints (s: string) : int[] =
    let out = ResizeArray<int>(s.Length)
    let mutable i = 0
    while i < s.Length do
        let c = int s.[i]
        if c >= 0xD800 && c <= 0xDBFF && i + 1 < s.Length && int s.[i + 1] >= 0xDC00 && int s.[i + 1] <= 0xDFFF then
            out.Add(0x10000 + ((c - 0xD800) <<< 10) + (int s.[i + 1] - 0xDC00))
            i <- i + 2
        else
            out.Add c
            i <- i + 1
    out.ToArray()

let appendCodePoint (sb: StringBuilder) (cp: int) =
    if cp >= 0x10000 then
        let v = cp - 0x10000
        sb.Append(char (0xD800 + (v >>> 10))).Append(char (0xDC00 + (v &&& 0x3FF))) |> ignore
    else sb.Append(char cp) |> ignore

/// The first index from `start` on where `cps` holds `value` (-1 when none).
let private indexFrom (cps: int[]) (value: int) (start: int) =
    let mutable i = start
    while i < cps.Length && cps.[i] <> value do i <- i + 1
    if i < cps.Length then i else -1

let textOf (cps: int[]) (start: int) (finish: int) : string =
    let sb = StringBuilder()
    for i in start .. finish - 1 do appendCodePoint sb cps.[i]
    sb.ToString()

// ---------------------------------------------------------------------------
// Character tests
// ---------------------------------------------------------------------------

let private isAsciiDigit c = c >= 48 && c <= 57
let private isAsciiWord c = (c >= 48 && c <= 57) || (c >= 65 && c <= 90) || (c >= 97 && c <= 122) || c = 95
let private isAsciiSpace c = c = 32 || (c >= 9 && c <= 13)
let private isHSpace c =
    c = 9 || c = 32 || c = 0xA0 || c = 0x1680 || c = 0x180E || (c >= 0x2000 && c <= 0x200A) || c = 0x202F || c = 0x205F || c = 0x3000
let private isVSpace c = (c >= 10 && c <= 13) || c = 0x85 || c = 0x2028 || c = 0x2029
let private isLetterCat k = k = catLu || k = catLl || k = catLt || k = categoryNumber "Lm" || k = categoryNumber "Lo"
let private isNumberCat k = k = catNd || k = categoryNumber "Nl" || k = categoryNumber "No"
let private isMarkCat k = k = catMn || k = catMc || k = catMe
/// \w under UCP: letters, numbers, marks and connector punctuation (PCRE2 10.43).
let private isUnicodeWord c =
    let k = category c
    isLetterCat k || isNumberCat k || isMarkCat k || k = catPc
let private isUnicodeSpace c =
    let k = category c
    isAsciiSpace c || c = 0x85 || k = categoryNumber "Zs" || k = categoryNumber "Zl" || k = categoryNumber "Zp"

/// The \w of a pattern: ASCII, or Unicode under (*UCP).
let isWordChar (ucp: bool) (c: int) = if ucp then isUnicodeWord c else isAsciiWord c

let private posixClass (ucp: bool) (name: string) : (int -> bool) option =
    match name with
    | "alpha" -> Some (if ucp then (fun c -> isLetterCat (category c)) else fun c -> (c >= 65 && c <= 90) || (c >= 97 && c <= 122))
    | "digit" -> Some (if ucp then (fun c -> category c = catNd) else isAsciiDigit)
    | "alnum" -> Some (if ucp then (fun c -> let k = category c in isLetterCat k || k = catNd) else fun c -> isAsciiWord c && c <> 95)
    | "upper" -> Some (if ucp then (fun c -> category c = catLu) else fun c -> c >= 65 && c <= 90)
    | "lower" -> Some (if ucp then (fun c -> category c = catLl) else fun c -> c >= 97 && c <= 122)
    | "space" -> Some (if ucp then isUnicodeSpace else isAsciiSpace)
    | "blank" -> Some (if ucp then isHSpace else fun c -> c = 32 || c = 9)
    | "word" -> Some (isWordChar ucp)
    | "xdigit" -> Some (fun c -> isAsciiDigit c || (c >= 65 && c <= 70) || (c >= 97 && c <= 102))
    | "cntrl" -> Some (fun c -> c < 32 || c = 127)
    | "punct" -> Some (fun c -> (c >= 33 && c <= 47) || (c >= 58 && c <= 64) || (c >= 91 && c <= 96) || (c >= 123 && c <= 126))
    | "graph" -> Some (fun c -> c >= 33 && c <= 126)
    | "print" -> Some (fun c -> c >= 32 && c <= 126)
    | "ascii" -> Some (fun c -> c < 128)
    | _ -> None

/// Loose matching of a property name: case, spaces, hyphens and underscores do not count.
let private looseName (s: string) =
    let sb = StringBuilder()
    for ch in s do
        if ch <> ' ' && ch <> '-' && ch <> '_' then sb.Append(Char.ToLowerInvariant ch) |> ignore
    sb.ToString()

let private categoryLongNames =
    [ "Other", "C"; "Control", "Cc"; "Format", "Cf"; "Unassigned", "Cn"; "Private_Use", "Co"; "Surrogate", "Cs"
      "Letter", "L"; "Lowercase_Letter", "Ll"; "Modifier_Letter", "Lm"; "Other_Letter", "Lo"; "Titlecase_Letter", "Lt"
      "Uppercase_Letter", "Lu"; "Cased_Letter", "L&"; "Mark", "M"; "Combining_Mark", "M"; "Spacing_Mark", "Mc"
      "Enclosing_Mark", "Me"; "Nonspacing_Mark", "Mn"; "Number", "N"; "Decimal_Number", "Nd"; "Digit", "Nd"
      "Letter_Number", "Nl"; "Other_Number", "No"; "Punctuation", "P"; "Connector_Punctuation", "Pc"
      "Dash_Punctuation", "Pd"; "Close_Punctuation", "Pe"; "Final_Punctuation", "Pf"; "Initial_Punctuation", "Pi"
      "Other_Punctuation", "Po"; "Open_Punctuation", "Ps"; "Symbol", "S"; "Currency_Symbol", "Sc"; "Modifier_Symbol", "Sk"
      "Math_Symbol", "Sm"; "Other_Symbol", "So"; "Separator", "Z"; "Line_Separator", "Zl"; "Paragraph_Separator", "Zp"
      "Space_Separator", "Zs"; "LC", "L&"; "Lc", "L&" ]

let private scriptTable =
    lazy (
        let table = Dictionary<string, int>()
        RegexUnicode.scriptNames.Split(';')
        |> Array.iteri (fun i entry -> for name in entry.Split(',') do table.[looseName name] <- i)
        table)

/// A \p{...} property as a test (None when PCRE2 has no such property).
let private property (ucp: bool) (name: string) : (int -> bool) option =
    let byCategories (codes: string) =
        if codes = "L&" then Some (fun c -> let k = category c in k = catLu || k = catLl || k = catLt)
        elif codes.Length = 1 then
            let members = RegexUnicode.categoryNames |> Array.mapi (fun i n -> i, n) |> Array.filter (fun (_, n) -> n.[0] = codes.[0]) |> Array.map fst
            if members.Length = 0 then None
            else
                let mask = members |> Array.fold (fun m i -> m ||| (1 <<< i)) 0
                Some (fun c -> (mask >>> category c) &&& 1 = 1)
        else
            let k = categoryNumber codes
            if k < 0 then None else Some (fun c -> category c = k)
    let n = name.Trim()
    let exact = RegexUnicode.categoryNames |> Array.contains n
    if exact || n = "L&" || (n.Length = 1 && "CLMNPSZ".IndexOf n >= 0) then byCategories n
    else
        let loose = looseName n
        match categoryLongNames |> List.tryFind (fun (long, short) -> looseName long = loose || looseName short = loose) with
        | Some (_, short) -> byCategories short
        | None ->
            match loose with
            | "any" -> Some (fun _ -> true)
            | "xan" -> Some (fun c -> let k = category c in isLetterCat k || isNumberCat k)
            | "xsp" | "xps" -> Some isUnicodeSpace
            | "xwd" -> Some isUnicodeWord
            | "xuc" -> Some (fun c -> c = 0x24 || c = 0x40 || c = 0x60 || (c >= 0xA0 && c <= 0xD7FF) || (c >= 0xE000))
            | _ ->
                let key = if loose.StartsWith "sc:" || loose.StartsWith "script:" then loose.Substring(loose.IndexOf ':' + 1) else loose
                match scriptTable.Force().TryGetValue key with
                | true, s -> Some (fun c -> Array.contains s (scripts c))
                | _ -> None

// ---------------------------------------------------------------------------
// The pattern tree
// ---------------------------------------------------------------------------

type AssertKind =
    | SubjectStart          // \A, ^
    | LineStart             // ^ multiline
    | SubjectEndOrNewline   // $, \Z
    | LineEnd               // $ multiline
    | SubjectEnd            // \z
    | WordBoundary of bool  // \b (ucp)
    | NotWordBoundary of bool
    | SearchStart           // \G

type Greed = Greedy | Lazy | Possessive

type Node =
    | NEmpty
    | NChar of int * bool             // code point, caseless
    | NAny of bool                    // dot-all
    | NSet of (int -> bool)
    | NSeq of Node list
    | NAlt of Node list
    | NGroup of int * Node            // a capture group
    | NRepeat of Node * int * int * Greed   // max -1 unbounded
    | NAssert of AssertKind
    | NBackref of int list * bool
    | NBackrefName of string * bool
    | NLook of bool * bool * Node     // ahead, negated
    | NAtomic of Node
    | NCondGroup of int * Node * Node
    | NCondName of string * Node * Node
    | NCondLook of bool * bool * Node * Node * Node
    | NCall of int
    | NCallName of string
    | NGrapheme
    | NKeep
    | NFail

type private Flags =
    { I: bool; M: bool; S: bool; X: bool; XX: bool; N: bool; U: bool; J: bool }

let private noFlags = { I = false; M = false; S = false; X = false; XX = false; N = false; U = false; J = false }

// ---------------------------------------------------------------------------
// The parser
// ---------------------------------------------------------------------------

type private Parser(pattern: int[], caseless: bool) =
    let n = pattern.Length
    let mutable pos = 0
    let mutable groupCount = 0
    let names = ResizeArray<string * int>()
    let mutable ucp = false
    let mutable maxBackref = 0
    let fail (message: string) = raise (PatternError message)
    let peekAt i = if i < n then pattern.[i] else -1
    let isDigitAt i = i < n && isAsciiDigit pattern.[i]
    let isPatternSpace c = c = 32 || (c >= 9 && c <= 13) || c = 0x85 || c = 0x200E || c = 0x200F || c = 0x2028 || c = 0x2029

    member _.GroupCount = groupCount
    member _.Names = names
    member _.Ucp = ucp
    member _.MaxBackref = maxBackref

    member private _.Lookahead (text: string) =
        let mutable ok = pos + text.Length <= n
        let mutable k = 0
        while ok && k < text.Length do
            if pattern.[pos + k] <> int text.[k] then ok <- false
            k <- k + 1
        ok

    /// The leading (*VERB) settings of a pattern.
    member this.ParseStartVerbs () =
        let mutable going = true
        while going && this.Lookahead "(*" do
            let close = indexFrom pattern (int ')') pos
            if close < 0 then going <- false
            else
                let verb = textOf pattern (pos + 2) close
                let isLetter (c: char) = (c >= 'A' && c <= 'Z') || c = '_'
                if verb.Length > 0 && (verb |> Seq.forall (fun c -> isLetter c || (c >= '0' && c <= '9') || c = '=')) && isLetter verb.[0] then
                    let name = (if verb.IndexOf '=' >= 0 then verb.Substring(0, verb.IndexOf '=') else verb)
                    match name with
                    | "UCP" -> ucp <- true; pos <- close + 1
                    | "UTF" | "UTF16" | "LF" | "NO_AUTO_POSSESS" | "NO_DOTSTAR_ANCHOR" | "NO_JIT" | "NO_START_OPT" | "NOTEMPTY"
                    | "NOTEMPTY_ATSTART" | "LIMIT_MATCH" | "LIMIT_DEPTH" | "LIMIT_HEAP" | "LIMIT_RECURSION" | "BSR_UNICODE" ->
                        pos <- close + 1
                    | "CR" | "CRLF" | "ANYCRLF" | "ANY" | "NUL" | "BSR_ANYCRLF" -> fail ("unsupported (*" + name + ")")
                    | _ -> going <- false
                else going <- false

    member this.SkipExtended (flags: Flags) =
        if flags.X then
            let mutable going = true
            while going && pos < n do
                if isPatternSpace pattern.[pos] then pos <- pos + 1
                elif pattern.[pos] = int '#' then
                    while pos < n && pattern.[pos] <> 10 do pos <- pos + 1
                elif this.Lookahead "(?#" then
                    while pos < n && pattern.[pos] <> int ')' do pos <- pos + 1
                    if pos >= n then fail "missing ) after (?# comment"
                    pos <- pos + 1
                else going <- false
        else
            // (?# comments are skipped in every mode
            while this.Lookahead "(?#" do
                while pos < n && pattern.[pos] <> int ')' do pos <- pos + 1
                if pos >= n then fail "missing ) after (?# comment"
                pos <- pos + 1

    /// A whole pattern.
    member this.Parse () : Node =
        this.ParseStartVerbs ()
        let flags = ref { noFlags with I = caseless }
        let node = this.ParseAlternatives flags false
        if pos < n then fail "unmatched closing parenthesis"
        node

    /// Alternatives up to a ")" (not consumed) or the end. `branchReset`: each alternative numbers its
    /// groups from the same number.
    member this.ParseAlternatives (flags: Flags ref) (branchReset: bool) : Node =
        let first = groupCount
        let mutable highest = groupCount
        let alternatives = ResizeArray<Node>()
        let mutable going = true
        while going do
            if branchReset then groupCount <- first
            alternatives.Add(this.ParseSequence flags)
            highest <- max highest groupCount
            if pos < n && pattern.[pos] = int '|' then pos <- pos + 1
            else going <- false
        if branchReset then groupCount <- highest
        if alternatives.Count = 1 then alternatives.[0] else NAlt (List.ofSeq alternatives)

    member this.ParseSequence (flags: Flags ref) : Node =
        let items = ResizeArray<Node>()
        let mutable going = true
        while going do
            this.SkipExtended flags.Value
            if pos >= n || pattern.[pos] = int '|' || pattern.[pos] = int ')' then going <- false
            else
                match this.ParseAtom flags with
                | None -> ()
                | Some (atom, repeatable) ->
                    let quantified = this.ParseQuantifiers flags.Value atom repeatable
                    items.Add quantified
        if items.Count = 0 then NEmpty elif items.Count = 1 then items.[0] else NSeq (List.ofSeq items)

    /// A quantifier `{n}`, `{n,}`, `{n,m}`, `{,m}` at `pos`, when there is one there.
    member _.TryBraces () : (int * int * int) option =
        if peekAt pos <> int '{' then None
        else
            let mutable i = pos + 1
            let skipSpaces () = while i < n && (pattern.[i] = 32 || pattern.[i] = 9) do i <- i + 1
            let readNumber () =
                let start = i
                let mutable v = 0
                while i < n && isAsciiDigit pattern.[i] do
                    v <- min 1000000 (v * 10 + pattern.[i] - 48)
                    i <- i + 1
                if i = start then None else Some v
            skipSpaces ()
            let low = readNumber ()
            skipSpaces ()
            if i < n && pattern.[i] = int '}' then
                match low with
                | Some v -> Some (v, v, i + 1)
                | None -> None
            elif i < n && pattern.[i] = int ',' then
                i <- i + 1
                skipSpaces ()
                let high = readNumber ()
                skipSpaces ()
                if i < n && pattern.[i] = int '}' then
                    match low, high with
                    | None, None -> None
                    | Some a, None -> Some (a, -1, i + 1)
                    | None, Some b -> Some (0, b, i + 1)
                    | Some a, Some b -> Some (a, b, i + 1)
                else None
            else None

    member this.ParseQuantifiers (flags: Flags) (atom: Node) (repeatable: bool) : Node =
        let mutable node = atom
        let mutable quantified = false
        let mutable going = true
        while going do
            this.SkipExtended flags
            let c = peekAt pos
            let bounds =
                if c = int '*' then pos <- pos + 1; Some (0, -1)
                elif c = int '+' then pos <- pos + 1; Some (1, -1)
                elif c = int '?' then pos <- pos + 1; Some (0, 1)
                else
                    match this.TryBraces () with
                    | Some (a, b, after) -> pos <- after; Some (a, b)
                    | None -> None
            match bounds with
            | None -> going <- false
            | Some (low, high) ->
                if quantified then fail "quantifier does not follow a repeatable item"
                if not repeatable then fail "quantifier does not follow a repeatable item"
                if low > 65535 || high > 65535 then fail "number too big in {} quantifier"
                if high >= 0 && high < low then fail "numbers out of order in {} quantifier"
                let greed =
                    if peekAt pos = int '?' then pos <- pos + 1; (if flags.U then Greedy else Lazy)
                    elif peekAt pos = int '+' then pos <- pos + 1; Possessive
                    else (if flags.U then Lazy else Greedy)
                node <-
                    match node with
                    // A quantified assertion matches at most once.
                    | NLook _ when high <> 0 -> if low = 0 then NRepeat (node, 0, 1, greed) else node
                    | _ -> NRepeat (node, low, high, greed)
                quantified <- true
        node

    /// One item; `None` for something that matches nothing and is no item (an option setting).
    member this.ParseAtom (flags: Flags ref) : (Node * bool) option =
        let f = flags.Value
        let c = pattern.[pos]
        if c = int '(' then this.ParseGroup flags
        elif c = int '[' then
            pos <- pos + 1
            Some (NSet (this.ParseClass f), true)
        elif c = int '.' then
            pos <- pos + 1
            Some (NAny f.S, true)
        elif c = int '^' then
            pos <- pos + 1
            Some (NAssert (if f.M then LineStart else SubjectStart), false)
        elif c = int '$' then
            pos <- pos + 1
            Some (NAssert (if f.M then LineEnd else SubjectEndOrNewline), false)
        elif c = int '\\' then this.ParseEscape f
        elif c = int '*' || c = int '+' || c = int '?' then fail "quantifier does not follow a repeatable item"
        elif c = int '{' && (this.TryBraces ()).IsSome then fail "quantifier does not follow a repeatable item"
        else
            pos <- pos + 1
            Some (NChar (c, f.I), true)

    member _.ReadName (terminator: int) : string =
        let start = pos
        while pos < n && isWordChar true pattern.[pos] do pos <- pos + 1
        let name = textOf pattern start pos
        if name.Length = 0 then fail "group name expected"
        if isAsciiDigit (int name.[0]) then fail "group name must not start with a digit"
        if name.Length > 128 then fail "group name too long"
        if peekAt pos <> terminator then fail "syntax error in group name"
        pos <- pos + 1
        name

    member _.AddName (name: string) (number: int) (flags: Flags) =
        match names |> Seq.tryFind (fun (other, _) -> other = name) with
        | Some (_, existing) when existing <> number && not flags.J -> fail "two named groups have the same name"
        | _ -> names.Add((name, number))

    member this.CaptureGroup (flags: Flags ref) (name: string option) : Node =
        groupCount <- groupCount + 1
        let number = groupCount
        name |> Option.iter (fun nm -> this.AddName nm number flags.Value)
        let inner = this.ParseAlternatives (ref flags.Value) false
        this.Close ()
        NGroup (number, inner)

    member _.Close () =
        if peekAt pos <> int ')' then fail "missing closing parenthesis"
        pos <- pos + 1

    /// The (?...) option letters at `pos` up to ")" or ":".
    member _.ParseOptions (start: Flags) : Flags =
        let mutable f = start
        let mutable on = true
        if peekAt pos = int '^' then
            pos <- pos + 1
            f <- { noFlags with J = start.J }
        let mutable going = true
        while going do
            let c = peekAt pos
            if c = int ')' || c = int ':' then going <- false
            else
                pos <- pos + 1
                match char c with
                | '-' -> on <- false
                | 'i' -> f <- { f with I = on }
                | 'm' -> f <- { f with M = on }
                | 's' -> f <- { f with S = on }
                | 'n' -> f <- { f with N = on }
                | 'U' -> f <- { f with U = on }
                | 'J' -> f <- { f with J = on }
                | 'x' ->
                    if peekAt pos = int 'x' then
                        pos <- pos + 1
                        f <- { f with X = on; XX = on }
                    else f <- { f with X = on; XX = (if on then f.XX else false) }
                | _ -> raise (PatternError "unrecognized character after (? or (?-")
        f

    member this.ParseGroup (flags: Flags ref) : (Node * bool) option =
        pos <- pos + 1
        let f = flags.Value
        let sub () = ref flags.Value
        if peekAt pos = int '*' then
            // (*VERB) and the alpha assertions
            let close = pos + 1
            let mutable i = close
            while i < n && ((pattern.[i] >= int 'a' && pattern.[i] <= int 'z') || (pattern.[i] >= int 'A' && pattern.[i] <= int 'Z') || pattern.[i] = int '_') do i <- i + 1
            let word = textOf pattern close i
            if peekAt i = int ')' && (word = "FAIL" || word = "F") then
                pos <- i + 1
                Some (NFail, false)
            elif peekAt i = int ':' then
                pos <- i + 1
                let assertion ahead negated =
                    let inner = this.ParseAlternatives (sub ()) false
                    this.Close ()
                    Some (NLook (ahead, negated, inner), true)
                match word with
                | "pla" | "positive_lookahead" -> assertion true false
                | "nla" | "negative_lookahead" -> assertion true true
                | "plb" | "positive_lookbehind" -> assertion false false
                | "nlb" | "negative_lookbehind" -> assertion false true
                | "atomic" ->
                    let inner = this.ParseAlternatives (sub ()) false
                    this.Close ()
                    Some (NAtomic inner, true)
                | _ -> fail ("unsupported (*" + word + ":")
            else fail ("unsupported (*" + word + ")")
        elif peekAt pos <> int '?' then
            if f.N then
                let inner = this.ParseAlternatives (sub ()) false
                this.Close ()
                Some (inner, true)
            else Some (this.CaptureGroup flags None, true)
        else
            pos <- pos + 1
            let c = peekAt pos
            let simple (make: Node -> Node) =
                pos <- pos + 1
                let inner = this.ParseAlternatives (sub ()) false
                this.Close ()
                Some (make inner, true)
            if c = int ':' then simple id
            elif c = int '|' then
                pos <- pos + 1
                let inner = this.ParseAlternatives (sub ()) true
                this.Close ()
                Some (inner, true)
            elif c = int '>' then simple NAtomic
            elif c = int '=' then simple (fun x -> NLook (true, false, x))
            elif c = int '!' then simple (fun x -> NLook (true, true, x))
            elif c = int '<' && peekAt (pos + 1) = int '=' then
                pos <- pos + 1
                simple (fun x -> NLook (false, false, x))
            elif c = int '<' && peekAt (pos + 1) = int '!' then
                pos <- pos + 1
                simple (fun x -> NLook (false, true, x))
            elif c = int '<' then
                pos <- pos + 1
                let name = this.ReadName (int '>')
                Some (this.CaptureGroup flags (Some name), true)
            elif c = int '\'' then
                pos <- pos + 1
                let name = this.ReadName (int '\'')
                Some (this.CaptureGroup flags (Some name), true)
            elif c = int 'P' then
                pos <- pos + 1
                let d = peekAt pos
                pos <- pos + 1
                if d = int '<' then
                    let name = this.ReadName (int '>')
                    Some (this.CaptureGroup flags (Some name), true)
                elif d = int '=' then
                    let name = this.ReadName (int ')')
                    Some (NBackrefName (name, f.I), true)
                elif d = int '>' then
                    let name = this.ReadName (int ')')
                    Some (NCallName name, true)
                else fail "unrecognized character after (?P"
            elif c = int '&' then
                pos <- pos + 1
                let name = this.ReadName (int ')')
                Some (NCallName name, true)
            elif c = int 'R' && peekAt (pos + 1) = int ')' then
                pos <- pos + 2
                Some (NCall 0, true)
            elif isAsciiDigit c || ((c = int '+' || c = int '-') && isDigitAt (pos + 1)) then
                let sign = if c = int '+' then 1 elif c = int '-' then -1 else 0
                if sign <> 0 then pos <- pos + 1
                let mutable v = 0
                while isDigitAt pos do
                    v <- min 100000 (v * 10 + pattern.[pos] - 48)
                    pos <- pos + 1
                this.Close ()
                let number =
                    if sign > 0 then groupCount + v
                    elif sign < 0 then groupCount - v + 1
                    else v
                if number < 0 || (sign < 0 && number <= 0) then fail "reference to non-existent subpattern"
                Some (NCall number, true)
            elif c = int '(' then
                pos <- pos + 1
                this.ParseConditional flags
            elif c = int 'C' then fail "callouts are not supported"
            elif c = int '*' then fail "non-atomic assertions are not supported"
            else
                // option setting: (?i) to the end of the group, (?i:...) for its own group
                let changed = this.ParseOptions f
                if peekAt pos = int ')' then
                    pos <- pos + 1
                    flags.Value <- changed
                    None
                elif peekAt pos = int ':' then
                    pos <- pos + 1
                    let inner = this.ParseAlternatives (ref changed) false
                    this.Close ()
                    Some (inner, true)
                else fail "missing closing parenthesis"

    member this.ParseConditional (flags: Flags ref) : (Node * bool) option =
        let f = flags.Value
        let branches () =
            let inner = this.ParseAlternatives (ref flags.Value) false
            this.Close ()
            match inner with
            | NAlt [ yes; no ] -> yes, no
            | NAlt _ -> fail "conditional subpattern contains more than two branches"
            | yes -> yes, NEmpty
        let c = peekAt pos
        if c = int '?' && (peekAt (pos + 1) = int '=' || peekAt (pos + 1) = int '!' || (peekAt (pos + 1) = int '<' && (peekAt (pos + 2) = int '=' || peekAt (pos + 2) = int '!'))) then
            pos <- pos + 1
            let ahead = peekAt pos <> int '<'
            if not ahead then pos <- pos + 1
            let negated = peekAt pos = int '!'
            pos <- pos + 1
            let test = this.ParseAlternatives (ref f) false
            this.Close ()
            let yes, no = branches ()
            Some (NCondLook (ahead, negated, test, yes, no), true)
        elif isAsciiDigit c || ((c = int '+' || c = int '-') && isDigitAt (pos + 1)) then
            let sign = if c = int '+' then 1 elif c = int '-' then -1 else 0
            if sign <> 0 then pos <- pos + 1
            let mutable v = 0
            while isDigitAt pos do
                v <- min 100000 (v * 10 + pattern.[pos] - 48)
                pos <- pos + 1
            this.Close ()
            let number = if sign > 0 then groupCount + v elif sign < 0 then groupCount - v + 1 else v
            if number <= 0 then fail "reference to non-existent subpattern"
            let yes, no = branches ()
            Some (NCondGroup (number, yes, no), true)
        elif c = int '<' || c = int '\'' then
            pos <- pos + 1
            let name = this.ReadName (if c = int '<' then int '>' else int '\'')
            this.Close ()
            let yes, no = branches ()
            Some (NCondName (name, yes, no), true)
        elif this.Lookahead "DEFINE)" then
            pos <- pos + 7
            let inner = this.ParseAlternatives (ref flags.Value) false
            this.Close ()
            (match inner with NAlt _ -> fail "DEFINE subpattern contains more than one branch" | _ -> ())
            // Never matched itself: only its groups are there to be called. (Kept in the tree, unreachable.)
            Some (NCondGroup (-1, NEmpty, inner), true)
        elif c = int 'R' then fail "recursion conditions are not supported"
        else
            // (?(name)...)
            let name = this.ReadName (int ')')
            let yes, no = branches ()
            Some (NCondName (name, yes, no), true)

    /// A hexadecimal / octal number at `pos` of at most `maxDigits` digits.
    member _.ReadDigits (radix: int) (maxDigits: int) : int * int =
        let mutable v = 0
        let mutable count = 0
        let digit c =
            if c >= 48 && c <= 57 then c - 48
            elif c >= 65 && c <= 70 then c - 55
            elif c >= 97 && c <= 102 then c - 87
            else 99
        while count < maxDigits && pos < n && digit pattern.[pos] < radix do
            v <- min 0x7FFFFFF (v * radix + digit pattern.[pos])
            pos <- pos + 1
            count <- count + 1
        v, count

    member this.Braced (radix: int) : int =
        // `{digits}` at pos (pos at the "{")
        pos <- pos + 1
        let v, count = this.ReadDigits radix 1000
        if count = 0 || peekAt pos <> int '}' then fail "invalid escape: digits in braces expected"
        pos <- pos + 1
        v

    member _.CheckCodePoint (v: int) =
        if v > 0x10FFFF then fail "character code point value in \\x{} or \\o{} is too large"
        if v >= 0xD800 && v <= 0xDFFF then fail "disallowed Unicode code point (>= 0xd800 && <= 0xdfff)"
        v

    /// An escape that stands for one character, `pos` after the backslash and at its letter (shared by
    /// classes and the rest; digits are read here as octal). `None` when it is not such an escape.
    member this.CharEscape (inClass: bool) : int option =
        let c = peekAt pos
        let ch = if c >= 0 && c < 128 then char c else '\000'
        match ch with
        | 'a' -> pos <- pos + 1; Some 7
        | 'e' -> pos <- pos + 1; Some 27
        | 'f' -> pos <- pos + 1; Some 12
        | 'n' -> pos <- pos + 1; Some 10
        | 'r' -> pos <- pos + 1; Some 13
        | 't' -> pos <- pos + 1; Some 9
        | 'b' when inClass -> pos <- pos + 1; Some 8
        | 'x' ->
            pos <- pos + 1
            if peekAt pos = int '{' then Some (this.CheckCodePoint (this.Braced 16))
            else
                let v, _ = this.ReadDigits 16 2
                Some v
        | 'o' ->
            pos <- pos + 1
            if peekAt pos <> int '{' then fail "missing opening brace after \\o"
            Some (this.CheckCodePoint (this.Braced 8))
        | 'c' ->
            pos <- pos + 1
            let d = peekAt pos
            if d < 32 || d > 126 then fail "\\c must be followed by a printable ASCII character"
            pos <- pos + 1
            let upper = if d >= 97 && d <= 122 then d - 32 else d
            Some (upper ^^^ 0x40)
        | '0' ->
            let v, _ = this.ReadDigits 8 3
            Some v
        | _ when c >= 49 && c <= 57 ->
            // octal (the caller has ruled out a backreference)
            if c >= 56 then fail "\\8 or \\9 is not a valid escape here"
            let v, _ = this.ReadDigits 8 3
            Some v
        | 'N' when peekAt (pos + 1) = int '{' && peekAt (pos + 2) = int 'U' && peekAt (pos + 3) = int '+' ->
            pos <- pos + 4
            let v, count = this.ReadDigits 16 8
            if count = 0 || peekAt pos <> int '}' then fail "invalid \\N{U+...}"
            pos <- pos + 1
            Some (this.CheckCodePoint v)
        | _ -> None

    /// \d \w \s \h \v \p and their negations as a test; `None` for another escape.
    member this.SetEscape () : (int -> bool) option =
        let c = peekAt pos
        let ch = if c >= 0 && c < 128 then char c else '\000'
        let one (test: int -> bool) = pos <- pos + 1; Some test
        match ch with
        | 'd' -> one (if ucp then (fun x -> category x = catNd) else isAsciiDigit)
        | 'D' -> one (if ucp then (fun x -> category x <> catNd) else fun x -> not (isAsciiDigit x))
        | 'w' -> one (isWordChar ucp)
        | 'W' -> one (fun x -> not (isWordChar ucp x))
        | 's' -> one (if ucp then isUnicodeSpace else isAsciiSpace)
        | 'S' -> one (if ucp then (fun x -> not (isUnicodeSpace x)) else fun x -> not (isAsciiSpace x))
        | 'h' -> one isHSpace
        | 'H' -> one (fun x -> not (isHSpace x))
        | 'v' -> one isVSpace
        | 'V' -> one (fun x -> not (isVSpace x))
        | 'p' | 'P' ->
            pos <- pos + 1
            let mutable negated = ch = 'P'
            let name =
                if peekAt pos = int '{' then
                    let close = indexFrom pattern (int '}') pos
                    if close < 0 then fail "malformed \\P or \\p sequence"
                    let text = textOf pattern (pos + 1) close
                    pos <- close + 1
                    if text.StartsWith "^" then
                        negated <- not negated
                        text.Substring 1
                    else text
                elif pos < n then
                    pos <- pos + 1
                    textOf pattern (pos - 1) pos
                else fail "malformed \\P or \\p sequence"
            match property ucp name with
            | Some test -> Some (if negated then (fun x -> not (test x)) else test)
            | None -> fail ("unknown property name after \\P or \\p: " + name)
        | _ -> None

    /// A character class, `pos` after the "[".
    member this.ParseClass (flags: Flags) : int -> bool =
        let negated = peekAt pos = int '^'
        if negated then pos <- pos + 1
        let ranges = ResizeArray<int * int>()
        let tests = ResizeArray<int -> bool>()
        let mutable first = true
        let mutable going = true
        let mutable quoted = false
        // One element: Choice1Of2 a character, Choice2Of2 a set.
        let element () : Choice<int, int -> bool> option =
            if pos >= n then fail "missing terminating ] for character class"
            let c = pattern.[pos]
            if quoted then
                if c = int '\\' && peekAt (pos + 1) = int 'E' then
                    quoted <- false
                    pos <- pos + 2
                    None
                else
                    pos <- pos + 1
                    Some (Choice1Of2 c)
            elif c = int '[' && (peekAt (pos + 1) = int ':' || peekAt (pos + 1) = int '.' || peekAt (pos + 1) = int '=') then
                let kind = peekAt (pos + 1)
                let mutable close = pos + 2
                while close + 1 < n && not (pattern.[close] = kind && pattern.[close + 1] = int ']') do close <- close + 1
                if close + 1 >= n then
                    pos <- pos + 1
                    Some (Choice1Of2 c)
                else
                    if kind <> int ':' then fail "POSIX collating elements are not supported"
                    let text = textOf pattern (pos + 2) close
                    pos <- close + 2
                    let neg = text.StartsWith "^"
                    let name = if neg then text.Substring 1 else text
                    match posixClass ucp name with
                    | Some test -> Some (Choice2Of2 (if neg then (fun x -> not (test x)) else test))
                    | None -> fail "unknown POSIX class name"
            elif c = int '\\' then
                pos <- pos + 1
                if pos >= n then fail "\\ at end of pattern"
                let d = pattern.[pos]
                if d = int 'Q' then
                    quoted <- true
                    pos <- pos + 1
                    None
                elif d = int 'E' then
                    pos <- pos + 1
                    None
                else
                    match this.SetEscape () with
                    | Some test -> Some (Choice2Of2 test)
                    | None ->
                        match this.CharEscape true with
                        | Some v -> Some (Choice1Of2 v)
                        | None ->
                            if (d >= 65 && d <= 90) || (d >= 97 && d <= 122) then fail "unrecognized escape in character class"
                            pos <- pos + 1
                            Some (Choice1Of2 d)
            elif flags.XX && (c = 32 || c = 9) then
                pos <- pos + 1
                None
            else
                pos <- pos + 1
                Some (Choice1Of2 c)
        while going do
            if pos >= n then fail "missing terminating ] for character class"
            if not quoted && pattern.[pos] = int ']' && not first then
                pos <- pos + 1
                going <- false
            else
                first <- false
                match element () with
                | None -> ()
                | Some (Choice2Of2 test) ->
                    if not quoted && peekAt pos = int '-' && peekAt (pos + 1) <> int ']' && pos + 1 < n then
                        fail "invalid range in character class"
                    tests.Add test
                | Some (Choice1Of2 low) ->
                    if peekAt pos = int '-' && pos + 1 < n && (quoted || peekAt (pos + 1) <> int ']') then
                        let save = pos
                        pos <- pos + 1
                        match element () with
                        | Some (Choice1Of2 high) ->
                            if high < low then fail "range out of order in character class"
                            ranges.Add((low, high))
                        | Some (Choice2Of2 _) -> fail "invalid range in character class"
                        | None ->
                            pos <- save + 1
                            ranges.Add((low, low))
                            ranges.Add((45, 45))
                    else ranges.Add((low, low))
        let rangeArray = ranges.ToArray()
        let testArray = tests.ToArray()
        let plain (x: int) =
            let mutable found = false
            let mutable i = 0
            while not found && i < rangeArray.Length do
                let (lo, hi) = rangeArray.[i]
                if x >= lo && x <= hi then found <- true
                i <- i + 1
            let mutable k = 0
            while not found && k < testArray.Length do
                if testArray.[k] x then found <- true
                k <- k + 1
            found
        let inClass =
            if flags.I then (fun x -> plain x || (caseVariants x |> List.exists (fun v -> v <> x && plain v)))
            else plain
        if negated then (fun x -> not (inClass x)) else inClass

    /// An escape outside a class, `pos` at the backslash.
    member this.ParseEscape (flags: Flags) : (Node * bool) option =
        pos <- pos + 1
        if pos >= n then fail "\\ at end of pattern"
        let c = pattern.[pos]
        let ch = if c < 128 then char c else '\000'
        let ci = flags.I
        let readGroupNumber () =
            let mutable v = 0
            while isDigitAt pos do
                v <- min 100000 (v * 10 + pattern.[pos] - 48)
                pos <- pos + 1
            v
        let backref (number: int) =
            if number <= 0 then fail "a numbered reference must not be zero"
            maxBackref <- max maxBackref number
            Some (NBackref ([ number ], ci), true)
        match ch with
        | 'Q' ->
            pos <- pos + 1
            let items = ResizeArray<Node>()
            let mutable going = true
            while going && pos < n do
                if pattern.[pos] = int '\\' && peekAt (pos + 1) = int 'E' then
                    pos <- pos + 2
                    going <- false
                else
                    items.Add(NChar (pattern.[pos], ci))
                    pos <- pos + 1
            if items.Count = 0 then None
            elif items.Count = 1 then Some (items.[0], true)
            else
                // the quantifier after \Q..\E applies to its last character
                let last = items.[items.Count - 1]
                items.RemoveAt(items.Count - 1)
                let quantified = this.ParseQuantifiers flags last true
                items.Add quantified
                Some (NSeq (List.ofSeq items), false)
        | 'E' -> pos <- pos + 1; None
        | 'A' -> pos <- pos + 1; Some (NAssert SubjectStart, false)
        | 'z' -> pos <- pos + 1; Some (NAssert SubjectEnd, false)
        | 'Z' -> pos <- pos + 1; Some (NAssert SubjectEndOrNewline, false)
        | 'G' -> pos <- pos + 1; Some (NAssert SearchStart, false)
        | 'b' -> pos <- pos + 1; Some (NAssert (WordBoundary ucp), false)
        | 'B' -> pos <- pos + 1; Some (NAssert (NotWordBoundary ucp), false)
        | 'K' -> pos <- pos + 1; Some (NKeep, false)
        | 'X' -> pos <- pos + 1; Some (NGrapheme, true)
        | 'C' -> pos <- pos + 1; Some (NAny true, true)
        | 'R' ->
            pos <- pos + 1
            Some (NAtomic (NAlt [ NSeq [ NChar (13, false); NChar (10, false) ]; NSet isVSpace ]), true)
        | 'N' when peekAt (pos + 1) <> int '{' ->
            pos <- pos + 1
            Some (NAny false, true)
        | 'g' ->
            pos <- pos + 1
            let d = peekAt pos
            if d = int '<' || d = int '\'' then
                // subroutine call \g<name> / \g<n>
                pos <- pos + 1
                let close = if d = int '<' then int '>' else int '\''
                if isDigitAt pos || ((peekAt pos = int '+' || peekAt pos = int '-') && isDigitAt (pos + 1)) then
                    let sign = if peekAt pos = int '+' then 1 elif peekAt pos = int '-' then -1 else 0
                    if sign <> 0 then pos <- pos + 1
                    let v = readGroupNumber ()
                    if peekAt pos <> close then fail "syntax error in subpattern number"
                    pos <- pos + 1
                    let number = if sign > 0 then groupCount + v elif sign < 0 then groupCount - v + 1 else v
                    if number < 0 then fail "reference to non-existent subpattern"
                    Some (NCall number, true)
                else
                    let name = this.ReadName close
                    Some (NCallName name, true)
            elif d = int '{' then
                pos <- pos + 1
                if isDigitAt pos || (peekAt pos = int '-' && isDigitAt (pos + 1)) || (peekAt pos = int '+' && isDigitAt (pos + 1)) then
                    let sign = if peekAt pos = int '-' then -1 elif peekAt pos = int '+' then 1 else 0
                    if sign <> 0 then pos <- pos + 1
                    let v = readGroupNumber ()
                    if peekAt pos <> int '}' then fail "\\g is not followed by a braced, angle-bracketed, or quoted name/number or by a plain number"
                    pos <- pos + 1
                    if sign > 0 then fail "\\g{+n} is not a backreference"
                    backref (if sign < 0 then groupCount - v + 1 else v)
                else
                    let name = this.ReadName (int '}')
                    Some (NBackrefName (name, ci), true)
            elif isDigitAt pos || (d = int '-' && isDigitAt (pos + 1)) then
                let negative = d = int '-'
                if negative then pos <- pos + 1
                let v = readGroupNumber ()
                backref (if negative then groupCount - v + 1 else v)
            else fail "\\g is not followed by a braced, angle-bracketed, or quoted name/number or by a plain number"
        | 'k' ->
            pos <- pos + 1
            let d = peekAt pos
            let close = if d = int '<' then int '>' elif d = int '\'' then int '\'' elif d = int '{' then int '}' else -1
            if close < 0 then fail "\\k is not followed by a braced, angle-bracketed, or quoted name"
            pos <- pos + 1
            let name = this.ReadName close
            Some (NBackrefName (name, ci), true)
        | _ when c >= 49 && c <= 57 ->
            // \1..\9 are backreferences; \10 on when there are that many groups before, else octal
            let start = pos
            let v = readGroupNumber ()
            if v < 10 || v <= groupCount || c >= 56 then backref v
            else
                pos <- start
                match this.CharEscape false with
                | Some x -> Some (NChar (x, ci), true)
                | None -> fail "invalid escape"
        | _ ->
            match this.SetEscape () with
            | Some test -> Some (NSet (if ci then (fun x -> test x || (caseVariants x |> List.exists test)) else test), true)
            | None ->
                match this.CharEscape false with
                | Some x -> Some (NChar (x, ci), true)
                | None ->
                    if (c >= 65 && c <= 90) || (c >= 97 && c <= 122) then
                        fail ("unrecognized character follows \\: " + string ch)
                    pos <- pos + 1
                    Some (NChar (c, ci), true)

// ---------------------------------------------------------------------------
// The program
// ---------------------------------------------------------------------------

type Inst =
    | IChar of int
    | ICharFold of int                 // compares folded code points
    | IAny
    | IAnyButNewline
    | ISet of (int -> bool)
    | ISplit of int * int              // try the first, then the second
    | IJmp of int
    | IOpen of int                     // group number
    | IClose of int
    | IAssert of AssertKind
    | IBackref of int[] * bool
    | IAtomic of int * int             // body, continuation
    | ILook of int * int * bool * bool * int * int * int   // body, continuation, ahead, negated, min, max length, register
    | ICondGroup of int * int          // group (-1 never), no-branch
    | ICondLook of int * int * int * bool * bool * int * int * int  // body, yes, no, ahead, negated, min, max, register
    | ICall of int                     // group number
    | IEndAt of int                    // register
    | IRepStart of int                 // register
    | IRepCheck of int * int           // register, exit
    | IGrapheme
    | IKeep
    | IFail
    | IMatch

/// A compiled pattern.
type Regex =
    { Program: Inst[]
      /// Capture groups (group 0, the match, apart).
      Groups: int
      Names: (string * int) list
      /// Where each group's code (for calls) starts: a run from there ends at an IMatch.
      CallStarts: int[]
      StateSize: int
      /// Matches only at the search start (`^` / `\A` first, outside multiline mode).
      Anchored: bool
      /// The code point every match starts with, when there is one (-1 otherwise).
      FirstChar: int }

let private maxProgram = 400000

/// The shortest and longest text a node matches (-1 for no limit).
let rec private lengths (groups: Dictionary<int, Node>) (depth: int) (node: Node) : int * int =
    let add (a1, b1) (a2, b2) = a1 + a2, (if b1 < 0 || b2 < 0 then -1 else b1 + b2)
    match node with
    | NEmpty | NAssert _ | NLook _ | NKeep -> 0, 0
    | NFail -> 0, 0
    | NChar _ | NAny _ | NSet _ -> 1, 1
    | NSeq items -> items |> List.fold (fun acc x -> add acc (lengths groups depth x)) (0, 0)
    | NAlt items ->
        let all = items |> List.map (lengths groups depth)
        (all |> List.map fst |> List.min), (if all |> List.exists (fun (_, b) -> b < 0) then -1 else all |> List.map snd |> List.max)
    | NGroup (_, inner) | NAtomic inner -> lengths groups depth inner
    | NRepeat (inner, low, high, _) ->
        let a, b = lengths groups depth inner
        a * low, (if high < 0 || b < 0 then (if b = 0 then 0 else -1) else b * high)
    | NCondGroup (g, yes, no) when g < 0 -> lengths groups depth no |> ignore; 0, 0
    | NCondGroup (_, yes, no) | NCondName (_, yes, no) | NCondLook (_, _, _, yes, no) ->
        let a1, b1 = lengths groups depth yes
        let a2, b2 = lengths groups depth no
        min a1 a2, (if b1 < 0 || b2 < 0 then -1 else max b1 b2)
    | NCall g ->
        if depth > 20 then 0, -1
        else
            match groups.TryGetValue g with
            | true, body -> lengths groups (depth + 1) body
            | _ -> 0, -1
    | NBackref _ | NBackrefName _ | NCallName _ | NGrapheme -> 0, -1

let private nullable (groups: Dictionary<int, Node>) (node: Node) = fst (lengths groups 0 node) = 0

/// Compile a pattern (raises PatternError).
let compileRegex (pattern: string) (caseless: bool) : Regex =
    let parser = Parser(codePoints pattern, caseless)
    let parsed = parser.Parse()
    let groupCount = parser.GroupCount
    let names = List.ofSeq parser.Names
    let numbersOf (name: string) =
        match names |> List.filter (fun (n, _) -> n = name) |> List.map snd |> List.distinct with
        | [] -> raise (PatternError ("reference to non-existent subpattern: " + name))
        | numbers -> numbers
    if parser.MaxBackref > groupCount then raise (PatternError "reference to non-existent subpattern")
    // Names resolved to numbers (a name may be used before its group).
    let rec resolve (node: Node) : Node =
        match node with
        | NSeq items -> NSeq (List.map resolve items)
        | NAlt items -> NAlt (List.map resolve items)
        | NGroup (g, inner) -> NGroup (g, resolve inner)
        | NRepeat (inner, a, b, greed) -> NRepeat (resolve inner, a, b, greed)
        | NLook (ahead, neg, inner) -> NLook (ahead, neg, resolve inner)
        | NAtomic inner -> NAtomic (resolve inner)
        | NBackrefName (name, ci) -> NBackref (numbersOf name, ci)
        | NBackref (numbers, _) when numbers |> List.exists (fun g -> g > groupCount) -> raise (PatternError "reference to non-existent subpattern")
        | NCondGroup (g, yes, no) ->
            if g > groupCount then raise (PatternError "reference to non-existent subpattern")
            NCondGroup (g, resolve yes, resolve no)
        | NCondName (name, yes, no) ->
            let numbers =
                match names |> List.filter (fun (n, _) -> n = name) with
                | [] -> raise (PatternError ("reference to non-existent subpattern: " + name))
                | found -> found |> List.map snd
            // (several groups of one name: true when any of them is set - the first number stands for them)
            match numbers with
            | [ g ] -> NCondGroup (g, resolve yes, resolve no)
            | g :: _ -> NCondGroup (g, resolve yes, resolve no)
            | [] -> NEmpty
        | NCondLook (ahead, neg, test, yes, no) -> NCondLook (ahead, neg, resolve test, resolve yes, resolve no)
        | NCallName name ->
            match numbersOf name with
            | g :: _ -> NCall g
            | [] -> NEmpty
        | NCall g when g > groupCount -> raise (PatternError "reference to non-existent subpattern")
        | other -> other
    let tree = resolve parsed
    let groupNodes = Dictionary<int, Node>()
    groupNodes.[0] <- tree
    let rec collect (node: Node) =
        match node with
        | NSeq items | NAlt items -> List.iter collect items
        | NGroup (g, inner) ->
            if not (groupNodes.ContainsKey g) then groupNodes.[g] <- node
            collect inner
        | NRepeat (inner, _, _, _) | NLook (_, _, inner) | NAtomic inner -> collect inner
        | NCondGroup (_, yes, no) | NCondName (_, yes, no) -> collect yes; collect no
        | NCondLook (_, _, test, yes, no) -> collect test; collect yes; collect no
        | _ -> ()
    collect tree
    let code = ResizeArray<Inst>()
    let emit (inst: Inst) =
        code.Add inst
        if code.Count > maxProgram then raise (PatternError "regular expression is too large")
        code.Count - 1
    let mutable registers = 3 * (groupCount + 1)
    let newRegister () =
        registers <- registers + 1
        registers - 1
    let calledGroups = HashSet<int>()
    let rec comp (node: Node) =
        match node with
        | NEmpty -> ()
        | NChar (c, ci) ->
            if ci && (caseVariants c).Length > 1 then emit (ICharFold (fold c)) |> ignore
            else emit (IChar c) |> ignore
        | NAny dotAll -> emit (if dotAll then IAny else IAnyButNewline) |> ignore
        | NSet test -> emit (ISet test) |> ignore
        | NSeq items -> List.iter comp items
        | NAlt items ->
            let jumps = ResizeArray<int>()
            let rec alternatives (rest: Node list) =
                match rest with
                | [] -> ()
                | [ last ] -> comp last
                | item :: more ->
                    let split = emit IFail
                    comp item
                    jumps.Add(emit IFail)
                    code.[split] <- ISplit (split + 1, code.Count)
                    alternatives more
            alternatives items
            for j in jumps do code.[j] <- IJmp code.Count
        | NGroup (g, inner) ->
            emit (IOpen g) |> ignore
            comp inner
            emit (IClose g) |> ignore
        | NRepeat (inner, low, high, Possessive) ->
            let at = emit IFail
            comp (NRepeat (inner, low, high, Greedy))
            emit IMatch |> ignore
            code.[at] <- IAtomic (at + 1, code.Count)
        | NRepeat (inner, low, high, greed) ->
            for _ in 1 .. low do comp inner
            if high < 0 then
                let guard = nullable groupNodes inner
                let loop = emit IFail
                let register = if guard then newRegister () else -1
                if guard then emit (IRepStart register) |> ignore
                comp inner
                let check = if guard then emit IFail else -1
                emit (IJmp loop) |> ignore
                let exit = code.Count
                code.[loop] <- (if greed = Greedy then ISplit (loop + 1, exit) else ISplit (exit, loop + 1))
                if guard then code.[check] <- IRepCheck (register, exit)
            elif high > low then
                let splits = ResizeArray<int>()
                for _ in low + 1 .. high do
                    splits.Add(emit IFail)
                    comp inner
                let exit = code.Count
                for s in splits do code.[s] <- (if greed = Greedy then ISplit (s + 1, exit) else ISplit (exit, s + 1))
        | NAssert kind -> emit (IAssert kind) |> ignore
        | NBackref (numbers, ci) -> emit (IBackref (List.toArray numbers, ci)) |> ignore
        | NLook (ahead, negated, inner) ->
            let at = emit IFail
            let low, high = if ahead then 0, 0 else lengths groupNodes 0 inner
            if not ahead && high < 0 then raise (PatternError "lookbehind assertion is not fixed length")
            if not ahead && high > 255 && high <> low then raise (PatternError "length of lookbehind assertion is not limited")
            let register = if ahead then -1 else newRegister ()
            comp inner
            if not ahead then emit (IEndAt register) |> ignore
            emit IMatch |> ignore
            code.[at] <- ILook (at + 1, code.Count, ahead, negated, low, high, register)
        | NAtomic inner ->
            let at = emit IFail
            comp inner
            emit IMatch |> ignore
            code.[at] <- IAtomic (at + 1, code.Count)
        | NCondGroup (g, yes, no) ->
            let at = emit IFail
            comp yes
            let jump = emit IFail
            let noStart = code.Count
            if g >= 0 then comp no
            code.[at] <- ICondGroup (g, noStart)
            code.[jump] <- IJmp code.Count
        | NCondLook (ahead, negated, test, yes, no) ->
            let at = emit IFail
            let low, high = if ahead then 0, 0 else lengths groupNodes 0 test
            if not ahead && high < 0 then raise (PatternError "lookbehind assertion is not fixed length")
            let register = if ahead then -1 else newRegister ()
            comp test
            if not ahead then emit (IEndAt register) |> ignore
            emit IMatch |> ignore
            let yesStart = code.Count
            comp yes
            let jump = emit IFail
            let noStart = code.Count
            comp no
            code.[jump] <- IJmp code.Count
            code.[at] <- ICondLook (at + 1, yesStart, noStart, ahead, negated, low, high, register)
        | NCall g ->
            calledGroups.Add g |> ignore
            emit (ICall g) |> ignore
        | NGrapheme -> emit IGrapheme |> ignore
        | NKeep -> emit IKeep |> ignore
        | NFail -> emit IFail |> ignore
        | NBackrefName _ | NCondName _ | NCallName _ -> raise (PatternError "unresolved name")
    emit (IOpen 0) |> ignore
    comp tree
    emit (IClose 0) |> ignore
    emit IMatch |> ignore
    // Each called group compiled once more on its own, ending in IMatch (calls of calls added as met).
    let callStarts = Array.create (groupCount + 1) -1
    let mutable pending = List.ofSeq calledGroups
    while not pending.IsEmpty do
        let g = pending.Head
        pending <- pending.Tail
        if callStarts.[g] < 0 then
            callStarts.[g] <- code.Count
            let before = HashSet<int>(calledGroups)
            comp (if g = 0 then tree else groupNodes.[g])
            emit IMatch |> ignore
            for c in calledGroups do
                if not (before.Contains c) then pending <- c :: pending
    for g in calledGroups do
        if callStarts.[g] < 0 then raise (PatternError "reference to non-existent subpattern")
    let rec startsAnchored (node: Node) =
        match node with
        | NAssert SubjectStart -> true
        | NSeq (first :: _) -> startsAnchored first
        | NGroup (_, inner) | NAtomic inner -> startsAnchored inner
        | NAlt items -> items |> List.forall startsAnchored
        | _ -> false
    let rec firstChar (node: Node) =
        match node with
        | NChar (c, false) -> c
        | NSeq (first :: _) -> firstChar first
        | NGroup (_, inner) | NAtomic inner -> firstChar inner
        | NRepeat (inner, low, _, _) when low > 0 -> firstChar inner
        | _ -> -1
    { Program = code.ToArray(); Groups = groupCount; Names = names; CallStarts = callStarts
      StateSize = registers; Anchored = startsAnchored tree; FirstChar = firstChar tree }

// ---------------------------------------------------------------------------
// The machine
// ---------------------------------------------------------------------------

/// PCRE2's default match limit.
let matchLimit = 10000000

let private isGraphemeExtend (c: int) =
    let k = category c
    k = catMn || k = catMe || k = catMc || c = 0x200D || (c >= 0xFE00 && c <= 0xFE0F) || (c >= 0x1F3FB && c <= 0x1F3FF) || (c >= 0xE0020 && c <= 0xE007F)

type private Machine(re: Regex, text: int[]) =
    let prog = re.Program
    let n = text.Length
    let state : int[] = Array.create re.StateSize -1
    let mutable steps = 0
    let mutable searchStart = 0
    let mutable notEmptyAt = -1
    let mutable depth = 0
    let tempOf g = 2 * (re.Groups + 1) + g

    member _.State = state
    member _.Steps = steps

    member private _.WordAt (ucp: bool) (i: int) = i >= 0 && i < n && isWordChar ucp text.[i]

    member private this.AssertHolds (kind: AssertKind) (pos: int) =
        match kind with
        | SubjectStart -> pos = 0
        | LineStart -> pos = 0 || (text.[pos - 1] = 10 && pos < n)
        | SubjectEndOrNewline -> pos = n || (pos = n - 1 && text.[pos] = 10)
        | LineEnd -> pos = n || text.[pos] = 10
        | SubjectEnd -> pos = n
        | WordBoundary ucp -> this.WordAt ucp (pos - 1) <> this.WordAt ucp pos
        | NotWordBoundary ucp -> this.WordAt ucp (pos - 1) = this.WordAt ucp pos
        | SearchStart -> pos = searchStart

    /// Run from `pc0` at `pos0` to an IMatch: the end position, or -1. Leaves the state as the match
    /// made it (a failed run leaves it as it found it).
    member this.Run (pc0: int) (pos0: int) (top: bool) : int =
        let stack = ResizeArray<int>()
        let mutable pc = pc0
        let mutable pos = pos0
        let mutable result = -2
        let mutable back = false
        let set (i: int) (v: int) =
            stack.Add(-i - 1)
            stack.Add(state.[i])
            state.[i] <- v
        while result = -2 do
            if back then
                back <- false
                let mutable going = true
                while going do
                    let k = stack.Count
                    if k = 0 then
                        result <- -1
                        going <- false
                    else
                        let b = stack.[k - 1]
                        let a = stack.[k - 2]
                        stack.RemoveAt(k - 1)
                        stack.RemoveAt(k - 2)
                        if a >= 0 then
                            pc <- a
                            pos <- b
                            going <- false
                        else state.[-a - 1] <- b
            else
                match prog.[pc] with
                | IChar c ->
                    if pos < n && text.[pos] = c then
                        pos <- pos + 1
                        pc <- pc + 1
                    else back <- true
                | ICharFold c ->
                    if pos < n && fold text.[pos] = c then
                        pos <- pos + 1
                        pc <- pc + 1
                    else back <- true
                | IAny ->
                    if pos < n then
                        pos <- pos + 1
                        pc <- pc + 1
                    else back <- true
                | IAnyButNewline ->
                    if pos < n && text.[pos] <> 10 then
                        pos <- pos + 1
                        pc <- pc + 1
                    else back <- true
                | ISet test ->
                    if pos < n && test text.[pos] then
                        pos <- pos + 1
                        pc <- pc + 1
                    else back <- true
                | ISplit (first, second) ->
                    steps <- steps + 1
                    if steps > matchLimit then raise MatchLimitExceeded
                    stack.Add second
                    stack.Add pos
                    pc <- first
                | IJmp target -> pc <- target
                | IOpen g ->
                    set (tempOf g) pos
                    pc <- pc + 1
                | IClose g ->
                    set (2 * g) state.[tempOf g]
                    set (2 * g + 1) pos
                    pc <- pc + 1
                | IAssert kind ->
                    if this.AssertHolds kind pos then pc <- pc + 1 else back <- true
                | IBackref (groups, ci) ->
                    let mutable chosen = -1
                    for g in groups do
                        if chosen < 0 && state.[2 * g + 1] >= 0 then chosen <- g
                    if chosen < 0 then back <- true
                    else
                        let s = state.[2 * chosen]
                        let e = state.[2 * chosen + 1]
                        let len = e - s
                        if pos + len > n then back <- true
                        else
                            let mutable ok = true
                            let mutable i = 0
                            while ok && i < len do
                                let a = text.[s + i]
                                let b = text.[pos + i]
                                if a <> b && (not ci || fold a <> fold b) then ok <- false
                                i <- i + 1
                            if ok then
                                pos <- pos + len
                                pc <- pc + 1
                            else back <- true
                | IAtomic (body, next) ->
                    let snapshot = Array.copy state
                    let e = this.Sub body pos
                    if e < 0 then back <- true
                    else
                        for i in 0 .. state.Length - 1 do
                            if state.[i] <> snapshot.[i] then
                                stack.Add(-i - 1)
                                stack.Add(snapshot.[i])
                        pos <- e
                        pc <- next
                | ILook (body, next, ahead, negated, low, high, register) ->
                    let snapshot = Array.copy state
                    let found = this.Look body pos ahead low high register
                    if found = negated then
                        Array.blit snapshot 0 state 0 state.Length
                        back <- true
                    else
                        if negated then Array.blit snapshot 0 state 0 state.Length
                        else
                            for i in 0 .. state.Length - 1 do
                                if state.[i] <> snapshot.[i] then
                                    stack.Add(-i - 1)
                                    stack.Add(snapshot.[i])
                        pc <- next
                | ICondGroup (g, noBranch) ->
                    if g >= 0 && state.[2 * g + 1] >= 0 then pc <- pc + 1 else pc <- noBranch
                | ICondLook (body, yes, no, ahead, negated, low, high, register) ->
                    let snapshot = Array.copy state
                    let found = this.Look body pos ahead low high register
                    if negated || not found then Array.blit snapshot 0 state 0 state.Length
                    else
                        for i in 0 .. state.Length - 1 do
                            if state.[i] <> snapshot.[i] then
                                stack.Add(-i - 1)
                                stack.Add(snapshot.[i])
                    pc <- (if found <> negated then yes else no)
                | ICall g ->
                    let snapshot = Array.copy state
                    depth <- depth + 1
                    if depth > 250 then raise MatchLimitExceeded
                    let e = this.Sub re.CallStarts.[g] pos
                    depth <- depth - 1
                    Array.blit snapshot 0 state 0 state.Length
                    if e < 0 then back <- true
                    else
                        pos <- e
                        pc <- pc + 1
                | IEndAt register ->
                    if pos = state.[register] then pc <- pc + 1 else back <- true
                | IRepStart register ->
                    set register pos
                    pc <- pc + 1
                | IRepCheck (register, exit) ->
                    if pos = state.[register] then pc <- exit else pc <- pc + 1
                | IGrapheme ->
                    if pos >= n then back <- true
                    else
                        let c = text.[pos]
                        pos <- pos + 1
                        if c = 13 && pos < n && text.[pos] = 10 then pos <- pos + 1
                        elif c >= 0x1F1E6 && c <= 0x1F1FF && pos < n && text.[pos] >= 0x1F1E6 && text.[pos] <= 0x1F1FF then pos <- pos + 1
                        elif c <> 10 && c <> 13 then
                            let mutable going = true
                            while going && pos < n do
                                if isGraphemeExtend text.[pos] then
                                    if text.[pos] = 0x200D && pos + 1 < n then pos <- pos + 2 else pos <- pos + 1
                                else going <- false
                        pc <- pc + 1
                | IKeep ->
                    set (tempOf 0) pos
                    pc <- pc + 1
                | IFail -> back <- true
                | IMatch ->
                    if top && notEmptyAt >= 0 && pos = notEmptyAt && state.[0] = pos then back <- true
                    else result <- pos
        result

    member this.Sub (body: int) (pos: int) : int = this.Run body pos false

    /// Does the lookaround's body match at `pos` (behind: ending at `pos`)?
    member this.Look (body: int) (pos: int) (ahead: bool) (low: int) (high: int) (register: int) : bool =
        if ahead then this.Sub body pos >= 0
        else
            let saved = state.[register]
            state.[register] <- pos
            let mutable found = false
            let mutable len = high
            while not found && len >= low do
                if pos - len >= 0 && this.Sub body (pos - len) >= 0 then found <- true
                len <- len - 1
            state.[register] <- saved
            found

    /// The first match at or after `start` (`anchored`: only at `start`; `notEmpty`: not an empty one at
    /// `start`): the capture slots (2 per group, group 0 the match; -1 unset), or None.
    member this.Search (start: int) (anchored: bool) (notEmpty: bool) : int[] option =
        searchStart <- start
        notEmptyAt <- if notEmpty then start else -1
        let groupsLength = 2 * (re.Groups + 1)
        let mutable found = None
        let mutable at = start
        let last = if anchored || re.Anchored then start else n
        while found.IsNone && at <= last do
            if re.FirstChar >= 0 && not anchored then
                while at < n && text.[at] <> re.FirstChar do at <- at + 1
                if at >= n then at <- n + 1
            if at <= last then
                for i in 0 .. state.Length - 1 do state.[i] <- -1
                let e = this.Run 0 at true
                if e >= 0 then found <- Some (Array.sub state 0 groupsLength)
                at <- at + 1
        found

/// The first match of `re` in `text` from `start` on.
let search (re: Regex) (text: int[]) (start: int) : int[] option =
    Machine(re, text).Search start false false

/// Every match, as PCRE2's global matching finds them: after an empty match the next one is tried at
/// the same place, not empty and anchored, and when there is none one code point further on.
let allMatches (re: Regex) (text: int[]) : int[] list =
    let machine = Machine(re, text)
    let found = ResizeArray<int[]>()
    let mutable start = 0
    let mutable afterEmpty = false
    let mutable going = true
    while going do
        match machine.Search start afterEmpty afterEmpty with
        | Some caps ->
            found.Add caps
            afterEmpty <- caps.[1] = caps.[0]
            start <- max caps.[1] start
        | None ->
            if afterEmpty && start < text.Length then
                afterEmpty <- false
                start <- start + 1
            else going <- false
    List.ofSeq found

/// The number of a named group (the first of that name), when the pattern has one.
let groupNumber (re: Regex) (name: string) : int option =
    re.Names |> List.tryFind (fun (n, _) -> n = name) |> Option.map snd

/// All the numbers of a named group (several with (?J) or (?|)).
let groupNumbers (re: Regex) (name: string) : int list =
    re.Names |> List.filter (fun (n, _) -> n = name) |> List.map snd
