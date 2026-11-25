/// The regular expression functions of Excel 365: REGEXTEST, REGEXEXTRACT and REGEXREPLACE, on the
/// PCRE2-flavoured engine of Pcre.fs (the same on .NET and in the browser). A pattern PCRE2 does not
/// compile is #VALUE!, as is a match that runs past PCRE2's match limit.
module Rbt.Spreadsheet.Functions.FnRegex

open System
open System.Collections.Generic
open System.Text
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Functions.Kit
open Rbt.Spreadsheet.Functions.Pcre

let moduleName = "Regex"

let private compiled = Dictionary<string, Result<Regex, string>>()

/// The pattern compiled (remembered: a formula filled down a column asks for the same one each row).
let private regexOf (pattern: string) (caseless: bool) : Result<Regex, string> =
    let key = (if caseless then "i" else "c") + pattern
    match compiled.TryGetValue key with
    | true, known -> known
    | _ ->
        let made =
            try Ok (compileRegex pattern caseless)
            with PatternError _ -> Result.Error Err.value
        if compiled.Count > 1000 then compiled.Clear()
        compiled.[key] <- made
        made

/// Run a match, a pattern that takes too long to match being #VALUE!.
let private guarded (f: unit -> 'a) : Result<'a, string> =
    try Ok (f ())
    with MatchLimitExceeded -> Result.Error Err.value

/// A whole-number option of a choice of values (case_sensitivity 0/1, return_mode 0/1/2).
let private choice (v: CellValue) (fallback: float) (allowed: float list) : Result<float, string> =
    res {
        let! x =
            match v with
            | CellValue.Empty -> Ok fallback
            | v -> coerceNumber v
        let x = Math.Truncate x
        do! check (List.contains x allowed) Err.value
        return x }

let private caseless (v: CellValue) : Result<bool, string> =
    res {
        let! x = choice v 0.0 [ 0.0; 1.0 ]
        return x = 1.0 }

let private groupText (text: int[]) (caps: int[]) (g: int) =
    if caps.[2 * g] < 0 || caps.[2 * g + 1] < 0 then "" else textOf text caps.[2 * g] caps.[2 * g + 1]

// ---------------------------------------------------------------------------
// The replacement string, as pcre2_substitute reads it with PCRE2_SUBSTITUTE_EXTENDED (Excel takes
// \U$1 and the like): $$, $n, ${n}, $name, ${name}, ${n:-default}, ${n:+set:unset}; \U \L \u \l \E
// case forcing; \Q..\E; the character escapes of patterns. An unset group inserts nothing; a group the
// pattern does not have, or a bad escape, is an error (#VALUE!).
// ---------------------------------------------------------------------------

exception ReplacementError

type private Expander(re: Regex, text: int[], caps: int[], repl: int[]) =
    let n = repl.Length
    let out = StringBuilder()
    let mutable pos = 0
    let mutable force = 0      // 1 upper, -1 lower, until \E
    let mutable single = 0     // 1 upper, -1 lower, the next character
    let peekAt i = if i < n then repl.[i] else -1
    let isDigit c = c >= 48 && c <= 57

    member _.Text = out.ToString()

    member _.Emit (cp: int) (writing: bool) =
        if writing then
            let cased =
                if single <> 0 then
                    let c = if single > 0 then toUpper cp else toLower cp
                    single <- 0
                    c
                elif force > 0 then toUpper cp
                elif force < 0 then toLower cp
                else cp
            appendCodePoint out cased

    member this.Group (g: int) (writing: bool) =
        if g < 0 || g > re.Groups then raise ReplacementError
        if writing && caps.[2 * g] >= 0 then
            for i in caps.[2 * g] .. caps.[2 * g + 1] - 1 do this.Emit text.[i] true

    /// A group reference's number (`pos` after the "$" or "${").
    member _.Reference () : int =
        if isDigit (peekAt pos) then
            let mutable v = 0
            while isDigit (peekAt pos) do
                v <- min 100000 (v * 10 + repl.[pos] - 48)
                pos <- pos + 1
            v
        else
            let start = pos
            while pos < n && isWordChar true repl.[pos] do pos <- pos + 1
            if pos = start then raise ReplacementError
            match groupNumbers re (textOf repl start pos) with
            | [] -> raise ReplacementError
            | numbers ->
                // the first of the name that is set
                match numbers |> List.tryFind (fun g -> caps.[2 * g] >= 0) with
                | Some g -> g
                | None -> numbers.Head

    member _.IsSet (g: int) = g >= 0 && g <= re.Groups && caps.[2 * g] >= 0

    /// Expand up to one of `stops` (not consumed) at this level, or the end.
    member this.Expand (stops: int list) (writing: bool) =
        let mutable going = true
        while going && pos < n do
            let c = repl.[pos]
            if List.contains c stops then going <- false
            elif c = int '$' then
                pos <- pos + 1
                let d = peekAt pos
                if d = int '$' then
                    pos <- pos + 1
                    this.Emit 36 writing
                elif d = int '{' then
                    pos <- pos + 1
                    let g = this.Reference ()
                    if peekAt pos = int '}' then
                        pos <- pos + 1
                        this.Group g writing
                    elif peekAt pos = int ':' && peekAt (pos + 1) = int '-' then
                        pos <- pos + 2
                        if this.IsSet g then
                            this.Group g writing
                            this.Expand [ int '}' ] false
                        else this.Expand [ int '}' ] writing
                        if peekAt pos <> int '}' then raise ReplacementError
                        pos <- pos + 1
                    elif peekAt pos = int ':' && peekAt (pos + 1) = int '+' then
                        pos <- pos + 2
                        let set = this.IsSet g
                        if g > re.Groups then raise ReplacementError
                        this.Expand [ int ':'; int '}' ] (writing && set)
                        if peekAt pos = int ':' then
                            pos <- pos + 1
                            this.Expand [ int '}' ] (writing && not set)
                        if peekAt pos <> int '}' then raise ReplacementError
                        pos <- pos + 1
                    else raise ReplacementError
                elif d >= 0 && (isDigit d || isWordChar true d) then
                    let g = this.Reference ()
                    this.Group g writing
                else raise ReplacementError
            elif c = int '\\' then
                pos <- pos + 1
                let d = peekAt pos
                if d < 0 then raise ReplacementError
                pos <- pos + 1
                let ch = if d < 128 then char d else '\000'
                let hexOrOctal (radix: int) (maxDigits: int) =
                    let digit x =
                        if x >= 48 && x <= 57 then x - 48
                        elif x >= 65 && x <= 70 then x - 55
                        elif x >= 97 && x <= 102 then x - 87
                        else 99
                    let mutable v = 0
                    let mutable count = 0
                    while count < maxDigits && pos < n && digit repl.[pos] < radix do
                        v <- v * radix + digit repl.[pos]
                        pos <- pos + 1
                        count <- count + 1
                    v, count
                let braced (radix: int) =
                    if peekAt pos <> int '{' then raise ReplacementError
                    pos <- pos + 1
                    let v, count = hexOrOctal radix 8
                    if count = 0 || peekAt pos <> int '}' || v > 0x10FFFF || (v >= 0xD800 && v <= 0xDFFF) then raise ReplacementError
                    pos <- pos + 1
                    v
                match ch with
                | 'U' -> force <- 1
                | 'L' -> force <- -1
                | 'u' -> single <- 1
                | 'l' -> single <- -1
                | 'E' -> force <- 0; single <- 0
                | 'Q' ->
                    let mutable quoting = true
                    while quoting && pos < n do
                        if repl.[pos] = int '\\' && peekAt (pos + 1) = int 'E' then
                            pos <- pos + 2
                            quoting <- false
                        else
                            this.Emit repl.[pos] writing
                            pos <- pos + 1
                | 'a' -> this.Emit 7 writing
                | 'b' -> this.Emit 8 writing
                | 'e' -> this.Emit 27 writing
                | 'f' -> this.Emit 12 writing
                | 'n' -> this.Emit 10 writing
                | 'r' -> this.Emit 13 writing
                | 't' -> this.Emit 9 writing
                | 'v' -> this.Emit 11 writing
                | 'x' ->
                    if peekAt pos = int '{' then this.Emit (braced 16) writing
                    else
                        let v, _ = hexOrOctal 16 2
                        this.Emit v writing
                | 'o' -> this.Emit (braced 8) writing
                | '0' ->
                    let v, _ = hexOrOctal 8 2
                    this.Emit v writing
                | 'c' ->
                    let x = peekAt pos
                    if x < 32 || x > 126 then raise ReplacementError
                    pos <- pos + 1
                    this.Emit ((if x >= 97 && x <= 122 then x - 32 else x) ^^^ 0x40) writing
                | _ ->
                    // a backreference (\1) or an unknown letter is an error; other characters stand for themselves
                    if (d >= 48 && d <= 57) || (d >= 65 && d <= 90) || (d >= 97 && d <= 122) then raise ReplacementError
                    this.Emit d writing
            else
                this.Emit c writing
                pos <- pos + 1

/// The replacement of one match.
let private expand (re: Regex) (text: int[]) (caps: int[]) (repl: int[]) : string =
    let e = Expander(re, text, caps, repl)
    e.Expand [] true
    e.Text

// ---------------------------------------------------------------------------
// The functions
// ---------------------------------------------------------------------------

let private regexTest (text: string) (pattern: string) (ci: bool) : Result<CellValue, string> =
    res {
        let! re = regexOf pattern ci
        let! found = guarded (fun () -> search re (codePoints text) 0)
        return CellValue.Boolean found.IsSome }

/// REGEXEXTRACT of one text: a row of the results (one for return mode 0).
let private regexExtract (text: string) (pattern: string) (mode: float) (ci: bool) : Result<CellValue[], string> =
    res {
        let! re = regexOf pattern ci
        let cps = codePoints text
        if mode = 1.0 then
            let! all = guarded (fun () -> allMatches re cps)
            if all.IsEmpty then return! fail Err.na
            else return all |> List.map (fun caps -> CellValue.String (groupText cps caps 0)) |> Array.ofList
        else
            let! found = guarded (fun () -> search re cps 0)
            match found with
            | None -> return! fail Err.na
            | Some caps ->
                if mode = 0.0 then return [| CellValue.String (groupText cps caps 0) |]
                elif re.Groups = 0 then return! fail Err.na
                else return [| for g in 1 .. re.Groups -> CellValue.String (groupText cps caps g) |] }

let private replaceMatches (re: Regex) (cps: int[]) (repl: int[]) (all: int[][]) (occurrence: int) (original: string) : Result<CellValue, string> =
    let chosen =
        if occurrence = 0 then -1
        elif occurrence > 0 then occurrence - 1
        else all.Length + occurrence
    if occurrence <> 0 && (chosen < 0 || chosen >= all.Length) then Ok (CellValue.String original)
    else
        let sb = StringBuilder()
        let mutable last = 0
        let mutable failed = false
        for i in 0 .. all.Length - 1 do
            if not failed && (occurrence = 0 || i = chosen) then
                let caps = all.[i]
                sb.Append(textOf cps last caps.[0]) |> ignore
                try sb.Append(expand re cps caps repl) |> ignore
                with ReplacementError -> failed <- true
                last <- max last caps.[1]
        if failed then Result.Error Err.value
        else
            sb.Append(textOf cps last cps.Length) |> ignore
            Ok (CellValue.String (sb.ToString()))

let private regexReplace (text: string) (pattern: string) (replacement: string) (occurrence: int) (ci: bool) : Result<CellValue, string> =
    res {
        let! re = regexOf pattern ci
        let cps = codePoints text
        let! all = guarded (fun () -> allMatches re cps |> Array.ofList)
        return! replaceMatches re cps (codePoints replacement) all occurrence text }

/// The value at (r, c) of an argument broadcast to the result's size: a single row or column repeats,
/// a cell past an argument's own size is #N/A (as Excel lifts a function over arrays).
let private broadcastAt (arr: CellValue[][]) (r: int) (c: int) : CellValue =
    let rows = arr.Length
    let cols = if rows > 0 then arr.[0].Length else 0
    let rr = if rows = 1 then 0 else r
    let cc = if cols = 1 then 0 else c
    if rr >= rows || cc >= cols then CellValue.Error Err.na else unwrapValue arr.[rr].[cc]

let functions : FunctionDef list =
    [ scalar "REGEXTEST" 2 3 "|S" (fun a -> res {
        let! text = a.Text 0
        let! pattern = a.Text 1
        let! ci = caseless (a.Value 2)
        return! regexTest text pattern ci })
      // Lifted by hand: return modes 1 and 2 give a row for one text; over an array of texts each
      // cell is the first of its row (as Excel's lifted array results are).
      arrayed "REGEXEXTRACT" 2 4 "" (fun a ->
        let args = [| for i in 0 .. 3 -> a.Array i |]
        let single = args |> Array.forall (fun arr -> arr.Length = 1 && arr.[0].Length = 1)
        let one (r: int) (c: int) : Result<CellValue[], string> =
            res {
                let at i = broadcastAt args.[i] r c
                let! text = coerceText (at 0)
                let! pattern = coerceText (at 1)
                let! mode = choice (if a.Supplied 2 then at 2 else CellValue.Empty) 0.0 [ 0.0; 1.0; 2.0 ]
                let! ci = caseless (if a.Supplied 3 then at 3 else CellValue.Empty)
                return! regexExtract text pattern mode ci }
        if single then
            match one 0 0 with
            | Ok row -> Ok [| row |]
            | Result.Error e -> Ok [| [| err e |] |]
        else
            let rows = args |> Array.map (fun arr -> arr.Length) |> Array.max
            let cols = args |> Array.map (fun arr -> if arr.Length > 0 then arr.[0].Length else 0) |> Array.max
            Ok [| for r in 0 .. rows - 1 ->
                    [| for c in 0 .. cols - 1 ->
                        match one r c with
                        | Ok row -> row.[0]
                        | Result.Error e -> err e |] |])
      scalar "REGEXREPLACE" 3 5 "|S" (fun a -> res {
        let! text = a.Text 0
        let! pattern = a.Text 1
        let! replacement = a.Text 2
        let! occurrence = a.NumberOr 3 0.0
        let! ci = caseless (a.Value 4)
        return! regexReplace text pattern replacement (int (Math.Truncate occurrence)) ci }) ]

let install () = registerFunctions moduleName functions
