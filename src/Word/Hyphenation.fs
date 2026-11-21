/// Automatic hyphenation (w:autoHyphenation): where a word may be broken at a line's end, and when
/// Word does (docs/layout-parity-flow.md, fixtures hyphen-words and hyphen-rules).
///
/// Word's hyphenator is its proofing tools' own data, not published. Measured on 8092 words (every
/// probe forcing a decision, Word 2019): Word breaks at the last of its points that fits with the hyphen,
/// at least 2 letters in and 3 from the end (2 in Word 2010's layout, compatibility mode 14, whose
/// points differ too). Liang's TeX patterns (hyph-en-us, HyphenPatternsEnUs.fs) give Word's points for
/// 84% of the words and Word's break for 91% of the probes (mode 14: 78% / 89% with 2 from the end);
/// more points than Word in 4% of the points, fewer in 12%. Word's points are a dictionary's: the Moby
/// hyphenation list's are Word's for 98.6% of the words it has of 2000 more checked in Word (the patterns':
/// 87.7%), so its points for the 50000 commonest words where they differ from the patterns
/// (HyphenWordsEnUs.fs) are installed over the patterns' exceptions: Word's break for 94.1% of the probes
/// (95.6% of the 2000 more's, 93.4% in mode 14). Patterns and words are loaded on use.
module Rbt.Word.Hyphenation

open System.Collections.Generic

/// A document's hyphenation, as Layout uses it.
type Settings = {
    /// The hyphenation zone (px): a word whose start is less than this from the line's end is not
    /// hyphenated (Word 2010's layout and before; 0 in Word 2013's, which has no zone).
    Zone: float
    /// The most lines in a row a paragraph may end with a hyphen (w:consecutiveHyphenLimit; 0: no limit).
    ConsecutiveLimit: int
    /// w:doNotHyphenateCaps.
    DoNotHyphenateCaps: bool
    /// The fewest letters after a break (3; 2 in Word 2010's layout).
    RightMin: int
}

/// The fewest letters before a break.
let leftMin = 2

/// Parsed patterns: each pattern's letters to its digits, and the exceptions' breaks.
type Patterns = {
    Values: Dictionary<string, int[]>
    Exceptions: Dictionary<string, int[]>
    MaxLength: int
}

/// Liang's patterns ("a1b", ".ab4") and exceptions ("as-so-ciate"), space-separated.
let parse (patterns: string) (exceptions: string) : Patterns =
    let values = Dictionary<string, int[]>()
    let mutable maxLength = 0
    for pattern in patterns.Split([| ' '; '\n'; '\r'; '\t' |], System.StringSplitOptions.RemoveEmptyEntries) do
        let letters = System.Text.StringBuilder()
        let digits = ResizeArray<int>()
        digits.Add 0
        for ch in pattern do
            if ch >= '0' && ch <= '9' then digits.[digits.Count - 1] <- int ch - int '0'
            else
                letters.Append ch |> ignore
                digits.Add 0
        let key = letters.ToString()
        values.[key] <- digits.ToArray()
        maxLength <- max maxLength key.Length
    let exceptionBreaks = Dictionary<string, int[]>()
    for word in exceptions.Split([| ' '; '\n'; '\r'; '\t' |], System.StringSplitOptions.RemoveEmptyEntries) do
        let breaks = ResizeArray<int>()
        let letters = System.Text.StringBuilder()
        for ch in word do
            if ch = '-' then breaks.Add letters.Length
            else letters.Append ch |> ignore
        exceptionBreaks.[letters.ToString()] <- breaks.ToArray()
    { Values = values; Exceptions = exceptionBreaks; MaxLength = maxLength }

/// Where the patterns break a lower-case word: the letters before each break, ascending, at least
/// `left` in and `right` from its end.
let pointsOf (patterns: Patterns) (word: string) (left: int) (right: int) : int[] =
    let n = word.Length
    if n < left + right then [||]
    else
        match patterns.Exceptions.TryGetValue word with
        | true, breaks -> breaks |> Array.filter (fun p -> p >= left && n - p >= right)
        | _ ->
            let s = "." + word + "."
            let values = Array.zeroCreate (s.Length + 1)
            for i in 0 .. s.Length - 1 do
                for j in i + 1 .. min s.Length (i + patterns.MaxLength) do
                    match patterns.Values.TryGetValue (s.Substring(i, j - i)) with
                    | true, digits ->
                        for k in 0 .. digits.Length - 1 do
                            if digits.[k] > values.[i + k] then values.[i + k] <- digits.[k]
                    | _ -> ()
            // (values.[p + 1]: between the word's p-th and p+1-th letters.)
            [| for p in left .. n - right do if values.[p + 1] % 2 = 1 then yield p |]

// (.NET: the patterns are compiled in; the browser imports them on use.)
let mutable private loaded : Patterns option =
#if FABLE_COMPILER
    None
#else
    Some (parse HyphenPatternsEnUs.patterns (HyphenPatternsEnUs.exceptions + " " + HyphenWordsEnUs.words))
#endif
let private cache = Dictionary<string, int[]>()

/// Installs the patterns and the words that are exceptions to them, later ones over earlier ones
/// (HyphenPatternsEnUs, HyphenWordsEnUs: imported on use in the browser, see Word.fs).
let register (patterns: string) (exceptions: string) =
    loaded <- Some (parse patterns exceptions)
    cache.Clear()

/// Whether the patterns are there (in the browser: once they are imported).
let isLoaded () = loaded.IsSome

#if FABLE_COMPILER
[<Fable.Core.Emit("Promise.all([import('./HyphenPatternsEnUs.js'), import('./HyphenWordsEnUs.js')]).then(([m, d]) => [m.patterns, m.exceptions + ' ' + d.words])")>]
let private importPatterns () : Fable.Core.JS.Promise<string array> = Fable.Core.Util.jsNative
[<Fable.Core.Emit("$0.then($1)")>]
let private andThen (_promise: Fable.Core.JS.Promise<'T>) (_next: 'T -> 'U) : Fable.Core.JS.Promise<'U> = Fable.Core.Util.jsNative
let mutable private importing : Fable.Core.JS.Promise<bool> option = None

/// Imports and installs the patterns (once): a promise that is true when this call installed them -
/// what was laid out before is laid out again by the caller - and false when they were there already.
let load () : Fable.Core.JS.Promise<bool> =
    if loaded.IsSome then Fable.Core.JS.Constructors.Promise.resolve false
    else
        match importing with
        | Some pending -> andThen pending (fun _ -> false)
        | None ->
            let pending = andThen (importPatterns ()) (fun parts -> register parts.[0] parts.[1]; true)
            importing <- Some pending
            pending
#endif

/// A lower-case word's breaks (at least 2 letters in and 2 from the end; none before the patterns
/// are loaded), remembered per word.
let points (word: string) : int[] =
    match loaded with
    | None -> [||]
    | Some patterns ->
        match cache.TryGetValue word with
        | true, known -> known
        | _ ->
            let found = pointsOf patterns word leftMin 2
            if cache.Count > 20000 then cache.Clear()
            cache.[word] <- found
            found
