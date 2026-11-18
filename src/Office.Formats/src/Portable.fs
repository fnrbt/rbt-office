/// F# that compiles to good code on .NET and on JavaScript (Fable): small
/// helpers for the operations whose obvious F# spelling is fine on .NET and
/// slow in the browser. Each does exactly what the .NET spelling does (the
/// .NET branch *is* that spelling); the JavaScript branch is the native
/// operation. See docs/portable-fsharp.md for when to use which.
module Portable

open System.Collections.Generic
#if FABLE_COMPILER
open Fable.Core.JsInterop
#endif

/// ResizeArray used as a stack. (`xs.RemoveAt (xs.Count - 1)` is
/// `xs.splice(i, 1)` in JavaScript - an array allocated per call, many
/// times slower than `pop()`.)
module Stack =
    /// Remove the last item and return it. The array must not be empty
    /// (.NET throws; JavaScript returns undefined).
    let inline pop (xs: ResizeArray<'T>) : 'T =
#if FABLE_COMPILER
        emitJsExpr xs "$0.pop()"
#else
        let last = xs.Count - 1
        let item = xs.[last]
        xs.RemoveAt last
        item
#endif

    /// Remove the last item. The array must not be empty.
    let inline dropLast (xs: ResizeArray<'T>) : unit =
#if FABLE_COMPILER
        ignore (emitJsExpr xs "$0.pop()" : obj)
#else
        xs.RemoveAt (xs.Count - 1)
#endif

/// Dictionary look-ups without the out-parameter. (`match d.TryGetValue k
/// with | true, v -> ...` allocates a reference cell, two closures and a
/// tuple in JavaScript, and looks the key up twice.)
module Dict =
    /// The value of a key, ValueNone when the dictionary does not have it.
    /// Allocates nothing on either backend (Fable compiles voption as an
    /// erased option).
    let inline tryFind (key: 'K) (dictionary: Dictionary<'K, 'V>) : 'V voption =
#if FABLE_COMPILER
        let d = dictionary
        let k = key
        // A dictionary with primitive keys is a JavaScript Map: one look-up,
        // a second only for a value that is undefined (unit, None). Fable's
        // own Dictionary (structural keys) throws from get: has first.
        let found : 'V = emitJsExpr (d, k) "($0 instanceof Map || $0.has($1)) ? $0.get($1) : undefined"
        if not (isNull (box found)) || d.ContainsKey k then ValueSome found else ValueNone
#else
        let mutable found = Unchecked.defaultof<'V>
        if dictionary.TryGetValue(key, &found) then ValueSome found else ValueNone
#endif

/// Strings.
module Str =
    /// `s.Replace(oldValue, newValue)` (ordinal). In JavaScript Fable's
    /// Replace builds a regular expression from an escaped copy of
    /// `oldValue` on every call (and reads `$&`, `$1` in `newValue` as
    /// patterns); this is indexOf, and split/join when it is there.
    /// `oldValue` must not be empty.
#if FABLE_COMPILER
    let replace (oldValue: string) (newValue: string) (s: string) : string =
        emitJsExpr (s, oldValue, newValue) "$0.indexOf($1) < 0 ? $0 : $0.split($1).join($2)"
#else
    let inline replace (oldValue: string) (newValue: string) (s: string) : string =
        s.Replace(oldValue, newValue)
#endif

/// Numbers as text.
module Num =
    /// `x` with `digits` decimals, invariant: what `sprintf "%.*f"` and
    /// `ToString("F<digits>", InvariantCulture)` give on .NET (`sprintf "%f"`
    /// is six decimals). Fable compiles sprintf to a format-string
    /// interpreter, and ToString with a culture to `String(x)` (all digits).
    let inline toFixed (digits: int) (x: float) : string =
#if FABLE_COMPILER
        emitJsExpr (x, digits) "$0.toFixed($1)"
#else
        x.ToString("F" + string digits, System.Globalization.CultureInfo.InvariantCulture)
#endif
