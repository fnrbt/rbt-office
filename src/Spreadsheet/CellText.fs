module RedBat.Spreadsheet.CellText

open System.Collections.Generic
open RedBat.Spreadsheet.Model

let private resizeArrayToList (items: List<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
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

let private addStringCacheCharsCapped maxTotalChars amount total =
    if amount <= 0 then total
    elif total > maxTotalChars || amount > maxTotalChars - total then maxTotalChars + 1
    else total + amount

let private addBoundedStringCacheEntryWithValueChars
    maxEntries
    maxTotalChars
    (valueTextCharCount: string -> 'V -> int)
    (totalChars: byref<int>)
    (cache: Dictionary<string, 'V>)
    (order: Queue<string>)
    (key: string)
    value =
    let entryCharCount (key: string) (value: 'V) =
        0
        |> addStringCacheCharsCapped maxTotalChars key.Length
        |> addStringCacheCharsCapped maxTotalChars (valueTextCharCount key value)

    let currentEntryChars = entryCharCount key value
    if currentEntryChars > maxTotalChars then
        let mutable previous = Unchecked.defaultof<'V>
        if cache.TryGetValue(key, &previous) then
            cache.Remove key |> ignore
            totalChars <- max 0 (totalChars - entryCharCount key previous)
        value
    else
        let mutable previous = Unchecked.defaultof<'V>
        if cache.TryGetValue(key, &previous) then
            totalChars <- max 0 (totalChars - entryCharCount key previous)
        else
            order.Enqueue key
        cache.[key] <- value
        totalChars <- totalChars + currentEntryChars
        while (cache.Count > maxEntries || totalChars > maxTotalChars) && order.Count > 0 do
            let oldest = order.Dequeue()
            if oldest <> key then
                let mutable oldestValue = Unchecked.defaultof<'V>
                if cache.TryGetValue(oldest, &oldestValue) then
                    cache.Remove oldest |> ignore
                    totalChars <- max 0 (totalChars - entryCharCount oldest oldestValue)
        if cache.Count > maxEntries || totalChars > maxTotalChars then
            let current = cache.[key]
            cache.Clear()
            order.Clear()
            order.Enqueue key
            cache.[key] <- current
            totalChars <- currentEntryChars
        value

let private addBoundedStringCacheEntry
    maxEntries
    maxTotalChars
    (totalChars: byref<int>)
    (cache: Dictionary<string, 'V>)
    (order: Queue<string>)
    (key: string)
    value =
    addBoundedStringCacheEntryWithValueChars
        maxEntries
        maxTotalChars
        (fun _ _ -> 0)
        &totalChars
        cache
        order
        key
        value

/// Interpret common Excel date/time format codes (m/d/yyyy, h:mm, etc.) and
/// render the Excel serial number as a date string. Excel serial 1 ==
/// 1900-01-01; the 1900 leap-year bug means dates after 1900-02-28 need an
/// extra day subtracted, but for typical post-1901 fixtures the simple
/// `baseDate.AddDays(serial - 2)` covers it.
/// Excel format codes use `;` to separate positive/negative/zero/text
/// sections (in that order). For date/time interpretation we only need the
/// positive section, but real numeric values need the right section picked
/// by sign so a format like `"TRUE";"TRUE";"FALSE"` renders 0 as FALSE.
/// Splitter respects quoted segments and backslash escapes.
let private splitSections (code: string) : string list =
    if code.IndexOf(';') < 0 then
        [ code ]
    else
        let sections = System.Collections.Generic.List<string>()
        let sb = System.Text.StringBuilder()
        let mutable i = 0
        let mutable inQuote = false
        while i < code.Length do
            let c = code.[i]
            if c = '\\' && i + 1 < code.Length then
                sb.Append(c).Append(code.[i + 1]) |> ignore
                i <- i + 2
            elif c = '"' then
                inQuote <- not inQuote
                sb.Append(c) |> ignore
                i <- i + 1
            elif c = ';' && not inQuote then
                sections.Add(sb.ToString())
                sb.Clear() |> ignore
                i <- i + 1
            else
                sb.Append(c) |> ignore
                i <- i + 1
        sections.Add(sb.ToString())
        resizeArrayToList sections

let private splitSectionsCache = Dictionary<string, string list>()
let private splitSectionsCacheOrder = Queue<string>()
let private splitSectionsCacheMaxEntries = 4096
let private splitSectionsCacheMaxTotalChars = 5000000
let private splitSectionsCacheMaxTotalSections = 200000
let mutable private splitSectionsCacheTotalChars = 0
let mutable private splitSectionsCacheTotalSections = 0

let private splitSectionsEntryMetrics (code: string) (sections: string list) =
    let mutable textChars = addStringCacheCharsCapped splitSectionsCacheMaxTotalChars code.Length 0
    let mutable sectionCount = 0
    let mutable remaining = sections
    while textChars <= splitSectionsCacheMaxTotalChars
          && sectionCount <= splitSectionsCacheMaxTotalSections
          && not remaining.IsEmpty do
        let section = remaining.Head
        if not (System.Object.ReferenceEquals(section, code)) then
            textChars <- addStringCacheCharsCapped splitSectionsCacheMaxTotalChars section.Length textChars
        sectionCount <- sectionCount + 1
        remaining <- remaining.Tail
    struct (textChars, sectionCount)

let private splitSectionsRemoveCacheEntry (code: string) =
    let mutable previous = Unchecked.defaultof<string list>
    if splitSectionsCache.TryGetValue(code, &previous) then
        splitSectionsCache.Remove code |> ignore
        let struct (entryChars, entrySections) = splitSectionsEntryMetrics code previous
        splitSectionsCacheTotalChars <- max 0 (splitSectionsCacheTotalChars - entryChars)
        splitSectionsCacheTotalSections <- max 0 (splitSectionsCacheTotalSections - entrySections)

let private addSplitSectionsCacheEntry (code: string) (sections: string list) =
    let struct (entryChars, entrySections) = splitSectionsEntryMetrics code sections
    if entryChars > splitSectionsCacheMaxTotalChars
       || entrySections > splitSectionsCacheMaxTotalSections then
        splitSectionsRemoveCacheEntry code
        sections
    else
        let mutable previous = Unchecked.defaultof<string list>
        if splitSectionsCache.TryGetValue(code, &previous) then
            let struct (previousChars, previousSections) = splitSectionsEntryMetrics code previous
            splitSectionsCacheTotalChars <- max 0 (splitSectionsCacheTotalChars - previousChars)
            splitSectionsCacheTotalSections <- max 0 (splitSectionsCacheTotalSections - previousSections)
        else
            splitSectionsCacheOrder.Enqueue code
        splitSectionsCache.[code] <- sections
        splitSectionsCacheTotalChars <- splitSectionsCacheTotalChars + entryChars
        splitSectionsCacheTotalSections <- splitSectionsCacheTotalSections + entrySections
        while (splitSectionsCache.Count > splitSectionsCacheMaxEntries
               || splitSectionsCacheTotalChars > splitSectionsCacheMaxTotalChars
               || splitSectionsCacheTotalSections > splitSectionsCacheMaxTotalSections)
              && splitSectionsCacheOrder.Count > 0 do
            let oldest = splitSectionsCacheOrder.Dequeue()
            if oldest <> code then
                splitSectionsRemoveCacheEntry oldest
        if splitSectionsCache.Count > splitSectionsCacheMaxEntries
           || splitSectionsCacheTotalChars > splitSectionsCacheMaxTotalChars
           || splitSectionsCacheTotalSections > splitSectionsCacheMaxTotalSections then
            let current = splitSectionsCache.[code]
            splitSectionsCache.Clear()
            splitSectionsCacheOrder.Clear()
            splitSectionsCacheOrder.Enqueue code
            splitSectionsCache.[code] <- current
            splitSectionsCacheTotalChars <- entryChars
            splitSectionsCacheTotalSections <- entrySections
        sections

let private splitSectionsCached (code: string) =
    let mutable sections = Unchecked.defaultof<string list>
    if splitSectionsCache.TryGetValue(code, &sections) then
        sections
    else
        sections <- splitSections code
        addSplitSectionsCacheEntry code sections

let private firstSection (code: string) =
    if code.IndexOf(';') < 0 then code
    else
        match splitSectionsCached code with
        | first :: _ -> first
        | [] -> code

/// A leading comparison tag like `[>=100]` or `[<0]` on a format section.
/// May sit before or after a colour tag (`[<0][Red]`, `[Red][>100]`). The
/// `[op N]` shape doesn't collide with `[h]`, `[$…]`, or `[Red]`, so matching
/// anywhere in the section is safe.
let private sectionCondition (section: string) : (string * float) option =
    let parseNumber start finish =
        let mutable i = start
        let mutable sign = 1.0
        if i <= finish && section.[i] = '-' then
            sign <- -1.0
            i <- i + 1
        let mutable whole = 0.0
        let mutable fraction = 0.0
        let mutable scale = 1.0
        let mutable afterDot = false
        let mutable sawDigit = false
        let mutable valid = i <= finish
        while valid && i <= finish do
            let ch = section.[i]
            if ch >= '0' && ch <= '9' then
                sawDigit <- true
                let digit = float (int ch - int '0')
                if afterDot then
                    scale <- scale * 10.0
                    fraction <- fraction + digit / scale
                else
                    whole <- whole * 10.0 + digit
                i <- i + 1
            elif ch = '.' && not afterDot && sawDigit && i < finish then
                afterDot <- true
                i <- i + 1
            else
                valid <- false
        if valid && sawDigit then Some (sign * (whole + fraction)) else None

    let parseBracket start finish =
        if start > finish then None
        else
            let op, numberStart =
                if start + 1 <= finish then
                    match section.[start], section.[start + 1] with
                    | '<', '=' -> Some "<=", start + 2
                    | '>', '=' -> Some ">=", start + 2
                    | '<', '>' -> Some "<>", start + 2
                    | '<', _ -> Some "<", start + 1
                    | '>', _ -> Some ">", start + 1
                    | '=', _ -> Some "=", start + 1
                    | _ -> None, start
                else
                    match section.[start] with
                    | '<' -> Some "<", start + 1
                    | '>' -> Some ">", start + 1
                    | '=' -> Some "=", start + 1
                    | _ -> None, start
            match op with
            | Some op ->
                match parseNumber numberStart finish with
                | Some threshold -> Some (op, threshold)
                | None -> None
            | None -> None

    let mutable i = 0
    let mutable result: (string * float) option = None
    while result.IsNone && i < section.Length do
        if section.[i] = '[' then
            let endBracket = section.IndexOf(']', i + 1)
            if endBracket < 0 then
                i <- section.Length
            else
                result <- parseBracket (i + 1) (endBracket - 1)
                i <- endBracket + 1
        else
            i <- i + 1
    result

let private sectionConditionCache = Dictionary<string, (string * float) option>()
let private sectionConditionCacheOrder = Queue<string>()
let private sectionConditionCacheMaxEntries = 4096
let private sectionConditionCacheMaxTotalChars = 5000000
let mutable private sectionConditionCacheTotalChars = 0

let private sectionConditionCached section =
    let mutable condition = Unchecked.defaultof<(string * float) option>
    if sectionConditionCache.TryGetValue(section, &condition) then
        condition
    else
        condition <- sectionCondition section
        addBoundedStringCacheEntry sectionConditionCacheMaxEntries sectionConditionCacheMaxTotalChars &sectionConditionCacheTotalChars sectionConditionCache sectionConditionCacheOrder section condition

let private conditionMatches (op: string) (threshold: float) (n: float) : bool =
    match op with
    | ">" -> n > threshold
    | ">=" -> n >= threshold
    | "<" -> n < threshold
    | "<=" -> n <= threshold
    | "=" -> n = threshold
    | "<>" -> n <> threshold
    | _ -> false

let private signSectionForValue n fallback secondSection thirdSection =
    if n > 0.0 then
        fallback
    elif n < 0.0 then
        match secondSection with
        | ValueSome section -> section
        | ValueNone -> fallback
    else
        match thirdSection with
        | ValueSome section -> section
        | ValueNone -> fallback

let private sectionSelectionBySign (sections: string list) (n: float) : struct (string * bool * int) =
    let fallback =
        match sections with
        | first :: _ -> first
        | [] -> ""
    let mutable sectionSlots = 0
    let mutable secondSection = ValueNone
    let mutable thirdSection = ValueNone
    let mutable remaining = sections
    while sectionSlots < 3 && not remaining.IsEmpty do
        match remaining with
        | section :: tail ->
            sectionSlots <- sectionSlots + 1
            if sectionSlots = 2 then secondSection <- ValueSome section
            elif sectionSlots = 3 then thirdSection <- ValueSome section
            remaining <- tail
        | [] -> ()
    struct (signSectionForValue n fallback secondSection thirdSection, false, sectionSlots)

let private sectionSelectionForValue (code: string) (sections: string list) (n: float) : struct (string * bool * int) =
    if code.IndexOf('[') < 0 then
        sectionSelectionBySign sections n
    else
        let fallback =
            match sections with
            | first :: _ -> first
            | [] -> ""

        let mutable hasCondition = false
        let mutable matched = ValueNone
        let mutable lastUnconditioned = ValueNone
        let mutable lastSection = fallback
        let mutable sectionSlots = 0
        let mutable secondSection = ValueNone
        let mutable thirdSection = ValueNone
        let mutable remaining = sections
        while not remaining.IsEmpty do
            match remaining with
            | section :: tail ->
                if sectionSlots < 3 then
                    sectionSlots <- sectionSlots + 1
                    if sectionSlots = 2 then secondSection <- ValueSome section
                    elif sectionSlots = 3 then thirdSection <- ValueSome section
                lastSection <- section
                match sectionConditionCached section with
                | Some (op, th) ->
                    hasCondition <- true
                    if matched.IsNone && conditionMatches op th n then
                        matched <- ValueSome section
                | None ->
                    lastUnconditioned <- ValueSome section
                remaining <- tail
            | [] -> ()

        if hasCondition then
            // Conditional format: pick the first section whose `[op N]` condition is
            // true; if none match, the last UNconditioned section is the default.
            match matched, lastUnconditioned with
            | ValueSome section, _ -> struct (section, true, sectionSlots)
            | ValueNone, ValueSome section -> struct (section, true, sectionSlots)
            | ValueNone, ValueNone -> struct (lastSection, true, sectionSlots)
        else
            // Excel sign grammar: [pos], [pos;neg], [pos;neg;zero], [pos;neg;zero;text].
            // With < 3 sections, zero falls back to the positive section — NOT the
            // negative one. With < 2 sections, negative also falls back to positive.
            struct (signSectionForValue n fallback secondSection thirdSection, false, sectionSlots)

let private selectedSectionForValue (code: string) (n: float) : string =
    if code.IndexOf(';') < 0 then code
    else
        let struct (section, _, _) = sectionSelectionForValue code (splitSectionsCached code) n
        section

let private equalsIgnoreCase (left: string) (right: string) =
    if left.Length <> right.Length then
        false
    else
        let mutable index = 0
        let mutable equal = true
        while equal && index < left.Length do
            let l = left.[index]
            let r = right.[index]
            if l <> r && System.Char.ToUpperInvariant l <> System.Char.ToUpperInvariant r then
                equal <- false
            index <- index + 1
        equal

let private sliceEqualsIgnoreCase (text: string) start length (expected: string) =
    if length <> expected.Length then
        false
    else
        let mutable index = 0
        let mutable equal = true
        while equal && index < length do
            let l = text.[start + index]
            let r = expected.[index]
            if l <> r && System.Char.ToUpperInvariant l <> System.Char.ToUpperInvariant r then
                equal <- false
            index <- index + 1
        equal

let private trimmedStartsWithIgnoreCase (text: string) (prefix: string) =
    let mutable startIndex = 0
    let mutable endIndex = text.Length - 1
    while startIndex <= endIndex && System.Char.IsWhiteSpace text.[startIndex] do
        startIndex <- startIndex + 1
    while endIndex >= startIndex && System.Char.IsWhiteSpace text.[endIndex] do
        endIndex <- endIndex - 1
    let trimmedLength = endIndex - startIndex + 1
    trimmedLength >= prefix.Length && sliceEqualsIgnoreCase text startIndex prefix.Length prefix

let private colorTagValueAt (text: string) start length =
    if sliceEqualsIgnoreCase text start length "red" then Some "#FF0000"
    elif sliceEqualsIgnoreCase text start length "black" then Some "#000000"
    elif sliceEqualsIgnoreCase text start length "green" then Some "#008000"
    elif sliceEqualsIgnoreCase text start length "white" then Some "#FFFFFF"
    elif sliceEqualsIgnoreCase text start length "blue" then Some "#0000FF"
    elif sliceEqualsIgnoreCase text start length "magenta" then Some "#FF00FF"
    elif sliceEqualsIgnoreCase text start length "yellow" then Some "#FFFF00"
    elif sliceEqualsIgnoreCase text start length "cyan" then Some "#00FFFF"
    else None

let private isConditionTagAt (text: string) (start: int) (length: int) =
    let stop = start + length
    let mutable i = start
    let has (prefix: string) =
        i + prefix.Length <= stop && sliceEqualsIgnoreCase text i prefix.Length prefix
    if has "<=" || has ">=" || has "<>" then
        i <- i + 2
    elif i < stop && (text.[i] = '=' || text.[i] = '<' || text.[i] = '>') then
        i <- i + 1
    else
        i <- stop + 1
    if i < stop && text.[i] = '-' then
        i <- i + 1
    let digitStart = i
    while i < stop && System.Char.IsDigit text.[i] do
        i <- i + 1
    let hasIntegerDigits = i > digitStart
    let mutable hasValidFraction = true
    if i < stop && text.[i] = '.' then
        i <- i + 1
        let fractionStart = i
        while i < stop && System.Char.IsDigit text.[i] do
            i <- i + 1
        hasValidFraction <- i > fractionStart
    i = stop && hasIntegerDigits && hasValidFraction

let private removeColorAndConditionTags (text: string) =
    if text.IndexOf('[') < 0 then
        text
    else
        let sb = System.Text.StringBuilder(text.Length)
        let mutable i = 0
        while i < text.Length do
            if text.[i] = '[' then
                let close = text.IndexOf(']', i + 1)
                if close > i then
                    let start = i + 1
                    let length = close - start
                    if (colorTagValueAt text start length).IsSome || isConditionTagAt text start length then
                        i <- close + 1
                    else
                        sb.Append(text.[i]) |> ignore
                        i <- i + 1
                else
                    sb.Append(text.[i]) |> ignore
                    i <- i + 1
            else
                sb.Append(text.[i]) |> ignore
                i <- i + 1
        sb.ToString()

let private unwrapTextPlaceholderLiterals (text: string) =
    let sb = System.Text.StringBuilder(text.Length)
    let mutable i = 0
    while i < text.Length do
        match text.[i] with
        | '"' ->
            let mutable j = i + 1
            while j < text.Length && text.[j] <> '"' do
                sb.Append(text.[j]) |> ignore
                j <- j + 1
            i <- if j < text.Length then j + 1 else j
        | '\\' when i + 1 < text.Length ->
            sb.Append(text.[i + 1]) |> ignore
            i <- i + 2
        | '[' ->
            let close = text.IndexOf(']', i + 1)
            if close > i && (colorTagValueAt text (i + 1) (close - i - 1)).IsSome then
                i <- close + 1
            else
                sb.Append(text.[i]) |> ignore
                i <- i + 1
        | ch ->
            sb.Append(ch) |> ignore
            i <- i + 1
    sb.ToString()

/// Map an Excel format-section colour tag like `[Red]` to a hex string.
/// Returns None when the tag is absent or unrecognised. Excel only honors
/// the eight named colours plus `[Color N]` (palette index 1..56); we map
/// the eight named cases — enough for the common `#,##0.00;[Red]-...`
/// idiom that the parity pipeline cares about.
let private numberFormatSectionColor (section: string) : string option =
    // The colour tag may not be the FIRST bracket — a section can lead with a
    // condition (`[<0][Red]0.0`) or have the colour first (`[Red][>100]0`), so
    // match the colour tag anywhere among the leading tags, not just at ^.
    let mutable i = 0
    let mutable result: string option = None
    while result.IsNone && i < section.Length do
        if section.[i] = '[' then
            let endBracket = section.IndexOf(']', i + 1)
            if endBracket < 0 then
                i <- section.Length
            else
                result <- colorTagValueAt section (i + 1) (endBracket - i - 1)
                i <- endBracket + 1
        else
            i <- i + 1
    result

/// Pick the colour tag for the section of `code` that applies to value `n`.
/// Default (no tag) → None.
let numberFormatColorForValue (code: string) (n: float) : string option =
    if code.IndexOf('[') < 0 then None
    else numberFormatSectionColor (selectedSectionForValue code n)

let private appendLower (sb: System.Text.StringBuilder) (ch: char) =
    sb.Append(System.Char.ToLowerInvariant ch) |> ignore

/// Strip everything that could carry a stray date/time letter without being a
/// real date/time token: quoted literals (`"High"`→h, `"Red"`→d), backslash
/// escapes (`\d`), and ALL bracket tags — locale `[$…]`, colour `[Red]`,
/// condition `[>=100]`, AND elapsed `[h]`/`[m]`/`[s]` (those route via
/// `hasElapsedBracket`, which is checked separately in the dispatch). What
/// remains is the bare format token stream we sniff for date/time chars.
let private dateTimeSniff (code: string) =
    let s = firstSection code
    let sb = System.Text.StringBuilder(s.Length)
    let mutable i = 0
    while i < s.Length do
        match s.[i] with
        | '"' ->
            let endQuote = s.IndexOf('"', i + 1)
            if endQuote >= 0 then
                i <- endQuote + 1
            else
                appendLower sb s.[i]
                i <- i + 1
        | '\\' when i + 1 < s.Length ->
            i <- i + 2
        | '[' ->
            let endBracket = s.IndexOf(']', i + 1)
            if endBracket >= 0 then
                i <- endBracket + 1
            else
                appendLower sb s.[i]
                i <- i + 1
        | ch ->
            appendLower sb ch
            i <- i + 1
    sb.ToString()

let private hasDateTimeSniffCandidate (code: string) =
    let section = firstSection code
    let mutable i = 0
    let mutable found = false
    while not found && i < section.Length do
        match section.[i] with
        | 'y' | 'Y' | 'd' | 'D' | 'm' | 'M' | 'h' | 'H' | 's' | 'S' -> found <- true
        | _ -> i <- i + 1
    found

let private containsIgnoringLocaleTags (needle: char) (s: string) =
    let expected = System.Char.ToUpperInvariant needle
    let mutable i = 0
    let mutable found = false
    while not found && i < s.Length do
        if s.[i] = '[' && i + 1 < s.Length && s.[i + 1] = '$' then
            let endBracket = s.IndexOf(']', i + 2)
            if endBracket >= 0 then i <- endBracket + 1
            else i <- i + 1
        else
            if System.Char.ToUpperInvariant s.[i] = expected then
                found <- true
            i <- i + 1
    found

let private isDateOrTimeFormat (code: string) =
    if not (hasDateTimeSniffCandidate code) then false
    else
        let lower = dateTimeSniff code
        lower.Contains "y"
        || lower.Contains "d"
        || lower.Contains "m/"
        || lower.Contains "mmm"
        || lower.Contains "h"
        || (lower.Contains "s" && lower.Contains "m")

let private excelSerialToDate (serial: float) : System.DateTime =
    // Excel's 1900 calendar: serial 1 is 1900-01-01 but Excel also treats
    // 1900 as a leap year (it isn't). For serials >= 60 we subtract one to
    // line up with the real Gregorian calendar.
    let baseDate = System.DateTime(1899, 12, 30)
    let days = if serial >= 60.0 then serial else serial + 1.0
    baseDate.AddDays(days)

/// Tokenize an Excel format code into .NET DateTime.ToString format. Tracks
/// context (after `h`/`s` → minute; after month-only marker → month) for the
/// shared `m` token. Recognizes fractional-second suffixes `.0`/`.00`/`.000`
/// (Excel) and emits the .NET equivalents `.f`/`.ff`/`.fff`. Quoted/escaped
/// literals are emitted unchanged. Unrecognized tokens are wrapped in single
/// quotes so .NET treats them as literals rather than throwing.
let private excelToNetFormat (code: string) : string =
    // Date/time tokens are matched case-insensitively, but quoted/escaped
    // literals must keep their original case — Excel's `"FY"yy` renders
    // "FY18", not "fy18".
    let orig = firstSection code
    let sb = System.Text.StringBuilder()
    let mutable i = 0
    // mInTimeContext is true after we just emitted h/H or are about to emit
    // s/S — Excel's `m` token resolves to minute in those positions.
    let mutable mInTimeContext = false
    let charLower index = System.Char.ToLowerInvariant orig.[index]
    let matchesAt (index: int) (expected: string) =
        if index + expected.Length > orig.Length then
            false
        else
            let mutable offset = 0
            let mutable matches = true
            while matches && offset < expected.Length do
                if charLower (index + offset) <> expected.[offset] then
                    matches <- false
                offset <- offset + 1
            matches
    let containsToken (expected: string) =
        let mutable index = 0
        let mutable found = false
        while not found && index <= orig.Length - expected.Length do
            if matchesAt index expected then
                found <- true
            index <- index + 1
        found
    let has12Hour = containsToken "am/pm" || containsToken "a/p"

    let runOf c =
        let mutable n = 0
        while i + n < orig.Length && charLower (i + n) = c do n <- n + 1
        n

    while i < orig.Length do
        match charLower i with
        | 'y' ->
            let n = runOf 'y'
            sb.Append(if n >= 4 then "yyyy" elif n >= 2 then "yy" else "yyyy") |> ignore
            i <- i + n
            mInTimeContext <- false
        | 'd' ->
            let n = runOf 'd'
            sb.Append(System.String('d', min 4 n)) |> ignore
            i <- i + n
            mInTimeContext <- false
        | 'h' ->
            let n = runOf 'h'
            let token = if has12Hour then (if n >= 2 then "hh" else "h") else (if n >= 2 then "HH" else "H")
            sb.Append(token: string) |> ignore
            i <- i + n
            mInTimeContext <- true
        | 's' ->
            let n = runOf 's'
            sb.Append(System.String('s', min 2 n)) |> ignore
            i <- i + n
            mInTimeContext <- true
        | 'm' ->
            let n = runOf 'm'
            // 3+ m's are always month (mmm/mmmm). 1-2 m's: minute if adjacent
            // to a time token, otherwise month.
            let isMinute =
                n <= 2 && (
                    mInTimeContext ||
                    (let j = i + n
                     j < orig.Length && (orig.[j] = ':' || charLower j = 'h' || charLower j = 's')))
            if isMinute then
                sb.Append(System.String('m', min 2 n)) |> ignore
                mInTimeContext <- true
            else
                let upper =
                    match n with
                    | 1 -> "M"
                    | 2 -> "MM"
                    | 3 -> "MMM"
                    | 4 -> "MMMM"
                    | _ -> "MMMMM"
                sb.Append(upper) |> ignore
                mInTimeContext <- false
            i <- i + n
        | 'a' when matchesAt i "am/pm" ->
            sb.Append("tt") |> ignore
            i <- i + 5
        | 'a' when matchesAt i "a/p" ->
            sb.Append("t") |> ignore
            i <- i + 3
        | '.' when i + 1 < orig.Length && orig.[i + 1] = '0' ->
            let mutable n = 0
            while i + 1 + n < orig.Length && orig.[i + 1 + n] = '0' do n <- n + 1
            sb.Append('.') |> ignore
            sb.Append(System.String('f', min 7 n)) |> ignore
            i <- i + 1 + n
        | '\\' when i + 1 < orig.Length ->
            // Escaped literal char.
            sb.Append('\'').Append(orig.[i + 1]).Append('\'') |> ignore
            i <- i + 2
        | '"' ->
            // Quoted literal.
            let endQuote = orig.IndexOf('"', i + 1)
            if endQuote < 0 then i <- orig.Length
            else
                sb.Append('\'').Append(orig.Substring(i + 1, endQuote - i - 1)).Append('\'') |> ignore
                i <- endQuote + 1
        | '[' ->
            // Elapsed-time bracket [h]/[m]/[s] — .NET can't express elapsed
            // hours past 24 with a stock format; emit a normal token and
            // accept the wraparound for now.
            let endBracket = orig.IndexOf(']', i + 1)
            if endBracket < 0 then i <- i + 1
            else i <- endBracket + 1
        | ':' | '/' | '-' | ' ' | ',' as c ->
            sb.Append(c) |> ignore
            i <- i + 1
        | c ->
            // Unrecognized: wrap as single-quote literal so .NET doesn't
            // throw on it.
            sb.Append('\'').Append(c).Append('\'') |> ignore
            i <- i + 1
    sb.ToString()

/// Excel rounds sub-second time values to the precision implied by the
/// format's `.0+` suffix. .NET's `fff` truncates instead, so we pre-round
/// to that precision before formatting. Returns -1 when no fractional
/// suffix is present (no rounding needed).
let private fractionalSecondPrecision (code: string) : int =
    let section = firstSection code
    let mutable i = 0
    let mutable maxRun = -1
    while i < section.Length do
        if section.[i] = '.' && i + 1 < section.Length && section.[i + 1] = '0' then
            let mutable n = 0
            while i + 1 + n < section.Length && section.[i + 1 + n] = '0' do n <- n + 1
            if n > maxRun then maxRun <- n
            i <- i + 1 + n
        else
            i <- i + 1
    maxRun

/// True when the format's first section contains an elapsed-time bracket
/// like `[h]`, `[mm]`, or `[ss]` (interior is only h/m/s letters). These
/// request a running total of that unit — hours past 24, minutes past 60,
/// seconds without limit — rather than a wall-clock component, which a
/// stock .NET DateTime format string cannot express. Colour / locale tags
/// such as `[Red]` or `[$-409]` are not matched (they carry other letters).
let private hasElapsedBracket (code: string) : bool =
    if code.IndexOf('[') < 0 then false
    else
        let section = firstSection code
        let mutable i = 0
        let mutable found = false
        while not found && i < section.Length do
            if section.[i] = '[' then
                let endBracket = section.IndexOf(']', i + 1)
                if endBracket < 0 then
                    i <- section.Length
                else
                    let mutable j = i + 1
                    let mutable valid = j < endBracket
                    while valid && j < endBracket do
                        match System.Char.ToUpperInvariant section.[j] with
                        | 'H' | 'M' | 'S' -> j <- j + 1
                        | _ -> valid <- false
                    if valid then
                        found <- true
                    i <- endBracket + 1
            else
                i <- i + 1
        found

/// Format an Excel time serial through an elapsed-time format code
/// (`[h]:mm:ss`, `[m]`, `[s].000`, ...). The leading bracket unit carries
/// the full elapsed total; any trailing `mm` / `ss` tokens are the
/// remainder within the parent unit. Matching Excel (verified via
/// `tools/excel-parity/probe-elapsed.vbs`), the serial is rounded to the
/// format's fractional-second precision, away from zero, *before* it is
/// decomposed — so a 32.5 s value prints `0:00:33` under `[h]:mm:ss` yet
/// `32.500` under `[s].000` and `0:00` under `[h]:mm`.
let private formatElapsedTime (code: string) (serial: float) : string =
    let section = firstSection code
    // Drop leading colour / locale tags so the token walk sees only the
    // time pattern (`[Red][h]:mm` -> `[h]:mm`); keep the elapsed brackets.
    let isElapsedBracket start finish =
        let mutable index = start
        let mutable valid = start <= finish
        while valid && index <= finish do
            match System.Char.ToUpperInvariant section.[index] with
            | 'H' | 'M' | 'S' -> index <- index + 1
            | _ -> valid <- false
        valid
    let patternBuilder = System.Text.StringBuilder(section.Length)
    let mutable patternIndex = 0
    while patternIndex < section.Length do
        if section.[patternIndex] = '[' then
            let endBracket = section.IndexOf(']', patternIndex + 1)
            if endBracket >= 0 then
                if isElapsedBracket (patternIndex + 1) (endBracket - 1) then
                    for j = patternIndex to endBracket do
                        patternBuilder.Append(section.[j]) |> ignore
                patternIndex <- endBracket + 1
            else
                patternBuilder.Append(section.[patternIndex]) |> ignore
                patternIndex <- patternIndex + 1
        else
            patternBuilder.Append(section.[patternIndex]) |> ignore
            patternIndex <- patternIndex + 1
    let pattern = patternBuilder.ToString()
    let prec = max 0 (fractionalSecondPrecision pattern)
    let scale = System.Math.Pow(10.0, float prec)
    let totalSec =
        System.Math.Round(serial * 86400.0 * scale, System.MidpointRounding.AwayFromZero) / scale
    let intSec = int64 (floor (abs totalSec))
    let frac = abs totalSec - float intSec
    let totalHours = intSec / 3600L
    let totalMinutes = intSec / 60L
    let minWithinHour = (intSec % 3600L) / 60L
    let secWithinMin = intSec % 60L
    let sign = if totalSec < 0.0 then "-" else ""
    let fracDigits () =
        if prec <= 0 then ""
        else
            let scaled = int64 (System.Math.Round(frac * scale, System.MidpointRounding.AwayFromZero))
            "." + scaled.ToString().PadLeft(prec, '0')
    let sb = System.Text.StringBuilder()
    sb.Append(sign) |> ignore
    let lower = pattern
    let mutable i = 0
    while i < lower.Length do
        let c = lower.[i]
        if c = '[' then
            let endB = lower.IndexOf(']', i + 1)
            if endB < 0 then i <- lower.Length
            else
                let width = endB - i - 1
                let v =
                    match System.Char.ToLowerInvariant lower.[i + 1] with
                    | 'h' -> totalHours
                    | 'm' -> totalMinutes
                    | _ -> intSec
                sb.Append(v.ToString().PadLeft(width, '0')) |> ignore
                i <- endB + 1
        elif c = 'h' || c = 'H' then
            let mutable n = 0
            while i + n < lower.Length && (lower.[i + n] = 'h' || lower.[i + n] = 'H') do n <- n + 1
            // Plain (non-bracketed) hours only appear without a [h] lead;
            // fall back to wall-clock hours so we never throw.
            sb.Append((totalHours % 24L).ToString().PadLeft(min 2 n, '0')) |> ignore
            i <- i + n
        elif c = 'm' || c = 'M' then
            let mutable n = 0
            while i + n < lower.Length && (lower.[i + n] = 'm' || lower.[i + n] = 'M') do n <- n + 1
            sb.Append(minWithinHour.ToString().PadLeft(min 2 n, '0')) |> ignore
            i <- i + n
        elif c = 's' || c = 'S' then
            let mutable n = 0
            while i + n < lower.Length && (lower.[i + n] = 's' || lower.[i + n] = 'S') do n <- n + 1
            sb.Append(secWithinMin.ToString().PadLeft(min 2 n, '0')) |> ignore
            i <- i + n
        elif c = '.' && i + 1 < lower.Length && lower.[i + 1] = '0' then
            let mutable n = 0
            while i + 1 + n < lower.Length && lower.[i + 1 + n] = '0' do n <- n + 1
            sb.Append(fracDigits ()) |> ignore
            i <- i + 1 + n
        elif c = '\\' && i + 1 < lower.Length then
            sb.Append(lower.[i + 1]) |> ignore
            i <- i + 2
        elif c = '"' then
            let endQ = lower.IndexOf('"', i + 1)
            if endQ < 0 then i <- lower.Length
            else
                for j = i + 1 to endQ - 1 do
                    sb.Append(lower.[j]) |> ignore
                i <- endQ + 1
        else
            sb.Append(c) |> ignore
            i <- i + 1
    sb.ToString()

/// Format an Excel serial number as a date string per the given Excel
/// format code. Returns None when the serial is outside DateTime's
/// representable range; caller should render `#####` overflow in that
/// case (Excel's behavior for serials it can't convert).
let private formatExcelDateOpt (code: string) (serial: float) : string option =
    if hasElapsedBracket code then Some (formatElapsedTime code serial) else
    let dt =
        try Some (excelSerialToDate serial)
        with _ -> None
    let inv = System.Globalization.CultureInfo.InvariantCulture
    match dt with
    | None -> None
    | Some dt ->
        let dtRounded =
            // Excel rounds a time to the precision of its finest displayed
            // unit before drawing. With an explicit `.0+` suffix that's the
            // fractional second; otherwise, when seconds are shown at all,
            // round to the whole second (Excel renders 59.9999996 s as the
            // next whole second, e.g. `0:01:00`, where .NET would truncate
            // to `0:00:59`). Date-only / minute-only formats keep `prec < 0`
            // so they are left untouched.
            let prec =
                let p = fractionalSecondPrecision code
                if p >= 0 then p
                elif containsIgnoringLocaleTags 's' (firstSection code) then 0
                else -1
            if prec < 0 then dt
            else
                // Keep the exponent in float space. Fable maps int64 `pown`
                // through BigInt, which browser Math.pow rejects.
                let unitsPerSecond = System.Math.Pow(10.0, float prec)
                let ticksPerUnit = int64 (float System.TimeSpan.TicksPerSecond / unitsPerSecond)
                if ticksPerUnit <= 0L then dt
                else
                    let rounded = (dt.Ticks + ticksPerUnit / 2L) / ticksPerUnit * ticksPerUnit
                    try System.DateTime(rounded) with _ -> dt
        let netFormat = excelToNetFormat code
        try Some (dtRounded.ToString(netFormat, inv))
        with _ -> Some (dtRounded.ToString("MM/dd/yyyy", inv))

/// Approximate rendered Maximum Digit Width (MDW) in CSS pixels for a given
/// XLSX font size (in points). Excel's MDW table is flat around the common
/// 10/11/12 pt sizes (all 7 px), then jumps to 8 px at 12-13 pt. Use the
/// `round(fontPt × 7/11)` formula but floor it at 7 — the value Excel uses
/// for every Calibri / Arial size from 8 pt up through 11 pt. Without the
/// floor, size-10 cells (Arial 10pt, the Excel 2003 legacy default in many
/// pandas-test_datetime_mi-style POI fixtures) under-count chars and trip
/// ##### overflow on values that Excel renders inside the cell.
/// Public so the paint layer can size column-width heuristics from the same
/// Maximum-Digit-Width table the formatter uses for ##### overflow.
let mdwPx (fontPt: float) : float =
    max 7.0 (System.Math.Round(fontPt * 7.0 / 11.0))

let private generalNumberFormats =
    [| "G11"; "G10"; "G9"; "G8"; "G7"; "G6"; "G5"; "G4"; "G3"; "G2"; "G1" |]

/// Width-aware Excel "General" number formatting. Picks the longest .NET
/// "G<n>" representation that fits the cell width. If even the shortest
/// representation overflows, returns `#####` like Excel does.
let private generalNumberDisplay (cellWidthPx: float) (fontPt: float) (n: float): string =
    // Excel's General-format digit budget = floor((cellWidthPx - 5) / MDW).
    // The 5 px matches OOXML's column-width padding (2 px each side + 1 px
    // gridline) so pixels and effective-char count round-trip via the same
    // formula. Verified against pandas-test1 col B (88 px / mdw=8 → 10
    // chars → "0.98026851", matches Excel COM cell.Text). Try the full
    // precision sweep so we surface G9 mid-trim instead of jumping from
    // G10 straight to G8.
    let usableWidth = max 0.0 (cellWidthPx - 5.0)
    let maxChars = max 1 (int (usableWidth / mdwPx fontPt))

    let inv = System.Globalization.CultureInfo.InvariantCulture
    let mutable result: string voption = ValueNone
    let mutable i = 0
    while result.IsNone && i < generalNumberFormats.Length do
        let s = n.ToString(generalNumberFormats.[i], inv)
        if s.Length <= maxChars then
            result <- ValueSome s
        i <- i + 1
    match result with
    | ValueSome s -> s
    | ValueNone -> System.String('#', maxChars)

let private fractionDenominatorDigits (code: string) : int option =
    let mutable i = 0
    while i < code.Length && (code.[i] = '#' || code.[i] = '0') do
        i <- i + 1
    while i < code.Length && System.Char.IsWhiteSpace code.[i] do
        i <- i + 1

    let numStart = i
    while i < code.Length && code.[i] = '?' do
        i <- i + 1
    let numDigits = i - numStart

    if numDigits = 0 || i >= code.Length || code.[i] <> '/' then None
    else
        i <- i + 1
        let denStart = i
        while i < code.Length && code.[i] = '?' do
            i <- i + 1
        let denDigits = i - denStart
        if denDigits = 0 || i <> code.Length then None else Some denDigits

/// Render an Excel fraction format like `# ?/?` (one-digit denominator),
/// `# ??/??` (two), `# ???/???` (three). Splits the number into its
/// integer part and the closest reduced fraction with denominator ≤
/// 10^digits, mirroring Excel's display. Returns None if the code isn't
/// a recognized fraction pattern.
let private tryFormatFraction (code: string) (n: float) : string option =
    // Match `<intPart>?<spaces>?/?` style: any number of ?'s on both sides
    // of the slash. The integer part can be empty or "#"/"0" placeholders.
    match fractionDenominatorDigits code with
    | None -> None
    | Some denDigits ->
        let maxDen = pown 10 denDigits - 1
        let sign = if n < 0.0 then "-" else ""
        let abs = System.Math.Abs n
        let whole = int (System.Math.Floor abs)
        let frac = abs - float whole
        // Stern-Brocot search for the best rational with denominator
        // ≤ maxDen. Exact enough for Excel's display.
        let mutable bestNum, bestDen = 0, 1
        let mutable bestErr = frac
        for d in 1 .. maxDen do
            let nn = int (System.Math.Round (frac * float d))
            let err = System.Math.Abs (frac - float nn / float d)
            if err < bestErr then
                bestErr <- err
                bestNum <- nn
                bestDen <- d
        // Reduce.
        let rec gcd a b = if b = 0 then a else gcd b (a % b)
        let g = gcd bestNum bestDen |> max 1
        let nn, dd = bestNum / g, bestDen / g
        let result =
            if nn = 0 && whole = 0 then "0"
            elif nn = 0 then sprintf "%s%d" sign whole
            elif whole = 0 then sprintf "%s%d/%d" sign nn dd
            else sprintf "%s%d %d/%d" sign whole nn dd
        Some result

/// Strip Excel's column-width placeholders that have no .NET equivalent.
/// Excel uses `_X` to mean "skip a column's worth of width = width(X)"
/// (typically used as `_(` / `_)` to align with paren-bracketed negatives)
/// and `*X` to mean "fill with character X" (typically `* ` to right-
/// justify a "-" placeholder for zero). Neither has a placement in
/// monospaced PNG/PDF output, so collapse `_X` to a single space and
/// drop the `*` filler entirely (rendering the next char once).
///
/// Also normalize Excel's `?` digit-or-space placeholder. .NET has no
/// direct equivalent — `?` looks like a literal in custom-format-string
/// land, which would print as `??` rather than space-padding. The
/// pragmatic mapping is `?` → `#` (digit or nothing): produces the
/// same visual for non-zero values and renders an empty string for
/// zero, which is what Excel's accounting "-" section needs anyway.
/// Skip `?` inside quoted literal segments.
let private stripExcelWidthPlaceholders (code: string) : string =
    let sb = System.Text.StringBuilder(code.Length)
    let mutable i = 0
    let mutable inQuote = false
    while i < code.Length do
        let c = code.[i]
        if c = '"' then
            sb.Append(c) |> ignore
            inQuote <- not inQuote
            i <- i + 1
        elif inQuote then
            sb.Append(c) |> ignore
            i <- i + 1
        elif (c = '_' || c = '*') && i + 1 < code.Length then
            // `_X` reserves the width of one X glyph (Excel uses this
            // for column-aligning open/close parens in accounting
            // formats); `*X` repeats X to fill remaining column width.
            // Without access to the cell width here we substitute a
            // single occurrence. Width-aware accounting fill expands this
            // marker later, after the overflow decision has been made.
            if c = '_' then sb.Append(' ') |> ignore
            else sb.Append(code.[i + 1]) |> ignore
            i <- i + 2
        elif c = '?' then
            sb.Append('#') |> ignore
            i <- i + 1
        else
            sb.Append(c) |> ignore
            i <- i + 1
    sb.ToString()

/// Excel scales a number by 1000 for each comma at the END of the format
/// (after the last digit placeholder) — `#,##0,`→thousands, `0.0,,`→millions.
/// .NET only honours a scaling comma immediately LEFT of the decimal point, so
/// a trailing comma after the decimals (`#,##0.0,`) is ignored. Strip the
/// trailing scaling commas and return the matching divisor so we scale the
/// value ourselves; the grouping commas inside the digit run are left intact.
let private trailingScale (section: string) : string * float =
    let isDigitPh c = c = '0' || c = '#' || c = '?'
    let mutable idx = section.Length - 1
    while idx >= 0 && not (isDigitPh section.[idx]) do
        idx <- idx - 1
    if idx < 0 then
        section, 1.0
    else
        let mutable i = idx + 1
        while i < section.Length && section.[i] = ',' do i <- i + 1
        let commas = i - (idx + 1)
        if commas = 0 then section, 1.0
        else section.Substring(0, idx + 1) + section.Substring(i), (1000.0 ** float commas)

/// Excel's General text for a number, used when a `@` text placeholder is
/// applied to a numeric value (Excel shows the number, e.g. `@`→"5"/"-3.5").
let private generalString (inv: System.Globalization.CultureInfo) (n: float) : string =
    if n = floor n && abs n < 1e15 then (int64 n).ToString(inv)
    else n.ToString("G15", inv)

let private isSimpleDotNetNumberFormat (section: string) =
    let mutable i = 0
    let mutable simple = true
    while simple && i < section.Length do
        match section.[i] with
        | ';' | '[' | ']' | '"' | '\\' | '_' | '*' | '?' | '@' | '/' -> simple <- false
        | _ -> i <- i + 1
    simple

let private isHexDigit c =
    (c >= '0' && c <= '9')
    || (c >= 'A' && c <= 'F')
    || (c >= 'a' && c <= 'f')

let private appendQuotedLiteral (sb: System.Text.StringBuilder) (text: string) start length =
    if length > 0 then
        sb.Append('\'') |> ignore
        for i = start to start + length - 1 do
            let c = text.[i]
            if c = '\'' then sb.Append("''") |> ignore
            else sb.Append(c) |> ignore
        sb.Append('\'') |> ignore

let private replaceCurrencyLocaleTags (text: string) =
    if text.IndexOf("[$", System.StringComparison.Ordinal) < 0 then
        text
    else
        let sb = System.Text.StringBuilder(text.Length)
        let mutable changed = false
        let mutable i = 0
        while i < text.Length do
            if i + 1 < text.Length && text.[i] = '[' && text.[i + 1] = '$' then
                let prefixStart = i + 2
                let mutable j = prefixStart
                while j < text.Length && text.[j] <> ']' && text.[j] <> '-' do
                    j <- j + 1
                let mutable close = -1
                if j < text.Length && text.[j] = ']' then
                    close <- j
                elif j < text.Length && text.[j] = '-' then
                    let hexStart = j + 1
                    let mutable k = hexStart
                    while k < text.Length && isHexDigit text.[k] do
                        k <- k + 1
                    if k > hexStart && k < text.Length && text.[k] = ']' then
                        close <- k
                if close >= 0 then
                    appendQuotedLiteral sb text prefixStart (j - prefixStart)
                    changed <- true
                    i <- close + 1
                else
                    sb.Append(text.[i]) |> ignore
                    i <- i + 1
            else
                sb.Append(text.[i]) |> ignore
                i <- i + 1
        if changed then sb.ToString() else text

type private NumberFormatSectionInfo = {
    Pos: string
    ScaleDivisor: float
    IsSimple: bool
    Cleaned: string
    ContainsTextPlaceholder: bool
    HasCondition: bool
}

let private numberFormatSectionInfoCache = Dictionary<string, NumberFormatSectionInfo>()
let private numberFormatSectionInfoCacheOrder = Queue<string>()
let private numberFormatSectionInfoCacheMaxEntries = 4096
let private numberFormatSectionInfoCacheMaxTotalChars = 5000000
let mutable private numberFormatSectionInfoCacheTotalChars = 0

let private numberFormatSectionInfoValueTextCharCount section (info: NumberFormatSectionInfo) =
    let mutable total = 0
    if not (System.Object.ReferenceEquals(info.Pos, section)) then
        total <- addStringCacheCharsCapped numberFormatSectionInfoCacheMaxTotalChars info.Pos.Length total
    if not (System.Object.ReferenceEquals(info.Cleaned, section))
       && not (System.Object.ReferenceEquals(info.Cleaned, info.Pos)) then
        total <- addStringCacheCharsCapped numberFormatSectionInfoCacheMaxTotalChars info.Cleaned.Length total
    total

let private numberFormatSectionInfo section =
    let mutable info = Unchecked.defaultof<NumberFormatSectionInfo>
    if numberFormatSectionInfoCache.TryGetValue(section, &info) then
        info
    else
        let pos, scaleDivisor = trailingScale section
        let isSimple = isSimpleDotNetNumberFormat pos
        let cleaned =
            if isSimple then
                pos
            else
                let withoutPlaceholders = stripExcelWidthPlaceholders pos
                let translated =
                    let sb = System.Text.StringBuilder(withoutPlaceholders.Length)
                    let mutable i = 0
                    let mutable inDquote = false
                    while i < withoutPlaceholders.Length do
                        let c = withoutPlaceholders.[i]
                        if c = '\\' && i + 1 < withoutPlaceholders.Length then
                            sb.Append('\'').Append(withoutPlaceholders.[i + 1]).Append('\'') |> ignore
                            i <- i + 2
                        elif c = '"' then
                            sb.Append('\'') |> ignore
                            inDquote <- not inDquote
                            i <- i + 1
                        else
                            sb.Append(c) |> ignore
                            i <- i + 1
                    sb.ToString()
                translated
                |> replaceCurrencyLocaleTags
                |> removeColorAndConditionTags
        info <-
            {
                Pos = pos
                ScaleDivisor = scaleDivisor
                IsSimple = isSimple
                Cleaned = cleaned
                ContainsTextPlaceholder = pos.Contains "@"
                HasCondition = (sectionConditionCached pos).IsSome
            }
        addBoundedStringCacheEntryWithValueChars
            numberFormatSectionInfoCacheMaxEntries
            numberFormatSectionInfoCacheMaxTotalChars
            numberFormatSectionInfoValueTextCharCount
            &numberFormatSectionInfoCacheTotalChars
            numberFormatSectionInfoCache
            numberFormatSectionInfoCacheOrder
            section
            info

/// Apply an Excel non-date number-format code (`0.000`, `#,##0.00`, `0%`,
/// `0.00E+00`, etc.) using .NET's custom-format-string support. Excel's
/// number-format grammar is largely compatible with .NET's; we strip the
/// section separator first and pass the positive section to `Double.ToString`,
/// falling back to `G` if .NET rejects it.
let private formatExcelNumber (code: string) (n: float) : string * string =
    let inv = System.Globalization.CultureInfo.InvariantCulture
    let selectedSection, isConditional, sectionCount =
        if code.IndexOf(';') < 0 then
            code, (code.IndexOf('[') >= 0 && (sectionConditionCached code).IsSome), 1
        else
            let sections = splitSectionsCached code
            let struct (section, conditional, sectionCount) = sectionSelectionForValue code sections n
            section, conditional, sectionCount
    let sectionInfo = numberFormatSectionInfo selectedSection
    let pos = sectionInfo.Pos
    let scaleDivisor = sectionInfo.ScaleDivisor
    // Excel formats the section against the ABSOLUTE value when the section
    // already accounts for the sign (a dedicated negative/zero section, or any
    // conditioned section), so .NET doesn't prepend a second `-` to something
    // like `(#,##0.00)`. When a NEGATIVE value falls through to a conditional
    // format's default (unconditioned) section, Excel prepends a `-` itself
    // (e.g. `[>=100]0;[>=50]0;0` renders -10 as `-10`, "Low" as "-Low").
    let baseValue, prependMinus =
        if isConditional then
            if sectionInfo.HasCondition then System.Math.Abs n, false
            else System.Math.Abs n, (n < 0.0)
        elif (n < 0.0 && sectionCount >= 2) || (n = 0.0 && sectionCount >= 3) then
            System.Math.Abs n, false
        else n, false
    let valueForSection = baseValue / scaleDivisor
    if System.String.IsNullOrWhiteSpace pos then
        // An empty section renders nothing — e.g. `0.00;;` hides zeros and
        // negatives, `0;;` hides zeros.
        "", selectedSection
    elif scaleDivisor = 1.0 && not prependMinus && sectionInfo.IsSimple then
        let formatted =
            try valueForSection.ToString(pos, inv)
            with _ ->
                try n.ToString("G", inv)
                with _ -> string n
        formatted, selectedSection
    else
    // Fractions (`# ?/?`, `# ??/??`, `?/??`) have no .NET equivalent and
    // need explicit numerator-denominator search. Try that first.
    match tryFormatFraction pos n with
    | Some s -> s, selectedSection
    | None ->
    if sectionInfo.ContainsTextPlaceholder then
        // `@` text placeholder on a number → its General text, with quoted
        // literals / escapes unwrapped and colour tags dropped.
        let s = pos.Replace("@", generalString inv n)
        unwrapTextPlaceholderLiterals s, selectedSection
    else
        let formatted =
            try valueForSection.ToString(sectionInfo.Cleaned, inv)
            with _ ->
                try n.ToString("G", inv)
                with _ -> string n
        (if prependMinus then "-" + formatted else formatted), selectedSection

/// Expand Excel's `* ` accounting fill when it follows a literal currency
/// marker. The fill characters are alignment padding, not overflow content:
/// Excel first decides whether the number fits, then repeats the fill so the
/// currency marker stays near the left edge and the number stays rightward.
let private expandAccountingFill
    (measure: (string -> float) option)
    (cellWidth: float)
    (section: string)
    (text: string) : string =
    let isHash =
        let mutable allHash = text.Length > 0
        let mutable i = 0
        while allHash && i < text.Length do
            if text.[i] <> '#' then allHash <- false
            i <- i + 1
        allHash
    match measure with
    | None -> text
    | Some m when isHash -> text
    | Some m ->
        let star = section.IndexOf("* ")
        if star <= 0 then text
        else
            let prev = section.[star - 1]
            // Covers Excel's common accounting patterns: literal `_($* #,##0)`
            // (prev = '$') and the locale currency tag `_-[$$-409]* #,##0`
            // (prev = ']', the close of a `[$…]` marker). Non-currency star
            // fills stay collapsed — right-aligned drawing already matches.
            if prev <> '$' && prev <> ']' then text
            else
                let marker = "$ "
                let markerAt = text.IndexOf(marker, System.StringComparison.Ordinal)
                if markerAt < 0 then text
                else
                    let suffixStart = markerAt + marker.Length
                    let appendWithoutMarker (sb: System.Text.StringBuilder) =
                        for i = 0 to markerAt do
                            sb.Append(text.[i]) |> ignore
                        for i = suffixStart to text.Length - 1 do
                            sb.Append(text.[i]) |> ignore
                    let fillWidth = max 1.0 (m " ")
                    let baseText = System.Text.StringBuilder(text.Length - 1)
                    appendWithoutMarker baseText
                    let baseWidth = m (baseText.ToString())
                    // Excel's text fill runs slightly past the visible cell
                    // box; cellWidth + 5 px matches the OOXML cell padding
                    // boundary and synthetic-numfmt's COM cell.Text exactly.
                    let target = cellWidth + 5.0
                    let count =
                        int (System.Math.Round((target - baseWidth) / fillWidth))
                        |> max 1
                    let result = System.Text.StringBuilder(text.Length + count - 1)
                    for i = 0 to markerAt do
                        result.Append(text.[i]) |> ignore
                    result.Append(System.String(' ', count)) |> ignore
                    for i = suffixStart to text.Length - 1 do
                        result.Append(text.[i]) |> ignore
                    result.ToString()

/// Greedy word-wrap inside a cell. Splits on hard line breaks (\n) first,
/// then on whitespace within each segment, packing as many words onto a line
/// as the cell's usable width allows. Falls back to char-level breaks when
/// a single word exceeds the width.
///
/// When `measure` is Some, line-width budget is computed via the real loaded
/// font (matches Excel's glyph-aware wrapping). When None — Fable web canvas
/// or other contexts without a measurer — fall back to `mdwPx × charCount`,
/// which over-estimates width and wraps earlier than Excel.
let wrapLines
    (measure: (string -> float) option)
    (cellWidthPx: float) (fontPt: float)
    (text: string) : string list =
    if System.String.IsNullOrEmpty text then [ text ]
    else
        // Excel's effective usable wrap width is cellWidth - 14 px. The
        // older `cellWidth - 8` budget under-broke wide-column wraps:
        // synthetic-wrap-autoheight's 154-px cell wraps to 8 lines in
        // Excel, but our text engine measures `width like the one Excel`
        // at 140.36 px which fits in 146 (= 154 - 8) and bled to a 7-line
        // wrap. With the tighter 14-px inset the budget drops to 140 and
        // that line breaks before "Excel", restoring the 8-line count.
        let usable = max 0.0 (cellWidthPx - 14.0)
        // Approximation budget for the measurer-less path.
        let mdw = mdwPx fontPt
        let approxMaxChars = max 1 (int (usable / mdw))
        match measure with
        | None ->
            if text.[0] <> ' ' && text.Length <= approxMaxChars && text.IndexOf('\n') < 0 then
                [ text ]
            else
                let lines = System.Collections.Generic.List<string>()
                let current = System.Text.StringBuilder()
                let appendSlice start length =
                    for offset = 0 to length - 1 do
                        current.Append(text.[start + offset]) |> ignore
                let addCurrent () =
                    if current.Length > 0 then
                        lines.Add(current.ToString())
                        current.Clear() |> ignore
                let appendOversizedWord start length =
                    let mutable offset = 0
                    while length - offset > approxMaxChars do
                        lines.Add(text.Substring(start + offset, approxMaxChars))
                        offset <- offset + approxMaxChars
                    let remaining = length - offset
                    if remaining > 0 then
                        appendSlice (start + offset) remaining
                let appendWord start length =
                    if current.Length = 0 then
                        if length > approxMaxChars then appendOversizedWord start length
                        else appendSlice start length
                    else
                        let candidateLength = current.Length + 1 + length
                        if candidateLength <= approxMaxChars then
                            current.Append(' ') |> ignore
                            appendSlice start length
                        else
                            addCurrent ()
                            if length > approxMaxChars then appendOversizedWord start length
                            else appendSlice start length
                let wrapSegment segmentStart segmentEnd =
                    let mutable wordStart = segmentStart
                    let mutable i = segmentStart
                    while i <= segmentEnd do
                        if i = segmentEnd || text.[i] = ' ' then
                            appendWord wordStart (i - wordStart)
                            wordStart <- i + 1
                        i <- i + 1
                    addCurrent ()
                let mutable segmentStart = 0
                let mutable i = 0
                while i <= text.Length do
                    if i = text.Length || text.[i] = '\n' then
                        wrapSegment segmentStart i
                        segmentStart <- i + 1
                    i <- i + 1
                resizeArrayToList lines
        | Some m ->
            let fits (s: string) = m s <= usable
            if text.[0] <> ' ' && text.IndexOf('\n') < 0 && fits text then
                [ text ]
            else
                let lines = System.Collections.Generic.List<string>()
                let current = System.Text.StringBuilder()
                let mutable currentWidth = 0.0
                let spaceWidth = m " "
                let exactFitSlack = 4.0
                let wordText start length =
                    if length = 0 then "" else text.Substring(start, length)
                let breakOversizedWord start length =
                    // Char-by-char until we'd overflow; emit each chunk as its
                    // own line. Used when a single word is wider than the cell.
                    let mutable buf = System.Text.StringBuilder()
                    let mutable bufWidth = 0.0
                    for i = start to start + length - 1 do
                        let ch = text.[i]
                        buf.Append(ch) |> ignore
                        let candidate = buf.ToString()
                        let candidateWidth = m candidate
                        if candidateWidth > usable && buf.Length > 1 then
                            // Last char overflowed; emit buffer minus last char,
                            // then start a new buffer with just that char.
                            lines.Add(candidate.Substring(0, candidate.Length - 1))
                            buf <- System.Text.StringBuilder()
                            buf.Append(ch) |> ignore
                            bufWidth <- m (string ch)
                        else
                            bufWidth <- candidateWidth
                    struct (buf, bufWidth)
                let candidate = System.Text.StringBuilder()
                let candidateFits (word: string) wordWidth =
                    let approximateWidth = currentWidth + spaceWidth + wordWidth
                    if approximateWidth <= usable - exactFitSlack then
                        struct (true, approximateWidth)
                    elif approximateWidth > usable + exactFitSlack then
                        struct (false, approximateWidth)
                    else
                        candidate.Clear() |> ignore
                        candidate.Append(current).Append(' ').Append(word) |> ignore
                        let exactWidth = m (candidate.ToString())
                        struct (exactWidth <= usable, exactWidth)
                let emitCurrent () =
                    lines.Add(current.ToString())
                    current.Clear() |> ignore
                    currentWidth <- 0.0
                let appendWord start length =
                    let word = wordText start length
                    let wordWidth = if length = 0 then 0.0 else m word
                    if current.Length = 0 then
                        if wordWidth > usable then
                            // word longer than cell — char-break, leaving the
                            // last chunk in `current` to continue with next word.
                            let struct (leftover, leftoverWidth) = breakOversizedWord start length
                            if leftover.Length > 0 then
                                current.Append(leftover) |> ignore
                                currentWidth <- leftoverWidth
                        else
                            current.Append(word) |> ignore
                            currentWidth <- wordWidth
                    else
                        let struct (doesFit, candidateWidth) = candidateFits word wordWidth
                        if doesFit then
                            current.Append(' ').Append(word) |> ignore
                            currentWidth <- candidateWidth
                        else
                            emitCurrent ()
                            if wordWidth > usable then
                                let struct (leftover, leftoverWidth) = breakOversizedWord start length
                                if leftover.Length > 0 then
                                    current.Append(leftover) |> ignore
                                    currentWidth <- leftoverWidth
                            else
                                current.Append(word) |> ignore
                                currentWidth <- wordWidth
                let wrapSegment segmentStart segmentEnd =
                    let mutable wordStart = segmentStart
                    let mutable i = segmentStart
                    while i <= segmentEnd do
                        if i = segmentEnd || text.[i] = ' ' then
                            appendWord wordStart (i - wordStart)
                            wordStart <- i + 1
                        i <- i + 1
                    if current.Length > 0 then
                        emitCurrent ()
                let mutable segmentStart = 0
                let mutable i = 0
                while i <= text.Length do
                    if i = text.Length || text.[i] = '\n' then
                        wrapSegment segmentStart i
                        segmentStart <- i + 1
                    i <- i + 1
                resizeArrayToList lines

/// Excel shows ##### when a number's display string is wider than the cell.
///
/// `measure` is the cell's real-font width function (Some on the native /
/// PDF paths, None on the Fable web canvas). When present we both decide
/// overflow AND size the `#` run against the *actual* glyph the renderer
/// will draw. The older `mdwPx`-only path under-counted: a date string
/// measured at ~7 px/char but rendered in Calibri (~9 px '#') produced a
/// `#####` run wider than the cell, and since numbers right-align the
/// excess spilled left into the row-header gutter. The None branch keeps
/// the approximation for the web build.
let private fitOrOverflow
    (measure: (string -> float) option)
    (cellWidth: float) (fontName: string option) (fontPt: float) (text: string) : string =
    // Two thresholds: the FIT check (cellWidth - 8) matches the 4-px
    // text inset on each side in buildCellPaintItems. The HASH BUDGET
    // (cellWidth - 5) matches OOXML's column-width padding so the
    // overflow-`#` count round-trips with Excel's chars-per-pixel math.
    // Verified against pandas-test1 Sheet1 col A (95 px / mdw=8 → 11 #s,
    // matches Excel COM cell.Text for the m/d/yyyy h:mm date that
    // overflows).
    //
    // Fit check uses the measured glyph width. Excel's effective fit
    // budget is `cellWidth - 3` (~2-px inset on the right with a tiny
    // left inset) — verified empirically against pandas-testskiprows C4
    // "2015-01-01 00:00:00" which Excel renders in a 128-px cell at our
    // measured 124.2 px. The 3-px budget lets it fit; the older 8-px
    // budget hashed it. pandas-test1 m/d/yyyy h:mm dates still hash
    // because the date's glyph width sits 3-4 px above the threshold
    // for the narrower Calibri-12 95-px cells.
    //
    // Calibri 11 is the one common case where Excel's COM `cell.Text`
    // boundary sits at the full cell width: synthetic-narrow-date fits
    // "1/10/2000 0:00" in a 94-px cell, while the same text at Calibri
    // 12 still hashes in a 95-px cell (pandas-test1). Keep the tighter
    // inset for the other sizes/families.
    let isCalibri11 =
        let name = fontName |> Option.defaultValue "Calibri"
        trimmedStartsWithIgnoreCase name "Calibri"
        && abs (fontPt - 11.0) < 0.01
    let fitInset = if isCalibri11 then 0.0 else 3.0
    let fitWidth = max 0.0 (cellWidth - fitInset)
    let hashBudget = max 0.0 (cellWidth - 5.0)
    // Use the actual font's digit width when we have a measurer — Excel's
    // hash budget rounds to chars-per-MDW where MDW is the digit advance
    // of the cell's font. The nominal `mdwPx fontPt` ignores font family
    // (always returns 7 for size 11), which under-counts hashes for Arial
    // workbooks where MDW=8 (pandas-testdateoverflow's Polish-locale
    // cells resolve to Arial; Excel renders 29 #s = floor(235/8), we used
    // to render 33 = floor(235/7)).
    let mdw =
        match measure with
        | Some m ->
            let zero = m "0"
            if zero <= 0.0 then mdwPx fontPt else max 7.0 (System.Math.Floor zero)
        | None -> mdwPx fontPt
    match measure with
    | Some m ->
        let textWidth = m text
        if textWidth <= fitWidth then text
        else
            let count = max 1 (int (hashBudget / mdw))
            System.String('#', count)
    | None ->
        let fitChars = max 1 (int (fitWidth / mdw))
        let hashChars = max 1 (int (hashBudget / mdw))
        if text.Length <= fitChars then text
        else System.String('#', hashChars)

/// Per-cell display text logic. Public so the VBA-style API surface
/// (Spreadsheet GetVbaLayoutCells) can return the same text Excel's
/// `cell.Text` would produce — including `#####` overflow, format-applied
/// numbers / dates, error markers.
let cellDisplayText
    (measure: (string -> float) option)
    (cellWidth: float) (format: CellFormat) (value: CellValue): string =
    let fontPt = format.FontSize |> Option.map float |> Option.defaultValue 11.0
    match unwrapFormulaValue value with
    | CellValue.Empty -> ""
    | Number n ->
        match format.NumberFormat with
        | Some code when isDateOrTimeFormat code || hasElapsedBracket code ->
            // Out-of-range serials (e.g. testdateoverflow's 1E+20) can't be
            // formatted as a date; Excel renders ##### overflow for those
            // rather than falling through to scientific notation. Treat
            // the None branch as "doesn't fit" by passing an oversized
            // text that fitOrOverflow will hash-render.
            match formatExcelDateOpt code n with
            | Some s -> fitOrOverflow measure cellWidth format.FontName fontPt s
            | None -> fitOrOverflow measure cellWidth format.FontName fontPt (System.String('M', 1000))
        | Some "currency"
        | Some "percentage"
        | Some "date"
        | Some "decimal2"
        | Some "number" ->
            // Legacy named format codes from the Fable web UI's formatting
            // commands. Keep these on the original Model.formatNumber path.
            fitOrOverflow measure cellWidth format.FontName fontPt (formatNumber format.NumberFormat n)
        | Some code when System.String.IsNullOrWhiteSpace code ||
                         code.Equals("general", System.StringComparison.OrdinalIgnoreCase) ->
            generalNumberDisplay cellWidth fontPt n
        | Some code ->
            // Real Excel format code (0.000, #,##0.00, 0%, 0.00E+00, etc.).
            let formatted, selectedSection = formatExcelNumber code n
            let fitted = fitOrOverflow measure cellWidth format.FontName fontPt formatted
            if selectedSection.IndexOf("* ") < 0 then fitted
            else expandAccountingFill measure cellWidth selectedSection fitted
        | None -> generalNumberDisplay cellWidth fontPt n
    | String s -> s
    | Boolean b -> if b then "TRUE" else "FALSE"
    | Error e -> sprintf "#%s" e
    | Formula _ -> ""
