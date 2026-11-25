// dotnet fsi tools/excel-parity/regex-compare.fsx [engine-out.jsonl]
// The regular expression engine (Functions/Pcre.fs, the Debug build) on the cases of regex-perl.jsonl
// (make-regex-cases.py, answered by regex-perl-oracle.pl): prints every case where the engine and perl
// differ - in compiling, in the first match's groups, in the matches of a global search - and the count.
// Lines starting "known:" are differences where PCRE2 is known to differ from perl (listed below), not
// counted. With an argument, writes the engine's answers there (as regex-perl.jsonl has them: the
// browser build is compared with the same file, see regex-js-check.mjs).
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"
open System
open System.IO
open System.Text.Json
open Rbt.Spreadsheet.Functions.Pcre

let here = __SOURCE_DIRECTORY__
let lines = File.ReadAllLines(Path.Combine(here, "regex-perl.jsonl"))
let output = fsi.CommandLineArgs |> Array.tryItem 1
let written = ResizeArray<string>()

/// Cases where PCRE2 is known to answer otherwise than perl (the engine answering as PCRE2 does).
let knownCases =
    [ "\\Q.*\\E+", "perl does \\Q..\\E only in a pattern written in the program, not in one read from a string"
      "(?:(a)|(b))+", "PCRE2 keeps a group's capture of an earlier iteration (pcre2compat); perl unsets it"
      "(?U)a+", "perl has no (?U)"
      "(?U)a+?", "perl has no (?U)"
      "straße", "perl folds case fully (ß = ss); PCRE2 simply"
      "(?<q>\")?\\w+(?(q)\")", "PCRE2 takes (?(name)...); perl wants (?(<name>)...)"
      "(a|b)c\\g<1>", "perl has no \\g<n> subroutine calls"
      "(?=a??\\d?\\w?).+?(?:\\b)[ab]*?", "perl's optimiser skips the match at 1 (.+?\\b alone finds it)"
      "(x*|(?:a|b))(^$(c++)|((?<!b).(?<!b)[^a]??|$\\s{0,2}?)[^a])*+x", "perl reports captures of a path it backtracked out of" ]

/// Where PCRE2 and perl part: a pattern perl takes and PCRE2 rejects, or the other way round.
let knownDifference (pattern: string) =
    knownCases |> List.exists (fun (p, _) -> p = pattern) ||
    // PCRE2 does not repeat an assertion (\b*, ^+); perl does (with a warning)
    pattern.Contains "\\b*" || pattern.Contains "\\b+" || pattern.Contains "\\b?" || pattern.Contains "\\b{"
    || pattern.Contains "^*" || pattern.Contains "^+" || pattern.Contains "^?" || pattern.Contains "^{"
    || pattern.Contains "$*" || pattern.Contains "$+" || pattern.Contains "$?" || pattern.Contains "${"

let mutable differ = 0
let mutable known = 0
for line in lines do
    let doc = JsonDocument.Parse line
    let root = doc.RootElement
    let pattern = root.GetProperty("pattern").GetString()
    let subject = root.GetProperty("subject").GetString()
    let ci = root.GetProperty("i").GetBoolean()
    let perlError = root.GetProperty("error").GetBoolean()
    let cps = codePoints subject
    let engine =
        try
            let re = compileRegex pattern ci
            let first = search re cps 0 |> Option.map List.ofArray
            let all = allMatches re cps |> List.truncate 1001 |> List.map (fun caps -> [ caps.[0]; caps.[1] ])
            Ok (first, all)
        with
        | PatternError message -> Result.Error ("pattern error: " + message)
        | MatchLimitExceeded -> Result.Error "match limit"
    let json (o: obj) = JsonSerializer.Serialize o
    let engineLine =
        match engine with
        | Ok (first, all) ->
            sprintf "{\"all\":%s,\"error\":false,\"first\":%s,\"i\":%b,\"pattern\":%s,\"subject\":%s}"
                (json all) (match first with Some f -> json f | None -> "null") ci (json pattern) (json subject)
        | Result.Error _ -> sprintf "{\"error\":true,\"i\":%b,\"pattern\":%s,\"subject\":%s}" ci (json pattern) (json subject)
    written.Add engineLine
    let report (what: string) =
        if knownDifference pattern then
            known <- known + 1
            printfn "known: %s  /%s/%s on %s" what pattern (if ci then "i" else "") (json subject)
        else
            differ <- differ + 1
            printfn "DIFF %s  /%s/%s on %s" what pattern (if ci then "i" else "") (json subject)
    match engine with
    | Result.Error message -> if not perlError then report ("engine " + message + ", perl compiles")
    | Ok (first, all) ->
        if perlError then report "perl rejects, engine compiles"
        else
            let perlFirst =
                match root.GetProperty("first").ValueKind with
                | JsonValueKind.Null -> None
                | _ -> Some [ for x in root.GetProperty("first").EnumerateArray() -> x.GetInt32() ]
            let perlAll = [ for m in root.GetProperty("all").EnumerateArray() -> [ for x in m.EnumerateArray() -> x.GetInt32() ] ]
            // perl's @- has no entries past the last group that took part; the engine has every group
            let trimmed (groups: int list option) =
                groups |> Option.map (fun g ->
                    let mutable l = g
                    while l.Length > 2 && l.[l.Length - 1] = -1 && l.[l.Length - 2] = -1 do l <- List.take (l.Length - 2) l
                    l)
            let sameMatch = (first |> Option.map (List.take 2)) = (perlFirst |> Option.map (List.take 2))
            if trimmed first <> trimmed perlFirst then
                // perl keeps what a group captured inside a negative lookaround; PCRE2 does not
                if sameMatch && (pattern.Contains "(?!" || pattern.Contains "(?<!") then
                    known <- known + 1
                    printfn "known: negative lookaround captures  /%s/ on %s" pattern (json subject)
                else report (sprintf "first engine %A perl %A" first perlFirst)
            elif all <> perlAll then report (sprintf "all engine %A perl %A" all perlAll)
printfn "%d cases: %d differ from perl (%d known PCRE2/perl differences)" lines.Length differ known
output |> Option.iter (fun path -> File.WriteAllLines(path, written))
