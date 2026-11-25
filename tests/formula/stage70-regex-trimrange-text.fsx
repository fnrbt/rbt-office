#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#load "function-table-kit.fsx"
open System
open System.IO
open System.Text.Json
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open ``Function-table-kit``
module Wb = Rbt.Spreadsheet.Workbook
module Recalc = Rbt.Spreadsheet.Recalc
module Bridge = Rbt.Spreadsheet.XlsxBridge
module Pcre = Rbt.Spreadsheet.Functions.Pcre

// REGEXTEST, REGEXEXTRACT, REGEXREPLACE (Functions/FnRegex.fs on the PCRE2-flavoured engine of
// Functions/Pcre.fs), TRIMRANGE and the trim references (Formula.fs), ASC, DBCS and PHONETIC
// (Functions/FnText.fs). The regular expression functions and TRIMRANGE are Excel 365's, which the
// parity VM (Excel 2019) does not have: their cases are every example of Microsoft's pages and of
// exceljet's, the rules of PCRE2, and the engine against perl on the cases of
// tools/excel-parity/make-regex-cases.py (regex-perl.jsonl). ASC, DBCS, JIS and PHONETIC are Excel
// 2019's: the VM's answers are in tools/excel-parity/widthconv-excel.tsv and phonetic-*.tsv.

let row (values: CellValue list) = [| Array.ofList values |]
let spills label m formula (expected: CellValue[][]) = eq (label + " " + formula) (arrayOf m formula) expected

// ===== REGEXTEST ======================================================================
// Microsoft's examples: "alfalfa" in A2 (here A1); the phone format.
let docs = modelOf [ (0, 0), S "alfalfa"; (1, 0), S "(378) 555-4195"; (2, 0), S "+1(878) 555-8622" ]
is "docs" docs "=REGEXTEST(A1,\"a\")" (B true)
is "docs" docs "=REGEXTEST(A1,\"[a-z]\")" (B true)
is "docs" docs "=REGEXTEST(A1,\"[A-Z]\")" (B false)
is "docs" docs "=REGEXTEST(A1,\"[aeiou]\")" (B true)
is "docs" docs "=REGEXTEST(A1,\"[0-9]\")" (B false)
is "docs" docs "=REGEXTEST(A2,\"^\\([0-9]{3}\\) [0-9]{3}-[0-9]{4}$\")" (B true)
is "docs" docs "=REGEXTEST(A3,\"^\\([0-9]{3}\\) [0-9]{3}-[0-9]{4}$\")" (B false)
// case_sensitivity: 0 (the default) case-sensitive, 1 not; anything else #VALUE!
is "case" empty "=REGEXTEST(\"ABC\",\"abc\")" (B false)
is "case" empty "=REGEXTEST(\"ABC\",\"abc\",0)" (B false)
is "case" empty "=REGEXTEST(\"ABC\",\"abc\",1)" (B true)
is "case" empty "=REGEXTEST(\"ABC\",\"abc\",TRUE)" (B true)
is "case" empty "=REGEXTEST(\"ABC\",\"abc\",1.9)" (B true)
error empty "=REGEXTEST(\"ABC\",\"abc\",2)" Err.value
error empty "=REGEXTEST(\"ABC\",\"abc\",-1)" Err.value
error empty "=REGEXTEST(\"ABC\",\"abc\",\"x\")" Err.value
is "case" empty "=REGEXTEST(\"ABC\",\"(?i)abc\")" (B true)
is "case" empty "=REGEXTEST(\"ÉCOLE\",\"école\",1)" (B true)
// the arguments: a number is its text, an error is the result, an empty pattern matches
is "text" empty "=REGEXTEST(12345,\"^\\d+$\")" (B true)
is "text" empty "=REGEXTEST(1.5,\"\\.\")" (B true)
is "text" empty "=REGEXTEST(TRUE,\"TRUE\")" (B true)
is "text" empty "=REGEXTEST(\"\",\"\")" (B true)
is "text" empty "=REGEXTEST(\"abc\",\"\")" (B true)
error empty "=REGEXTEST(#N/A,\"a\")" Err.na
error empty "=REGEXTEST(\"a\",#DIV/0!)" Err.div0
error empty "=REGEXTEST(\"a\")" Err.value
// patterns PCRE2 does not compile: #VALUE!
for bad in [ "("; ")"; "[a"; "a{2,1}"; "*a"; "+"; "a**"; "(?<n>a)(?<n>b)"; "\\"; "[z-a]"; "(?<=a+)b"; "\\k<nope>"
             "\\9"; "(?P<1a>x)"; "\\x{110000}"; "[[:nope:]]"; "\\p{Nope}"; "(?z)"; "(*SKIP)a"; "a(?#x"; "\\c"; "(?(1)a|b|c)"
             "\\i"; "[\\d-z]"; "x{70000}"; "\\g{0}"; "(?<=\\d+)" ] do
    error empty (sprintf "=REGEXTEST(\"a\",\"%s\")" (bad.Replace("\"", "\"\""))) Err.value
// arrays: lifted over the text and the pattern
let column = modelOf [ (0, 0), S "abc"; (1, 0), S "ABC"; (2, 0), S "123" ]
spills "lifted" column "=REGEXTEST(A1:A3,\"[a-z]\")" [| [| B true |]; [| B false |]; [| B false |] |]
spills "lifted" column "=REGEXTEST(A1:A3,\"[a-z]\",1)" [| [| B true |]; [| B true |]; [| B false |] |]
spills "lifted" empty "=REGEXTEST(\"a1\",{\"\\d\",\"[a-z]\",\"x\"})" (row [ B true; B true; B false ])
spills "lifted" empty "=REGEXTEST(\"aB\",\"b\",{0;1})" [| [| B false |]; [| B true |] |]

// ===== REGEXEXTRACT ===================================================================
let names = modelOf [ (0, 0), S "DylanWilliams"
                      (1, 0), S "Sonia Rees (378) 555-4195\nAngel Brown (878) 555-8622\nBlake Martin (437) 555-8987\nWilliam Kirby (619) 555-4212\nAvery Smith (579) 555-5658\nParker Jones (346) 555-0925\nLiam Small (405) 555-0887\nHollie Rees (666) 555-1872" ]
text names "=REGEXEXTRACT(A1,\"[A-Z][a-z]+\")" "Dylan"
// return mode 1: every match, spilled to the right (exceljet: "spills to the right")
spills "docs" names "=REGEXEXTRACT(A1,\"[A-Z][a-z]+\",1)" (row [ S "Dylan"; S "Williams" ])
spills "docs" names "=REGEXEXTRACT(A2,\"[0-9()]+ [0-9-]+\",1)"
    (row [ S "(378) 555-4195"; S "(878) 555-8622"; S "(437) 555-8987"; S "(619) 555-4212"; S "(579) 555-5658"; S "(346) 555-0925"; S "(405) 555-0887"; S "(666) 555-1872" ])
// return mode 2: the groups of the first match (an unset group ""), to the right
spills "groups" empty "=REGEXEXTRACT(\"2024-05-17 and 2023-12-31\",\"(\\d{4})-(\\d\\d)-(\\d\\d)\",2)" (row [ S "2024"; S "05"; S "17" ])
spills "groups" empty "=REGEXEXTRACT(\"ab\",\"(a)(x)?(b)\",2)" (row [ S "a"; S ""; S "b" ])
spills "groups" empty "=REGEXEXTRACT(\"John Smith\",\"(?<first>\\w+) (?<last>\\w+)\",2)" (row [ S "John"; S "Smith" ])
spills "groups" empty "=REGEXEXTRACT(\"ab\",\"(?:a)(b)\",2)" (row [ S "b" ])
// (no groups to give: #N/A - not documented)
error empty "=REGEXEXTRACT(\"ab\",\"ab\",2)" Err.na
// no match: #N/A, in every mode
error empty "=REGEXEXTRACT(\"abc\",\"\\d\")" Err.na
error empty "=REGEXEXTRACT(\"abc\",\"\\d\",1)" Err.na
error empty "=REGEXEXTRACT(\"abc\",\"(\\d)\",2)" Err.na
error empty "=REGEXEXTRACT(\"abc\",\"a\",3)" Err.value
error empty "=REGEXEXTRACT(\"abc\",\"a\",-1)" Err.value
error empty "=REGEXEXTRACT(\"abc\",\"[\")" Err.value
error empty "=REGEXEXTRACT(\"abc\",\"a\",0,2)" Err.value
text empty "=REGEXEXTRACT(\"ABC\",\"b\",0,1)" "B"
text empty "=REGEXEXTRACT(\"ABC\",\"b\",,1)" "B"
text empty "=REGEXEXTRACT(\"abc\",\"\")" ""
// what it gives is text (VALUE makes a number of it)
text empty "=REGEXEXTRACT(\"cost 42\",\"\\d+\")" "42"
number empty "=VALUE(REGEXEXTRACT(\"cost 42\",\"\\d+\"))" 42.0
text empty "=REGEXEXTRACT(12345,\"3.\")" "34"
// empty matches of mode 1 are matches
spills "empty matches" empty "=REGEXEXTRACT(\"ab\",\"x*\",1)" (row [ S ""; S ""; S "" ])
spills "empty matches" empty "=REGEXEXTRACT(\"axb\",\"x*\",1)" (row [ S ""; S "x"; S ""; S "" ])
// over an array of texts: each cell the first of its result
spills "lifted" column "=REGEXEXTRACT(A1:A3,\"[a-z0-9]\",0,1)" [| [| S "a" |]; [| S "A" |]; [| S "1" |] |]
spills "lifted" column "=REGEXEXTRACT(A1:A3,\"[a-z]\")" [| [| S "a" |]; [| E Err.na |]; [| E Err.na |] |]
spills "lifted" column "=REGEXEXTRACT(A1:A3,\".\",1)" [| [| S "a" |]; [| S "A" |]; [| S "1" |] |]
spills "lifted" empty "=REGEXEXTRACT({\"a1\",\"b2\"},\"\\d\")" (row [ S "1"; S "2" ])
spills "lifted" empty "=REGEXEXTRACT(\"a1b2\",\"\\d\",{0,1})" (row [ S "1"; S "1" ])
number empty "=COUNTA(REGEXEXTRACT(\"a1b22c333\",\"\\d+\",1))" 3.0
text empty "=INDEX(REGEXEXTRACT(\"a1b22c333\",\"\\d+\",1),3)" "333"
text empty "=TEXTJOIN(\"|\",,REGEXEXTRACT(\"a1b22c333\",\"\\d+\",1))" "1|22|333"
text empty "=REGEXEXTRACT(\"😀 smile\",\".\")" "😀"
number empty "=LEN(REGEXEXTRACT(\"x😀y\",\"..\"))" 3.0

// ===== REGEXREPLACE ===================================================================
let phones = modelOf [ (0, 0), S "Sonia Rees(378) 555-4195\nAngel Brown(878) 555-8622\nBlake Martin(437) 555-8987\nWilliam Kirby(619) 555-4212\nAvery Smith(579) 555-5658\nParker Jones(346) 555-0925\nLiam Small(405) 555-0887\nHollie Rees(666) 555-1872"
                       (1, 0), S "SoniaBrown"; (2, 0), S "0234568765" ]
text phones "=REGEXREPLACE(A1,\"[0-9]+-\",\"***-\")" "Sonia Rees(378) ***-4195\nAngel Brown(878) ***-8622\nBlake Martin(437) ***-8987\nWilliam Kirby(619) ***-4212\nAvery Smith(579) ***-5658\nParker Jones(346) ***-0925\nLiam Small(405) ***-0887\nHollie Rees(666) ***-1872"
text phones "=REGEXREPLACE(A2,\"([A-Z][a-z]+)([A-Z][a-z]+)\",\"$2, $1\")" "Brown, Sonia"
// exceljet's
text empty "=REGEXREPLACE(\"tuttle\",\"t\",\"b\")" "bubble"
text empty "=REGEXREPLACE(\"Tuttle\",\"t\",\"b\")" "Tubble"
text empty "=REGEXREPLACE(\"Tuttle\",\"t\",\"b\",,1)" "bubble"
text empty "=REGEXREPLACE(\"Tuttle\",\"[Tt]\",\"b\")" "bubble"
text empty "=REGEXREPLACE(\"(555) 123-4567 x8\",\"[^0-9]\",\"\")" "55512345678"
text empty "=REGEXREPLACE(\"hello world\",\"^(.)\",\"\\U$1\")" "Hello world"
text empty "=REGEXREPLACE(\"my cats and a dog, birds\",\"\\b((cat|dog|bird)s?)\\b\",\"\\U$1\",,1)" "my CATS and a DOG, BIRDS"
text phones "=REGEXREPLACE(A3,\"(\\d{3})(\\d{3})(\\d{4})\",\"($1)-$2-$3\")" "(023)-456-8765"
text empty "=REGEXREPLACE(REGEXREPLACE(\"023.456.8765\",\"[^0-9]\",\"\"),\"(\\d{3})(\\d{3})(\\d{4})\",\"($1)-$2-$3\")" "(023)-456-8765"
text empty "=REGEXREPLACE(\"abc\",\"x\",\"y\")" "abc"
// occurrence: 0 (the default) every match, n the n-th, -n the n-th from the end, past them the text as it is
text empty "=REGEXREPLACE(\"a-b-c-d\",\"-\",\"+\",0)" "a+b+c+d"
text empty "=REGEXREPLACE(\"a-b-c-d\",\"-\",\"+\",1)" "a+b-c-d"
text empty "=REGEXREPLACE(\"a-b-c-d\",\"-\",\"+\",2)" "a-b+c-d"
text empty "=REGEXREPLACE(\"a-b-c-d\",\"-\",\"+\",3)" "a-b-c+d"
text empty "=REGEXREPLACE(\"a-b-c-d\",\"-\",\"+\",4)" "a-b-c-d"
text empty "=REGEXREPLACE(\"a-b-c-d\",\"-\",\"+\",-1)" "a-b-c+d"
text empty "=REGEXREPLACE(\"a-b-c-d\",\"-\",\"+\",-3)" "a+b-c-d"
text empty "=REGEXREPLACE(\"a-b-c-d\",\"-\",\"+\",-4)" "a-b-c-d"
text empty "=REGEXREPLACE(\"a-b-c-d\",\"-\",\"+\",2.9)" "a-b+c-d"
error empty "=REGEXREPLACE(\"a-b\",\"-\",\"+\",\"x\")" Err.value
text empty "=REGEXREPLACE(\"tuttle\",\"T\",\"b\",-2,1)" "tubtle"
error empty "=REGEXREPLACE(\"a\",\"a\",\"b\",0,3)" Err.value
error empty "=REGEXREPLACE(\"a\",\"(\",\"b\")" Err.value
// the replacement: $n ${n} $name ${name} $0 $$, ${n:-default} ${n:+set:unset}, \U \L \u \l \E, \Q..\E, escapes
text empty "=REGEXREPLACE(\"John Smith\",\"(\\w+) (\\w+)\",\"${2}_$1\")" "Smith_John"
text empty "=REGEXREPLACE(\"John Smith\",\"(?<f>\\w+) (?<l>\\w+)\",\"${l}, ${f}\")" "Smith, John"
text empty "=REGEXREPLACE(\"John Smith\",\"(?<f>\\w+) (?<l>\\w+)\",\"$l $f\")" "Smith John"
text empty "=REGEXREPLACE(\"abc\",\"b\",\"[$0]\")" "a[b]c"
text empty "=REGEXREPLACE(\"5\",\"\\d\",\"$$$0\")" "$5"
text empty "=REGEXREPLACE(\"ab\",\"(a)(x)?\",\"[$2]\")" "[]b"
text empty "=REGEXREPLACE(\"ab\",\"(a)(x)?\",\"${2:-none}\")" "noneb"
text empty "=REGEXREPLACE(\"ab\",\"(a)(x)?\",\"${1:-none}\")" "ab"
text empty "=REGEXREPLACE(\"ab\",\"(a)(x)?\",\"${2:+yes:no}\")" "nob"
text empty "=REGEXREPLACE(\"ab\",\"(a)(x)?\",\"${1:+yes:no}\")" "yesb"
text empty "=REGEXREPLACE(\"hello world\",\"(\\w+)\",\"\\u$1\")" "Hello World"
text empty "=REGEXREPLACE(\"HELLO WORLD\",\"(\\w+)\",\"\\l$1\")" "hELLO wORLD"
text empty "=REGEXREPLACE(\"hello WORLD\",\"(\\w+) (\\w+)\",\"\\U$1\\E \\L$2\")" "HELLO world"
text empty "=REGEXREPLACE(\"hello\",\"(\\w+)\",\"\\u\\L$1\")" "Hello"
text empty "=REGEXREPLACE(\"ab\",\"b\",\"\\n\")" "a\n"
text empty "=REGEXREPLACE(\"ab\",\"b\",\"\\t\\x{41}\\\\\")" "a\tA\\"
text empty "=REGEXREPLACE(\"ab\",\"b\",\"\\Q$1\\E\")" "a$1"
text empty "=REGEXREPLACE(\"a.b\",\"\\.\",\"\\.\")" "a.b"
text empty "=REGEXREPLACE(\"café\",\"(.+)\",\"\\U$1\")" "CAFÉ"
error empty "=REGEXREPLACE(\"ab\",\"(a)\",\"$2\")" Err.value
error empty "=REGEXREPLACE(\"ab\",\"(a)\",\"${nope}\")" Err.value
error empty "=REGEXREPLACE(\"ab\",\"(a)\",\"\\1\")" Err.value
error empty "=REGEXREPLACE(\"ab\",\"(a)\",\"$\")" Err.value
error empty "=REGEXREPLACE(\"ab\",\"(a)\",\"\\q\")" Err.value
// a replacement is only read where there is a match
text empty "=REGEXREPLACE(\"ab\",\"x\",\"$9\")" "ab"
// empty matches: between every character (PCRE2's global substitution)
text empty "=REGEXREPLACE(\"abc\",\"\",\"-\")" "-a-b-c-"
text empty "=REGEXREPLACE(\"axb\",\"x*\",\"-\")" "-a--b-"
text empty "=REGEXREPLACE(\"a b\",\"\\b\",\"|\")" "|a| |b|"
text empty "=REGEXREPLACE(\"   padded  \",\"^\\s+|\\s+$\",\"\")" "padded"
text empty "=REGEXREPLACE(\"a  b     c\",\"\\s{2,}\",\" \")" "a b c"
text empty "=REGEXREPLACE(\"aaa\",\"a*?\",\"-\")" "-------"
text empty "=REGEXREPLACE(\"x😀y\",\".\",\"[$0]\")" "[x][😀][y]"
text empty "=REGEXREPLACE(\"Line1\nLine2\",\"(?m)^\",\"> \")" "> Line1\n> Line2"
spills "lifted" column "=REGEXREPLACE(A1:A3,\"[a-z]\",\"*\")" [| [| S "***" |]; [| S "ABC" |]; [| S "123" |] |]
spills "lifted" empty "=REGEXREPLACE(\"a-b-c\",\"-\",{\"+\",\"_\"})" (row [ S "a+b+c"; S "a_b_c" ])
spills "lifted" empty "=REGEXREPLACE(\"a-b-c\",\"-\",\"+\",{1;-1})" [| [| S "a+b-c" |]; [| S "a-b+c" |] |]

// ===== the engine =====================================================================
let matchOf (pattern: string) (subject: string) =
    let re = Pcre.compileRegex pattern false
    let cps = Pcre.codePoints subject
    Pcre.search re cps 0 |> Option.map (fun caps -> Pcre.textOf cps caps.[0] caps.[1])
eq "lazy" (matchOf "<.+?>" "<a><b>") (Some "<a>")
eq "possessive gives nothing back" (matchOf "a*+a" "aaaa") None
eq "atomic" (matchOf "(?>a+)b" "aaab") (Some "aaab")
eq "backreference" (matchOf "(\\w)\\1" "hello") (Some "ll")
eq "lookbehind of several lengths" (matchOf "(?<=ab|c)d" "xcd") (Some "d")
eq "conditional" (matchOf "(?<q>\")?\\w+(?(q)\")" "\"abc\"") (Some "\"abc\"")
eq "recursion" (matchOf "\\((?:[^()]|(?R))*\\)" "x(a(b)c)y") (Some "(a(b)c)")
eq "\\K" (matchOf "foo\\Kbar" "foobar") (Some "bar")
eq "\\w is ASCII" (matchOf "\\w+" "naïve") (Some "na")
eq "(*UCP) \\w is Unicode" (matchOf "(*UCP)\\w+" "naïve") (Some "naïve")
eq "\\p{Katakana} by script extensions" (matchOf "\\p{Katakana}+" "カタカナー") (Some "カタカナー")
eq "$ before a final newline" (matchOf "c$" "abc\n") (Some "c")
eq "dot and newline" (matchOf "a.b" "a\nb") None
eq "dot-all" (matchOf "(?s)a.b" "a\nb") (Some "a\nb")
eq "a supplementary character is one" (matchOf "^.$" "😀") (Some "😀")
// catastrophic backtracking runs into the match limit: #VALUE!, as PCRE2's error
error empty "=REGEXTEST(\"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaab\",\"^(a+)+$\")" Err.value
// and against perl, every case of regex-perl.jsonl (the known PCRE2 / perl differences apart)
let perlCases = Path.Combine(__SOURCE_DIRECTORY__, "..", "..", "tools", "excel-parity", "regex-perl.jsonl")
let knownPerl =
    set [ "\\Q.*\\E+"; "(?:(a)|(b))+"; "(?U)a+"; "(?U)a+?"; "straße"; "(?<q>\")?\\w+(?(q)\")"; "(a|b)c\\g<1>"
          "(?=a??\\d?\\w?).+?(?:\\b)[ab]*?"; "(x*|(?:a|b))(^$(c++)|((?<!b).(?<!b)[^a]??|$\\s{0,2}?)[^a])*+x" ]
let mutable perlCompared = 0
let mutable perlDiffer = 0
for line in File.ReadAllLines perlCases do
    let root = (JsonDocument.Parse line).RootElement
    let pattern = root.GetProperty("pattern").GetString()
    if not (knownPerl.Contains pattern) then
        let cps = Pcre.codePoints (root.GetProperty("subject").GetString())
        let engine =
            try
                let re = Pcre.compileRegex pattern (root.GetProperty("i").GetBoolean())
                Some (Pcre.search re cps 0 |> Option.map (fun c -> c.[0], c.[1]), Pcre.allMatches re cps |> List.map (fun c -> [ c.[0]; c.[1] ]))
            with Pcre.PatternError _ -> None
        let perl =
            if root.GetProperty("error").GetBoolean() then None
            else
                let first =
                    match root.GetProperty("first").ValueKind with
                    | JsonValueKind.Null -> None
                    | _ -> let g = [ for x in root.GetProperty("first").EnumerateArray() -> x.GetInt32() ] in Some (g.[0], g.[1])
                Some (first, [ for m in root.GetProperty("all").EnumerateArray() -> [ for x in m.EnumerateArray() -> x.GetInt32() ] ])
        perlCompared <- perlCompared + 1
        if engine <> perl then
            perlDiffer <- perlDiffer + 1
            if perlDiffer <= 5 then printfn "FAIL perl /%s/ engine %A perl %A" pattern engine perl
eq (sprintf "the engine as perl, on %d cases" perlCompared) perlDiffer 0

// ===== TRIMRANGE and the trim references ==============================================
// B3:D5 hold values (B3 1, C3 2, D4 "x", B5 5); A1:E10 has blank rows and columns at both ends.
let grid = modelOf [ (2, 1), N 1.0; (2, 2), N 2.0; (3, 3), S "x"; (4, 1), N 5.0 ]
let refOf (formula: string) =
    match value grid ("=CELL(\"address\"," + formula + ")"), value grid ("=ROWS(" + formula + ")"), value grid ("=COLUMNS(" + formula + ")") with
    | CellValue.String a, CellValue.Number r, CellValue.Number c -> sprintf "%s %gx%g" a r c
    | a, r, c -> sprintf "%A %A %A" a r c
eq "TRIMRANGE both ends" (refOf "TRIMRANGE(A1:E10)") "$B$3 3x3"
eq "TRIMRANGE 3,3" (refOf "TRIMRANGE(A1:E10,3,3)") "$B$3 3x3"
eq "TRIMRANGE leading rows" (refOf "TRIMRANGE(A1:E10,1,0)") "$A$3 8x5"
eq "TRIMRANGE trailing rows" (refOf "TRIMRANGE(A1:E10,2,0)") "$A$1 5x5"
eq "TRIMRANGE no rows, columns both" (refOf "TRIMRANGE(A1:E10,0)") "$B$1 10x3"
eq "TRIMRANGE leading columns" (refOf "TRIMRANGE(A1:E10,0,1)") "$B$1 10x4"
eq "TRIMRANGE trailing columns" (refOf "TRIMRANGE(A1:E10,0,2)") "$A$1 10x4"
eq "TRIMRANGE nothing" (refOf "TRIMRANGE(A1:E10,0,0)") "$A$1 10x5"
eq "TRIMRANGE omitted arguments are 3" (refOf "TRIMRANGE(A1:E10,,)") "$B$3 3x3"
eq "trim all A1.:.E10" (refOf "A1.:.E10") "$B$3 3x3"
eq "trim trailing A1:.E10 = TRIMRANGE(,2,2)" (refOf "A1:.E10") "$A$1 5x4"
eq "trim leading A1.:E10 = TRIMRANGE(,1,1)" (refOf "A1.:E10") "$B$3 8x4"
eq "a whole column, trailing" (refOf "B:.B") "$B$1 5x1"
eq "a whole column, both" (refOf "A.:.D") "$B$3 3x3"
eq "whole rows" (refOf "1:.10") "$A$1 5x4"
eq "whole rows, both" (refOf "1.:.10") "$B$3 3x3"
eq "a whole column kept whole" (refOf "TRIMRANGE(B:B,0,0)") "$B$1 1.04858e+06x1"
eq "absolute operands" (refOf "$A$1.:.$E$10") "$B$3 3x3"
eq "the file's form" (refOf "_xlfn._TRO_ALL(A1:E10)") "$B$3 3x3"
eq "a single cell" (refOf "TRIMRANGE(D1:D10)") "$D$4 1x1"
number grid "=SUM(TRIMRANGE(A1:E10))" 8.0
number grid "=SUM(A:.E)" 8.0
number grid "=COUNTA(A1.:.E10)" 4.0
number grid "=COUNTBLANK(A1.:.E10)" 5.0
is "ISREF" grid "=ISREF(TRIMRANGE(A1:E10))" (B true)
number grid "=ROWS(TRIMRANGE(A1:E10)) * COLUMNS(A1.:.E10)" 9.0
number grid "=ROWS(DROP(TRIMRANGE(B:B),1))" 2.0
spills "spill" grid "=TRIMRANGE(B1:C10)" [| [| N 1.0; N 2.0 |]; [| CellValue.Empty; CellValue.Empty |]; [| N 5.0; CellValue.Empty |] |]
spills "spill" grid "=LEN(B:.B)" [| [| N 0.0 |]; [| N 0.0 |]; [| N 1.0 |]; [| N 0.0 |]; [| N 1.0 |] |]
error grid "=TRIMRANGE(F1:G5)" Err.ref
error grid "=ROWS(TRIMRANGE(F1:G5))" Err.ref
error grid "=SUM(F:.G)" Err.ref
number grid "=ROWS(TRIMRANGE(F1:G5,0,0))" 5.0
error grid "=TRIMRANGE(A1:E10,4)" Err.value
error grid "=TRIMRANGE(A1:E10,1,-1)" Err.value
error grid "=TRIMRANGE(A1:E10,\"x\")" Err.value
error grid "=TRIMRANGE(#N/A)" Err.na
// an array: its empty values trimmed at the edges
spills "array" grid "=TRIMRANGE({1,2;3,4})" [| [| N 1.0; N 2.0 |]; [| N 3.0; N 4.0 |] |]
// a formula giving "" is not blank
let withEmptyText = modelOf [ (0, 0), N 1.0; (5, 0), CellValue.Formula ("=\"\"", S "") ]
number withEmptyText "=ROWS(A:.A)" 6.0
// the trim references parse to their function and print back
for formula in [ "A1:.E10"; "A1.:E10"; "A1.:.E10"; "Sheet2!A1.:.B5"; "'My Sheet'!A:.A"; "SUM($A$1:.$B$9)+1"; "1.:.10"; "A:.C"; "TRIMRANGE(A:A,2,0)" ] do
    match parseFormula ("=" + formula) with
    | Ok e -> eq ("prints back " + formula) (exprToFormula e) formula
    | Result.Error message -> eq ("parses " + formula) message ""
eq "parses to its function" (parseFormula "=A1.:E10" |> Result.map (function Function (name, [ RangeRef _ ]) -> name | _ -> "?")) (Ok "_TRO_LEADING")
eq "the file's name prints as the operator" (parseFormula "=_xlfn._TRO_TRAILING(A1:E10)" |> Result.map exprToFormula) (Ok "A1:.E10")
// in the file: TRIMRANGE calls (what Excel stores for the operators is not documented)
eq "file form" (FormulaNames.fileFormula "=SUM(A1.:.E10)") "=SUM(_xlfn.TRIMRANGE(A1:E10,3,3))"
eq "file form, sheet" (FormulaNames.fileFormula "='My S'!A:.A") "=_xlfn.TRIMRANGE('My S'!A:A,2,2)"
eq "file form, leading" (FormulaNames.fileFormula "=ROWS(Data!$A$1.:$C$9)") "=ROWS(_xlfn.TRIMRANGE(Data!$A$1:$C$9,1,1))"
eq "file form, text left alone" (FormulaNames.fileFormula "=\"A1.:B2\"&A1") "=\"A1.:B2\"&A1"
eq "file form, TRIMRANGE" (FormulaNames.fileFormula "=TRIMRANGE(A:A)") "=_xlfn.TRIMRANGE(A:A)"

// as the data grows, what reads a trim reference is calculated again
let cellOf (data: CellValue) = { Value = data; Format = defaultFormat; Comment = None }
let formulaCell (body: string) = cellOf (CellValue.Formula (body, CellValue.Empty))
let setCells (cells: (int * int * CellData) list) (wb: Wb.Workbook) =
    let mutable model = resolveSheet wb.Sheets.["sheet1"].Model
    for (r, c, data) in cells do model <- setCell { Row = r; Col = c } data model
    Wb.updateSheetModel "sheet1" model wb
let valueAtCell (wb: Wb.Workbook) r c =
    match getCell { Row = r; Col = c } (resolveSheet wb.Sheets.["sheet1"].Model) with
    | Some cell -> unwrapFormulaValue cell.Value
    | None -> CellValue.Empty
let book =
    Wb.createWorkbook ()
    |> setCells [ 0, 0, cellOf (N 1.0); 1, 0, cellOf (N 2.0); 2, 0, cellOf (N 3.0)
                  0, 2, formulaCell "=ROWS(A:.A)"; 1, 2, formulaCell "=SUM(TRIMRANGE(A1:A100))"; 0, 4, formulaCell "=A:.A*10" ]
let calculated, _ = Recalc.recalculateAll book
eq "ROWS(A:.A)" (valueAtCell calculated 0 2) (N 3.0)
eq "SUM(TRIMRANGE(A1:A100))" (valueAtCell calculated 1 2) (N 6.0)
eq "spilled A:.A*10" (valueAtCell calculated 2 4, valueAtCell calculated 3 4) (N 30.0, CellValue.Empty)
let engine = Recalc.build calculated
let grown = calculated |> setCells [ 9, 0, cellOf (N 10.0) ]
let incremental = Recalc.recalculate engine grown (Some [ "sheet1", { Row = 9; Col = 0 } ])
eq "the incremental recalculation takes the edit" (incremental.IsSome, Recalc.lastDecline engine) (true, "")
match incremental with
| Some after ->
    eq "grown: ROWS(A:.A)" (valueAtCell after 0 2) (N 10.0)
    eq "grown: SUM(TRIMRANGE(A1:A100))" (valueAtCell after 1 2) (N 16.0)
    eq "grown: the spill" (valueAtCell after 9 4) (N 100.0)
| None ->
    let after, _ = Recalc.recalculateAll grown
    eq "grown (full): ROWS(A:.A)" (valueAtCell after 0 2) (N 10.0)
    eq "grown (full): the spill" (valueAtCell after 9 4) (N 100.0)
let fullyGrown, _ = Recalc.recalculateAll grown
eq "grown, full recalculation" (valueAtCell fullyGrown 0 2, valueAtCell fullyGrown 1 2, valueAtCell fullyGrown 9 4) (N 10.0, N 16.0, N 100.0)

// ===== ASC, DBCS, JIS, PHONETIC =======================================================
// Excel 2019 on the VM (English, no East Asian editing language: phonetic-vm-language.tsv): ASC and
// DBCS give every text as it is (widthconv-excel.tsv: every code point of the blocks they would
// touch), JIS is no function (#NAME?), PHONETIC is #N/A for any reference - furigana in the file
// (synthetic-phonetic.xlsx, phonetic-excel.tsv) or not (phonetic-cells-excel.tsv).
let widthCases = File.ReadAllLines(Path.Combine(__SOURCE_DIRECTORY__, "..", "..", "tools", "excel-parity", "widthconv-cells.txt"))
let widthExcel = File.ReadAllLines(Path.Combine(__SOURCE_DIRECTORY__, "..", "..", "tools", "excel-parity", "widthconv-excel.tsv"))
let mutable widthDiffer = 0
for (case, excel) in Array.zip widthCases widthExcel do
    let formula = case.Split('\t').[2]
    let parts = excel.Split('\t')
    let expected =
        match parts.[2] with
        | "s" -> S parts.[3]
        | "e" -> E (parts.[3].Substring 1)
        | "b" -> B (parts.[3] = "True")
        | "n" -> N (float parts.[4])
        | _ -> CellValue.Empty
    // (column A and B of the probe sheet are empty)
    let got = value empty formula
    if got <> expected then
        widthDiffer <- widthDiffer + 1
        if widthDiffer <= 5 then printfn "FAIL VM %s expected %A got %A" formula expected got
eq (sprintf "as Excel 2019 on the VM, %d cases" widthCases.Length) widthDiffer 0
error empty "=JIS(\"ABC\")" Err.name
text empty "=ASC(\"ＡＢＣ\")" "ＡＢＣ"
text empty "=DBCS(\"ABC\")" "ABC"
let probeCells = modelOf [ (0, 0), S "abc"; (1, 0), N 123.0; (5, 0), S "漢字" ]
error probeCells "=PHONETIC(A1)" Err.na
error probeCells "=PHONETIC(A1:A6)" Err.na
error probeCells "=PHONETIC(A7)" Err.na
error probeCells "=PHONETIC((A1,A2))" Err.na
error probeCells "=PHONETIC(A7:A9 A1:A6)" Err.nul
let phoneticBook =
    let bytes = File.ReadAllBytes(Path.Combine(__SOURCE_DIRECTORY__, "..", "test-documents-excel", "synthetic-phonetic.xlsx"))
    Bridge.Import.importWorkbook (XlsxReader.loadBytesNative bytes |> Option.get)
let phoneticSheet = resolveSheet phoneticBook.Sheets.[0].Model
eq "the reader keeps the furigana" (Map.tryFind { Row = 0; Col = 0 } phoneticSheet.Phonetics) (Some { Text = "漢字"; Reading = "カンジ"; Kind = "" })
eq "a reading of part of the text" (Map.tryFind { Row = 2; Col = 0 } phoneticSheet.Phonetics |> Option.map (fun p -> p.Reading)) (Some "ヤマダさん")
eq "no reading, no entry" (Map.tryFind { Row = 3; Col = 0 } phoneticSheet.Phonetics) None
eq "an inline string's reading is not its text" (getCell { Row = 13; Col = 0 } phoneticSheet |> Option.map (fun c -> c.Value)) (Some (S "漢字"))
let phoneticOf (formula: string) = value phoneticSheet formula
let phoneticExcel = File.ReadAllLines(Path.Combine(__SOURCE_DIRECTORY__, "..", "..", "tools", "excel-parity", "phonetic-excel.tsv"))
for line in phoneticExcel do
    let parts = line.Split('\t')
    if parts.Length > 3 && parts.[1].StartsWith "B" then
        match getCell { Row = int (parts.[1].Substring 1) - 1; Col = 1 } phoneticSheet with
        | Some { Value = CellValue.Formula (formula, _) } -> eq ("VM " + formula) (phoneticOf formula) (E (parts.[3].Substring 1))
        | _ -> ()

// With a double-byte editing language (Japanese): what Microsoft documents, the conversions of
// Unicode's width forms (as Windows' LCMapString) - not checked against such an Excel.
setDoubleByteLanguage true
text empty "=ASC(\"ＡＢＣ１２３\")" "ABC123"
text empty "=ASC(\"Ｅｘｃｅｌ　２０１９\")" "Excel 2019"
text empty "=ASC(\"アイウ\")" "ｱｲｳ"
text empty "=ASC(\"ガギグパ\")" "ｶﾞｷﾞｸﾞﾊﾟ"
text empty "=ASC(\"ヴァイオリン\")" "ｳﾞｧｲｵﾘﾝ"
text empty "=ASC(\"「データ」、。・ー\")" "｢ﾃﾞｰﾀ｣､｡･ｰ"
text empty "=ASC(\"ひらがな漢字\")" "ひらがな漢字"
text empty "=ASC(\"ヰヱヵヶ\")" "ヰヱヵヶ"
text empty "=ASC(\"￥１００\")" "¥100"
text empty "=DBCS(\"ABC123\")" "ＡＢＣ１２３"
text empty "=DBCS(\"Excel 2019\")" "Ｅｘｃｅｌ　２０１９"
text empty "=DBCS(\"ｶﾞｷﾞ\")" "ガギ"
text empty "=DBCS(\"ﾊﾟﾋﾟﾌﾟ\")" "パピプ"
text empty "=DBCS(\"ｳﾞｧｲｵﾘﾝ\")" "ヴァイオリン"
text empty "=DBCS(\"ｱﾞ\")" "ア゛"
text empty "=DBCS(\"ﾞ\")" "゛"
text empty "=DBCS(\"｢ﾃﾞｰﾀ｣､｡\")" "「データ」、。"
text empty "=DBCS(\"~\\'\"\"\")" "～＼＇＂"
text empty "=DBCS(123)" "１２３"
text empty "=ASC(DBCS(\"Hello, World! ｶﾞ\"))" "Hello, World! ｶﾞ"
spills "lifted" empty "=ASC({\"Ａ\",\"Ｂ\"})" (row [ S "A"; S "B" ])
error empty "=ASC(#N/A)" Err.na
// PHONETIC: the furigana; a cell without any its text; a range all its cells' (numbers left out)
let ph (formula: string) (expected: CellValue) = eq ("PHONETIC " + formula) (phoneticOf formula) expected
ph "=PHONETIC(A1)" (S "カンジ")
ph "=PHONETIC(A2)" (S "トウキョウトミナトク")
ph "=PHONETIC(A3)" (S "ヤマダサン")
ph "=PHONETIC(A4)" (S "abc")
ph "=PHONETIC(A5)" (S "plain")
ph "=PHONETIC(A6)" (S "にほんご")
ph "=PHONETIC(A7)" (S "かんじ")
ph "=PHONETIC(A8)" (S "ワタシノナマエ")
ph "=PHONETIC(A9)" (S "ﾊﾝｶｸ")
ph "=PHONETIC(A10)" (S "")
ph "=PHONETIC(A12)" (S "")
ph "=PHONETIC(A13)" (S "漢字")
ph "=PHONETIC(A1:A2)" (S "カンジトウキョウトミナトク")
ph "=PHONETIC(A4:A5)" (S "abcplain")
ph "=PHONETIC(A10:A11)" (S "")
ph "=PHONETIC(A3:A1)" (S "カンジトウキョウトミナトクヤマダサン")
ph "=PHONETIC((A1,A2))" (E Err.na)
// an edited cell's text is no longer the one its furigana was of
let editedSheet = setCell { Row = 0; Col = 0 } { Value = S "感じ"; Format = defaultFormat; Comment = None } phoneticSheet
eq "edited" (value editedSheet "=PHONETIC(A1)") (S "感じ")
setDoubleByteLanguage false

printfn "stage70 regex, trimrange, text: %d passed, %d failed" pass fail
if fail > 0 then exit 1
