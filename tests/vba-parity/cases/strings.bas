Sub Main()
    P "len", Len("hello")
    P "mid", Mid("abcdef", 2, 3)
    P "mid no length", Mid("abcdef", 4)
    P "mid past end", Mid("abc", 5)
    P "left", Left("abc", 10)
    P "instr", InStr("hello world", "o")
    P "instr start", InStr(6, "hello world", "o")
    P "instr text compare", InStr(1, "HELLO", "l", vbTextCompare)
    P "instr none", InStr("abc", "z")
    P "instrrev", InStrRev("a,b,c", ",")
    P "instrrev start", InStrRev("a,b,c", ",", 3)
    P "replace", Replace("aaaa", "a", "b")
    P "replace start count", Replace("aaaa", "a", "b", 2, 2)
    P "replace text", Replace("ABab", "a", "x", , , vbTextCompare)
    P "split ubound", UBound(Split("a,b,c", ","))
    P "split limit", Join(Split("a,b,c,d", ",", 2), "|")
    P "split empty", UBound(Split("", ","))
    P "join", Join(Array(1, 2, 3), "-")
    P "strconv proper", StrConv("hELLO wORLD", vbProperCase)
    P "ucase", UCase("mIxEd")
    P "trim", "[" & Trim("  x  ") & "]"
    P "ltrim", "[" & LTrim("  x  ") & "]"
    P "string", String(3, "ab")
    P "space", "[" & Space(2) & "]"
    P "strcomp", StrComp("a", "B")
    P "strcomp text", StrComp("a", "A", vbTextCompare)
    P "like", "abc" Like "a*"
    P "like class", "b7" Like "[a-c]#"
    P "like not", "x" Like "[!a-c]"
    P "asc", Asc("A")
    P "chr", Chr(65) & Chr(97)
    P "chrw", AscW(ChrW(8364))
    P "strreverse", StrReverse("abc")
    P "concat number", "a" & 1.5
    P "plus strings", "1" + "2"
    P "plus mixed", "1" + 2
    P "compare string number", "10" > 9
    P "lcase$", LCase$("ABC")
    P "format @", Format("abc", "@@@@@")
    P "format >", Format("abc", ">")
    P "space in like", "a b" Like "a?b"
End Sub
