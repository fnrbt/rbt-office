Attribute VB_Name = "HighlightList"
' Name: Highlight words from a word list
' Source: https://word.tips.net/T000502_Highlight_Words_from_a_Word_List (Allen Wyatt, "CompareWordList":
'         a Replace All with "^&" per word of a list, whole words, case-sensitive)
' Author: written for this showcase after that tip's approach
' Licence: the page is "Copyright Sharon Parq Associates" with no licence for the code, so it is not
'          copied; this module is the repository's (MIT)
' What it does: three lists of words, each in its own highlight colour (Options.DefaultHighlightColorIndex,
'               Replacement.Highlight, Replacement.Text = "^&", MatchWholeWord, wdReplaceAll), and a
'               count of each word's hits (Find.Execute in a loop) written at the end.
' Adaptations: the word lists are arrays in the macro instead of a checklist document opened from disk.
' Needs: text (Setup types a paragraph about programming languages)
' Entry point: Main
Sub Setup()
    ActiveDocument.Content.Text = "Programming languages" & vbCr & _
        "Python and JavaScript are the most popular languages today, while C and C++ still run " & _
        "the operating systems underneath them. Java and C# power the enterprise; Go and Rust are the " & _
        "newer systems languages. F# and OCaml bring functional programming to the mainstream, Haskell " & _
        "keeps it pure, and VBA still automates Word and Excel on millions of desks. Python, Rust and Go " & _
        "grow fastest; Java, Python and JavaScript have the most jobs."
    ActiveDocument.Paragraphs(1).Style = wdStyleHeading1
    ActiveDocument.Paragraphs(2).Range.Font.Size = 14
    ActiveDocument.Paragraphs(2).LineSpacingRule = wdLineSpace1pt5
End Sub

Sub HighlightWords(words As Variant, colorIndex As Integer)
    Dim w
    Options.DefaultHighlightColorIndex = colorIndex
    For Each w In words
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = w
            .Replacement.Text = "^&"
            .Replacement.Highlight = True
            .MatchWholeWord = True
            .MatchCase = True
            .Format = True
            .Execute Replace:=wdReplaceAll
        End With
    Next w
End Sub

Function CountOf(word As String) As Integer
    Dim r As Range, n As Integer
    Set r = ActiveDocument.Paragraphs(2).Range
    With r.Find
        .Text = word
        .MatchWholeWord = True
        .MatchCase = True
        Do While .Execute
            n = n + 1
            r.Collapse wdCollapseEnd
        Loop
    End With
    CountOf = n
End Function

Sub Main()
    Dim summary As String, w
    Setup
    HighlightWords Array("Python", "JavaScript", "Java"), wdYellow
    HighlightWords Array("Rust", "Go", "C", "C++"), wdBrightGreen
    HighlightWords Array("F#", "OCaml", "Haskell", "VBA"), wdTurquoise
    For Each w In Array("Python", "Java", "Rust", "Go")
        summary = summary & w & " " & CountOf(CStr(w)) & "   "
    Next w
    ActiveDocument.Content.InsertParagraphAfter
    ActiveDocument.Content.InsertAfter "Mentions: " & summary
    With ActiveDocument.Paragraphs.Last.Range.Font
        .Italic = True
        .Color = RGB(89, 89, 89)
    End With
End Sub
