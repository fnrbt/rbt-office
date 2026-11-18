#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Vb.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"

// Word's object model for macros (WordMacroHost) on the native path: macros run on a copy of a document
// and give back the document they leave; text, character and paragraph formatting, styles, Find and
// Replace, tables, Selection, bookmarks, errors; and the showcase macros of tests/macros/word-showcase.

open System
open Rbt.Word.WordMacroApi
module Doc = Rbt.Word.Document
module TB = Rbt.Word.TextBuffer
module Project = Rbt.Word.WordMacroProject

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-40s expected %A got %A" label expected got

let output = ResizeArray<string>()
let hooks = { defaultHooks with Output = output.Add; MessageBox = (fun prompt _ _ -> output.Add ("MsgBox " + prompt); 1) }

let paragraph (text: string) : Doc.ParagraphBlock = { (Doc.newEmptyParagraph ()) with Buffer = TB.ofString text }
let docOf (texts: string list) =
    let paras = texts |> List.map paragraph
    { Doc.empty with Blocks = paras |> List.map Doc.Paragraph } |> Doc.setCursor paras.Head.Id 0

/// Run Sub `name` of `code` on `doc`: the document it leaves, or the failure's text.
let runOn (doc: Doc.Document) (code: string) (name: string) =
    output.Clear()
    let project = Project.ofModules [ "Module1", code ]
    match project.Problems with
    | p :: _ -> Error p.Text
    | [] ->
        match project.Run(doc, hooks, name) with
        | MacroDone d -> Ok d
        | MacroFailed f -> Error f.Text
let run doc code = runOn doc code "Main"
let ok label result =
    match result with
    | Ok d -> d
    | Error text -> fail <- fail + 1; printfn "FAIL %-40s %s" label text; Doc.empty

let paras (doc: Doc.Document) = Doc.allParagraphs doc
let texts doc = paras doc |> List.map (fun p -> TB.toString p.Buffer)
let styleAt (doc: Doc.Document) (index: int) (offset: int) = TB.getStyleAt offset (List.item index (paras doc)).Buffer
let bodyTexts (doc: Doc.Document) =
    doc.Blocks |> List.choose (function Doc.Paragraph p | Doc.Heading (_, p) -> Some (TB.toString p.Buffer) | _ -> None)

// ---- text: Content, InsertAfter, InsertParagraphAfter, Range.Text -----------------------------------------
let d1 =
    run (docOf [ "" ]) """
Sub Main()
    ActiveDocument.Content.InsertAfter "Hello"
    ActiveDocument.Content.InsertParagraphAfter
    ActiveDocument.Content.InsertAfter "World"
End Sub""" |> ok "insertafter"
check "insertafter: paragraphs" (texts d1) [ "Hello"; "World" ]

let d2 =
    run (docOf [ "one two three" ]) """
Sub Main()
    Dim r As Range
    Set r = ActiveDocument.Range(4, 7)
    Debug.Print r.Text
    r.Text = "2"
    Debug.Print r.Start & "-" & r.End
    Debug.Print ActiveDocument.Content.Text
    Debug.Print Len(ActiveDocument.Content.Text) & " " & ActiveDocument.Content.End
End Sub""" |> ok "range text"
check "range text: replaced" (texts d2) [ "one 2 three" ]
check "range text: output" (List.ofSeq output) [ "two"; "4-5"; "one 2 three\r"; "12 12" ]

let d3 =
    run (docOf [ "abc" ]) """
Sub Main()
    ActiveDocument.Content.Text = "First" & vbCr & "Second" & vbCrLf & "Third"
    Debug.Print ActiveDocument.Paragraphs.Count
    Debug.Print ActiveDocument.Paragraphs(2).Range.Text
    ActiveDocument.Paragraphs(1).Range.InsertBefore "1. "
End Sub""" |> ok "content text"
check "content text: paragraphs" (texts d3) [ "1. First"; "Second"; "Third" ]
check "content text: output" (List.ofSeq output) [ " 3 "; "Second\r" ]

// TypeName of the globals, as Word gives them; an unknown member names the type.
run (docOf [ "x" ]) """
Sub Main()
    Debug.Print TypeName(ActiveDocument); " "; TypeName(ThisDocument); " "; TypeName(ActiveWindow); " "; TypeName(ActiveDocument.Paragraphs(1))
    On Error Resume Next
    ActiveDocument.Nope = 1
    Debug.Print Err.Number; Err.Description
End Sub""" |> ok "type names" |> ignore
check "type names" (List.ofSeq output) [ "Document Document Window Paragraph"; " 438 Object doesn't support this property or method (Document.Nope)" ]

// ---- character formatting -------------------------------------------------------------------------------------
let d4 =
    run (docOf [ "Hello world" ]) """
Sub Main()
    With ActiveDocument.Range(0, 5).Font
        .Bold = True
        .Size = 20
        .Name = "Arial"
        .Color = RGB(255, 0, 0)
    End With
    ActiveDocument.Range(6, 11).Font.ColorIndex = wdBlue
    ActiveDocument.Range(6, 11).Italic = True
    ActiveDocument.Range(6, 11).HighlightColorIndex = wdYellow
    Debug.Print ActiveDocument.Range(0, 5).Font.Bold & " " & ActiveDocument.Content.Font.Bold & " " & ActiveDocument.Range(0, 2).Font.Size
End Sub""" |> ok "font"
let s4 = styleAt d4 0 1
check "font: bold" s4.Bold true
// (The model's sizes are layout pixels: 20 pt is 26.67 px.)
check "font: size" (System.Math.Round(s4.FontSize * 0.75, 2)) 20.0
check "font: name" s4.FontFamily (Some "Arial")
check "font: color" s4.Color (Some "#FF0000")
let s4b = styleAt d4 0 7
check "font: colorindex" s4b.Color (Some "#0000FF")
check "font: italic" s4b.Italic true
check "font: highlight" s4b.Highlight (Some "yellow")
check "font: space not bold" (styleAt d4 0 5).Bold false
check "font: reads" (List.ofSeq output) [ "-1 9999999 20" ]

// Rainbow: a colour per character with For Each over Characters.
let d5 =
    run (docOf [ "Rainbow" ]) """
Sub Main()
    Dim c As Range, i As Integer
    Dim colors
    colors = Array(wdColorRed, wdColorOrange, wdColorYellow, wdColorGreen, wdColorBlue, wdColorIndigo, wdColorViolet)
    For Each c In ActiveDocument.Paragraphs(1).Range.Characters
        c.Font.Color = colors(i Mod 7)
        i = i + 1
    Next c
    Debug.Print i
End Sub""" |> ok "rainbow"
check "rainbow: characters (with the mark)" (List.ofSeq output) [ " 8 " ]
check "rainbow: colours" ([ 0 .. 6 ] |> List.map (fun i -> (styleAt d5 0 i).Color)) [ Some "#FF0000"; Some "#FF6600"; Some "#FFFF00"; Some "#008000"; Some "#0000FF"; Some "#333399"; Some "#800080" ]

// ---- paragraphs, styles -------------------------------------------------------------------------------------
let d6 =
    run (docOf [ "Title"; "Body text"; "More" ]) """
Sub Main()
    ActiveDocument.Paragraphs(1).Style = wdStyleHeading1
    ActiveDocument.Paragraphs(1).Alignment = wdAlignParagraphCenter
    With ActiveDocument.Paragraphs(2).Format
        .SpaceAfter = 12
        .LeftIndent = 36
    End With
    ActiveDocument.Paragraphs(3).Range.Style = "Quote"
    Dim p As Paragraph
    For Each p In ActiveDocument.Paragraphs
        Debug.Print p.Range.Style & "|" & p.Alignment
    Next
End Sub""" |> ok "styles"
check "styles: heading block" (match d6.Blocks.Head with Doc.Heading (1, p) -> p.Metadata.Style | _ -> None) (Some "Heading1")
check "styles: centre" ((paras d6).Head.Metadata.Alignment) Doc.Center
check "styles: space after" ((paras d6).[1].Metadata.SpacingAfter) (Some 12.0)
check "styles: indent px" ((paras d6).[1].Metadata.IndentLeft) (Some 48.0)
check "styles: quote" ((paras d6).[2].Metadata.Style) (Some "Quote")
check "styles: defined" (d6.Styles |> List.map (fun s -> s.Id) |> List.sort) [ "Heading1"; "Quote" ]
check "styles: reads" (List.ofSeq output) [ "Heading 1|1"; "Normal|0"; "Quote|1" ]

// TypeParagraph after a heading gives the next style (Normal), as in Word.
let d7 =
    run (docOf [ "" ]) """
Sub Main()
    Selection.Style = ActiveDocument.Styles("Heading 2")
    Selection.TypeText "Chapter"
    Selection.TypeParagraph
    Selection.TypeText "Text"
    Selection.Font.Bold = True
    Selection.TypeText " bold"
    Selection.Font.Bold = False
    Selection.TypeText " plain"
End Sub""" |> ok "selection typing"
check "typing: texts" (texts d7) [ "Chapter"; "Text bold plain" ]
check "typing: heading" (match d7.Blocks.Head with Doc.Heading (2, _) -> true | _ -> false) true
check "typing: next style normal" ((paras d7).[1].Metadata.Style) None
check "typing: bold run" ((styleAt d7 1 6).Bold, (styleAt d7 1 11).Bold, (styleAt d7 1 1).Bold) (true, false, false)

// ---- Find and Replace ---------------------------------------------------------------------------------------------
let d8 =
    run (docOf [ "The cat sat on the mat."; "Another cat." ]) """
Sub Main()
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "cat"
        .Replacement.Text = "dog"
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    With ActiveDocument.Content.Find
        .Text = "the"
        .MatchCase = True
        .Replacement.Text = "a"
        .Execute Replace:=wdReplaceAll
    End With
End Sub""" |> ok "replace all"
// (Without MatchWholeWord "the" is found in "Another" too, as in Word.)
check "replace all: text" (texts d8) [ "The dog sat on a mat."; "Anoar dog." ]

let d9 =
    run (docOf [ "red green blue red" ]) """
Sub Main()
    Dim r As Range, n As Integer
    Set r = ActiveDocument.Content
    With r.Find
        .Text = "red"
        Do While .Execute
            n = n + 1
            r.Font.Bold = True
            r.Collapse wdCollapseEnd
        Loop
    End With
    Debug.Print n
    Options.DefaultHighlightColorIndex = wdBrightGreen
    With ActiveDocument.Content.Find
        .Text = "green"
        .Replacement.Text = ""
        .Replacement.Highlight = True
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
    With ActiveDocument.Content.Find
        .Text = "blue"
        .Replacement.Text = "^&!"
        .Replacement.Font.Italic = True
        .Execute Replace:=wdReplaceAll, Format:=True
    End With
End Sub""" |> ok "find loop"
check "find loop: count" (List.ofSeq output) [ " 2 " ]
check "find loop: text" (texts d9) [ "red green blue! red" ]
check "find loop: bold" ((styleAt d9 0 1).Bold, (styleAt d9 0 5).Bold, (styleAt d9 0 17).Bold) (true, false, true)
check "find loop: highlight" ((styleAt d9 0 5).Highlight) (Some "green")
check "find loop: italic" ((styleAt d9 0 11).Italic) true

let d10 =
    run (docOf [ "Call 555-1234 or 555-9876 now" ]) """
Sub Main()
    With ActiveDocument.Content.Find
        .Text = "([0-9]{3})-([0-9]{4})"
        .MatchWildcards = True
        .Replacement.Text = "\2-\1"
        .Execute Replace:=wdReplaceAll
    End With
    With ActiveDocument.Content.Find
        .MatchWildcards = False
        .Text = "now"
        .Replacement.Text = "^pnow"
        .Execute Replace:=wdReplaceAll
    End With
End Sub""" |> ok "wildcards"
check "wildcards: text" (texts d10) [ "Call 1234-555 or 9876-555 "; "now" ]

// ---- tables ---------------------------------------------------------------------------------------------------
let d11 =
    run (docOf [ "Before"; "" ]) """
Sub Main()
    Dim t As Table, r As Integer, c As Integer
    Set t = ActiveDocument.Tables.Add(ActiveDocument.Paragraphs(2).Range, 3, 3)
    For r = 1 To 3
        For c = 1 To 3
            t.Cell(r, c).Range.Text = r * c
        Next
    Next
    t.Rows(1).Shading.BackgroundPatternColor = wdColorGray25
    t.Cell(2, 2).Shading.BackgroundPatternColor = RGB(0, 0, 255)
    t.Borders.Enable = True
    t.Rows(1).Range.Font.Bold = True
    Debug.Print t.Rows.Count & " " & t.Columns.Count & " " & ActiveDocument.Tables.Count
    Debug.Print Len(t.Cell(1, 1).Range.Text) & " " & (t.Cell(3, 3).Range.Text = "9" & vbCr & Chr(7))
    t.Rows.Add
    Debug.Print CStr(t.Rows.Count)
    Debug.Print CStr(ActiveDocument.Paragraphs.Count)
End Sub""" |> ok "tables"
let table11 = d11.Blocks |> List.pick (function Doc.Table t -> Some t | _ -> None)
let cellText (t: Doc.TableBlock) r c = t.Rows.[r].Cells.[c].Content |> List.map (fun p -> TB.toString p.Buffer) |> String.concat "\r"
check "tables: placed" (d11.Blocks |> List.map (function Doc.Table _ -> "T" | _ -> "P")) [ "P"; "T"; "P" ]
check "tables: cells" ([ for r in 0 .. 2 -> [ for c in 0 .. 2 -> cellText table11 r c ] ]) [ [ "1"; "2"; "3" ]; [ "2"; "4"; "6" ]; [ "3"; "6"; "9" ] ]
check "tables: row shading" (table11.Rows.[0].Cells |> List.map (fun c -> c.BackgroundColor)) [ Some "#C0C0C0"; Some "#C0C0C0"; Some "#C0C0C0" ]
check "tables: cell shading" (table11.Rows.[1].Cells.[1].BackgroundColor) (Some "#0000FF")
check "tables: header bold" ((TB.getStyleAt 0 table11.Rows.[0].Cells.[0].Content.Head.Buffer).Bold) true
// (Word counts each end-of-row mark as a paragraph: 1 + 4 x (3 + 1) + 1.)
check "tables: output" (List.ofSeq output) [ "3 3 1"; "3 True"; "4"; "18" ]
check "tables: added row" table11.Rows.Length 4

// ---- Selection movement --------------------------------------------------------------------------------------------
let d12 =
    run (docOf [ "abc"; "def" ]) """
Sub Main()
    Selection.EndKey Unit:=wdStory
    Selection.TypeText "!"
    Selection.HomeKey Unit:=wdStory
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Underline = wdUnderlineDouble
    Debug.Print Selection.Text & " " & Selection.Start & " " & Selection.End
    Selection.MoveDown Unit:=wdParagraph, Count:=1
    Selection.TypeText "<"
    Debug.Print Selection.Start
End Sub""" |> ok "selection"
check "selection: text" (texts d12) [ "abc"; "<def!" ]
check "selection: underline" ((styleAt d12 0 0).UnderlineStyle, (styleAt d12 0 0).Underline) (Some "double", true)
check "selection: output" (List.ofSeq output) [ "a 0 1"; " 5 " ]

// ---- words, sentences, statistics ------------------------------------------------------------------------------
let d13 =
    run (docOf [ "Hello, big world. It's here!" ]) """
Sub Main()
    Debug.Print ActiveDocument.Words.Count
    Debug.Print ActiveDocument.Words(1).Text & "|" & ActiveDocument.Words(2).Text & "|" & ActiveDocument.Words(6).Text
    Debug.Print ActiveDocument.Sentences.Count & "|" & ActiveDocument.Sentences(2).Text
    Debug.Print ActiveDocument.ComputeStatistics(wdStatisticWords)
    Dim w
    For Each w In ActiveDocument.Words
        If Len(Trim(w.Text)) > 3 Then w.Font.Bold = True
    Next
    ActiveDocument.Sentences(1).Case = wdUpperCase
End Sub""" |> ok "words"
check "words: output" (List.ofSeq output) [ " 9 "; "Hello|, |It's "; "2|It's here!\r"; " 5 " ]
check "words: bold" ((styleAt d13 0 0).Bold, (styleAt d13 0 7).Bold, (styleAt d13 0 11).Bold) (true, false, true)
check "words: case" (texts d13) [ "HELLO, BIG WORLD. It's here!" ]

// ---- bookmarks -----------------------------------------------------------------------------------------------
let d14 =
    run (docOf [ "Dear NAME, thanks." ]) """
Sub Main()
    ActiveDocument.Bookmarks.Add "Name", ActiveDocument.Range(5, 9)
    Debug.Print ActiveDocument.Bookmarks.Exists("Name") & " " & ActiveDocument.Bookmarks("Name").Range.Text
    ActiveDocument.Bookmarks("Name").Range.Text = "Ada"
End Sub""" |> ok "bookmarks"
check "bookmarks: text" (texts d14) [ "Dear Ada, thanks." ]
check "bookmarks: output" (List.ofSeq output) [ "True NAME" ]

// ---- errors ---------------------------------------------------------------------------------------------------
let e1 = run (docOf [ "x" ]) "Sub Main()\n    ActiveDocument.Content.Frobnicate\nEnd Sub"
check "errors: 438" (match e1 with Error t -> t.Contains "438" && t.Contains "Frobnicate" | Ok _ -> false) true
let e2 = run (docOf [ "x" ]) "Sub Main()\n    ActiveDocument.Content.Text = \"changed\"\n    Err.Raise 5\nEnd Sub"
check "errors: nothing kept" (match e2 with Error _ -> true | Ok _ -> false) true
let e3 =
    run (docOf [ "x" ]) """
Sub Main()
    On Error Resume Next
    Dim t
    Set t = ActiveDocument.Tables(1)
    Debug.Print Err.Number
End Sub""" |> ok "on error"
check "errors: 5941" (List.ofSeq output) [ " 5941 " ]

// ---- Documents.Add, ConvertToTable, lists ------------------------------------------------------------------------
let d15 =
    run (docOf [ "old" ]) """
Sub Main()
    Dim doc As Document
    Set doc = Documents.Add
    doc.Content.InsertAfter "Name,Score" & vbCr & "Ada,90" & vbCr & "Bob,85"
    doc.Content.ConvertToTable Separator:=",", NumColumns:=2
    doc.Tables(1).Style = "Grid Table 4 - Accent 1"
    doc.Content.InsertParagraphAfter
    doc.Content.InsertAfter "Item one" & vbCr & "Item two"
    doc.Paragraphs(doc.Paragraphs.Count).Range.ListFormat.ApplyBulletDefault
End Sub""" |> ok "documents add"
let table15 = d15.Blocks |> List.pick (function Doc.Table t -> Some t | _ -> None)
check "convert: cells" ([ for r in 0 .. 2 -> [ for c in 0 .. 1 -> cellText table15 r c ] ]) [ [ "Name"; "Score" ]; [ "Ada"; "90" ]; [ "Bob"; "85" ] ]
check "convert: header fill" (table15.Rows.[0].Cells.[0].BackgroundColor) (Some "#4472C4")
check "convert: after" (bodyTexts d15) [ ""; "Item one"; "Item two" ]
check "convert: bullet" ((paras d15 |> List.last).Metadata.ListStyle) Doc.Bullet

// ---- the showcase macros (tests/macros/word-showcase): each Main runs on Word's blank document --------
// (Their results were compared with Word 2019's on the VM: tools/word-parity/macro-parity.sh.)
let blank =
    match Rbt.Word.Docx.loadBytesNative (IO.File.ReadAllBytes(IO.Path.Combine(__SOURCE_DIRECTORY__, "../macros/word-blank.docx"))) with
    | Some d -> d
    | None -> failwith "word-blank.docx"
let showcase = IO.Path.Combine(__SOURCE_DIRECTORY__, "../macros/word-showcase")
for file in IO.Directory.GetFiles(showcase, "*.bas") |> Array.sort do
    let name = IO.Path.GetFileNameWithoutExtension file
    let code = snd (Rbt.Word.WordMacroHost.splitHeader (IO.File.ReadAllText file))
    let watch = Diagnostics.Stopwatch.StartNew()
    match runOn blank code "Main" with
    | Ok d -> check (sprintf "showcase %s: runs" name) (Doc.allParagraphs d |> List.isEmpty |> not) true
    | Error text -> check (sprintf "showcase %s: runs" name) text "ok"
    if watch.ElapsedMilliseconds > 20000L then printfn "     (showcase %s took %d ms)" name watch.ElapsedMilliseconds

let calendar =
    runOn blank (snd (Rbt.Word.WordMacroHost.splitHeader (IO.File.ReadAllText(IO.Path.Combine(showcase, "month-calendar.bas"))))) "Main" |> ok "calendar"
let calendarTables = calendar.Blocks |> List.choose (function Doc.Table t -> Some t | _ -> None)
check "calendar: three months" calendarTables.Length 3
check "calendar: weekend shading" (calendarTables.Head.Rows.[1].Cells |> List.map (fun c -> c.BackgroundColor)) [ None; None; None; None; None; Some "#DEEBF7"; Some "#DEEBF7" ]
check "calendar: page breaks as Word makes them" (Doc.allParagraphs calendar |> List.filter (fun p -> p.Metadata.PageBreakBefore) |> List.length) 2

// ---- a project saved with a document: a .docx made macro-enabled, a .docm's project replaced ----------
let unzip (bytes: byte[]) =
    use archive = new IO.Compression.ZipArchive(new IO.MemoryStream(bytes))
    [ for e in archive.Entries ->
        use s = e.Open()
        use m = new IO.MemoryStream()
        s.CopyTo m
        e.FullName, m.ToArray() ] |> Map.ofList
let save (doc: Doc.Document) = DocumentWriter.writeTypedPackageBytes (Rbt.Word.Docx.exportTypedDocumentPackage doc)
let project = Rbt.Word.WordMacroHost.modulesBytes [ Rbt.Vb.Api.ModuleSource.standard "Module1" "Sub Main()\r\nEnd Sub\r\n" ]
check "a plain document: no project" (hasProject blank) false
let saved = unzip (save (withProjectBytes project blank))
check "docx -> docm: the project part" (saved |> Map.containsKey "word/vbaProject.bin") true
let types = Text.Encoding.UTF8.GetString saved.["[Content_Types].xml"]
check "docx -> docm: macro-enabled main part" (types.Contains "application/vnd.ms-word.document.macroEnabled.main+xml") true
check "docx -> docm: the project's content type" (types.Contains "/word/vbaProject.bin") true
check "docx -> docm: the relationship" (Text.Encoding.UTF8.GetString(saved.["word/_rels/document.xml.rels"]).Contains "relationships/vbaProject") true
let reopened = (Rbt.Word.Docx.loadBytesNative (save (withProjectBytes project blank))).Value
check "docx -> docm: read again" ((Rbt.Word.WordMacroHost.macroModulesOf reopened) |> List.map (fun m -> m.Name)) [ "Module1" ]
let docm = (Rbt.Word.Docx.loadBytesNative (IO.File.ReadAllBytes(IO.Path.Combine(showcase, "docm/rainbow-text.docm")))).Value
check "docm: its modules" (Rbt.Word.WordMacroHost.macroModulesOf docm |> List.map (fun m -> m.Name)) [ "ThisDocument"; "RainbowText" ]
let edited = (Project.openProject docm).WithModule("RainbowText", "Sub Other()\nEnd Sub")
let resaved = (Rbt.Word.Docx.loadBytesNative (save (withProjectBytes (edited.ProjectBytes()) docm))).Value
check "docm: the edited module saved" ((Project.openProject resaved).Macros) [ "Other" ]
check "docm: kept as it was when not edited" (unzip (save docm)).["word/vbaProject.bin"] (projectBytes docm).Value

printfn "word macros: %d passed, %d failed" pass fail
if fail > 0 then exit 1
