#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"

open Rbt.Word.TextBuffer
module Doc = Rbt.Word.Document
module TB = Rbt.Word.TextBuffer

// Editing commands of the model as the keys use them.

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-30s expected %A got %A" label expected got

let paragraph id (text: string) : Doc.ParagraphBlock = { Id = id; Buffer = TB.empty |> TB.insert 0 text; Metadata = Doc.defaultMetadata }
let textOf (para: Doc.ParagraphBlock) = TB.toString para.Buffer
let docOf (blocks: Doc.Block list) = { Doc.empty with Blocks = blocks; Selection = None }
let caret id offset (doc: Doc.Document) = Doc.setCursor id offset doc
let texts (doc: Doc.Document) = Doc.allParagraphs doc |> List.map textOf

let sectioned =
    docOf [ Doc.Paragraph { paragraph "s1" "end of section" with Metadata = { Doc.defaultMetadata with SectionBreak = Some "nextPage" } }
            Doc.Paragraph (paragraph "s2" "next section") ]

// ---- Enter in the last paragraph of a section: one section end, after the second part ----
let entered = sectioned |> caret "s1" 3 |> Doc.splitBlock
check "Enter: the section still ends once, with the second part"
    (Doc.allParagraphs entered |> List.map (fun p -> textOf p, p.Metadata.SectionBreak))
    [ "end", None; " of section", Some "nextPage"; "next section", None ]
let broken =
    docOf [ Doc.Paragraph { paragraph "b1" "before break" with Metadata = { Doc.defaultMetadata with PageBreakAfter = true; ColumnBreakAfter = true } } ]
    |> caret "b1" 6 |> Doc.splitBlock
check "Enter: a column break after the text stays after it"
    (Doc.allParagraphs broken |> List.map (fun p -> textOf p, p.Metadata.PageBreakAfter, p.Metadata.ColumnBreakAfter))
    [ "before", false, false; " break", true, true ]

printfn ""
printfn "Word editing tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
