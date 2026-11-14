// Which columns of a section Word balances (makes about as tall as each other).
// Checked against Word 2019 with a two-column section of twelve paragraphs after a
// one-column paragraph:
// - followed by a continuous section break: balanced (six and six);
// - followed by a next-page section break: not balanced, column 1 is filled to the
//   bottom of the page first (eleven and one);
// - the document's last section (its body w:sectPr, itself started by a continuous
//   break): not balanced either;
// - a document that is one two-column section: not balanced.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open Rbt.Word.TextBuffer
module Doc = Rbt.Word.Document
module Layout = Rbt.Word.Layout

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-60s expected %A got %A" label expected got

let w = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
let page = "<w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>"
let sectPr (kind: string) (columns: int) = sprintf "<w:sectPr><w:type w:val=\"%s\"/>%s<w:cols w:num=\"%d\" w:space=\"720\"/></w:sectPr>" kind page columns
let package (body: string) (finalSectPr: string) =
    use stream = new IO.MemoryStream()
    do
        use archive = new IO.Compression.ZipArchive(stream, IO.Compression.ZipArchiveMode.Create, true)
        let add (name: string) (text: string) =
            use writer = new IO.StreamWriter(archive.CreateEntry(name).Open(), Text.UTF8Encoding(false))
            writer.Write text
        add "[Content_Types].xml" """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"""
        add "word/document.xml" (sprintf "<w:document %s><w:body>%s%s</w:body></w:document>" w body finalSectPr)
    stream.ToArray()
let import body finalSectPr = package body finalSectPr |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let para (text: string) = sprintf "<w:p><w:r><w:t xml:space=\"preserve\">%s</w:t></w:r></w:p>" text
let sectionEnd kind columns = sprintf "<w:p><w:pPr>%s</w:pPr></w:p>" (sectPr kind columns)
let columnText = String.concat "" [ for i in 1 .. 22 -> para (sprintf "Column paragraph %d, %s" i (String.replicate 8 "word ")) ]

/// How many of the column paragraphs stand in the first and in the second column (of page 1).
let split (doc: Doc.Document) =
    let settings = Layout.documentSettings 16.0 24.0 3.5 doc
    let layouts = Layout.layoutDocumentPages settings measureText measureRange doc |> fst
    let middle = settings.Width / 2.0
    let xs =
        doc.Blocks
        |> List.choose Doc.tryGetParagraph
        |> List.filter (fun p -> (toString p.Buffer).StartsWith "Column paragraph")
        |> List.map (fun p ->
            let layout = layouts |> List.find (fun l -> l.BlockId = p.Id)
            let x = match layout.Lines.Head.Segments with Layout.TextSegment (_, _, x, _, _) :: _ -> x | _ -> nan
            settings.PaddingLeft + x, layout.Y < settings.PageHeight)
    let onFirst = xs |> List.filter snd
    (onFirst |> List.filter (fun (x, _) -> x < middle) |> List.length), (onFirst |> List.filter (fun (x, _) -> x >= middle) |> List.length)

let intro = para "Intro in one column." + sectionEnd "nextPage" 1
let lastSection = import (intro + columnText) (sectPr "continuous" 2)
let beforeContinuous = import (intro + columnText + sectionEnd "continuous" 2 + para "After.") (sectPr "continuous" 1)
let beforeNextPage = import (intro + columnText + sectionEnd "continuous" 2 + para "After.") (sectPr "nextPage" 1)
let oneSection = import columnText (sectPr "nextPage" 2)

let isFilledFirst (first, second) = first > second + 2 && second > 0 && first + second = 22
let isBalanced (first, second) = abs (first - second) <= 1 && first + second = 22
check "the last section: column 1 filled first" (isFilledFirst (split lastSection)) true
check "before a continuous break: balanced" (isBalanced (split beforeContinuous)) true
check "before a next-page break: column 1 filled first" (isFilledFirst (split beforeNextPage)) true
check "one two-column section: column 1 filled first" (isFilledFirst (split oneSection)) true

printfn "word column balance: %d passed, %d failed" pass fail
if fail > 0 then exit 1
