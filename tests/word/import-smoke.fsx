#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open Rbt.Word.Document

let mutable pass = 0
let mutable fail = 0

let check label got expected =
    if got = expected then
        pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-34s expected %A got %A" label expected got

let ok label cond detail =
    if cond then
        pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-34s %s" label detail

let textOfPara (para: ParagraphBlock) =
    Rbt.Word.TextBuffer.toString para.Buffer

let firstParagraphText (doc: Document) =
    doc.Blocks
    |> List.tryPick (function
        | Paragraph para
        | Heading (_, para) -> Some (textOfPara para)
        | Table table ->
            table.Rows
            |> List.tryPick (fun row ->
                row.Cells
                |> List.tryPick (fun cell ->
                    cell.Content |> List.tryHead |> Option.map textOfPara))
        | TextBox textBox ->
            textBox.Content |> List.tryHead |> Option.map textOfPara)
    |> Option.defaultValue ""

let blockKind = function
    | Paragraph _ -> "Paragraph"
    | Heading _ -> "Heading"
    | Table _ -> "Table"
    | TextBox _ -> "TextBox"

let hasSectionPageSetup (doc: Document) =
    let paraHasSetup (para: ParagraphBlock) =
        para.Metadata.SectionPageSetup.IsSome

    doc.Blocks
    |> List.exists (function
        | Paragraph para
        | Heading (_, para) -> paraHasSetup para
        | Table table ->
            table.Rows
            |> List.exists (fun row ->
                row.Cells
                |> List.exists (fun cell -> cell.Content |> List.exists paraHasSetup))
        | TextBox textBox -> textBox.Content |> List.exists paraHasSetup)

type Expected = {
    Path: string
    Blocks: int
    FirstText: string
    Kinds: string list
    HasSectionSetup: bool
}

let fixtures =
    [
        {
            Path = "tests/repros/word/two-column-section.docx"
            Blocks = 11
            FirstText = "Two column section repro"
            Kinds = List.replicate 11 "Paragraph"
            HasSectionSetup = false
        }
        {
            Path = "tests/repros/word/landscape-section.docx"
            Blocks = 9
            FirstText = "Landscape section (page 1)"
            Kinds = List.replicate 9 "Paragraph"
            HasSectionSetup = true
        }
        {
            Path = "tests/repros/word/table-import-preserves-structure.docx"
            Blocks = 3
            FirstText = "Table import repro"
            Kinds = [ "Paragraph"; "Table"; "Paragraph" ]
            HasSectionSetup = false
        }
    ]

let importFixture path importer =
    let root = IO.Path.GetFullPath(IO.Path.Combine(__SOURCE_DIRECTORY__, "../.."))
    let bytes = IO.File.ReadAllBytes(IO.Path.Combine(root, path))
    match DocxReader.loadBytesNative bytes with
    | Some parsed -> importer parsed
    | None -> failwithf "Could not parse %s" path

let assertFixture (expected: Expected) (doc: Document) =
    ok (expected.Path + " non-empty") (not (List.isEmpty doc.Blocks)) "document has no blocks"
    check (expected.Path + " block count") (List.length doc.Blocks) expected.Blocks
    check (expected.Path + " first text") (firstParagraphText doc) expected.FirstText
    check (expected.Path + " page width") doc.PageSetup.PageWidth 816.0
    check (expected.Path + " page height") doc.PageSetup.PageHeight 1056.0
    check (expected.Path + " margins") (doc.PageSetup.MarginLeft, doc.PageSetup.MarginRight, doc.PageSetup.MarginTop, doc.PageSetup.MarginBottom) (96.0, 96.0, 96.0, 96.0)
    check (expected.Path + " block kinds") (doc.Blocks |> List.map blockKind) expected.Kinds
    check (expected.Path + " section setup") (hasSectionPageSetup doc) expected.HasSectionSetup

for expected in fixtures do
    let doc = importFixture expected.Path Rbt.Word.Import.importDoc
    assertFixture expected doc

let summary (doc: Document) =
    List.length doc.Blocks,
    firstParagraphText doc,
    doc.PageSetup,
    (doc.Blocks |> List.map blockKind),
    hasSectionPageSetup doc

let libraryDoc = importFixture fixtures.[1].Path Rbt.Word.Import.importDoc
let compatDoc = importFixture fixtures.[1].Path NativeRunner.importDoc
check "compat alias summary" (summary compatDoc) (summary libraryDoc)

printfn "word import smoke: %d passed, %d failed" pass fail
if fail > 0 then exit 1
