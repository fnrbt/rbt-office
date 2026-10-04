// Opens documents saved by roundtrip.mjs (--out DIR) with the .NET reader, the
// other consumer of what DocumentWriter produces. For every <name>.pass1.docx /
// <name>.pass2.docx pair it checks that both parse and hold the same text.
//
//   dotnet build src/Office.Formats/Rbt.OfficeFormat.fsproj
//   dotnet fsi tools/docx-roundtrip/native-check.fsx DIR
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"

open System
open System.IO

let dir = fsi.CommandLineArgs |> Array.tryItem 1 |> Option.defaultValue "."
let mutable failed = 0
let load (path: string) =
    try File.ReadAllBytes path |> DocxReader.loadBytesNative
    with ex ->
        printfn "  %s: %s" (Path.GetFileName path) ex.Message
        None
let stats (doc: DocxReader.WordDocument) =
    let text = DocxReader.extractText doc
    let tables = doc.Content |> List.filter (function DocxReader.Tbl _ -> true | _ -> false) |> List.length
    // Note and comment bodies: how many, and how many blocks / tables they hold.
    let noteBlocks = (doc.Footnotes @ doc.Endnotes |> List.collect (fun note -> note.Blocks)) @ (doc.Comments |> List.collect (fun comment -> comment.Blocks))
    let noteTables = noteBlocks |> List.filter (function DocxReader.Tbl _ -> true | _ -> false) |> List.length
    let notes = doc.Footnotes.Length + doc.Endnotes.Length + doc.Comments.Length
    (text, notes, noteBlocks.Length, noteTables), doc.Content.Length, tables, doc.Headers.Length + doc.Footers.Length, doc.Images.Count

for first in Directory.GetFiles(dir, "*.pass1.docx") |> Array.sort do
    let name = Path.GetFileName(first).Replace(".pass1.docx", "")
    match load first, load (first.Replace(".pass1.docx", ".pass2.docx")) with
    | Some a, Some b ->
        let (textA, blocksA, tablesA, partsA, imagesA), (textB, blocksB, tablesB, partsB, imagesB) = stats a, stats b
        if textA = textB && blocksA = blocksB && tablesA = tablesB && partsA = partsB && imagesA = imagesB then
            let (text, notes, noteBlocks, noteTables) = textA
            printfn "ok   %s  (%d blocks, %d tables, %d header/footer parts, %d images, %d chars, %d notes/comments with %d blocks, %d tables)" name blocksA tablesA partsA imagesA text.Length notes noteBlocks noteTables
        else
            failed <- failed + 1
            let chars (text: string, _, _, _) = text.Length
            printfn "FAIL %s  save 1: %d blocks %d tables %d parts %d images %d chars %A; save 2: %d blocks %d tables %d parts %d images %d chars %A" name blocksA tablesA partsA imagesA (chars textA) textA blocksB tablesB partsB imagesB (chars textB) textB
    | _ ->
        failed <- failed + 1
        printfn "FAIL %s  the .NET reader could not open a saved document" name

printfn ""
printfn (if failed = 0 then "NATIVE READ OF SAVED DOCUMENTS: ALL PASSED" else "NATIVE READ OF SAVED DOCUMENTS: FAILED")
if failed > 0 then exit 1
