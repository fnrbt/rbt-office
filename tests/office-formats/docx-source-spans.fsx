// The source of what the reader reads (docs/save-architecture.md, step 2): every paragraph of
// a document's body knows the w:p it was read from and every run the w:r, as places in the
// main part's text - read at once or front first, the same places.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"

open System
open System.IO
open System.IO.Compression
open System.Text.RegularExpressions

let mutable pass = 0
let mutable fail = 0
let failures = ResizeArray<string>()
let ok label cond detail =
    if cond then pass <- pass + 1
    else
        fail <- fail + 1
        if failures.Count < 40 then failures.Add (sprintf "FAIL %s  %s" label detail)

let root = Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, "..", ".."))
let files =
    Directory.GetFiles(Path.Combine(root, "tests"), "*.docx", SearchOption.AllDirectories) |> Array.sort

let mainText (bytes: byte[]) =
    use zip = new ZipArchive(new MemoryStream(bytes))
    let read (name: string) =
        match zip.GetEntry name with
        | null -> None
        | entry -> use r = new StreamReader(entry.Open(), Text.Encoding.UTF8, true) in Some (r.ReadToEnd())
    let target =
        read "_rels/.rels"
        |> Option.bind (fun rels -> let m = Regex.Match(rels, "Type=\"[^\"]*/officeDocument\"[^>]*Target=\"/?([^\"]+)\"|Target=\"/?([^\"]+)\"[^>]*Type=\"[^\"]*/officeDocument\"") in if m.Success then Some (if m.Groups.[1].Success then m.Groups.[1].Value else m.Groups.[2].Value) else None)
        |> Option.defaultValue "word/document.xml"
    read target

let isTag (local: string) (text: string) (at: int) =
    let m = Regex.Match(text.Substring(at, min 80 (text.Length - at)), "^<([A-Za-z_][\\w.\\-]*:)?" + local + "[\\s/>]")
    m.Success

let bodyParagraphs (blocks: DocxReader.DocumentContent list) =
    let found = ResizeArray<DocxReader.Paragraph>()
    blocks |> DocxReader.mapBlockParagraphs (fun para -> found.Add para; para) |> ignore
    List.ofSeq found

let mutable documents = 0
let mutable paragraphs = 0
let mutable known = 0
for file in files do
    let name = Path.GetRelativePath(root, file)
    let bytes = File.ReadAllBytes file
    match (try DocxReader.loadBytesNative bytes with _ -> None), mainText bytes with
    | Some doc, Some text ->
        documents <- documents + 1
        let paras = bodyParagraphs doc.Content
        for para in paras do
            paragraphs <- paragraphs + 1
            let span = para.Source
            if XmlDsl.Span.isKnown span then
                known <- known + 1
                ok (name + ": a paragraph's source is a w:p") (isTag "p" text span.Start) (text.Substring(span.Start, min 60 (text.Length - span.Start)))
                for run in para.Runs do
                    if XmlDsl.Span.isKnown run.Source then
                        ok (name + ": a run's source is a w:r") (isTag "r" text run.Source.Start) (text.Substring(run.Source.Start, min 60 (text.Length - run.Source.Start)))
                        ok (name + ": a run's source is in its paragraph's") (run.Source.Start > span.Start && run.Source.End < span.End) (sprintf "%A in %A" run.Source span)
        // Front first: the same places.
        match (try DocxReader.openProgressiveNative false bytes with _ -> None) with
        | Some reader when not reader.AtOnce ->
            let blocks = ResizeArray<DocxReader.DocumentContent>()
            while not (reader.IsDone()) do blocks.AddRange (reader.Read 200).Blocks
            let front = bodyParagraphs (List.ofSeq blocks) |> List.map (fun p -> p.Source.Start)
            let whole = paras |> List.map (fun p -> p.Source.Start)
            ok (name + ": read front first, the paragraphs have the places they have read at once") (front = whole)
                (sprintf "%d against %d paragraphs" front.Length whole.Length)
        | _ -> ()
    | _ -> ()

for line in failures do printfn "%s" line
printfn "%d documents, %d paragraphs, %d with their source" documents paragraphs known
ok "nearly every paragraph knows its source" (known * 100 >= paragraphs * 99) (sprintf "%d of %d" known paragraphs)
printfn "DOCX source-span tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
