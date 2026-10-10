#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"

open System
open System.IO

// A package whose main part is read front first (DocxReader.ProgressiveReader):
// whatever the size of the steps it is read in, what it has given at the end
// is the document DocxReader.parseFromSource gives - the same blocks, shapes,
// page setup, revisions, pictures and kept parts - for every test document.

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-40s  %s" label detail)

let root = Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, "../.."))
let documents =
    Directory.GetDirectories(Path.Combine(root, "tests"), "test-documents*")
    |> Array.append [| Path.Combine(root, "tests/repros/word") |]
    |> Array.collect (fun dir -> Directory.GetFiles(dir, "*.docx", SearchOption.AllDirectories))
    |> Array.sort

/// The first field in which two documents differ.
let difference (a: DocxReader.WordDocument) (b: DocxReader.WordDocument) : string =
    let content =
        if a.Content.Length <> b.Content.Length then sprintf "Content: %d blocks against %d" a.Content.Length b.Content.Length
        else
            match List.zip a.Content b.Content |> List.tryFindIndex (fun (x, y) -> x <> y) with
            | Some index -> sprintf "Content: block %d differs" index
            | None -> ""
    let preserved =
        let p, q = a.Preserved, b.Preserved
        [ "SourceXml", p.SourceXml = q.SourceXml; "SourceRelationships", p.SourceRelationships = q.SourceRelationships; "SourceParts", p.SourceParts = q.SourceParts
          "BodyContentChunks", p.BodyContentChunks = q.BodyContentChunks; "BodyChunkCount", p.BodyChunkCount = q.BodyChunkCount; "Parts", p.Parts = q.Parts
          "PartRelationships", p.PartRelationships = q.PartRelationships; "RootRelationships", p.RootRelationships = q.RootRelationships
          "PartPaths", p.PartPaths = q.PartPaths; "OriginalText", p.OriginalText = q.OriginalText ]
        |> List.filter (fun (_, same) -> not same) |> List.map fst |> String.concat ", "
    [ content
      (if a.Shapes <> b.Shapes then sprintf "Shapes (%d against %d)" a.Shapes.Length b.Shapes.Length else "")
      (if a.PageSetup <> b.PageSetup then "PageSetup" else "")
      (if a.Revisions <> b.Revisions then "Revisions" else "")
      (if a.Images <> b.Images then "Images" else "")
      (if a.Headers <> b.Headers then "Headers" else "")
      (if a.Footers <> b.Footers then "Footers" else "")
      (if a.Footnotes <> b.Footnotes || a.Endnotes <> b.Endnotes then "notes" else "")
      (if a.Comments <> b.Comments then "Comments" else "")
      (if a.Styles <> b.Styles || a.Theme <> b.Theme || a.NumberingDefinitions <> b.NumberingDefinitions then "styles / theme / lists" else "")
      (if preserved <> "" then "Preserved: " + preserved else "") ]
    |> List.filter (fun text -> text <> "")
    |> String.concat "; "

let mutable streamed = 0
let mutable atOnce = 0
let mutable compared = 0
for path in documents do
    let name = Path.GetFileName path
    let bytes = File.ReadAllBytes path
    match (try Choice1Of2 (DocxReader.loadBytesNative bytes) with e -> Choice2Of2 e.Message) with
    | Choice2Of2 message ->
        // What cannot be read at once is not read front first either.
        let progressive = try (DocxReader.openProgressiveNative true bytes |> Option.map (fun reader -> reader.Finish ())) |> ignore; "read" with e -> "failed"
        ok (name + ": unreadable both ways") (progressive = "failed") message
    | Choice1Of2 None ->
        ok (name + ": not a document both ways") (DocxReader.openProgressiveNative true bytes).IsNone ""
    | Choice1Of2 (Some whole) ->
        compared <- compared + 1
        for step in [ 1; 7; 64; Int32.MaxValue ] do
            match DocxReader.openProgressiveNative true bytes with
            | None -> ok (name + ": opens") false "the progressive reader does not open it"
            | Some reader ->
                let blocks = ResizeArray<DocxReader.DocumentContent>()
                let shapes = ResizeArray<DocxReader.Shape>()
                let mutable steps = 0
                let mutable firsts = true
                let mutable images = true
                let mutable openSections = true
                while not (reader.IsDone ()) do
                    let batch = reader.Read step
                    // (The section the blocks end in: a page setup there always is.)
                    if (reader.OpenSection ()).IsNone && whole.PageSetup.IsSome then openSections <- false
                    if batch.First <> blocks.Count then firsts <- false
                    // (The pictures of the blocks, as the whole document's map has them.)
                    for KeyValue (key, source) in reader.ImagesOf batch.Blocks do
                        if Map.tryFind key whole.Images <> Some source then images <- false
                    blocks.AddRange batch.Blocks
                    shapes.AddRange batch.Shapes
                    steps <- steps + 1
                    if steps > 1000000 then failwith "the reader does not end"
                if step = 1 then (if steps > 2 then streamed <- streamed + 1 else atOnce <- atOnce + 1)
                let final = reader.Finish ()
                let label = sprintf "%s (%s at a time)" name (if step = Int32.MaxValue then "all" else string step)
                ok (label + ": the blocks as they were given") (List.ofSeq blocks = whole.Content && List.ofSeq shapes = whole.Shapes && firsts) (sprintf "%d blocks given, %d in the document" blocks.Count whole.Content.Length)
                ok (label + ": pictures") images "a picture of a block is not the document's"
                ok (label + ": the open section") (openSections && reader.OpenSection () = whole.PageSetup) "no page setup for the section being read, or not the document's at the end"
                ok (label + ": the document") (final = whole) (difference whole final)
                ok (label + ": the front") (reader.Front.Styles = whole.Styles && reader.Front.Headers = whole.Headers && reader.Front.Footers = whole.Footers
                                            && reader.Front.Footnotes = whole.Footnotes && reader.Front.Endnotes = whole.Endnotes && reader.Front.Comments = whole.Comments
                                            && reader.Front.PageSetup = whole.PageSetup && reader.Front.Theme = whole.Theme) "styles, headers, footers, notes, comments or page setup differ"

// A block is not given until the one after it is read: an empty paragraph that ends a section on a
// new page gives way to a continuous break after it, whatever the step.
printfn "docx progressive read: %d documents compared (%d read front first, %d at once); %d passed, %d failed" compared streamed atOnce pass fail
ok "most documents are read front first" (streamed * 10 >= compared * 7) (sprintf "%d of %d" streamed compared)
if fail > 0 then exit 1
