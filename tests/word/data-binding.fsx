// Content controls bound to custom XML, against Word 2019 on the fixtures
// tools/word-parity/fixtures/data-binding.py generates (tests/word/fixtures/data-binding; Word's side in
// word/: <name>.tsv - every paragraph's text as Word shows it on opening the document - and
// <name>.edited.docx - each bound control's text changed in Word, then saved). What Word does
// (docs/layout-parity-flow.md, the data-binding rows; DataBinding.fs):
// - every paragraph reads as Word's: the controls hold their nodes' values (the main story and the header);
// - DataBinding.writeBack given Word's edited document gives the custom XML Word saved with it.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
open System.IO.Compression
open Rbt.Word.Document

let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/data-binding")

let mutable pass = 0
let mutable fail = 0
let check (label: string) (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-28s Word %A ours %A" label expected got

// Word's: (story, paragraph text) in order; a line break \v, a tab \t.
let wordParagraphs (name: string) =
    File.ReadAllLines(Path.Combine(folder, "word", name + ".tsv"))
    |> Array.map (fun l -> l.TrimStart('\uFEFF').Split('\t'))
    |> Array.filter (fun c -> c.[0] = "para" && (c.[1] = "1" || c.[1] = "7"))
    |> Array.map (fun c ->
        let text = String.Join("\t", c.[3..]).Replace("\\t", "\t")
        let text = if text.EndsWith "\\r" then text.Substring(0, text.Length - 2) else text
        c.[1], text)
    |> Array.toList

let ourParagraphs (file: string) =
    let doc = File.ReadAllBytes file |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let texts (blocks: Block list) =
        blocks |> List.choose (function
            | Paragraph p | Heading (_, p) -> Some (Rbt.Word.TextBuffer.toString p.Buffer |> fun t -> t.Replace("\n", "\\v").Replace("\u000B", "\\v"))
            | _ -> None)
    (texts doc.Blocks |> List.map (fun t -> "1", t))
    @ (doc.HeaderParts |> List.filter (fun part -> part.Kind = "default") |> List.collect (fun part -> texts part.Blocks) |> List.map (fun t -> "7", t))

for name in [ "cases"; "cases2"; "header"; "repeat"; "repeat1"; "repeatfewer" ] do
    let word = wordParagraphs name
    let ours = ourParagraphs (Path.Combine(folder, name + ".docx"))
    // (Word lists the header's empty paragraphs of other stories; only the default header is compared.)
    let word = word |> List.filter (fun (story, text) -> story = "1" || text <> "")
    check (name + ": paragraphs") ours.Length word.Length
    for (story, w), (_, o) in List.zip word (List.truncate word.Length ours) do
        check (sprintf "%s: %s" name (w.Split(':').[0])) o w

// The runs of each control (by its w:tag) as Word saved the document it opened (<name>.saved.docx) and as
// ours reads it (DocxReader.withDataBindings): each run's text and bold / italic / colour.
let partOf (file: string) (path: string) =
    use zip = ZipFile.OpenRead file
    use r = new StreamReader(zip.GetEntry(path).Open())
    r.ReadToEnd()
let controlRuns (xml: string) =
    let w = Xml.Linq.XNamespace.Get "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    let doc = Xml.Linq.XDocument.Parse(xml.TrimStart('﻿'))
    [ for sdt in doc.Descendants(w + "sdt") do
        let tag = sdt.Element(w + "sdtPr").Element(w + "tag")
        if not (isNull tag) then
            let runs =
                [ for r in sdt.Element(w + "sdtContent").Descendants(w + "r") ->
                    let rPr = r.Element(w + "rPr")
                    let has (name: string) = not (isNull rPr) && not (isNull (rPr.Element(w + name)))
                    String.Join("", r.Elements(w + "t") |> Seq.map (fun t -> t.Value)), has "b", has "i", has "color" ]
            // (Runs of one format one after another as one: Word ends a run at a tab or a line break.)
            let merged =
                runs |> List.fold (fun acc (t, b, i, c) ->
                    match acc with
                    | (t0, b0, i0, c0) :: rest when (b0, i0, c0) = (b, i, c) -> (t0 + t, b, i, c) :: rest
                    | _ -> (t, b, i, c) :: acc) [] |> List.rev
            yield tag.Attribute(w + "val").Value, merged ]
for name in [ "cases"; "cases2"; "header" ] do
    let file = Path.Combine(folder, name + ".docx")
    let source : DocxReader.PartSource =
        { Text = (fun path -> try Some (partOf file path) with _ -> None); Bytes = fun _ -> None }
    let ours = (DocxReader.withDataBindings source).Text "word/document.xml" |> Option.get |> controlRuns
    let word = partOf (Path.Combine(folder, "word", name + ".saved.docx")) "word/document.xml" |> controlRuns |> Map.ofList
    for (tag, runs) in ours do
        match Map.tryFind tag word with
        | Some wordRuns -> check (sprintf "%s: %s runs" name tag) runs wordRuns
        | None -> ()

// What a store holds, by its itemID: each element's text and attribute (path -> value).
let storeValues (zipBytes: byte[]) =
    use zip = new ZipArchive(new MemoryStream(zipBytes))
    let parts = [ for e in zip.Entries -> e.FullName, (use r = new StreamReader(e.Open()) in r.ReadToEnd()) ]
    let text (path: string) =
        parts |> List.tryPick (fun (name, content) -> if name.Equals(path, StringComparison.OrdinalIgnoreCase) then Some content else None)
    let stores = DataBinding.storesOf text "word/document.xml"
    let stories =
        parts |> List.map fst |> List.filter (fun name -> name.StartsWith "word/" && name.EndsWith ".xml" && not (name.Contains "/_rels/"))
    stores, text, stories

let flatten (xml: string) =
    let doc = Xml.Linq.XDocument.Parse(xml.TrimStart('\uFEFF'))
    [ for e in doc.Descendants() do
        let path = e.AncestorsAndSelf() |> Seq.rev |> Seq.map (fun a -> a.Name.LocalName) |> String.concat "/"
        if not e.HasElements then yield path, e.Value
        for a in e.Attributes() do if not a.IsNamespaceDeclaration then yield path + "/@" + a.Name.LocalName, a.Value ]

for name in [ "cases"; "cases2"; "header" ] do
    let original, originalText, _ = storeValues (File.ReadAllBytes(Path.Combine(folder, name + ".docx")))
    let edited, editedText, editedStories = storeValues (File.ReadAllBytes(Path.Combine(folder, "word", name + ".edited.docx")))
    // Our package: the original's parts, Word's edited stories.
    let textOf (path: string) =
        if editedStories |> List.contains path then editedText path else originalText path
    let changed = DataBinding.writeBack textOf "word/document.xml" (editedStories |> List.choose editedText) |> Map.ofList
    for store in original do
        let ours = Map.tryFind store.Path changed |> Option.defaultValue store.Text
        match edited |> List.tryFind (fun s -> s.Id = store.Id) with
        | Some word -> check (sprintf "%s: store %s" name store.Id) (flatten ours) (flatten word.Text)
        | None -> check (sprintf "%s: store %s in Word's" name store.Id) false true

printfn "Data binding against Word: %d checks as Word, %d not" pass fail
if fail > 0 then exit 1
