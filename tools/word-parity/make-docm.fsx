// dotnet fsi tools/word-parity/make-docm.fsx <base.docx> <out.docm> <module.bas>...
//
// A macro-enabled copy of a document: its package with a source-only word/vbaProject.bin of the given
// modules (VbaProject.writeProjectModules; Word compiles the source when it opens the file), the
// relationship to it and the macro-enabled content type. Each .bas is a standard module named after
// its Attribute VB_Name (else the file's name). For running the showcase macros in Word and in ours.
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"

open System
open System.IO
open System.IO.Compression

let args = fsi.CommandLineArgs |> Array.skip 1
if args.Length < 3 then
    eprintfn "usage: make-docm.fsx <base.docx> <out.docm> <module.bas>..."
    exit 2
let basePath, outPath, modulePaths = args.[0], args.[1], args.[2..]

let moduleOf (path: string) =
    let source = File.ReadAllText path
    let m = Text.RegularExpressions.Regex.Match(source, "Attribute\\s+VB_Name\\s*=\\s*\"([^\"]+)\"")
    let name = if m.Success then m.Groups.[1].Value else Path.GetFileNameWithoutExtension path
    let source = if m.Success then source else sprintf "Attribute VB_Name = \"%s\"\r\n%s" name source
    name, VbaProject.StandardModule, source.Replace("\r\n", "\n").Replace("\n", "\r\n")
// (Word's projects have the document's own module, ThisDocument.)
let thisDocument =
    "ThisDocument", VbaProject.DocumentModule,
    String.concat "\r\n"
        [ "Attribute VB_Name = \"ThisDocument\""; "Attribute VB_Base = \"1Normal.ThisDocument\""
          "Attribute VB_GlobalNameSpace = False"; "Attribute VB_Creatable = False"; "Attribute VB_PredeclaredId = True"
          "Attribute VB_Exposed = True"; "Attribute VB_TemplateDerived = True"; "Attribute VB_Customizable = True"; "" ]
let project = VbaProject.writeProjectModules (thisDocument :: [ for p in modulePaths -> moduleOf p ])

let entries =
    use archive = ZipFile.OpenRead basePath
    [ for e in archive.Entries ->
        use s = e.Open()
        use m = new MemoryStream()
        s.CopyTo m
        e.FullName, m.ToArray() ]
let text (bytes: byte[]) = Text.Encoding.UTF8.GetString bytes
let bytes (s: string) = Text.Encoding.UTF8.GetBytes s
let patched =
    entries
    |> List.filter (fun (name, _) -> name <> "word/vbaProject.bin")
    |> List.map (fun (name, data) ->
        match name with
        | "[Content_Types].xml" ->
            let s = text data
            let s = s.Replace("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "application/vnd.ms-word.document.macroEnabled.main+xml")
            let s =
                if s.Contains "Extension=\"bin\"" then s
                else s.Replace("<Default Extension=\"xml\"", "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/><Default Extension=\"xml\"")
            name, bytes s
        | "word/_rels/document.xml.rels" ->
            let s = text data
            let s =
                if s.Contains "relationships/vbaProject" then s
                else s.Replace("</Relationships>", "<Relationship Id=\"rIdVbaProject\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" Target=\"vbaProject.bin\"/></Relationships>")
            name, bytes s
        | _ -> name, data)
let all = patched @ [ "word/vbaProject.bin", project ]
if File.Exists outPath then File.Delete outPath
do
    use stream = File.Create outPath
    use archive = new ZipArchive(stream, ZipArchiveMode.Create)
    for name, data in all do
        let e = archive.CreateEntry(name, CompressionLevel.Optimal)
        use s = e.Open()
        s.Write(data, 0, data.Length)
printfn "%s: %d module(s)" outPath modulePaths.Length
