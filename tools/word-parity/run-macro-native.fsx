// dotnet fsi tools/word-parity/run-macro-native.fsx <input.docm> <macro> <out.tsv> [out.docx]
//
// Our side of word-macro-on-vm.sh: opens the macro-enabled document natively, runs the macro with our
// object model (WordMacroHost), and writes the document as run-word-macro.vbs dumps it from Word (the
// dump is tests/macros/rbt-dump.bas, run by our object model), and the saved document when asked.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Vb.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"

open System.IO
open Rbt.Word.WordMacroApi

let args = fsi.CommandLineArgs |> Array.skip 1
let input, macro, outPath = args.[0], args.[1], args.[2]
let doc =
    match Rbt.Word.Docx.loadBytesNative (File.ReadAllBytes input) with
    | Some d -> d
    | None -> failwithf "cannot read %s" input
let lines = ResizeArray<string>()
let log = ResizeArray<string>()
let hooks = { defaultHooks with Output = log.Add; FileName = Path.GetFileName input }
let project = Rbt.Word.WordMacroProject.openProject doc
for p in project.Problems do eprintfn "%s" p.Text
let dumpCode = File.ReadAllText(Path.Combine(__SOURCE_DIRECTORY__, "../../tests/macros/rbt-dump.bas"))
let withDump = project.WithModule("RbtDump", snd (Rbt.Word.WordMacroHost.splitHeader dumpCode))
let result =
    match withDump.Run(doc, hooks, macro) with
    | MacroDone d -> d
    | MacroFailed f ->
        lines.Add(sprintf "error\t%d\t%s" f.Number (f.Description.Replace("\t", "\\t")))
        eprintfn "%s" f.Text
        doc
let printed = log.ToArray()
log.Clear()
match withDump.Run(result, { hooks with Output = lines.Add }, "RbtDump.Dump") with
| MacroDone _ -> ()
| MacroFailed f -> eprintfn "dump: %s" f.Text
File.WriteAllLines(outPath, lines)
if printed.Length > 0 then File.WriteAllLines(Path.ChangeExtension(outPath, ".log"), printed)
if args.Length > 3 then
    let package = Rbt.Word.Docx.exportTypedDocumentPackage result
    File.WriteAllBytes(args.[3], DocumentWriter.writeTypedPackageBytes package)
printfn "%s" outPath
