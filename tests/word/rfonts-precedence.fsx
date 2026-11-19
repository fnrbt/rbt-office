#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"

// Which w:rFonts wins (docs/layout-parity-flow.md ledger): the fixtures of
// tools/word-parity/fixtures/rfonts-precedence.py, against the face Word 2019 reported for each
// (fixtures/rfonts-precedence/word-fonts.tsv: Range.Font.Name of the paragraph's text).
open System.IO
open Rbt.Word

let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures", "rfonts-precedence")
let mutable pass = 0
let mutable fail = 0
for line in File.ReadAllLines(Path.Combine(folder, "word-fonts.tsv")) do
    match line.Split('\t') with
    | [| name; word |] ->
        let doc = Docx.loadBytesNative (File.ReadAllBytes(Path.Combine(folder, name + ".docx"))) |> Option.get
        let ours =
            doc.Blocks |> List.tryPick (function
                | Document.Paragraph p | Document.Heading (_, p) ->
                    p.Buffer.Runs |> List.tryPick (function TextBuffer.Text t -> Some (t.Style.FontFamily |> Option.defaultValue "") | _ -> None)
                | _ -> None)
            |> Option.defaultValue ""
        // (A theme face resolves to the theme's name for it: Calibri minor, Cambria major.)
        if ours = word then pass <- pass + 1
        else
            fail <- fail + 1
            printfn "FAIL %s: Word %s, ours %s" name word ours
    | _ -> ()
printfn "rfonts precedence: %d passed, %d failed" pass fail
if fail > 0 then exit 1
