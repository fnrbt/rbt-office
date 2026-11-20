// The native hosts set text in the fonts a document embeds as the browser does (tests/word/embedded-fonts.fsx
// for Word's rule): Rbt.WordNative.Render.measureTextFor / toRendererCommandFor take a document's
// EmbeddedFonts for the families this machine has no font of (Rbt.OfficeNative.Fonts.hasFamily), and keep
// the machine's font of a family it has. The fixtures' embedded face is Liberation Mono renamed: every
// character as wide as every other.
#r "nuget: Fable.Core, 4.3.0"
#r "../../office-wisp/word-render/bin/Debug/net10.0/Rbt.Renderer.dll"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../office-wisp/word-render/bin/Debug/net10.0/Rbt.Word.dll"
#r "../../office-wisp/word-render/bin/Debug/net10.0/Rbt.OfficeNative.dll"
#r "../../office-wisp/word-render/bin/Debug/net10.0/Rbt.WordNative.dll"

open System.IO
module TB = Rbt.Word.TextBuffer
module WR = Rbt.WordNative.Render
module Paint = Rbt.Word.Paint

let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/embedded-fonts")
let load name = File.ReadAllBytes(Path.Combine(folder, name + ".docx")) |> Rbt.Word.Docx.loadBytesNative |> Option.get
let mutable failures = 0
let mutable checks = 0
let check label ok detail =
    checks <- checks + 1
    if not ok then
        failures <- failures + 1
        printfn "FAIL %s %s" label detail
let style family bold = { TB.defaultStyle with FontFamily = Some family; FontSize = 10.0; Bold = bold }
let widths name family bold =
    let measure = WR.measureTextFor (load name)
    measure "MMMMMMMMMM" (style family bold), measure "iiiiiiiiii" (style family bold)
for name, family, bold in [ "reg", "Rbt Embed Mono", false; "reg", "Rbt Embed Mono", true; "bold-only", "Rbt Embed Mono", false; "names", "Rbt Outer", false ] do
    let m, i = widths name family bold
    check (sprintf "%s %s%s: the embedded face (monospaced)" name family (if bold then " bold" else "")) (abs (m - i) < 0.01 && m > 0.0) (sprintf "M %.3f i %.3f" m i)
for name, family in [ "restricted", "Rbt Embed Mono"; "names", "Rbt Inner" ] do
    let m, i = widths name family false
    check (sprintf "%s %s: not embedded - another font" name family) (m - i > 1.0) (sprintf "M %.3f i %.3f" m i)
// (Calibri: the machine's - or its stand-in's - not the embedded monospaced one, when this machine has it.)
if Rbt.OfficeNative.Fonts.hasFamily "Calibri" then
    let m, i = widths "calibri" "Calibri" false
    check "calibri: the machine's Calibri" (m - i > 1.0) (sprintf "M %.3f i %.3f" m i)
// A character the subset face lacks: in Calibri - Q and z of "Quiz" alone with one face embedded, all of
// "Quiz" (the run from Q) with four.
if Rbt.OfficeNative.Fonts.hasFamily "Calibri" then
    let calibri = WR.measureText "Quiz" (style "Calibri" false)
    let mono = WR.measureTextFor (load "reg") "uu" (style "Rbt Embed Mono" false)
    let calibriQz = WR.measureText "Q" (style "Calibri" false) + WR.measureText "z" (style "Calibri" false)
    let one = WR.measureTextFor (load "subset") "Quiz" (style "Rbt Embed Mono" false)
    let four = WR.measureTextFor (load "subset-all4") "Quiz" (style "Rbt Embed Mono" false)
    check "subset: Q and z in Calibri, ui embedded" (abs (one - (calibriQz + mono)) < 0.01) (sprintf "%.3f, expected %.3f" one (calibriQz + mono))
    check "subset-all4: the run from Q in Calibri" (abs (four - calibri) < 0.01) (sprintf "%.3f, expected %.3f" four calibri)
let paintStyle : Paint.TextPaintStyle =
    { Bold = false; Italic = false; Underline = false; AllCaps = false; SmallCaps = false; Strike = false
      FontSize = 10.0; CharacterSpacing = 0.0; FontFamily = Some "Rbt Embed Mono"; Color = "#000000"; Highlight = None; VertAlign = None }
let commands = WR.toRendererCommandFor (load "reg") (WR.createImageResolver ()) (Paint.DrawText ("Mi", paintStyle, 0.0, 10.0, 12.0))
check "reg: painted" (not commands.IsEmpty) ""
if failures > 0 then exit 1
printfn "native embedded fonts: %d checks ok" checks
