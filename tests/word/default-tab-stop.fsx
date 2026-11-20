// w:defaultTabStop of the settings part: the distance of the default tab
// stops. Word's is 720 twips (half an inch) when the part says none; files
// from Europe often say 708 (1.25 cm). It was not read: every document had
// its default stops every 48 px.
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
let package (settings: string option) (body: string) =
    use stream = new IO.MemoryStream()
    do
        use archive = new IO.Compression.ZipArchive(stream, IO.Compression.ZipArchiveMode.Create, true)
        let add (name: string) (text: string) =
            use writer = new IO.StreamWriter(archive.CreateEntry(name).Open(), Text.UTF8Encoding(false))
            writer.Write text
        add "[Content_Types].xml" """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"""
        add "word/document.xml" (sprintf "<w:document %s><w:body>%s<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr></w:body></w:document>" w body)
        match settings with
        | Some inner -> add "word/settings.xml" (sprintf "<w:settings %s><w:zoom w:percent=\"100\"/>%s<w:characterSpacingControl w:val=\"doNotCompress\"/></w:settings>" w inner)
        | None -> ()
    stream.ToArray()
let import settings body = package settings body |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let settingsOf (doc: Doc.Document) : Layout.LayoutSettings =
    let s = doc.PageSetup
    { Width = s.PageWidth; PageHeight = s.PageHeight; FontSize = 16.0; LineHeight = 24.0; PaddingLeft = s.MarginLeft; PaddingTop = s.MarginTop
      MarginRight = s.MarginRight; MarginBottom = s.MarginBottom; InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 3.5
      DefaultTabStop = s.DefaultTabStop; CompatibilityMode = s.CompatibilityMode; HangingPunctuation = Rbt.Word.Layout.NoHanging; SpaceShrink = 0.0 }
let layoutsOf (doc: Doc.Document) = Layout.layoutDocumentPages (settingsOf doc) measureText measureRange doc |> fst

/// The x of the text after each tab of the first paragraph.
let afterTabs (doc: Doc.Document) =
    let para = doc.Blocks |> List.choose Doc.tryGetParagraph |> List.head
    let layout = layoutsOf doc |> List.find (fun layout -> layout.BlockId = para.Id)
    layout.Lines.Head.Segments
    |> List.choose (function
        | Layout.TextSegment (text, _, x, _, _) when text.Length > 0 && text.[0] = 'x' -> Some (Math.Round(x, 2))
        | _ -> None)

let body = "<w:p><w:r><w:t>a</w:t></w:r><w:r><w:tab/><w:t>x</w:t></w:r><w:r><w:tab/><w:t>x</w:t></w:r><w:r><w:tab/><w:t>x</w:t></w:r></w:p>"
let european = import (Some "<w:defaultTabStop w:val=\"708\"/>") body
let plain = import None body
let unstated = import (Some "") body

check "read: w:defaultTabStop 708 twips is 47.2 px" (Math.Round(european.PageSetup.DefaultTabStop, 4)) 47.2
check "read: no settings part, Word's 720 twips (48 px)" plain.PageSetup.DefaultTabStop 48.0
check "read: a settings part without one, 48 px" unstated.PageSetup.DefaultTabStop 48.0
check "laid out: the text after each tab at the next stop of 708 twips" (afterTabs european) [47.2; 94.4; 141.6]
check "laid out: without one, of 720 twips" (afterTabs plain) [48.0; 96.0; 144.0]

// A stop of 1440 twips (an inch).
let wide = import (Some "<w:defaultTabStop w:val=\"1440\"/>") body
check "laid out: stops of 1440 twips" (afterTabs wide) [96.0; 192.0; 288.0]

printfn "word default tab stop: %d passed, %d failed" pass fail
if fail > 0 then exit 1
