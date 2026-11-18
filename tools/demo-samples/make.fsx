// The macro samples of the rbt-office demo page (github.com/fnrbt/rbt-office-demo, public/samples):
//
//   dotnet build src/Word && dotnet fsi tools/demo-samples/make.fsx <out dir>
//
// snake.xlsm, sierpinski.xlsm, mandelbrot.xlsm and rainbow.docm: blank macro-enabled files with
// source-only VBA projects (VbaProject.writeProjectModules; Office compiles the source when it opens them).
// Each workbook sets itself up in Workbook_Open (its buttons: shapes whose OnAction runs a macro).
// The macros from the web (tests/macros/showcase: snake CC0, Sierpinski GFDL 1.3) keep their source and
// licence lines; the test harness's own header lines are left out. The rest is ours (MIT).
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"

open System
open System.IO
open System.IO.Compression

let out = fsi.CommandLineArgs |> Array.tryItem 1 |> Option.defaultWith (fun () -> failwith "usage: make.fsx <out dir>")
Directory.CreateDirectory out |> ignore
let here = __SOURCE_DIRECTORY__
let showcase = Path.Combine(here, "../../tests/macros/showcase")

/// A module's source: the test harness's directive lines (' Entry:, ' Ticks: ...) left out.
let source (path: string) =
    let harness = [ "' Entry:"; "' Before:"; "' Ticks:"; "' Press:"; "' Autopilot:"; "' Input:"; "' Limit:" ]
    File.ReadAllLines path
    |> Array.filter (fun line -> not (harness |> List.exists (fun h -> line.StartsWith h)))
    |> String.concat "\n"

let standard (path: string) =
    let text = source path
    let m = Text.RegularExpressions.Regex.Match(text, "Attribute\\s+VB_Name\\s*=\\s*\"([^\"]+)\"")
    // (The Attribute line first, as VBA exports have it: comments before it would be module text.)
    let name = if m.Success then m.Groups.[1].Value else Path.GetFileNameWithoutExtension(path).Split([|'.'; '-'|]) |> Array.map (fun w -> string (Char.ToUpperInvariant w.[0]) + w.Substring 1) |> String.concat ""
    let body = (if m.Success then text.Replace(m.Value, "") else text).TrimStart('\n')
    name, VbaProject.StandardModule, sprintf "Attribute VB_Name = \"%s\"\n%s" name body
let document (name: string) (path: string) = name, VbaProject.DocumentModule, source path

let replaceEntry (package: byte[]) (entry: string) (data: byte[]) =
    use stream = new MemoryStream()
    stream.Write(package, 0, package.Length)
    (
        use archive = new ZipArchive(stream, ZipArchiveMode.Update, true)
        archive.GetEntry(entry) |> Option.ofObj |> Option.iter (fun e -> e.Delete())
        use target = archive.CreateEntry(entry).Open()
        target.Write(data, 0, data.Length)
    )
    stream.ToArray()

let xlsm (file: string) modules =
    let blank = File.ReadAllBytes(Path.Combine(here, "../../tests/test-documents-excel/pandas-blank.xlsx"))
    let injected = VbaProject.injectIntoXlsx blank "Module1" ""
    File.WriteAllBytes(Path.Combine(out, file), replaceEntry injected "xl/vbaProject.bin" (VbaProject.writeProjectModules modules))
    printfn "%s" file

xlsm "snake.xlsm"
    [ document "ThisWorkbook" (Path.Combine(here, "Snake.ThisWorkbook.bas"))
      standard (Path.Combine(showcase, "snake.Main.bas"))
      standard (Path.Combine(showcase, "snake.KeyBind.bas"))
      standard (Path.Combine(here, "Snake.Board.bas")) ]
xlsm "sierpinski.xlsm"
    [ document "ThisWorkbook" (Path.Combine(here, "Sierpinski.ThisWorkbook.bas"))
      standard (Path.Combine(showcase, "sierpinski-pentagon.bas")) ]
xlsm "mandelbrot.xlsm"
    [ document "ThisWorkbook" (Path.Combine(here, "Mandelbrot.ThisWorkbook.bas"))
      standard (Path.Combine(here, "Mandelbrot.bas")) ]

// The Word sample: a short document of its own and the macro.
let docm (file: string) (paragraphs: string list) modules =
    let escape (s: string) = s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
    let body =
        paragraphs |> List.map (fun p -> sprintf "<w:p><w:r><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\"/><w:sz w:val=\"28\"/></w:rPr><w:t xml:space=\"preserve\">%s</w:t></w:r></w:p>" (escape p))
        |> String.concat ""
    let thisDocument =
        "ThisDocument", VbaProject.DocumentModule,
        String.concat "\n"
            [ "Attribute VB_Name = \"ThisDocument\""; "Attribute VB_Base = \"1Normal.ThisDocument\""
              "Attribute VB_GlobalNameSpace = False"; "Attribute VB_Creatable = False"; "Attribute VB_PredeclaredId = True"
              "Attribute VB_Exposed = True"; "Attribute VB_TemplateDerived = True"; "Attribute VB_Customizable = True"; "" ]
    let parts =
        [ "[Content_Types].xml",
          "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.ms-word.document.macroEnabled.main+xml\"/></Types>"
          "_rels/.rels",
          "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>"
          "word/_rels/document.xml.rels",
          "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" Target=\"vbaProject.bin\"/></Relationships>"
          "word/document.xml",
          sprintf "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>%s<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr></w:body></w:document>" body ]
    use stream = new MemoryStream()
    (
        use archive = new ZipArchive(stream, ZipArchiveMode.Create, true)
        for name, text in parts do
            use w = new StreamWriter(archive.CreateEntry(name).Open())
            w.Write(text)
        let project = VbaProject.writeProjectModules (thisDocument :: modules)
        use target = archive.CreateEntry("word/vbaProject.bin").Open()
        target.Write(project, 0, project.Length)
    )
    File.WriteAllBytes(Path.Combine(out, file), stream.ToArray())
    printfn "%s" file

docm "rainbow.docm"
    [ "Macros, in the browser"
      "This document carries a VBA macro, Rainbow. Enable macros and run it: every character takes a colour of a rainbow, and the title grows."
      "The macro is the VBA a Word user would write - ActiveDocument, Characters, Paragraphs, Font - run in the page by rbt-office's VBA engine against its object model of this document. Nothing is installed and nothing is sent anywhere." ]
    [ standard (Path.Combine(here, "Rainbow.bas")) ]
