#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"

open System

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label detail)

let xml = DocxWriter.generateDocumentXml [ "A & <x> \"q\" 's'"; "second" ]
ok "docx-minimal-escapes-text"
   (xml.Contains("A &amp; &lt;x&gt; &quot;q&quot; &apos;s&apos;"))
   xml
ok "docx-minimal-keeps-order"
   (xml.IndexOf("A &amp;") >= 0 && xml.IndexOf("second") > xml.IndexOf("A &amp;"))
   xml

let paragraphs = [ for i in 1 .. 10000 -> sprintf "Paragraph %d alpha & <beta> gamma" i ]
DocxWriter.generateDocumentXml paragraphs |> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let before = GC.GetAllocatedBytesForCurrentThread()
let generated = DocxWriter.generateDocumentXml paragraphs
let allocatedMb = float (GC.GetAllocatedBytesForCurrentThread() - before) / 1024.0 / 1024.0
ok "docx-minimal-linear-alloc"
   (generated.Length > 500000 && allocatedMb < 64.0)
   (sprintf "chars=%d allocated=%.1f MB" generated.Length allocatedMb)

// Control characters XML 1.0 cannot carry must not reach a written part.
// (the lone surrogate is built at run time: a literal one is replaced by U+FFFD at compile time)
let dirty = "bell\u0007 vt\u000B ff\u000C esc\u001B nul\u0000 tab\t lf\n ok \uFFFE lone" + string (char 0xD83D) + " pair\uD83D\uDE00 end"
let cleaned = XmlGenerator.stripInvalidXmlChars dirty
ok "xml-strips-control-chars"
   (cleaned = "bell vt ff esc nul tab\t lf\n ok  lone pair\uD83D\uDE00 end")
   (sprintf "%A" cleaned)
let controlXml = DocxWriter.generateDocumentXml [ dirty ]
let controlXmlParses =
    try (System.Xml.Linq.XDocument.Parse controlXml).Root <> null
    with ex -> (printfn "  %s" ex.Message; false)
ok "docx-control-chars-well-formed" controlXmlParses controlXml
let attrXml = XmlGenerator.selfClosing "x" [ "a", "one\ttwo\nthree\u0001" ]
ok "xml-attribute-whitespace-as-refs"
   (attrXml = "<x a=\"one&#9;two&#10;three\"/>")
   attrXml
ok "xml-escape-null-is-empty" (XmlGenerator.escapeXml null = "") "null"

printfn ""
printfn "OfficeFormat DOCX writer tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
