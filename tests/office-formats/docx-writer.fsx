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

printfn ""
printfn "OfficeFormat DOCX writer tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
