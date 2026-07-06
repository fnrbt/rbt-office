#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office/src/Office.Formats/bin/Debug/netstandard2.0/OfficeFormat.dll"

open System

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label detail)

let cell value = XlsxWriter.simpleCell value
let formulaCell formula value = XlsxWriter.cellWithFormula formula value

let smallSheet: XlsxWriter.FormattedSheet =
    { Name = "Data"
      Rows =
        [ { Cells = [ cell "A & <x> \"q\" 's'"; cell "42" ] }
          { Cells = [ formulaCell "SUM(A1:B1)" "42" ] } ]
      MergedCells = [ XlsxWriter.createMergedRange 0 0 1 1 ] }

let xml = XlsxWriter.generateFormattedSheetXml smallSheet
ok "xlsx-escapes-inline-text"
   (xml.Contains("A &amp; &lt;x&gt; &quot;q&quot; &apos;s&apos;"))
   xml
ok "xlsx-formula-cell"
   (xml.Contains("<f>SUM(A1:B1)</f><v>42</v>"))
   xml
ok "xlsx-merge-cell"
   (xml.Contains("<mergeCell ref=\"A1:B2\"/>"))
   xml

let minimalXml = XlsxWriter.generateSheetXml [ [ "A & <x> \"q\" 's'"; "42" ] ]
ok "xlsx-minimal-escapes-inline-text"
   (minimalXml.Contains("A &amp; &lt;x&gt; &quot;q&quot; &apos;s&apos;"))
   minimalXml
ok "xlsx-minimal-number-cell"
   (minimalXml.Contains("<c r=\"B1\"><v>42</v></c>"))
   minimalXml

let workbookXml = XlsxWriter.generateWorkbookXml [ "One & <x>"; "Two" ]
ok "xlsx-workbook-escapes-sheet-name"
   (workbookXml.Contains("name=\"One &amp; &lt;x&gt;\"") && workbookXml.IndexOf("One &amp;") < workbookXml.IndexOf("Two"))
   workbookXml

let workbookRels = XlsxWriter.generateWorkbookRels 2
ok "xlsx-workbook-rels"
   (workbookRels.Contains("Id=\"rId2\"") && workbookRels.Contains("Target=\"worksheets/sheet2.xml\""))
   workbookRels

let contentTypes = XlsxWriter.generateContentTypesXlsx 2
ok "xlsx-content-types"
   (contentTypes.Contains("PartName=\"/xl/worksheets/sheet2.xml\""))
   contentTypes

let perfSheet: XlsxWriter.FormattedSheet =
    { Name = "Perf"
      Rows =
        [ for r in 1 .. 2000 ->
            { Cells =
                [ for c in 1 .. 20 ->
                    if c % 3 = 0 then cell (string (r * c))
                    else cell (sprintf "R%d C%d & <text>" r c) ] } ]
      MergedCells = [] }

XlsxWriter.generateFormattedSheetXml perfSheet |> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let before = GC.GetAllocatedBytesForCurrentThread()
let generated = XlsxWriter.generateFormattedSheetXml perfSheet
let allocatedMb = float (GC.GetAllocatedBytesForCurrentThread() - before) / 1024.0 / 1024.0
ok "xlsx-sheet-linear-alloc"
   (generated.Length > 1000000 && allocatedMb < 96.0)
   (sprintf "chars=%d allocated=%.1f MB" generated.Length allocatedMb)

let minimalPerfData =
    [ for r in 1 .. 2000 ->
        [ for c in 1 .. 20 ->
            if c % 3 = 0 then string (r * c)
            else sprintf "R%d C%d & <text>" r c ] ]

XlsxWriter.generateSheetXml minimalPerfData |> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let beforeMinimal = GC.GetAllocatedBytesForCurrentThread()
let generatedMinimal = XlsxWriter.generateSheetXml minimalPerfData
let allocatedMinimalMb = float (GC.GetAllocatedBytesForCurrentThread() - beforeMinimal) / 1024.0 / 1024.0
ok "xlsx-minimal-sheet-linear-alloc"
   (generatedMinimal.Length > 1000000 && allocatedMinimalMb < 96.0)
   (sprintf "chars=%d allocated=%.1f MB" generatedMinimal.Length allocatedMinimalMb)

let manySheets = [ for i in 1 .. 5000 -> sprintf "Sheet %d & <x>" i ]
XlsxWriter.generateWorkbookXml manySheets |> ignore
XlsxWriter.generateWorkbookRels manySheets.Length |> ignore
XlsxWriter.generateContentTypesXlsx manySheets.Length |> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let beforeMetadata = GC.GetAllocatedBytesForCurrentThread()
let metadataChars =
    (XlsxWriter.generateWorkbookXml manySheets).Length
    + (XlsxWriter.generateWorkbookRels manySheets.Length).Length
    + (XlsxWriter.generateContentTypesXlsx manySheets.Length).Length
let allocatedMetadataMb = float (GC.GetAllocatedBytesForCurrentThread() - beforeMetadata) / 1024.0 / 1024.0
ok "xlsx-metadata-linear-alloc"
   (metadataChars > 1000000 && allocatedMetadataMb < 48.0)
   (sprintf "chars=%d allocated=%.1f MB" metadataChars allocatedMetadataMb)

printfn ""
printfn "OfficeFormat XLSX writer tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
