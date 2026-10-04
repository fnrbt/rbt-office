#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// What a save keeps of the opened package besides the workbook's own parts:
// the document properties (docProps/core.xml, app.xml, custom.xml), printer
// settings (with the pageSetup r:id that names them), the calculation chain
// (only while no formula cell was added, removed or moved), sparklines and
// x14 conditional formats (while the sheet was not restructured), and the
// source range of a pivot cache (which follows inserted rows). And that every
// written package is whole: every relationship target exists, every part has
// a content type, nothing is left without a relationship leading to it.

open System
open System.IO
open System.IO.Compression
open System.Text.RegularExpressions
open Rbt.Spreadsheet.Model
module Bridge = Rbt.Spreadsheet.XlsxBridge
module Wb = Rbt.Spreadsheet.Workbook

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-34s  %s" label detail)
let eq label (got: 'a) (expected: 'a) =
    ok label (got = expected) (sprintf "expected %A  got %A" expected got)

let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let bytesOf (relative: string) = File.ReadAllBytes(Path.Combine(root, relative))
let loadBytes (bytes: byte[]) = XlsxReader.loadBytesNative bytes |> Option.get
let partsOf (bytes: byte[]) : Map<string, byte[]> =
    use archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read)
    [ for entry in archive.Entries ->
        use stream = entry.Open()
        use buffer = new MemoryStream()
        stream.CopyTo buffer
        entry.FullName, buffer.ToArray() ]
    |> Map.ofList
let textOf (parts: Map<string, byte[]>) (path: string) =
    match Map.tryFind path parts with
    | Some data -> Text.Encoding.UTF8.GetString data
    | None -> ""
let zipOf (parts: Map<string, byte[]>) : byte[] =
    use buffer = new MemoryStream()
    (use archive = new ZipArchive(buffer, ZipArchiveMode.Create, true)
     for KeyValue(path, data) in parts do
         let entry = archive.CreateEntry path
         use stream = entry.Open()
         stream.Write(data, 0, data.Length))
    buffer.ToArray()
let inputOf (imported: Bridge.ImportedWorkbook) (sheets: SpreadsheetModel array) (names: string array) : Bridge.ExportInput =
    { Sheets = sheets
      Names = names
      Visible = Array.create sheets.Length true
      ActiveSheetIndex = 0
      DefinedNames = imported.Context.Names
      Source = Some imported.Source
      Macros = None }
let modelsOf (imported: Bridge.ImportedWorkbook) = imported.Sheets |> Array.map (fun sheet -> sheet.Model)
let namesOf (imported: Bridge.ImportedWorkbook) = imported.Sheets |> Array.map (fun sheet -> sheet.Name)
let export (imported: Bridge.ImportedWorkbook) (sheets: SpreadsheetModel array) (names: string array) =
    XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (inputOf imported sheets names))
let addr row col : CellAddress = { Row = row; Col = col }
let text value = { Value = CellValue.String value; Format = defaultFormat; Comment = None }

// ---- package integrity ----------------------------------------------------------
let attribute (name: string) (tag: string) =
    let m = Regex.Match(tag, "\\s" + Regex.Escape name + "=\"([^\"]*)\"")
    if m.Success then Some (m.Groups.[1].Value.Replace("&amp;", "&")) else None
let resolve (baseDir: string) (target: string) =
    let full = if target.StartsWith "/" || baseDir = "" then target else baseDir + "/" + target
    let segments = ResizeArray<string>()
    for segment in full.Split('/') do
        if segment = ".." then (if segments.Count > 0 then segments.RemoveAt(segments.Count - 1))
        elif segment <> "" && segment <> "." then segments.Add segment
    String.Join("/", segments)
/// What is wrong with a package: relationships to parts that are not there,
/// parts without a content type, content types of parts that are not there,
/// parts nothing leads to.
let integrityProblems (parts: Map<string, byte[]>) : string list =
    let problems = ResizeArray<string>()
    let types = textOf parts "[Content_Types].xml"
    let defaults = [ for m in Regex.Matches(types, "<Default\\b[^>]*>") -> (attribute "Extension" m.Value |> Option.defaultValue "").ToLowerInvariant() ] |> Set.ofList
    let overrides = [ for m in Regex.Matches(types, "<Override\\b[^>]*>") -> (attribute "PartName" m.Value |> Option.defaultValue "").TrimStart('/') ]
    for path in overrides do
        if not (parts.ContainsKey path) then problems.Add ("content type for a missing part: " + path)
    if (List.distinct overrides).Length <> overrides.Length then problems.Add "a part has two content types"
    let overridden = Set.ofList overrides
    let reached = Collections.Generic.HashSet<string>()
    for KeyValue(path, data) in parts do
        if path.EndsWith ".rels" then
            let directory = Path.GetDirectoryName(path).Replace('\\', '/')
            let ownerDirectory = if directory.EndsWith "_rels" then directory.Substring(0, directory.Length - 5).TrimEnd('/') else directory
            let owner = (if ownerDirectory = "" then "" else ownerDirectory + "/") + Path.GetFileName(path).Replace(".rels", "")
            if owner <> "" && not (parts.ContainsKey owner) then problems.Add ("relationships of a missing part: " + path)
            let ids = Collections.Generic.HashSet<string>()
            for m in Regex.Matches(Text.Encoding.UTF8.GetString data, "<Relationship\\b[^>]*>") do
                match attribute "Id" m.Value with
                | Some id when not (ids.Add id) -> problems.Add (sprintf "%s: relationship id %s twice" path id)
                | _ -> ()
                if attribute "TargetMode" m.Value <> Some "External" then
                    let target = resolve ownerDirectory (attribute "Target" m.Value |> Option.defaultValue "")
                    reached.Add target |> ignore
                    if not (parts.ContainsKey target) then problems.Add (sprintf "%s: target %s is missing" path target)
        elif path <> "[Content_Types].xml" then
            let extension = Path.GetExtension(path).TrimStart('.').ToLowerInvariant()
            if not (overridden.Contains path) && not (defaults.Contains extension) then problems.Add ("no content type: " + path)
    for KeyValue(path, _) in parts do
        if not (path.EndsWith ".rels") && path <> "[Content_Types].xml" && not (reached.Contains path) then problems.Add ("orphan: " + path)
    List.ofSeq problems
let whole label (bytes: byte[]) =
    let problems = integrityProblems (partsOf bytes)
    ok (label + ": package is whole") problems.IsEmpty (String.concat "; " (problems |> List.truncate 6))
    // The relationship ids a sheet's XML names are relationships of the sheet.
    let parts = partsOf bytes
    for KeyValue(path, data) in parts do
        if path.StartsWith "xl/worksheets/sheet" then
            let xml = Text.Encoding.UTF8.GetString data
            let tail = xml.Substring(max 0 (xml.LastIndexOf "</sheetData>"))
            let rels = textOf parts (path.Replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels")
            for m in Regex.Matches(tail, "\\sr:id=\"([^\"]*)\"") do
                let id = m.Groups.[1].Value
                ok (sprintf "%s: %s r:id %s has a relationship" label path id) (rels.Contains (sprintf "Id=\"%s\"" id)) ""

// ---- the dashboard: document properties, printer settings, calculation chain ----
do
    let original = bytesOf "real-world/dashboard-sales-analysis.xlsx"
    let before = partsOf original
    eq "the fixture is whole itself" (integrityProblems before) []
    let imported = Bridge.Import.importWorkbook (loadBytes original)
    let models = modelsOf imported
    let names = namesOf imported
    let dataIndex = names |> Array.findIndex (fun name -> name = "Data")

    // A value edit: nothing about formulas changed.
    let edited = Array.copy models
    edited.[0] <- setCell (addr 40 0) (text "edited") edited.[0]
    let written = export imported edited names
    let after = partsOf written
    whole "value edit" written
    for part in [ "docProps/core.xml"; "docProps/app.xml"; "xl/printerSettings/printerSettings1.bin"; "xl/calcChain.xml" ] do
        eq (part + " carried over") (Map.tryFind part after) (Map.tryFind part before)
    let rootRels = textOf after "_rels/.rels"
    ok "root relationships name the properties" (rootRels.Contains "docProps/core.xml" && rootRels.Contains "docProps/app.xml" && rootRels.Contains "xl/workbook.xml") rootRels
    let types = textOf after "[Content_Types].xml"
    ok "content types of the properties" (types.Contains "/docProps/core.xml" && types.Contains "/docProps/app.xml") ""
    ok "content type and relationship of the chain" (types.Contains "/xl/calcChain.xml" && (textOf after "xl/_rels/workbook.xml.rels").Contains "calcChain.xml") ""
    let dataSheet = sprintf "xl/worksheets/sheet%d.xml" (dataIndex + 1)
    let pageSetup = Regex.Match(textOf after dataSheet, "<pageSetup[^>]*>").Value
    let relId = attribute "r:id" pageSetup |> Option.defaultValue ""
    ok "pageSetup names the printer settings" (relId <> "") pageSetup
    let dataRels = textOf after (sprintf "xl/worksheets/_rels/sheet%d.xml.rels" (dataIndex + 1))
    ok "and the sheet has that relationship" (Regex.IsMatch(dataRels, sprintf "Id=\"%s\"[^>]*printerSettings1\\.bin" (Regex.Escape relId))) dataRels
    // x14 conditional formats of the first sheet: kept, and still linked from its rules.
    let firstBefore = textOf before "xl/worksheets/sheet1.xml"
    let firstAfter = textOf after "xl/worksheets/sheet1.xml"
    ok "the fixture has x14 conditional formats" (firstBefore.Contains "x14:conditionalFormattings") ""
    ok "x14 conditional formats kept" (firstAfter.Contains "x14:conditionalFormattings") ""
    let x14Count (xml: string) = Regex.Matches(xml, "<x14:cfRule\\b").Count
    eq "all of them" (x14Count firstAfter) (x14Count firstBefore)

    // A formula added: the chain goes, with its relationship and content type.
    let withFormula = Array.copy models
    withFormula.[0] <- setCell (addr 41 0) { Value = CellValue.Formula ("=1+1", CellValue.Number 2.0); Format = defaultFormat; Comment = None } withFormula.[0]
    let written = export imported withFormula names
    let after = partsOf written
    whole "formula added" written
    ok "formula added: no calcChain part" (not (after.ContainsKey "xl/calcChain.xml")) ""
    ok "formula added: no calcChain relationship or content type"
        (not ((textOf after "xl/_rels/workbook.xml.rels").Contains "calcChain") && not ((textOf after "[Content_Types].xml").Contains "calcChain")) ""
    ok "formula added: the properties stay" (after.ContainsKey "docProps/core.xml" && after.ContainsKey "docProps/app.xml") ""

    // Rows inserted in the first sheet: its x14 formats name cells that moved, so they go
    // (the editor's own rules are written); formulas moved, so the chain goes.
    let wb : Wb.Workbook = imported.Context
    let restructured, _ = Wb.applyStructureEdit (List.item 0 wb.SheetOrder) (InsertRows (2, 3)) wb
    let restructuredModels = wb.SheetOrder |> List.map (fun id -> (Map.find id restructured.Sheets).Model) |> List.toArray
    let written = export imported restructuredModels names
    let after = partsOf written
    whole "rows inserted" written
    ok "rows inserted: x14 formats of that sheet are not written stale" (not ((textOf after "xl/worksheets/sheet1.xml").Contains "x14:conditionalFormattings")) ""
    ok "rows inserted: the properties and printer settings stay" (after.ContainsKey "docProps/core.xml" && after.ContainsKey "xl/printerSettings/printerSettings1.bin") ""

    // The Data sheet removed: its printer settings go with it, nothing is orphaned.
    let keep = [| for i in 0 .. models.Length - 1 do if i <> dataIndex then yield i |]
    let written = export imported (keep |> Array.map (fun i -> models.[i])) (keep |> Array.map (fun i -> names.[i]))
    let after = partsOf written
    whole "sheet removed" written
    ok "sheet removed: its printer settings are gone" (not (after.ContainsKey "xl/printerSettings/printerSettings1.bin")) ""
    ok "sheet removed: no content type for them" (not ((textOf after "[Content_Types].xml").Contains "printerSettings1.bin")) ""
    ok "sheet removed: the calculation chain is gone" (not (after.ContainsKey "xl/calcChain.xml")) ""

    // Only the first sheet kept: the pivot tables of the others go.
    let written = export imported [| models.[0] |] [| names.[0] |]
    whole "one sheet kept" written

// ---- custom properties, several printer settings ----------------------------------
for file in [ "real-world/milk-street-office.xlsx"; "real-world/lyric-multifamily.xlsx"; "real-world/jumeirah-hotel.xlsx"; "real-world/saopaulo-condos.xlsx"
              "real-world/dcf-stock-analysis.xlsx"; "real-world/chart-zoo.xlsx"; "real-world/conditional-formatting-tutorial.xlsx"; "test-documents-excel/synthetic-sparklines.xlsx" ] do
    let original = bytesOf file
    let before = partsOf original
    let imported = Bridge.Import.importWorkbook (loadBytes original)
    let models = modelsOf imported
    let names = namesOf imported
    let written = export imported models names
    let after = partsOf written
    let label = Path.GetFileNameWithoutExtension file
    whole label written
    for KeyValue(path, data) in before do
        if path.StartsWith "docProps/" || path.StartsWith "xl/printerSettings/" || path = "xl/calcChain.xml" then
            ok (sprintf "%s: %s carried over" label path) (Map.tryFind path after = Some data) ""
    let count (pattern: string) (parts: Map<string, byte[]>) =
        parts |> Map.toList |> List.sumBy (fun (path, data) -> if path.StartsWith "xl/worksheets/sheet" then Regex.Matches(Text.Encoding.UTF8.GetString data, pattern).Count else 0)
    eq (label + ": pageSetup r:id count") (count "<pageSetup[^>]*r:id=" after) (count "<pageSetup[^>]*r:id=" before)
    eq (label + ": sparkline groups") (count "<x14:sparklineGroup\\b" after) (count "<x14:sparklineGroup\\b" before)
    // Unedited and reopened: the same again (the kept things survive a second save).
    let again = Bridge.Import.importWorkbook (loadBytes written)
    let rewritten = export again (modelsOf again) (namesOf again)
    whole (label + " saved twice") rewritten
    eq (label + ": kept parts after a second save")
        (partsOf rewritten |> Map.toList |> List.map fst |> List.filter (fun p -> p.StartsWith "docProps/" || p.StartsWith "xl/printerSettings/" || p = "xl/calcChain.xml"))
        (after |> Map.toList |> List.map fst |> List.filter (fun p -> p.StartsWith "docProps/" || p.StartsWith "xl/printerSettings/" || p = "xl/calcChain.xml"))

// ---- sparklines go when their sheet is restructured ---------------------------------
do
    let original = bytesOf "test-documents-excel/synthetic-sparklines.xlsx"
    let imported = Bridge.Import.importWorkbook (loadBytes original)
    let wb = imported.Context
    let sheetXml (bytes: byte[]) = partsOf bytes |> Map.toList |> List.filter (fun (path, _) -> path.StartsWith "xl/worksheets/sheet") |> List.map (fun (_, data) -> Text.Encoding.UTF8.GetString data) |> String.concat ""
    ok "the fixture has sparklines" ((sheetXml original).Contains "sparklineGroups") ""
    let edited = modelsOf imported
    edited.[0] <- setCell (addr 30 8) (text "edited") edited.[0]
    ok "a cell edit keeps them" ((sheetXml (export imported edited (namesOf imported))).Contains "sparklineGroups") ""
    let restructured, undo = Wb.applyStructureEdit (List.item 0 wb.SheetOrder) (DeleteCols (0, 1)) wb
    let models = wb.SheetOrder |> List.map (fun id -> (Map.find id restructured.Sheets).Model) |> List.toArray
    let written = export imported models (namesOf imported)
    whole "sparklines, column deleted" written
    ok "a deleted column drops them" (not ((sheetXml written).Contains "sparklineGroups")) ""
    let undone = Wb.undoStructureEdit undo restructured
    let models = wb.SheetOrder |> List.map (fun id -> (Map.find id undone.Sheets).Model) |> List.toArray
    ok "undoing the deletion keeps them" ((sheetXml (export imported models (namesOf imported))).Contains "sparklineGroups") ""

// ---- the source range of a pivot cache follows its rows ----------------------------
do
    // The fixture's cache reads a table; give it a plain range of the Data sheet instead.
    let original = partsOf (bytesOf "real-world/dashboard-ecommerce-sales.xlsx")
    let cachePath = "xl/pivotCache/pivotCacheDefinition1.xml"
    let cacheXml = textOf original cachePath
    let source = Regex.Match(cacheXml, "<worksheetSource[^>]*>").Value
    ok "the fixture's cache has a worksheet source" (source <> "") ""
    let patched = original |> Map.add cachePath (Text.Encoding.UTF8.GetBytes (cacheXml.Replace(source, "<worksheetSource ref=\"A1:H200\" sheet=\"Data\"/>")))
    let workbook = loadBytes (zipOf patched)
    eq "the reader finds the source" (workbook.Preserved.PivotCacheSources |> List.filter (fun (path, _, _) -> path = cachePath)) [ cachePath, "Data", "A1:H200" ]
    eq "pivotCacheSourceOf: a named source is not a range" (XlsxReader.pivotCacheSourceOf "<x><worksheetSource name=\"Table_1\"/></x>") None
    eq "withPivotCacheSource" (XlsxReader.withPivotCacheSource "B3:C9" "<x><worksheetSource ref=\"A1:H200\" sheet=\"Data\"/></x>") "<x><worksheetSource ref=\"B3:C9\" sheet=\"Data\"/></x>"
    let imported = Bridge.Import.importWorkbook workbook
    let wb = imported.Context
    let names = namesOf imported
    let dataId = wb.SheetOrder |> List.find (fun id -> (Map.find id wb.Sheets).Name = "Data")
    let sourceOf (bytes: byte[]) = Regex.Match(textOf (partsOf bytes) cachePath, "<worksheetSource[^>]*>").Value
    eq "unedited: the range as it was" (sourceOf (export imported (modelsOf imported) names)) "<worksheetSource ref=\"A1:H200\" sheet=\"Data\"/>"
    let inserted, _ = Wb.applyStructureEdit dataId (InsertRows (0, 2)) wb
    let models = wb.SheetOrder |> List.map (fun id -> (Map.find id inserted.Sheets).Model) |> List.toArray
    let written = export imported models names
    eq "rows inserted above: the range moves down" (sourceOf written) "<worksheetSource ref=\"A3:H202\" sheet=\"Data\"/>"
    whole "pivot cache source moved" written
    let widened, _ = Wb.applyStructureEdit dataId (InsertCols (2, 1)) wb
    let models = wb.SheetOrder |> List.map (fun id -> (Map.find id widened.Sheets).Model) |> List.toArray
    eq "a column inserted inside: the range grows" (sourceOf (export imported models names)) "<worksheetSource ref=\"A1:I200\" sheet=\"Data\"/>"

printfn "xlsx-package: %d passed, %d failed" pass fail
if fail > 0 then exit 1
