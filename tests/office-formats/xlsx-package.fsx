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
      NameInfo = Map.empty; DefinedNames = imported.Context.Names
      Source = Some imported.Source
      Macros = None; MacroProject = None }
let modelsOf (imported: Bridge.ImportedWorkbook) = imported.Sheets |> Array.map (fun sheet -> sheet.Model)
let namesOf (imported: Bridge.ImportedWorkbook) = imported.Sheets |> Array.map (fun sheet -> sheet.Name)
let export (imported: Bridge.ImportedWorkbook) (sheets: SpreadsheetModel array) (names: string array) =
    XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (inputOf imported sheets names))
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

    // Rows inserted in the first sheet: its x14 data bar is a rule of the model and moves
    // with its cells; formulas moved, so the chain goes.
    let wb : Wb.Workbook = imported.Context
    let restructured, _ = Wb.applyStructureEdit (List.item 0 wb.SheetOrder) (InsertRows (2, 3)) wb
    let restructuredModels = wb.SheetOrder |> List.map (fun id -> (Map.find id restructured.Sheets).Model) |> List.toArray
    let written = export imported restructuredModels names
    let after = partsOf written
    whole "rows inserted" written
    ok "rows inserted: the x14 data bar of that sheet follows its cells"
        ((textOf after "xl/worksheets/sheet1.xml").Contains "<xm:sqref>E18:E25</xm:sqref>" && not ((textOf after "xl/worksheets/sheet1.xml").Contains "<xm:sqref>E15:E22</xm:sqref>")) ""
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

// ---- what kept parts say of a sheet follows the sheet -------------------------------
// The series of a chart, the source of a pivot cache and the formulas of kept extensions
// name sheets: they follow a rename, say #REF! of a deleted sheet, and a copied sheet
// has charts of its own that read the copy.
do
    let chartPaths (parts: Map<string, byte[]>) = parts |> Map.toList |> List.map fst |> List.filter (fun path -> Regex.IsMatch(path, "^xl/charts/chart\\d+\\.xml$")) |> List.sort
    let formulasOf (parts: Map<string, byte[]>) (path: string) = [ for m in Regex.Matches(textOf parts path, "<c:f>([^<]*)</c:f>") -> m.Groups.[1].Value ]
    let original = partsOf (bytesOf "real-world/chart-zoo.xlsx")
    let imported = Bridge.Import.importWorkbook (loadBytes (bytesOf "real-world/chart-zoo.xlsx"))
    let wb = imported.Context
    let zooId = wb.SheetOrder.Head
    let charts = chartPaths original
    ok "the fixture's charts read the sheet Zoo" (charts.Length >= 6 && charts |> List.forall (fun path -> formulasOf original path |> List.exists (fun f -> f.StartsWith "Zoo!"))) (sprintf "%A" charts)
    let unchanged = partsOf (export imported (modelsOf imported) (namesOf imported))
    ok "not renamed: the chart parts go back byte for byte" (charts |> List.forall (fun path -> Map.tryFind path unchanged = Map.tryFind path original)) ""

    // -- renamed (a name that has to be quoted, and escaped in XML)
    let renamedBytes = export imported (modelsOf imported) [| "Animals & Co" |]
    let renamed = partsOf renamedBytes
    whole "a sheet with charts renamed" renamedBytes
    let all = charts |> List.collect (formulasOf renamed)
    ok "renamed: every series reads the sheet by its new name" (all.Length > 10 && all |> List.forall (fun f -> f.StartsWith "'Animals &amp; Co'!$")) (sprintf "%A" (all |> List.truncate 3))
    eq "renamed: the references themselves stand" (all |> List.map (fun f -> f.Substring(f.IndexOf '!'))) (charts |> List.collect (formulasOf original) |> List.map (fun f -> f.Substring(f.IndexOf '!')))
    let again = loadBytes renamedBytes
    eq "renamed: read again - the sheet, with its charts" (again.Sheets |> List.map (fun sheet -> sheet.Name, sheet.Extras.Objects |> List.filter (fun item -> item.Kind = "chart") |> List.length)) [ "Animals & Co", charts.Length ]
    // ... and renamed a second time, from the written file
    let second = Bridge.Import.importWorkbook again
    let twice = partsOf (export second (modelsOf second) [| "Beasts" |])
    ok "renamed again from the written file" (chartPaths twice |> List.collect (formulasOf twice) |> List.forall (fun f -> f.StartsWith "Beasts!$")) (sprintf "%A" (chartPaths twice |> List.collect (formulasOf twice) |> List.truncate 3))

    // (A name that reads as a cell address is quoted, as Excel writes it.)
    let asCell = partsOf (export imported (modelsOf imported) [| "Zoo2" |])
    ok "renamed to a name that reads as a cell: quoted" (chartPaths asCell |> List.collect (formulasOf asCell) |> List.forall (fun f -> f.StartsWith "'Zoo2'!$")) ""

    // -- copied: the copy has chart parts of its own that read the copy
    let copied = Wb.duplicateSheet zooId "copy1" "Zoo (2)" 1 wb
    let sheetsOf (book: Wb.Workbook) = book.SheetOrder |> List.map (fun id -> (Map.find id book.Sheets).Model) |> List.toArray
    let namesIn (book: Wb.Workbook) = book.SheetOrder |> List.map (fun id -> (Map.find id book.Sheets).Name) |> List.toArray
    eq "copied: the copy's model has the kept objects" ((Map.find "copy1" copied.Sheets).Model.Drawings.Length) ((Map.find zooId wb.Sheets).Model.Drawings.Length)
    let copiedBytes = export imported (sheetsOf copied) (namesIn copied)
    let copiedParts = partsOf copiedBytes
    whole "a sheet with charts copied" copiedBytes
    let allCharts = chartPaths copiedParts
    eq "copied: twice the chart parts" allCharts.Length (2 * charts.Length)
    let copies = allCharts |> List.filter (fun path -> not (List.contains path charts))
    ok "copied: the source's charts are as they were" (charts |> List.forall (fun path -> Map.tryFind path copiedParts = Map.tryFind path original)) ""
    ok "copied: the copies read the copy" (copies |> List.collect (formulasOf copiedParts) |> List.forall (fun f -> f.StartsWith "'Zoo (2)'!$")) (sprintf "%A" (copies |> List.collect (formulasOf copiedParts) |> List.truncate 3))
    ok "copied: each copy has its relationships part and a content type" (copies |> List.forall (fun path ->
        let rels = path.Replace("xl/charts/", "xl/charts/_rels/") + ".rels"
        Map.containsKey rels copiedParts && (textOf copiedParts "[Content_Types].xml").Contains (sprintf "PartName=\"/%s\"" path))) ""
    let reread = loadBytes copiedBytes
    eq "copied: read again - both sheets have the charts" (reread.Sheets |> List.map (fun sheet -> sheet.Name, sheet.Extras.Objects |> List.filter (fun item -> item.Kind = "chart") |> List.length)) [ "Zoo", charts.Length; "Zoo (2)", charts.Length ]
    // The copy and the source are drawn from different drawing parts, each with relationships to its own charts.
    let drawingRels = copiedParts |> Map.toList |> List.filter (fun (path, _) -> path.StartsWith "xl/drawings/_rels/") |> List.map (fun (path, _) -> [ for m in Regex.Matches(textOf copiedParts path, "Target=\"/?(?:\\.\\./)?(?:xl/)?charts/(chart\\d+\\.xml)\"") -> m.Groups.[1].Value ] |> List.sort)
    ok "copied: two drawings, no chart shared between them" (drawingRels.Length = 2 && (Set.intersect (Set.ofList drawingRels.[0]) (Set.ofList drawingRels.[1])).IsEmpty && drawingRels |> List.forall (fun targets -> targets.Length = charts.Length)) (sprintf "%A" drawingRels)
    // A copy of the copy, the source renamed, the first copy deleted.
    let more = Wb.duplicateSheet "copy1" "copy2" "Third" 2 copied |> Wb.renameSheetEverywhere zooId "First"
    let moreBytes = export imported (sheetsOf more) (namesIn more)
    let moreParts = partsOf moreBytes
    whole "a copy of a copy, the source renamed" moreBytes
    eq "three sheets of charts" (chartPaths moreParts).Length (3 * charts.Length)
    let byName = chartPaths moreParts |> List.collect (formulasOf moreParts) |> List.countBy (fun f -> f.Substring(0, f.IndexOf '!'))  |> List.sortBy fst
    ok "each sheet's charts read that sheet" (byName |> List.map fst = [ "'Zoo (2)'"; "First"; "Third" ] && (byName |> List.map snd |> List.distinct).Length = 1) (sprintf "%A" byName)
    let without = match Wb.removeSheet "copy1" more with Some (after, _) -> after | None -> more
    let withoutBytes = export imported (sheetsOf without) (namesIn without)
    whole "a copy removed" withoutBytes
    eq "its chart parts are not written" (chartPaths (partsOf withoutBytes)).Length (2 * charts.Length)

// A sheet with pivot tables, and one with a slicer, copied: the copy has pivot tables of its own (parts
// of their own under another name, on the source's cache), the slicer stays with the source, the copy
// is a sheet of its own with its own drawing, and the package is whole - no slicer frame without a
// slicer, no relationship into nothing.
do
    let bytes = bytesOf "real-world/dashboard-ecommerce-sales.xlsx"
    let original = partsOf bytes
    let imported = Bridge.Import.importWorkbook (loadBytes bytes)
    let wb = imported.Context
    let sheetsOf (book: Wb.Workbook) = book.SheetOrder |> List.map (fun id -> (Map.find id book.Sheets).Model) |> List.toArray
    let namesIn (book: Wb.Workbook) = book.SheetOrder |> List.map (fun id -> (Map.find id book.Sheets).Name) |> List.toArray
    let pivotParts (parts: Map<string, byte[]>) = parts |> Map.toList |> List.map fst |> List.filter (fun path -> Regex.IsMatch(path, "^xl/pivotTables/pivotTable\\d+\\.xml$"))
    let slicerParts (parts: Map<string, byte[]>) = parts |> Map.toList |> List.map fst |> List.filter (fun path -> Regex.IsMatch(path, "^xl/slicers/slicer\\d+\\.xml$"))
    for name in [ "Waterfall Chart"; "KPI"; "Dashboard" ] do
        let index = imported.Sheets |> Array.findIndex (fun sheet -> sheet.Name = name)
        let sourceId = List.item index wb.SheetOrder
        let copied = Wb.duplicateSheet sourceId "copyOfIt" (name + " (2)") (index + 1) wb
        let copiedBytes = export imported (sheetsOf copied) (namesIn copied)
        let parts = partsOf copiedBytes
        whole (sprintf "%s copied" name) copiedBytes
        // The pivot tables of the source sheet, by the file's relationships.
        let sourcePivots =
            [ for m in Regex.Matches(textOf original ((fst (List.item index (loadBytes bytes).Sheets).Extras.SourcePart.Value).Replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels"), "Target=\"\\.\\./pivotTables/(pivotTable\\d+\\.xml)\"") -> "xl/pivotTables/" + m.Groups.[1].Value ]
        let newPivots = pivotParts parts |> List.filter (fun path -> not (Map.containsKey path original))
        eq (sprintf "%s copied: a pivot table part more for each of the source's, the slicers are the file's" name) (newPivots.Length, slicerParts parts) (sourcePivots.Length, slicerParts original)
        ok (sprintf "%s copied: the file's pivot table parts are as they were" name) (pivotParts original |> List.forall (fun path -> Map.tryFind path parts = Map.tryFind path original)) ""
        let rootOf (from: Map<string, byte[]>) (path: string) = Regex.Match(textOf from path, "<pivotTableDefinition\\b[^>]*>").Value
        let attr (name: string) (tag: string) = Regex.Match(tag, "\\s" + name + "=\"([^\"]*)\"").Groups.[1].Value
        // (Pivot tables on different sheets of the file may have the same name: a copy's is no other's.)
        let fileNames = pivotParts original |> List.map (fun path -> attr "name" (rootOf original path))
        let newNames = newPivots |> List.map (fun path -> attr "name" (rootOf parts path))
        ok (sprintf "%s copied: each copy has a name no other pivot table has" name)
            ((List.distinct newNames).Length = newNames.Length && newNames |> List.forall (fun copyName -> not (List.contains copyName fileNames))) (sprintf "%A" newNames)
        ok (sprintf "%s copied: each copy is on a cache of the file, has another uid, its relationships part and a content type" name)
            (newPivots |> List.forall (fun path ->
                let root = rootOf parts path
                let twin = sourcePivots |> List.tryFind (fun known -> attr "name" root = attr "name" (rootOf original known) + "_2")
                match twin with
                | Some known ->
                    attr "cacheId" root = attr "cacheId" (rootOf original known) && attr "xr:uid" root <> attr "xr:uid" (rootOf original known)
                    && Map.tryFind (path.Replace("xl/pivotTables/", "xl/pivotTables/_rels/") + ".rels") parts = Map.tryFind (known.Replace("xl/pivotTables/", "xl/pivotTables/_rels/") + ".rels") original
                    && (textOf parts "[Content_Types].xml").Contains (sprintf "PartName=\"/%s\"" path)
                    && Regex.Match(textOf parts path, "<location[^>]*>").Value = Regex.Match(textOf original known, "<location[^>]*>").Value
                | None -> false))
            (sprintf "%A" (newPivots |> List.map (rootOf parts)))
        // The copy's sheet part and what it leads to.
        let reread = loadBytes copiedBytes
        let copy = reread.Sheets |> List.find (fun sheet -> sheet.Name = name + " (2)")
        let source = reread.Sheets |> List.find (fun sheet -> sheet.Name = name)
        let partOf (sheet: XlsxReader.Worksheet) = fst sheet.Extras.SourcePart.Value
        let relsOf (path: string) = textOf parts (path.Replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels")
        ok (sprintf "%s copied: the copy lists its pivot tables, and no slicer" name)
            (Regex.Matches(relsOf (partOf copy), "/pivotTable\"").Count = sourcePivots.Length && newPivots |> List.forall (fun path -> (relsOf (partOf copy)).Contains ("/" + path))
             && not ((relsOf (partOf copy)).Contains "slicer") && not ((textOf parts (partOf copy)).Contains "slicerList"))
            (relsOf (partOf copy))
        eq (sprintf "%s copied: read again, the copy has its pivot tables where the source has them" name)
            (copy.Extras.PivotLocations |> List.map snd |> List.sort) (source.Extras.PivotLocations |> List.map snd |> List.sort)
        ok (sprintf "%s copied: the source still lists what it listed" name)
            (Regex.Matches(relsOf (partOf source), "pivotTable\"|slicer\"").Count = Regex.Matches(textOf original ((partOf (List.item index (loadBytes bytes).Sheets)).Replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels"), "pivotTable\"|slicer\"").Count) ""
        // No frame of a slicer in the copy's drawing; the source's drawing has its frames.
        let drawingOf (sheet: XlsxReader.Worksheet) = sheet.Extras.DrawingPart |> Option.map (textOf parts) |> Option.defaultValue ""
        let hadSlicers = Regex.Matches(drawingOf source, "<sle:slicer ").Count
        eq (sprintf "%s copied: no slicer frame in the copy (the source has %d)" name hadSlicers) (Regex.Matches(drawingOf copy, "<sle:slicer |timeslicer ").Count) 0
        eq (sprintf "%s copied: the copy has the source's other objects" name)
            (copy.Extras.Objects |> List.filter (fun item -> item.Kind = "chart") |> List.length, copy.Extras.Images.Length)
            (source.Extras.Objects |> List.filter (fun item -> item.Kind = "chart") |> List.length, source.Extras.Images.Length)
        eq (sprintf "%s copied: the copy has the cells" name) (Bridge.Import.importWorkbook reread).Sheets.[index + 1].Model.Cells.Count imported.Sheets.[index].Model.Cells.Count
    ok "copied: the workbook has slicer frames to leave behind"
        (imported.Sheets |> Array.exists (fun sheet -> sheet.Model.Drawings |> List.exists (fun d -> match d.Content with Rbt.Spreadsheet.Model.KeptObject (_, xml) -> xml.Contains "<sle:slicer " | _ -> false))) ""

// A pivot cache that reads a range of a sheet, a chart on another sheet, a link into a sheet, a sparkline:
// renamed and deleted sheets.
do
    let original = partsOf (bytesOf "real-world/dashboard-ecommerce-sales.xlsx")
    let cachePath = "xl/pivotCache/pivotCacheDefinition1.xml"
    let cacheXml = textOf original cachePath
    let source = Regex.Match(cacheXml, "<worksheetSource[^>]*>").Value
    let patched = original |> Map.add cachePath (Text.Encoding.UTF8.GetBytes (cacheXml.Replace(source, "<worksheetSource ref=\"A1:H200\" sheet=\"Data\"/>")))
    let workbook = loadBytes (zipOf patched)
    let imported = Bridge.Import.importWorkbook workbook
    let names = namesOf imported
    let renamedNames = names |> Array.map (fun name -> if name = "Data" then "Raw <Data>" else name)
    let bytes = export imported (modelsOf imported) renamedNames
    let parts = partsOf bytes
    whole "a sheet a pivot cache reads renamed" bytes
    eq "the pivot cache reads the sheet by its new name" (Regex.Match(textOf parts cachePath, "<worksheetSource[^>]*>").Value) "<worksheetSource ref=\"A1:H200\" sheet=\"Raw &lt;Data&gt;\"/>"
    eq "... and is read back so" ((loadBytes bytes).Preserved.PivotCacheSources |> List.filter (fun (path, _, _) -> path = cachePath)) [ cachePath, "Raw <Data>", "A1:H200" ]
    // The charts of the workbook that read the Data sheet follow too.
    let chartFormulas (from: Map<string, byte[]>) = [ for KeyValue (path, _) in from do if Regex.IsMatch(path, "^xl/charts/chart(Ex)?\\d+\\.xml$") then for m in Regex.Matches(textOf from path, "<c(?:x)?:f[^>]*>([^<]*)</c(?:x)?:f>") -> m.Groups.[1].Value ]
    let before = chartFormulas original
    let after = chartFormulas parts
    let readData (f: string) = f.StartsWith "Data!" || f.StartsWith "'Data'!"
    ok "charts: no series names the old sheet" (not (after |> List.exists readData) && after.Length = before.Length) (sprintf "%A" (after |> List.filter readData |> List.truncate 3))
    eq "charts: those that read it read the renamed sheet" (after |> List.filter (fun f -> f.StartsWith "'Raw &lt;Data&gt;'!") |> List.length) (before |> List.filter readData |> List.length)
    eq "charts: the other series are untouched" (after |> List.filter (fun f -> not (f.StartsWith "'Raw &lt;Data&gt;'!"))) (before |> List.filter (fun f -> not (readData f)))

// A sheet with pivot tables deleted or renamed: what names it elsewhere follows - the slicers that
// filter its pivot tables (by `tabId`), the charts drawn from them (`c:pivotSource`).
do
    let bytes = bytesOf "real-world/dashboard-sales-analysis.xlsx"
    let original = partsOf bytes
    let workbook = loadBytes bytes
    let imported = Bridge.Import.importWorkbook workbook
    let names = namesOf imported
    let caches (from: Map<string, byte[]>) = from |> Map.toList |> List.map fst |> List.filter (fun path -> Regex.IsMatch(path, "^xl/slicerCaches/slicerCache\\d+\\.xml$"))
    let tabIds (from: Map<string, byte[]>) = caches from |> List.collect (fun path -> [ for m in Regex.Matches(textOf from path, "<pivotTable tabId=\"(\\d+)\"") -> m.Groups.[1].Value ])
    let pivotSources (from: Map<string, byte[]>) =
        [ for KeyValue (path, _) in from do
            if Regex.IsMatch(path, "^xl/charts/chart\\d+\\.xml$") then
                let m = Regex.Match(textOf from path, "<c:pivotSource><c:name>([^<]*)</c:name>")
                if m.Success then yield path, m.Groups.[1].Value ]
    let idOf (name: string) = (workbook.Sheets |> List.find (fun sheet -> sheet.Name = name)).Extras.SheetId |> Option.get |> string
    ok "the file has what the check is about" (tabIds original |> List.contains (idOf "Line Pivot") && pivotSources original |> List.exists (fun (_, name) -> name.Contains "]Line Pivot!")) (sprintf "%A" (pivotSources original))
    // -- deleted
    let keep = names |> Array.mapi (fun index name -> index, name) |> Array.filter (fun (_, name) -> name <> "Line Pivot") |> Array.map fst
    let deletedBytes = export imported (keep |> Array.map (fun index -> (modelsOf imported).[index])) (keep |> Array.map (fun index -> names.[index]))
    let deleted = partsOf deletedBytes
    whole "a sheet with a pivot table deleted" deletedBytes
    eq "deleted: no slicer names a pivot table of the sheet, the others are as they were" (tabIds deleted) (tabIds original |> List.filter (fun id -> id <> idOf "Line Pivot"))
    ok "deleted: every tabId a slicer names is a sheet of the workbook"
        (let ids = [ for m in Regex.Matches(textOf deleted "xl/workbook.xml", "<sheet [^>]*sheetId=\"(\\d+)\"") -> m.Groups.[1].Value ] in tabIds deleted |> List.forall (fun id -> List.contains id ids)) ""
    eq "deleted: the charts of its pivot table name no pivot table, the others theirs"
        (pivotSources deleted) (pivotSources original |> List.filter (fun (_, name) -> not (name.Contains "]Line Pivot!")))
    ok "deleted: those charts are still there" (pivotSources original |> List.forall (fun (path, _) -> Map.containsKey path deleted)) ""
    // -- every sheet with a pivot table deleted: the slicers filter nothing, and say so
    let fewer = names |> Array.mapi (fun index name -> index, name) |> Array.filter (fun (_, name) -> not (name.EndsWith " Pivot")) |> Array.map fst
    let bareBytes = export imported (fewer |> Array.map (fun index -> (modelsOf imported).[index])) (fewer |> Array.map (fun index -> names.[index]))
    let bare = partsOf bareBytes
    whole "every sheet with a pivot table deleted" bareBytes
    ok "all deleted: the slicer caches list no pivot table" (caches bare = caches original && caches bare |> List.forall (fun path -> not ((textOf bare path).Contains "<pivotTable"))) (textOf bare (List.head (caches bare)))
    eq "all deleted: no chart names a pivot table" (pivotSources bare) []
    // -- renamed
    let renamedBytes = export imported (modelsOf imported) (names |> Array.map (fun name -> if name = "Line Pivot" then "Lines & Co" else name))
    let renamed = partsOf renamedBytes
    whole "a sheet with a pivot table renamed" renamedBytes
    eq "renamed: the charts name the pivot table on the renamed sheet"
        (pivotSources renamed) (pivotSources original |> List.map (fun (path, name) -> path, name.Replace("]Line Pivot!", "]Lines &amp; Co!")))
    eq "renamed: the slicers are as they were" (tabIds renamed) (tabIds original)

printfn "xlsx-package: %d passed, %d failed" pass fail
if fail > 0 then exit 1
