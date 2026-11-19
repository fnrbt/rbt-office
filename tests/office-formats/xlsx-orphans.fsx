#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// A deleted sheet takes with it what only it used: a slicer cache whose slicers were all on
// deleted sheets, a pivot cache whose pivot tables were - the part (with its records), the
// workbook's relationship to it, its entry in workbook.xml (`<pivotCaches>`, the list of slicer
// caches in the extension list), its content type, and a slicer cache's defined name. What a
// sheet that stays still uses stays, as the file has it.
//
// Usage: dotnet fsi tests/office-formats/xlsx-orphans.fsx

open System
open System.IO
open System.IO.Compression
open System.Text.RegularExpressions
module Bridge = Rbt.Spreadsheet.XlsxBridge
module M = Rbt.Spreadsheet.Model

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-34s  %s" label detail)

let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let partsOf (bytes: byte[]) : Map<string, byte[]> =
    use archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read)
    [ for entry in archive.Entries do
        if not (entry.FullName.EndsWith "/") then
            use stream = entry.Open()
            use buffer = new MemoryStream()
            stream.CopyTo buffer
            yield entry.FullName, buffer.ToArray() ]
    |> Map.ofList
let text (data: byte[]) = Text.Encoding.UTF8.GetString data
let values (pattern: string) (xml: string) = [ for found in Regex.Matches(xml, pattern) -> found.Groups.[1].Value ]
let short (value: string) = value.Substring(0, min 300 value.Length)

let bytes = File.ReadAllBytes(Path.Combine(root, "real-world/dashboard-ecommerce-sales.xlsx"))
let wb = XlsxReader.loadBytesNative bytes |> Option.get
let book = Bridge.Import.importWorkbook wb
let file0 = partsOf bytes
/// The workbook saved without the sheets named.
let without (names: string list) =
    let keep = [| for at in 0 .. book.Sheets.Length - 1 do if not (List.contains book.Sheets.[at].Name names) then yield at |]
    let input : Bridge.ExportInput =
        { Sheets = keep |> Array.map (fun at -> book.Sheets.[at].Model); Names = keep |> Array.map (fun at -> book.Sheets.[at].Name); Visible = keep |> Array.map (fun at -> book.Sheets.[at].Visible)
          ActiveSheetIndex = 0; NameInfo = book.Context.NameInfo; DefinedNames = book.Context.Names; Source = Some book.Source; Macros = None; MacroProject = None }
    XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut input)
let slicerCaches (parts: Map<string, byte[]>) =
    parts |> Map.toList |> List.choose (fun (path, data) -> if path.StartsWith "xl/slicerCaches/" && path.EndsWith ".xml" then Some (values "<slicerCacheDefinition\\s[^>]*?\\sname=\"([^\"]*)\"" (text data) |> List.head) else None) |> List.sort
let pivotCaches (parts: Map<string, byte[]>) = parts |> Map.toList |> List.map fst |> List.filter (fun path -> path.StartsWith "xl/pivotCache/" && not (path.Contains "_rels")) |> List.sort
/// The package says what it holds: what workbook.xml names by relationship is in its relationships, every
/// relationship leads to a part that is there, the content types name parts that are there, every part has a type.
let sound (label: string) (parts: Map<string, byte[]>) =
    let workbookXml = text parts.["xl/workbook.xml"]
    let rels = text parts.["xl/_rels/workbook.xml.rels"]
    let ids = values "\\sId=\"([^\"]*)\"" rels
    let named = values "\\sr:id=\"([^\"]*)\"" workbookXml
    ok (label + ": every relationship workbook.xml names is there") (named |> List.forall (fun id -> List.contains id ids)) (sprintf "%A not among %A" (named |> List.filter (fun id -> not (List.contains id ids))) ids)
    let targets = [ for found in Regex.Matches(rels, "<Relationship\\b[^>]*>") do if not (found.Value.Contains "TargetMode=\"External\"") then yield (values "\\sTarget=\"([^\"]*)\"" found.Value |> List.head) ]
    let resolved = targets |> List.map (fun target -> if target.StartsWith "/" then target.TrimStart '/' else "xl/" + target)
    ok (label + ": every relationship of the workbook leads to a part") (resolved |> List.forall parts.ContainsKey) (resolved |> List.filter (parts.ContainsKey >> not) |> String.concat " ")
    let types = text parts.["[Content_Types].xml"]
    let overrides = values "<Override\\b[^>]*PartName=\"/([^\"]*)\"" types
    ok (label + ": every part the content types name is there") (overrides |> List.forall parts.ContainsKey) (overrides |> List.filter (parts.ContainsKey >> not) |> String.concat " ")
    // Every slicer cache, pivot cache definition and pivot cache records part that is there is reached.
    let reached = Collections.Generic.HashSet<string>(resolved)
    for path in resolved do
        let relsPath = path.Substring(0, path.LastIndexOf '/') + "/_rels/" + path.Substring(path.LastIndexOf '/' + 1) + ".rels"
        match parts.TryFind relsPath with
        | Some data -> for target in values "\\sTarget=\"([^\"]*)\"" (text data) do reached.Add (if target.StartsWith "/" then target.TrimStart '/' else path.Substring(0, path.LastIndexOf '/' + 1) + target) |> ignore
        | None -> ()
    let caches = parts |> Map.toList |> List.map fst |> List.filter (fun path -> (path.StartsWith "xl/slicerCaches/" || path.StartsWith "xl/pivotCache/") && not (path.Contains "_rels"))
    ok (label + ": no cache part is left that nothing leads to") (caches |> List.forall reached.Contains) (caches |> List.filter (reached.Contains >> not) |> String.concat " ")

ok "the file: four slicer caches, three pivot caches, a name per slicer cache"
    (slicerCaches file0 = [ "Slicer_Region"; "Slicer_Segment"; "Slicer_Ship_Mode"; "Slicer_Year" ] && (pivotCaches file0 |> List.filter (fun path -> path.Contains "Definition")).Length = 3
     && values "<definedName name=\"(Slicer_[^\"]*)\">#N/A" (text file0.["xl/workbook.xml"]) |> List.sort = [ "Slicer_Region"; "Slicer_Segment"; "Slicer_Ship_Mode"; "Slicer_Year" ])
    (sprintf "%A %A" (slicerCaches file0) (pivotCaches file0))

// ---- a sheet without slicers or pivot tables deleted: every cache stays -------------------
do
    let now = partsOf (without [ "Data" ])
    ok "Data deleted: the slicer caches and the pivot caches are the file's" (slicerCaches now = slicerCaches file0 && pivotCaches now = pivotCaches file0) (sprintf "%A %A" (slicerCaches now) (pivotCaches now))
    sound "Data deleted" now

// ---- one of the two sheets with slicers deleted: the caches the other still uses stay ------
// The dashboard's slicers are Region and Ship Mode; "Waterfall Chart" has those and Segment and Year.
do
    let now = partsOf (without [ "Dashboard" ])
    ok "Dashboard deleted: every slicer cache is still used" (slicerCaches now = slicerCaches file0) (sprintf "%A" (slicerCaches now))
    sound "Dashboard deleted" now
    let now = partsOf (without [ "Waterfall Chart" ])
    let workbookXml = text now.["xl/workbook.xml"]
    ok "Waterfall Chart deleted: the slicer caches of Segment and Year go with it" (slicerCaches now = [ "Slicer_Region"; "Slicer_Ship_Mode" ]) (sprintf "%A" (slicerCaches now))
    ok "... workbook.xml lists two slicer caches" ((values "<x14:slicerCache r:id=\"([^\"]*)\"" workbookXml).Length = 2) (short (Regex.Match(workbookXml, "<x14:slicerCaches>[\\s\\S]*?</x14:slicerCaches>").Value))
    ok "... their names are gone, the others stay" (values "<definedName name=\"(Slicer_[^\"]*)\"" workbookXml |> List.sort = [ "Slicer_Region"; "Slicer_Ship_Mode" ]) (sprintf "%A" (values "<definedName name=\"(Slicer_[^\"]*)\"" workbookXml))
    ok "... the pivot caches stay (other sheets have pivot tables of them)" (pivotCaches now = pivotCaches file0) (sprintf "%A" (pivotCaches now))
    sound "Waterfall Chart deleted" now

// ---- both deleted: no slicer cache is left, and no list of them ------------------------------
do
    let saved = without [ "Dashboard"; "Waterfall Chart" ]
    let now = partsOf saved
    let workbookXml = text now.["xl/workbook.xml"]
    ok "both sheets with slicers deleted: no slicer cache part is left" (slicerCaches now = []) (sprintf "%A" (slicerCaches now))
    ok "... no list of slicer caches in workbook.xml, the other extensions stay" (not (workbookXml.Contains "slicerCache") && workbookXml.Contains "<x15:dataModel>" && workbookXml.Contains "<x14:workbookPr/>") (short (Regex.Match(workbookXml, "<extLst>[\\s\\S]*").Value))
    ok "... no relationship, no content type, no name of a slicer cache"
        (not ((text now.["xl/_rels/workbook.xml.rels"]).Contains "slicerCache") && not ((text now.["[Content_Types].xml"]).Contains "slicerCache") && not (workbookXml.Contains "name=\"Slicer_"))
        (short (text now.["xl/_rels/workbook.xml.rels"]))
    sound "both sheets with slicers deleted" now
    let again = Bridge.Import.importWorkbook (XlsxReader.loadBytesNative saved |> Option.get)
    ok "... read back: the eight other sheets" (again.Sheets |> Array.map (fun sheet -> sheet.Name) = (book.Sheets |> Array.map (fun sheet -> sheet.Name) |> Array.filter (fun name -> name <> "Dashboard" && name <> "Waterfall Chart"))) (sprintf "%A" (again.Sheets |> Array.map (fun sheet -> sheet.Name)))

// ---- the sheet with the only pivot tables of two caches deleted -----------------------------
// "KPI" has the pivot tables of caches 1 and 2 (and five of cache 0, which other sheets use too).
do
    let pivotTablesOf (parts: Map<string, byte[]>) = parts |> Map.toList |> List.choose (fun (path, data) -> if path.StartsWith "xl/pivotTables/pivotTable" then Some (values "<pivotTableDefinition\\s[^>]*?\\scacheId=\"([0-9]+)\"" (text data) |> List.head) else None) |> List.countBy id |> List.sort
    ok "the file: pivot tables of three caches" (pivotTablesOf file0 |> List.map fst = [ "0"; "1"; "2" ]) (sprintf "%A" (pivotTablesOf file0))
    let saved = without [ "KPI" ]
    let now = partsOf saved
    let workbookXml = text now.["xl/workbook.xml"]
    ok "KPI deleted: the pivot tables left are of cache 0" (pivotTablesOf now |> List.map fst = [ "0" ]) (sprintf "%A" (pivotTablesOf now))
    ok "... workbook.xml lists that cache only" (values "<pivotCache cacheId=\"([0-9]+)\"" workbookXml = [ "0" ]) (Regex.Match(workbookXml, "<pivotCaches>[\\s\\S]*?</pivotCaches>").Value)
    ok "... one cache definition and its records are left" ((pivotCaches now |> List.filter (fun path -> path.Contains "Definition")).Length = 1 && (pivotCaches now |> List.filter (fun path -> path.Contains "Records")).Length <= 1) (sprintf "%A" (pivotCaches now))
    ok "... the slicer caches stay (their slicers are on other sheets)" (slicerCaches now = slicerCaches file0) (sprintf "%A" (slicerCaches now))
    sound "KPI deleted" now
    ok "... read back" ((Bridge.Import.importWorkbook (XlsxReader.loadBytesNative saved |> Option.get)).Sheets.Length = book.Sheets.Length - 1) ""

printfn "\nXLSX orphaned-cache tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
