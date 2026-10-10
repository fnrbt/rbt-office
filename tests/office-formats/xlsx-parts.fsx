#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// Nothing is lost on a save: after an unedited export every part of the
// opened package is in the written one (under its own name), with a content
// type and reachable through relationships, for every workbook under
// tests/real-world and tests/test-documents-excel.
//
// Usage: dotnet fsi tests/office-formats/xlsx-parts.fsx [--list] [file.xlsx ...]
// (--list prints what is missing per workbook instead of the first few.)

open System
open System.IO
open System.IO.Compression
module Bridge = Rbt.Spreadsheet.XlsxBridge

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-34s  %s" label detail)

let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let args = fsi.CommandLineArgs |> Array.skip 1
let listAll = args |> Array.contains "--list"
let named = args |> Array.filter (fun a -> a <> "--list")
let files =
    if named.Length > 0 then named |> Array.map Path.GetFullPath
    else
        [| for dir in [ "real-world"; "test-documents-excel" ] do
            for file in Directory.GetFiles(Path.Combine(root, dir)) |> Array.sort do
                if file.EndsWith ".xlsx" || file.EndsWith ".xlsm" then yield file |]

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

/// The targets of a .rels part, resolved to part names (internal ones).
let relTargets (relsPath: string) (xml: string) : (string * string) list =
    let baseDir =
        let dir = Path.GetDirectoryName(relsPath).Replace('\\', '/')            // ".../_rels"
        let parent = if dir.EndsWith "_rels" then dir.Substring(0, dir.Length - 5).TrimEnd('/') else dir
        parent
    let resolve (target: string) =
        if target.StartsWith "/" then target.TrimStart('/')
        else
            let stack = Collections.Generic.List<string>(baseDir.Split([| '/' |], StringSplitOptions.RemoveEmptyEntries))
            for piece in target.Split('/') do
                if piece = ".." then (if stack.Count > 0 then stack.RemoveAt(stack.Count - 1))
                elif piece <> "." && piece <> "" then stack.Add piece
            String.Join("/", stack)
    [ for m in Text.RegularExpressions.Regex.Matches(xml, "<Relationship\\b[^>]*>") do
        let attr (name: string) =
            let found = Text.RegularExpressions.Regex.Match(m.Value, "\\b" + name + "=\"([^\"]*)\"")
            if found.Success then found.Groups.[1].Value else ""
        if attr "TargetMode" <> "External" then yield attr "Type", resolve (Uri.UnescapeDataString (attr "Target")) ]

/// Every part reachable from the package's root relationships.
let reachable (parts: Map<string, byte[]>) : Set<string> =
    let seen = Collections.Generic.HashSet<string>()
    let rec visit (relsPath: string) =
        match Map.tryFind relsPath parts with
        | Some data ->
            for relType, target in relTargets relsPath (text data) do
                ignore relType
                if seen.Add target then
                    let dir = Path.GetDirectoryName(target).Replace('\\', '/')
                    visit ((if dir = "" then "" else dir + "/") + "_rels/" + Path.GetFileName target + ".rels")
        | None -> ()
    visit "_rels/.rels"
    Set.ofSeq seen

/// Every part reachable from one part.
let reachableFrom (parts: Map<string, byte[]>) (start: string) : Set<string> =
    let seen = Collections.Generic.HashSet<string>()
    let rec visit (part: string) =
        let dir = Path.GetDirectoryName(part).Replace('\\', '/')
        let relsPath = (if dir = "" then "" else dir + "/") + "_rels/" + Path.GetFileName part + ".rels"
        match Map.tryFind relsPath parts with
        | Some data -> for _, target in relTargets relsPath (text data) do if seen.Add target then visit target
        | None -> ()
    visit start
    Set.ofSeq seen

let hasContentType (types: string) (part: string) =
    types.Contains ("PartName=\"/" + part + "\"")
    || (let ext = Path.GetExtension(part).TrimStart('.') in ext <> "" && Text.RegularExpressions.Regex.IsMatch(types, "<Default\\b[^>]*Extension=\"" + Text.RegularExpressions.Regex.Escape ext + "\"", Text.RegularExpressions.RegexOptions.IgnoreCase))

// Parts the writer makes itself (their content is the writer's; calcChain is
// derived data Excel rebuilds).
let regenerated (part: string) = part = "xl/calcChain.xml"

let mutable books = 0
let mutable skipped = 0
for file in files do
    let name = Path.GetFileName file
    let original = try Some (File.ReadAllBytes file |> partsOf) with _ -> None
    match original with
    | None -> skipped <- skipped + 1          // not a zip package (an encrypted or deliberately broken test file)
    | Some original ->
        match (try XlsxReader.loadBytesNative (File.ReadAllBytes file) with _ -> None) with
        | None -> skipped <- skipped + 1
        | Some workbook ->
            books <- books + 1
            try
                let imported = Bridge.Import.importWorkbook workbook
                // (The sheet on display stays the one the file shows.)
                let activeTab =
                    let found = Text.RegularExpressions.Regex.Match(text original.["xl/workbook.xml"], "activeTab=\"(\\d+)\"")
                    if found.Success then min (int found.Groups.[1].Value) (imported.Sheets.Length - 1) else 0
                let input : Bridge.ExportInput =
                    { Sheets = imported.Sheets |> Array.map (fun sheet -> sheet.Model)
                      Names = imported.Sheets |> Array.map (fun sheet -> sheet.Name)
                      Visible = imported.Sheets |> Array.map (fun sheet -> sheet.Visible)
                      ActiveSheetIndex = activeTab
                      NameInfo = Map.empty; DefinedNames = imported.Context.Names
                      Source = Some imported.Source
                      Macros = None }
                let written = XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut input) |> partsOf
                // Only what the original package itself reaches counts (a stray entry nothing refers to is not a part).
                let wanted = reachable original |> Set.filter (fun part -> Map.containsKey part original && not (regenerated part))
                let missing = wanted |> Set.filter (fun part -> not (Map.containsKey part written)) |> Set.toList
                let show (items: string list) = (if listAll then items else List.truncate 6 items) |> String.concat ", " |> fun s -> sprintf "%d: %s" items.Length s
                ok (name + ": parts") missing.IsEmpty (show missing)
                // What the writer does not make from the model goes back byte for byte: drawings, tables,
                // charts, media, pivot and slicer parts, the theme, document properties, embedded data ...
                let rewritten (part: string) =
                    [ "[Content_Types].xml"; "_rels/.rels"; "xl/workbook.xml"; "xl/_rels/workbook.xml.rels"; "xl/styles.xml"; "xl/sharedStrings.xml"; "xl/metadata.xml" ] |> List.contains part
                    || [ "xl/worksheets/"; "xl/comments"; "xl/threadedComments/"; "xl/persons/" ] |> List.exists (fun prefix -> part.StartsWith prefix)
                let changed =
                    wanted |> Set.toList
                    |> List.filter (fun part -> not (rewritten part) && (match Map.tryFind part written with Some data -> data <> original.[part] | None -> false))
                ok (name + ": untouched parts identical") changed.IsEmpty (show changed)
                // The sheets, chart sheets among them, are listed in the file's order, each kind of sheet as what it was.
                let sheetList (parts: Map<string, byte[]>) =
                    let rels = relTargets "xl/_rels/workbook.xml.rels" (text parts.["xl/_rels/workbook.xml.rels"])
                    let ids =
                        [ for m in Text.RegularExpressions.Regex.Matches(text parts.["xl/_rels/workbook.xml.rels"], "<Relationship\\b[^>]*>") do
                            let attr (name: string) = Text.RegularExpressions.Regex.Match(m.Value, "\\b" + name + "=\"([^\"]*)\"").Groups.[1].Value
                            yield attr "Id", attr "Type" ] |> Map.ofList
                    ignore rels
                    [ for m in Text.RegularExpressions.Regex.Matches(text parts.["xl/workbook.xml"], "<sheet\\b[^>]*>") do
                        let attr (name: string) = Text.RegularExpressions.Regex.Match(m.Value, "\\b" + name + "=\"([^\"]*)\"").Groups.[1].Value
                        let relType = Map.tryFind (attr "r:id") ids |> Option.defaultValue "?"
                        yield attr "name", relType.Substring(relType.LastIndexOf '/' + 1) ]
                if original.ContainsKey "xl/workbook.xml" && original.ContainsKey "xl/_rels/workbook.xml.rels" then
                    ok (name + ": sheets") (sheetList original = sheetList written) (sprintf "%A / %A" (sheetList original) (sheetList written))
                let reached = reachable written
                let unreached = wanted |> Set.filter (fun part -> Map.containsKey part written && not (Set.contains part reached)) |> Set.toList
                ok (name + ": related") unreached.IsEmpty (show unreached)
                let types = match Map.tryFind "[Content_Types].xml" written with Some data -> text data | None -> ""
                let untyped = written |> Map.toList |> List.map fst |> List.filter (fun part -> part <> "[Content_Types].xml" && not (hasContentType types part))
                ok (name + ": content types") untyped.IsEmpty (show untyped)
                // One cell of the first worksheet edited: every other worksheet goes back as the part it was
                // read from, byte for byte, and still means what it meant - the styles and shared strings
                // its cells index are the file's, added to. (Not for a workbook in the 1904 date system: its
                // serials are converted on import, so every sheet is written from the model.)
                let isWorksheet (sheet: XlsxReader.Worksheet) = match sheet.Extras.SourcePart with Some (_, relType) -> relType.EndsWith "/worksheet" | None -> false
                let sheetsRead = List.toArray workbook.Sheets
                match sheetsRead |> Array.tryFindIndex isWorksheet with
                | Some target when sheetsRead.Length > 1 ->
                    let models = imported.Sheets |> Array.map (fun sheet -> sheet.Model)
                    let at : Rbt.Spreadsheet.Model.CellAddress = { Row = 2; Col = 1 }
                    let cell : Rbt.Spreadsheet.Model.CellData =
                        { Value = Rbt.Spreadsheet.Model.CellValue.String "edited here"; Comment = None
                          Format = { Rbt.Spreadsheet.Model.defaultFormat with Bold = true; BgColor = Some "#ABCDEF"; NumberFormat = Some "0.0000\"u\"" } }
                    models.[target] <- { models.[target] with Cells = Map.add at cell models.[target].Cells }
                    let editedBytes = XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut { input with Sheets = models })
                    let edited = partsOf editedBytes
                    let others = sheetsRead |> Array.mapi (fun i sheet -> i, sheet) |> Array.filter (fun (i, sheet) -> i <> target && isWorksheet sheet)
                    if not workbook.Date1904 then
                        let differing =
                            others |> Array.toList
                            |> List.choose (fun (_, sheet) ->
                                let path = fst sheet.Extras.SourcePart.Value
                                match Map.tryFind path edited with
                                | Some data when data = original.[path] -> None
                                | _ -> Some path)
                        ok (name + ": other sheets identical after an edit") differing.IsEmpty (show differing)
                        // ... and everything those sheets reach is still there, as it was.
                        let lost =
                            wanted |> Set.toList
                            |> List.filter (fun part ->
                                not (part.StartsWith "xl/worksheets/sheet") && not (rewritten part && not (part.StartsWith "xl/comments" || part.StartsWith "xl/threadedComments/"))
                                && (others |> Array.exists (fun (_, sheet) -> reachableFrom original (fst sheet.Extras.SourcePart.Value) |> Set.contains part))
                                && (match Map.tryFind part edited with Some data -> data <> original.[part] | None -> true))
                        ok (name + ": what the other sheets reach is identical") lost.IsEmpty (show lost)
                    // Read again: every cell of every sheet has the value and the format it had (but the edited one).
                    match (try XlsxReader.loadBytesNative editedBytes with _ -> None) with
                    | None -> ok (name + ": edited export opens") false ""
                    | Some again ->
                        let describe (book: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) =
                            [ for row in sheet.Rows do
                                for c in row.Cells do
                                    let value =
                                        match c.Value with
                                        | XlsxReader.SharedString index when index >= 0 && index < book.SharedStrings.Length -> XlsxReader.String book.SharedStrings.[index]
                                        | other -> other
                                    let format = XlsxWriter.formatOfReaderStyle c.Style
                                    if value <> XlsxReader.Empty || c.Formula.IsSome || format <> XlsxWriter.defaultXlsxCellFormat then
                                        yield (c.Reference.Row, c.Reference.Column), (value, c.Formula, format) ]
                            |> Map.ofList
                        let wrong =
                            [ for i in 0 .. sheetsRead.Length - 1 do
                                if i < again.Sheets.Length then
                                    let was, now = describe workbook sheetsRead.[i], describe again again.Sheets.[i]
                                    let was, now = (if i = target then Map.remove (3, "B") was, Map.remove (3, "B") now else was, now)
                                    if was <> now then
                                        let first = was |> Map.toSeq |> Seq.tryFind (fun (key, cell) -> Map.tryFind key now <> Some cell) |> Option.map (fun ((row, col), _) -> sprintf "%s%d" col row) |> Option.defaultValue "+"
                                        yield sprintf "%s!%s" sheetsRead.[i].Name first ]
                        ok (name + ": cells mean what they meant after an edit") wrong.IsEmpty (show wrong)
                        let editedCell = again.Sheets.[target].Rows |> List.collect (fun row -> row.Cells) |> List.tryFind (fun c -> c.Reference.Row = 3 && c.Reference.Column = "B")
                        ok (name + ": the edited cell")
                            (match editedCell with
                             | Some c ->
                                (match c.Value with XlsxReader.SharedString index -> again.SharedStrings.[index] = "edited here" | XlsxReader.String text -> text = "edited here" | _ -> false)
                                && (match c.Style with Some style -> style.Bold && style.NumberFormat = Some "0.0000\"u\"" && style.FillColor |> Option.exists (fun color -> color.ToUpperInvariant().EndsWith "ABCDEF") | None -> false)
                             | None -> false) (sprintf "%A" (editedCell |> Option.map (fun c -> c.Value, c.Style)))
                | _ -> ()
            with error -> ok (name + ": export") false (error.Message.Split('\n').[0])

printfn "\nXLSX part carry-over tests: %d passed, %d failed (%d workbooks, %d files skipped)" pass fail books skipped
if fail > 0 then exit 1
