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
                // (Chart sheets are not sheets the editor has: they, and what only they lead to, are not carried. A known gap.)
                if not (relType.EndsWith "/chartsheet") && seen.Add target then
                    let dir = Path.GetDirectoryName(target).Replace('\\', '/')
                    visit ((if dir = "" then "" else dir + "/") + "_rels/" + Path.GetFileName target + ".rels")
        | None -> ()
    visit "_rels/.rels"
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
                let input : Bridge.ExportInput =
                    { Sheets = imported.Sheets |> Array.map (fun sheet -> sheet.Model)
                      Names = imported.Sheets |> Array.map (fun sheet -> sheet.Name)
                      Visible = imported.Sheets |> Array.map (fun sheet -> sheet.Visible)
                      ActiveSheetIndex = 0
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
                let reached = reachable written
                let unreached = wanted |> Set.filter (fun part -> Map.containsKey part written && not (Set.contains part reached)) |> Set.toList
                ok (name + ": related") unreached.IsEmpty (show unreached)
                let types = match Map.tryFind "[Content_Types].xml" written with Some data -> text data | None -> ""
                let untyped = written |> Map.toList |> List.map fst |> List.filter (fun part -> part <> "[Content_Types].xml" && not (hasContentType types part))
                ok (name + ": content types") untyped.IsEmpty (show untyped)
            with error -> ok (name + ": export") false (error.Message.Split('\n').[0])

printfn "\nXLSX part carry-over tests: %d passed, %d failed (%d workbooks, %d files skipped)" pass fail books skipped
if fail > 0 then exit 1
