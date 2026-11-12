// dotnet fsi scripts/stage-word-pdf-data.fsx <word-pdf-output>
// Use the resolved package, not a sibling checkout or a different Unicode version.
open System
open System.IO
open System.Reflection
open System.Text.Json

if fsi.CommandLineArgs.Length <> 2 then failwith "Usage: stage-word-pdf-data.fsx <word-pdf-output>"
let stage () =
    let repository = Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, ".."))
    use assets = JsonDocument.Parse(File.ReadAllText(Path.Combine(repository, "src/Word.Pdf/obj/project.assets.json")))
    let library =
        assets.RootElement.GetProperty("libraries").EnumerateObject()
        |> Seq.find (fun item -> item.Name.StartsWith("Rbt.Unicode.Core/", StringComparison.OrdinalIgnoreCase))
    let relativePackage = library.Value.GetProperty("path").GetString()
    let dll =
        library.Value.GetProperty("files").EnumerateArray()
        |> Seq.map (fun item -> item.GetString())
        |> Seq.find (fun name -> name.StartsWith("lib/netstandard2.0/") && name.EndsWith(".dll"))
    let assemblyPath =
        assets.RootElement.GetProperty("packageFolders").EnumerateObject()
        |> Seq.map (fun folder -> Path.Combine(folder.Name, relativePackage, dll))
        |> Seq.find File.Exists
    let assembly = Assembly.LoadFile(assemblyPath)
    use resource = assembly.GetManifestResourceStream("ucd.dat")
    if isNull resource then failwith "Resolved Unicode package has no ucd.dat resource"
    let output = Path.GetFullPath(fsi.CommandLineArgs.[1])
    Directory.CreateDirectory(output) |> ignore
    use target = File.Create(Path.Combine(output, "ucd.dat"))
    resource.CopyTo(target)
    printfn "Staged %s Unicode data (%d bytes)" library.Name target.Length

stage ()
