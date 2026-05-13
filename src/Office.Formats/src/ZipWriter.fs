module ZipWriter

open Fable.Core
open Fable.Core.JsInterop
open ZipInterop

#if !FABLE_COMPILER
open System.IO
open System.IO.Compression
#endif

/// Represents a file to be added to the ZIP archive
type PackageFile = {
    Path: string
    Content: string  // XML content as string
}

// JSZip write methods
[<Emit("$0.file($1, $2)")>]
let private addFile (zip: obj) (path: string) (content: string) : obj = jsNative

[<Emit("$0.file($1, $2)")>]
let private addBinaryFile (zip: obj) (path: string) (content: byte[]) : obj = jsNative

[<Emit("$0.generateAsync({type: 'blob', compression: 'DEFLATE'})")>]
let private generateBlob (zip: obj) : Fable.Core.JS.Promise<obj> = jsNative

/// Create a new ZIP archive
[<Emit("new ($0)()")>]
let private createZip (jszip: obj) : obj = jsNative

/// Create ZIP archive from file list and return as Blob
let createPackage (files: PackageFile list) : Fable.Core.JS.Promise<obj> =
    promise {
        let zip = createZip JSZip

        // Add all files to ZIP
        files |> List.iter (fun file ->
            addFile zip file.Path file.Content |> ignore
        )

        // Generate blob
        let! blob = generateBlob zip
        return blob
    }

/// Create ZIP archive from mixed content (text and binary files)
let createPackageWithBinary (textFiles: PackageFile list) (binaryFiles: (string * byte[]) list) : Fable.Core.JS.Promise<obj> =
    promise {
        let zip = createZip JSZip

        // Add text files
        textFiles |> List.iter (fun file ->
            addFile zip file.Path file.Content |> ignore
        )

        // Add binary files
        binaryFiles |> List.iter (fun (path, content) ->
            addBinaryFile zip path content |> ignore
        )

        // Generate blob
        let! blob = generateBlob zip
        return blob
    }

#if !FABLE_COMPILER
/// Native version of `createPackage` that returns the zipped bytes
/// directly (no JS Blob involved). Apps can then write to disk or send
/// over the wire.
let createPackageBytes (files: PackageFile list) : byte[] =
    use stream = new MemoryStream()
    do
        use archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen = true)
        for file in files do
            let entry = archive.CreateEntry(file.Path, CompressionLevel.Optimal)
            use entryStream = entry.Open()
            let bytes = System.Text.Encoding.UTF8.GetBytes(file.Content)
            entryStream.Write(bytes, 0, bytes.Length)
    stream.ToArray()

let createPackageFileFromDirectory (sourceDirectory: string) (outputPath: string) =
    let root = DirectoryInfo(sourceDirectory)
    if not root.Exists then
        invalidArg (nameof sourceDirectory) (sprintf "Source directory does not exist: %s" sourceDirectory)

    let outputFullPath = Path.GetFullPath(outputPath)
    let outputDirectory = Path.GetDirectoryName(outputFullPath)
    if not (isNull outputDirectory) && outputDirectory <> "" then
        Directory.CreateDirectory(outputDirectory) |> ignore

    if File.Exists(outputFullPath) then
        File.Delete(outputFullPath)

    use stream = new FileStream(outputFullPath, FileMode.CreateNew, FileAccess.Write, FileShare.None)
    use archive = new ZipArchive(stream, ZipArchiveMode.Create)

    let rootPath = root.FullName.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + string Path.DirectorySeparatorChar

    root.EnumerateFiles("*", SearchOption.AllDirectories)
    |> Seq.iter (fun file ->
        let relativePath = file.FullName.Substring(rootPath.Length).Replace(Path.DirectorySeparatorChar, '/').Replace(Path.AltDirectorySeparatorChar, '/')
        archive.CreateEntryFromFile(file.FullName, relativePath, CompressionLevel.Optimal) |> ignore)
#endif
