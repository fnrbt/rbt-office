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

[<Emit("(function (zip, source, path) { var file = source.files[path]; if (!file || file.dir) return false; zip.files[path] = file; return true; })($0, $1, $2)")>]
let private reuseEntry (zip: obj) (source: obj) (path: string) : bool = jsNative

/// `createPackageWithBinary`, with entries of an opened package put into the
/// new one as they are - the compressed data itself, neither inflated nor
/// deflated again: every binary file that is the very array `reuse` names
/// for its path (bytes read from `source`, see ZipInterop.rawArchiveBytes).
let createPackageReusing (source: obj) (reuse: (string * byte[]) list) (textFiles: PackageFile list) (binaryFiles: (string * byte[]) list) : Fable.Core.JS.Promise<obj> =
    promise {
        let zip = createZip JSZip
        textFiles |> List.iter (fun file -> addFile zip file.Path file.Content |> ignore)
        binaryFiles |> List.iter (fun (path, content) ->
            let same = reuse |> List.exists (fun (known, data) -> known = path && obj.ReferenceEquals (data, content))
            if not (same && reuseEntry zip source path) then addBinaryFile zip path content |> ignore)
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

let createPackageBytesWithBinary (textFiles: PackageFile list) (binaryFiles: (string * byte[]) list) : byte[] =
    use stream = new MemoryStream()
    do
        use archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen = true)
        for file in textFiles do
            let entry = archive.CreateEntry(file.Path, CompressionLevel.Optimal)
            use entryStream = entry.Open()
            let bytes = System.Text.Encoding.UTF8.GetBytes(file.Content)
            entryStream.Write(bytes, 0, bytes.Length)
        for path, content in binaryFiles do
            let entry = archive.CreateEntry(path, CompressionLevel.Optimal)
            use entryStream = entry.Open()
            entryStream.Write(content, 0, content.Length)
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
