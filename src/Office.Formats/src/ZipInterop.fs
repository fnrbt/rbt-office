module ZipInterop

open Fable.Core
open Fable.Core.JsInterop
open System
#if !FABLE_COMPILER
open System.IO
open System.IO.Compression
#endif

// JSZip interop for Fable
[<Import("default", "jszip")>]
let JSZip: obj = jsNative

type ZipFile = {
    Name: string
    GetText: unit -> Fable.Core.JS.Promise<string>
    GetArrayBuffer: unit -> Fable.Core.JS.Promise<obj>
    GetTextNative: (unit -> string) option
    GetBytesNative: (unit -> byte[]) option
}

type ZipArchive = {
    Files: Map<string, ZipFile>
}

// Normalize path separators for Office documents
let normalizePath (path: string) =
    path.TrimStart('/').Replace('\\', '/')

/// Resolve a relationship target to a package path. A target starting with
/// "/" is relative to the package root; anything else is relative to
/// `baseDir`, the directory of the part that owns the relationship.
let resolveTarget (baseDir: string) (target: string) : string =
    let target = target.Replace('\\', '/')
    let full =
        if target.StartsWith "/" || baseDir = "" then target
        else baseDir + "/" + target
    let parts = ResizeArray<string>()
    for segment in full.Split('/') do
        if segment = ".." then
            if parts.Count > 0 then parts.RemoveAt(parts.Count - 1)
        elif segment <> "" && segment <> "." then
            parts.Add segment
    System.String.Join("/", parts)

// Promise helpers
module Promise =
    [<Emit("$1.then($0)")>]
    let bind (f: 'a -> Fable.Core.JS.Promise<'b>) (p: Fable.Core.JS.Promise<'a>) : Fable.Core.JS.Promise<'b> = jsNative

    [<Emit("Promise.resolve($0)")>]
    let ret (x: 'a) : Fable.Core.JS.Promise<'a> = jsNative

    [<Emit("Promise.all($0)")>]
    let all (promises: Fable.Core.JS.Promise<'a> array) : Fable.Core.JS.Promise<'a array> = jsNative

// Computation expression for promises
type PromiseBuilder() =
    member _.Bind(p: Fable.Core.JS.Promise<'a>, f: 'a -> Fable.Core.JS.Promise<'b>) = Promise.bind f p
    member _.Return(x) = Promise.ret x
    member _.ReturnFrom(p) = p
    member _.Zero() = Promise.ret ()

let promise = PromiseBuilder()

// Load ZIP archive from blob/array buffer
[<Emit("$0.loadAsync($1)")>]
let private loadZip (jszip: obj) (data: obj) : Fable.Core.JS.Promise<obj> = jsNative

[<Emit("$0.file($1)")>]
let private getZipFile (archive: obj) (path: string) : obj = jsNative

[<Emit("$0.async('text')")>]
let private fileAsText (file: obj) : Fable.Core.JS.Promise<string> = jsNative

[<Emit("$0.async('arraybuffer')")>]
let private fileAsArrayBuffer (file: obj) : Fable.Core.JS.Promise<obj> = jsNative

let loadArchive (data: obj) : Fable.Core.JS.Promise<ZipArchive> =
    promise {
        let! archive = loadZip JSZip data
        let files: obj = archive?files

        let fileMap =
            JS.Constructors.Object.keys(files)
            |> Seq.toArray
            |> Array.choose (fun name ->
                let fileObj = files?(name)
                let isDir: bool = fileObj?dir

                if not isDir then
                    let zipFile = {
                        Name = name
                        GetText = fun () -> fileAsText fileObj
                        GetArrayBuffer = fun () -> fileAsArrayBuffer fileObj
                        GetTextNative = None
                        GetBytesNative = None
                    }
                    Some (name, zipFile)
                else
                    None)
            |> Map.ofArray

        return { Files = fileMap }
    }

// Helper to get file content
let getFileText (archive: ZipArchive) (path: string) : Fable.Core.JS.Promise<string option> =
    promise {
        match Map.tryFind path archive.Files with
        | Some file ->
            let! text = file.GetText()
            return Some text
        | None -> return None
    }

let getFileTextNative (archive: ZipArchive) (path: string) : string option =
    match Map.tryFind path archive.Files with
    | Some file -> file.GetTextNative |> Option.map (fun read -> read())
    | None -> None

let getFileBytesNative (archive: ZipArchive) (path: string) : byte[] option =
    match Map.tryFind path archive.Files with
    | Some file -> file.GetBytesNative |> Option.map (fun read -> read())
    | None -> None

#if FABLE_COMPILER
let loadArchiveNative (_data: byte[]) : ZipArchive =
    failwith "Native ZIP loading is unavailable in Fable"
#else
let private readEntryBytesNative (entry: System.IO.Compression.ZipArchiveEntry) =
    use entryStream = entry.Open()
    let initialCapacity =
        if entry.Length > 0L && entry.Length <= int64 Int32.MaxValue then int entry.Length else 0
    use ms =
        if initialCapacity > 0 then new MemoryStream(initialCapacity)
        else new MemoryStream()
    entryStream.CopyTo(ms)
    ms.ToArray()

let loadArchiveNative (data: byte[]) : ZipArchive =
    let stream = new MemoryStream(data, writable = false)
    let zip = new System.IO.Compression.ZipArchive(stream, ZipArchiveMode.Read, leaveOpen = false)
    let readLock = obj()
    let entries =
        zip.Entries
        |> Seq.filter (fun entry -> not (entry.FullName.EndsWith("/")))
        |> Seq.map (fun entry ->
            let name = normalizePath entry.FullName
            let bytes = lazy (lock readLock (fun () -> readEntryBytesNative entry))
            let file = {
                Name = name
                GetText = fun () -> failwith "JSZip text loading is unavailable on native .NET"
                GetArrayBuffer = fun () -> failwith "JSZip binary loading is unavailable on native .NET"
                GetTextNative = Some (fun () -> System.Text.Encoding.UTF8.GetString(bytes.Value))
                GetBytesNative = Some (fun () -> bytes.Value)
            }
            name, file)
        |> Map.ofSeq
    { Files = entries }
#endif

// Create Blob from ArrayBuffer
[<Emit("new Blob([$0], { type: $1 })")>]
let private createBlob (buffer: obj) (mimeType: string) : obj = jsNative

// Create object URL from Blob
[<Emit("URL.createObjectURL($0)")>]
let private createObjectUrl (blob: obj) : string = jsNative

[<Emit("(() => { const bytes = new Uint8Array($0); let binary = ''; for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]); return btoa(binary); })()")>]
let private arrayBufferToBase64 (_buffer: obj) : string = jsNative

// Get file as Blob URL (more efficient than base64)
let getFileAsBlobUrl (archive: ZipArchive) (path: string) (mimeType: string) : Fable.Core.JS.Promise<string option> =
    promise {
        match Map.tryFind path archive.Files with
        | Some file ->
            let! buffer = file.GetArrayBuffer()
            let blob = createBlob buffer mimeType
            let blobUrl = createObjectUrl blob
            return Some blobUrl
        | None -> return None
    }

let getFileAsDataUrl (archive: ZipArchive) (path: string) (mimeType: string) : Fable.Core.JS.Promise<string option> =
    promise {
        match Map.tryFind path archive.Files with
        | Some file ->
            let! buffer = file.GetArrayBuffer()
            return Some ("data:" + mimeType + ";base64," + arrayBufferToBase64 buffer)
        | None -> return None
    }
