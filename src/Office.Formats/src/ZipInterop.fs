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

    [<Emit("$1.then($0)")>]
    let map (f: 'a -> 'b) (p: Fable.Core.JS.Promise<'a>) : Fable.Core.JS.Promise<'b> = jsNative

    [<Emit("Promise.all($0)")>]
    let all (promises: Fable.Core.JS.Promise<'a> array) : Fable.Core.JS.Promise<'a array> = jsNative

// Computation expression for promises
type PromiseBuilder() =
    member _.Bind(p: Fable.Core.JS.Promise<'a>, f: 'a -> Fable.Core.JS.Promise<'b>) = Promise.bind f p
    member _.Return(x) = Promise.ret x
    member _.ReturnFrom(p) = p
    member _.Zero() = Promise.ret ()

let promise = PromiseBuilder()

/// Caps applied to every archive this library opens. Office packages come
/// from untrusted uploads, and a ZIP header can claim (or hide) arbitrary
/// decompressed sizes, so both the declared sizes and the bytes actually
/// produced by decompression are checked. Hosts with unusual needs can raise
/// or lower the values before loading.
type ZipLimits = {
    /// Maximum number of entries in the archive.
    mutable MaxEntries: int
    /// Maximum decompressed size of a single entry, in bytes.
    mutable MaxEntryBytes: float
    /// Maximum decompressed bytes across all entries read from one archive.
    mutable MaxTotalBytes: float
}

let limits : ZipLimits = {
    MaxEntries = 20000
    MaxEntryBytes = 512.0 * 1024.0 * 1024.0
    MaxTotalBytes = 1024.0 * 1024.0 * 1024.0
}

let private tooManyEntries (count: int) =
    failwith (sprintf "ZIP archive rejected: %d entries exceeds the limit of %d" count limits.MaxEntries)

let private entryTooLarge (name: string) =
    failwith (sprintf "ZIP archive rejected: entry '%s' decompresses to more than %.0f bytes" name limits.MaxEntryBytes)

let private archiveTooLarge () =
    failwith (sprintf "ZIP archive rejected: decompressed content exceeds %.0f bytes" limits.MaxTotalBytes)

/// Find an entry by part name. OPC part names are case-insensitive, and some
/// producers write `xl/SharedStrings.xml` while referring to it as
/// `sharedStrings.xml`, so fall back to a case-insensitive match.
let tryFindFile (archive: ZipArchive) (path: string) : ZipFile option =
    match Map.tryFind path archive.Files with
    | Some file -> Some file
    | None ->
        let wanted = path.ToLowerInvariant()
        archive.Files
        |> Map.tryPick (fun name file -> if name.ToLowerInvariant() = wanted then Some file else None)

// Load ZIP archive from blob/array buffer
[<Emit("$0.loadAsync($1)")>]
let private loadZip (jszip: obj) (data: obj) : Fable.Core.JS.Promise<obj> = jsNative

/// Declared decompressed size from the entry header, or -1 when JSZip does
/// not expose it.
[<Emit("(($0 && $0._data && typeof $0._data.uncompressedSize === 'number') ? $0._data.uncompressedSize : -1)")>]
let private declaredSize (file: obj) : float = jsNative

/// Decompress one entry with a hard byte budget. JSZip only compares the
/// produced size with the declared one after inflating everything, so a
/// lying header would otherwise be buffered in full; streaming lets us stop
/// as soon as the budget is exceeded. `$2(n)` accounts n bytes against the
/// archive total and returns false when that budget is exhausted. Resolves
/// to a Uint8Array.
[<Emit("""new Promise(function (resolve, reject) {
    var file = $0, limit = $1, account = $2, fail = $3;
    if (typeof file.internalStream !== 'function') {
        file.async('uint8array').then(function (bytes) {
            if (bytes.length > limit || !account(bytes.length)) { reject(fail(bytes.length > limit)); } else { resolve(bytes); }
        }, reject);
        return;
    }
    var chunks = [], total = 0, done = false;
    var stream = file.internalStream('uint8array');
    stream.on('data', function (chunk) {
        if (done) { return; }
        total += chunk.length;
        var overEntry = total > limit;
        if (overEntry || !account(chunk.length)) {
            done = true;
            try { stream.pause(); } catch (e) { }
            reject(fail(overEntry));
            return;
        }
        chunks.push(chunk);
    }).on('error', function (e) { if (!done) { done = true; reject(e); } })
      .on('end', function () {
        if (done) { return; }
        done = true;
        var out = new Uint8Array(total), offset = 0;
        for (var i = 0; i < chunks.length; i++) { out.set(chunks[i], offset); offset += chunks[i].length; }
        resolve(out);
    }).resume();
})""")>]
let private readEntryBounded (file: obj) (limit: float) (account: float -> bool) (fail: bool -> obj) : Fable.Core.JS.Promise<obj> = jsNative

[<Emit("new Error($0)")>]
let private jsError (message: string) : obj = jsNative

[<Emit("new TextDecoder('utf-8').decode($0)")>]
let private utf8Decode (bytes: obj) : string = jsNative

[<Emit("$0.buffer.slice($0.byteOffset, $0.byteOffset + $0.byteLength)")>]
let private toArrayBuffer (bytes: obj) : obj = jsNative

let private buildArchive (archive: obj) : ZipArchive =
    let files: obj = archive?files
    let names = JS.Constructors.Object.keys(files) |> Seq.toArray

    if names.Length > limits.MaxEntries then tooManyEntries names.Length

    // Reject on the declared sizes first: cheap, and catches honest
    // oversized archives before any decompression happens.
    let mutable declaredTotal = 0.0
    for name in names do
        let size = declaredSize (files?(name))
        if size > limits.MaxEntryBytes then entryTooLarge name
        if size > 0.0 then declaredTotal <- declaredTotal + size
    if declaredTotal > limits.MaxTotalBytes then archiveTooLarge ()

    // Bytes actually produced so far, across every entry of this archive.
    let producedTotal = ref 0.0
    let account (count: float) =
        producedTotal.Value <- producedTotal.Value + count
        producedTotal.Value <= limits.MaxTotalBytes

    let readBytes (name: string) (fileObj: obj) : Fable.Core.JS.Promise<obj> =
        let fail (overEntry: bool) =
            if overEntry then
                jsError (sprintf "ZIP archive rejected: entry '%s' decompresses to more than %.0f bytes" name limits.MaxEntryBytes)
            else
                jsError (sprintf "ZIP archive rejected: decompressed content exceeds %.0f bytes" limits.MaxTotalBytes)
        readEntryBounded fileObj limits.MaxEntryBytes account fail

    let fileMap =
        names
        |> Array.choose (fun name ->
            let fileObj = files?(name)
            let isDir: bool = fileObj?dir

            if not isDir then
                let zipFile = {
                    Name = normalizePath name
                    GetText = fun () -> readBytes name fileObj |> Promise.map utf8Decode
                    GetArrayBuffer = fun () -> readBytes name fileObj |> Promise.map toArrayBuffer
                    GetTextNative = None
                    GetBytesNative = None
                }
                Some (zipFile.Name, zipFile)
            else
                None)
        |> Map.ofArray

    { Files = fileMap }

let loadArchive (data: obj) : Fable.Core.JS.Promise<ZipArchive> =
    loadZip JSZip data |> Promise.map buildArchive

// Helper to get file content
let getFileText (archive: ZipArchive) (path: string) : Fable.Core.JS.Promise<string option> =
    promise {
        match tryFindFile archive path with
        | Some file ->
            let! text = file.GetText()
            return Some text
        | None -> return None
    }

let getFileTextNative (archive: ZipArchive) (path: string) : string option =
    match tryFindFile archive path with
    | Some file -> file.GetTextNative |> Option.map (fun read -> read())
    | None -> None

let getFileBytesNative (archive: ZipArchive) (path: string) : byte[] option =
    match tryFindFile archive path with
    | Some file -> file.GetBytesNative |> Option.map (fun read -> read())
    | None -> None

#if FABLE_COMPILER
let loadArchiveNative (_data: byte[]) : ZipArchive =
    failwith "Native ZIP loading is unavailable in Fable"
#else
/// Read one entry without trusting its header: the declared length is only
/// used to fail fast (never to size a buffer), and the copy stops as soon as
/// the entry or the archive budget is exceeded.
let private readEntryBytesNative (account: int -> bool) (entry: System.IO.Compression.ZipArchiveEntry) =
    if float entry.Length > limits.MaxEntryBytes then entryTooLarge entry.FullName
    use entryStream = entry.Open()
    use ms = new MemoryStream()
    let buffer : byte[] = Array.zeroCreate 81920
    let mutable total = 0.0
    let mutable read = entryStream.Read(buffer, 0, buffer.Length)
    while read > 0 do
        total <- total + float read
        if total > limits.MaxEntryBytes then entryTooLarge entry.FullName
        if not (account read) then archiveTooLarge ()
        ms.Write(buffer, 0, read)
        read <- entryStream.Read(buffer, 0, buffer.Length)
    ms.ToArray()

/// UTF-8 text of a part, without the byte-order mark some producers emit.
let private decodeUtf8 (bytes: byte[]) =
    if bytes.Length >= 3 && bytes.[0] = 0xEFuy && bytes.[1] = 0xBBuy && bytes.[2] = 0xBFuy then
        System.Text.Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3)
    else
        System.Text.Encoding.UTF8.GetString(bytes)

let loadArchiveNative (data: byte[]) : ZipArchive =
    let stream = new MemoryStream(data, writable = false)
    let zip = new System.IO.Compression.ZipArchive(stream, ZipArchiveMode.Read, leaveOpen = false)
    if zip.Entries.Count > limits.MaxEntries then tooManyEntries zip.Entries.Count
    let mutable declaredTotal = 0.0
    for entry in zip.Entries do
        if float entry.Length > limits.MaxEntryBytes then entryTooLarge entry.FullName
        declaredTotal <- declaredTotal + float entry.Length
    if declaredTotal > limits.MaxTotalBytes then archiveTooLarge ()
    let readLock = obj()
    // Bytes actually produced so far, across every entry of this archive.
    let producedTotal = ref 0.0
    let account (count: int) =
        producedTotal.Value <- producedTotal.Value + float count
        producedTotal.Value <= limits.MaxTotalBytes
    let entries =
        zip.Entries
        |> Seq.filter (fun entry -> not (entry.FullName.EndsWith("/")))
        |> Seq.map (fun entry ->
            let name = normalizePath entry.FullName
            let bytes = lazy (lock readLock (fun () -> readEntryBytesNative account entry))
            let file = {
                Name = name
                GetText = fun () -> failwith "JSZip text loading is unavailable on native .NET"
                GetArrayBuffer = fun () -> failwith "JSZip binary loading is unavailable on native .NET"
                GetTextNative = Some (fun () -> decodeUtf8 bytes.Value)
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
        match tryFindFile archive path with
        | Some file ->
            let! buffer = file.GetArrayBuffer()
            let blob = createBlob buffer mimeType
            let blobUrl = createObjectUrl blob
            return Some blobUrl
        | None -> return None
    }

let getFileAsDataUrl (archive: ZipArchive) (path: string) (mimeType: string) : Fable.Core.JS.Promise<string option> =
    promise {
        match tryFindFile archive path with
        | Some file ->
            let! buffer = file.GetArrayBuffer()
            return Some ("data:" + mimeType + ";base64," + arrayBufferToBase64 buffer)
        | None -> return None
    }
