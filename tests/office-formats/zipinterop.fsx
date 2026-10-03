#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"

open System
open System.IO
open System.IO.Compression
open System.Text

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label detail)

let createArchive entryCount entryBytes =
    use stream = new MemoryStream()
    do
        use archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen = true)

        let writeTextEntry path (text: string) =
            let entry = archive.CreateEntry(path, CompressionLevel.NoCompression)
            do
                use entryStream = entry.Open()
                let bytes = Encoding.UTF8.GetBytes(text)
                entryStream.Write(bytes, 0, bytes.Length)

        let writeBytesEntry path (bytes: byte[]) =
            let entry = archive.CreateEntry(path, CompressionLevel.NoCompression)
            do
                use entryStream = entry.Open()
                entryStream.Write(bytes, 0, bytes.Length)

        writeTextEntry "/docProps/core.xml" "<core>ok</core>"
        writeBytesEntry "word/media/used.bin" [| 1uy; 2uy; 3uy; 4uy |]

        let payload = Array.zeroCreate<byte> entryBytes
        for i in 1 .. entryCount do
            writeBytesEntry (sprintf "word/media/unused-%04d.bin" i) payload

    stream.ToArray()

let data = createArchive 24 (512 * 1024)
let archive = ZipInterop.loadArchiveNative data

ok "native-entry-count"
   (archive.Files.Count = 26)
   (sprintf "%d" archive.Files.Count)
ok "native-normalized-path"
   (archive.Files.ContainsKey "docProps/core.xml")
   (archive.Files |> Seq.map (fun kvp -> kvp.Key) |> String.concat ",")
ok "native-text-read"
   (ZipInterop.getFileTextNative archive "docProps/core.xml" = Some "<core>ok</core>")
   (sprintf "%A" (ZipInterop.getFileTextNative archive "docProps/core.xml"))
ok "native-binary-read"
   (ZipInterop.getFileBytesNative archive "word/media/used.bin" |> Option.exists ((=) [| 1uy; 2uy; 3uy; 4uy |]))
   (sprintf "%A" (ZipInterop.getFileBytesNative archive "word/media/used.bin" |> Option.map Array.length))

ZipInterop.loadArchiveNative data |> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let before = GC.GetAllocatedBytesForCurrentThread()
let lazyArchive = ZipInterop.loadArchiveNative data
let coreText = ZipInterop.getFileTextNative lazyArchive "docProps/core.xml"
let allocatedMb = float (GC.GetAllocatedBytesForCurrentThread() - before) / 1024.0 / 1024.0
ok "native-load-lazy-alloc"
   (coreText = Some "<core>ok</core>" && allocatedMb < 4.0)
   (sprintf "allocated=%.1f MB" allocatedMb)

// ----- limits for untrusted archives -----------------------------------------

let rejected (action: unit -> unit) =
    try
        action ()
        None
    with ex -> Some ex.Message

let savedEntries = ZipInterop.limits.MaxEntries
let savedEntryBytes = ZipInterop.limits.MaxEntryBytes
let savedTotalBytes = ZipInterop.limits.MaxTotalBytes
let restoreLimits () =
    ZipInterop.limits.MaxEntries <- savedEntries
    ZipInterop.limits.MaxEntryBytes <- savedEntryBytes
    ZipInterop.limits.MaxTotalBytes <- savedTotalBytes

ZipInterop.limits.MaxEntries <- 10
let tooMany = rejected (fun () -> ZipInterop.loadArchiveNative data |> ignore)
restoreLimits ()
ok "limit-entry-count" (tooMany |> Option.exists (fun m -> m.Contains "entries")) (sprintf "%A" tooMany)

ZipInterop.limits.MaxEntryBytes <- 1024.0
let tooBigEntry = rejected (fun () -> ZipInterop.loadArchiveNative data |> ignore)
restoreLimits ()
ok "limit-entry-bytes" (tooBigEntry |> Option.exists (fun m -> m.Contains "unused-0001")) (sprintf "%A" tooBigEntry)

ZipInterop.limits.MaxTotalBytes <- 1024.0 * 1024.0
let tooBigTotal = rejected (fun () -> ZipInterop.loadArchiveNative data |> ignore)
restoreLimits ()
ok "limit-total-bytes" (tooBigTotal |> Option.exists (fun m -> m.Contains "decompressed content")) (sprintf "%A" tooBigTotal)

// A header that under-reports its size must not get past the limits: patch
// the declared uncompressed size of a 512 KiB stored entry down to 16 bytes
// in both the local header and the central directory.
let lyingArchive =
    let bytes = createArchive 1 (512 * 1024)
    let patch (signature: byte[]) (sizeOffset: int) =
        let mutable i = 0
        while i < bytes.Length - 4 do
            if bytes.[i] = signature.[0] && bytes.[i + 1] = signature.[1] && bytes.[i + 2] = signature.[2] && bytes.[i + 3] = signature.[3] then
                let size = BitConverter.ToInt32(bytes, i + sizeOffset)
                if size = 512 * 1024 then
                    Array.blit (BitConverter.GetBytes 16) 0 bytes (i + sizeOffset) 4
            i <- i + 1
    patch [| 0x50uy; 0x4Buy; 0x03uy; 0x04uy |] 22
    patch [| 0x50uy; 0x4Buy; 0x01uy; 0x02uy |] 24
    bytes
ZipInterop.limits.MaxEntryBytes <- 64.0 * 1024.0
let lying =
    rejected (fun () ->
        let archive = ZipInterop.loadArchiveNative lyingArchive
        // Loading passes (the header claims 16 bytes); the read must not
        // hand back more than the limit whatever the stream produces.
        match ZipInterop.getFileBytesNative archive "word/media/unused-0001.bin" with
        | Some bytes when float bytes.Length > ZipInterop.limits.MaxEntryBytes -> ()
        | _ -> failwith "bounded")
restoreLimits ()
ok "limit-ignores-declared-size" lying.IsSome (sprintf "%A" lying)

ok "limits-restored-archive-loads"
   ((ZipInterop.loadArchiveNative data).Files.Count = 26)
   "archive no longer loads with default limits"

ok "case-insensitive-part-lookup"
   (ZipInterop.getFileTextNative archive "DocProps/Core.xml" = Some "<core>ok</core>")
   (sprintf "%A" (ZipInterop.getFileTextNative archive "DocProps/Core.xml"))

// ----- XML from untrusted parts ----------------------------------------------

let dtd = rejected (fun () -> XmlDsl.parse "<!DOCTYPE r [<!ENTITY e \"boom\">]><r>&e;</r>" |> ignore)
ok "xml-dtd-prohibited" dtd.IsSome (sprintf "%A" dtd)
let malformed = rejected (fun () -> XmlDsl.parse "<r><unclosed></r>" |> ignore)
ok "xml-malformed-throws" malformed.IsSome (sprintf "%A" malformed)
ok "xml-cdata-kept"
   (XmlDsl.Xml.text (XmlDsl.parse "<r><![CDATA[a<b]]></r>") = Some "a<b")
   (sprintf "%A" (XmlDsl.parse "<r><![CDATA[a<b]]></r>"))
ok "relationships-by-local-name"
   ((OfficeDsl.parseRelationships (XmlDsl.parse "<p:Relationships xmlns:p=\"urn:x\"><p:Relationship Id=\"rId1\" Type=\"t\" Target=\"a.xml\"/></p:Relationships>") |> List.map (fun r -> r.Id)) = [ "rId1" ])
   "prefixed Relationship element not parsed"

printfn ""
printfn "OfficeFormat ZipInterop tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
