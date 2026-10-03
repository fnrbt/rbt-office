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

printfn ""
printfn "OfficeFormat ZipInterop tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
