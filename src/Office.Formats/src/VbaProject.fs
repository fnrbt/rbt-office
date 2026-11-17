module VbaProject

// Read and write `xl/vbaProject.bin` — the OLE2 Compound File (CFB) that holds
// a workbook's VBA. We extract / emit the *source text* only (no p-code): on a
// version mismatch Office recompiles from source, and a file we author never
// matches, so source-only is the portable form. Two pieces: an [MS-CFB]
// compound-file reader/writer and the [MS-OVBA] run-length codec for the
// module source streams. Native (System.IO via byte arrays only; pure).

open System

// VBA module source is MBCS (Western = Windows-1252), but the codepage
// provider isn't on .NET Core by default and source is effectively ASCII, so
// we map bytes 0–255 ↔ chars directly (Latin-1). Dependency-free, exact for
// ASCII, and round-trips on netstandard2.0.
let private latin1Bytes (s: string) : byte[] = s.ToCharArray() |> Array.map (fun c -> byte (int c &&& 0xFF))
let private latin1String (b: byte[]) : string = String(b |> Array.map (fun x -> char (int x)))

// ===========================================================================
// MS-OVBA 2.4.1 compression
// ===========================================================================

let private numBits (n: int) =
    let mutable v = n
    let mutable b = 0
    while v > 0 do b <- b + 1; v <- v >>> 1
    b

/// Decompress an OVBA CompressedContainer starting at `start` (handles real
/// Office output: literal tokens + LZ CopyTokens, plus raw chunks).
let ovbaDecompress (input: byte[]) (start: int) : byte[] =
    let res = System.Collections.Generic.List<byte>()
    let mutable pos = start
    if pos >= input.Length || input.[pos] <> 0x01uy then failwith "bad OVBA signature"
    pos <- pos + 1
    while pos + 1 < input.Length do
        let header = int input.[pos] ||| (int input.[pos + 1] <<< 8)
        pos <- pos + 2
        let chunkSize = (header &&& 0x0FFF) + 3
        let compressed = (header &&& 0x8000) <> 0
        let chunkEnd = min input.Length (pos + chunkSize - 2)
        if not compressed then
            for k in 0 .. 4095 do (if pos + k < input.Length then res.Add input.[pos + k])
            pos <- pos + 4096
        else
            let chunkStart = res.Count
            while pos < chunkEnd do
                let flag = int input.[pos]
                pos <- pos + 1
                let mutable bit = 0
                while bit < 8 && pos < chunkEnd do
                    if (flag >>> bit) &&& 1 = 0 then
                        res.Add input.[pos]; pos <- pos + 1
                    else
                        let token = int input.[pos] ||| (int input.[pos + 1] <<< 8)
                        pos <- pos + 2
                        let diff = res.Count - chunkStart
                        let bitCount = max (numBits (diff - 1)) 4
                        let lenMask = 0xFFFF >>> bitCount
                        let length = (token &&& lenMask) + 3
                        let offset = (token >>> (16 - bitCount)) + 1
                        let src = res.Count - offset
                        for c in 0 .. length - 1 do res.Add res.[src + c]
                    bit <- bit + 1
    res.ToArray()

/// Compress to an OVBA CompressedContainer using all-literal chunks (no LZ
/// matching). Chunks cap decompressed size so compressed data fits the 12-bit
/// size field; Office decompresses these the same as its own.
let ovbaCompress (data: byte[]) : byte[] =
    let out = System.Collections.Generic.List<byte>()
    out.Add 0x01uy
    let chunkMax = 3072
    let mutable i = 0
    if data.Length = 0 then
        // One empty compressed chunk so the container is well-formed.
        out.Add 0x00uy; out.Add 0xB0uy
    while i < data.Length do
        let n = min chunkMax (data.Length - i)
        let cd = System.Collections.Generic.List<byte>()
        let mutable j = 0
        while j < n do
            cd.Add 0x00uy
            for k in 0 .. min 7 (n - j - 1) do cd.Add data.[i + j + k]
            j <- j + 8
        let size = cd.Count - 1
        let header = 0x8000 ||| 0x3000 ||| (size &&& 0x0FFF)
        out.Add (byte (header &&& 0xFF))
        out.Add (byte ((header >>> 8) &&& 0xFF))
        out.AddRange cd
        i <- i + n
    out.ToArray()

// ===========================================================================
// MS-CFB compound file
// ===========================================================================

let private cfbSig = [| 0xD0uy; 0xCFuy; 0x11uy; 0xE0uy; 0xA1uy; 0xB1uy; 0x1Auy; 0xE1uy |]
let private ENDOFCHAIN = 0xFFFFFFFE
let private FREESECT = 0xFFFFFFFF
let private FATSECT = 0xFFFFFFFD
let private NOSTREAM = 0xFFFFFFFF

let private u16 (b: byte[]) o = int b.[o] ||| (int b.[o + 1] <<< 8)
let private u32 (b: byte[]) o = int (uint32 b.[o] ||| (uint32 b.[o + 1] <<< 8) ||| (uint32 b.[o + 2] <<< 16) ||| (uint32 b.[o + 3] <<< 24))

/// A flattened CFB stream: leaf entry name + its bytes.
type CfbStream = { Name: string; Data: byte[] }

/// Parse a compound file into its streams (by directory leaf name). Handles
/// the common case: FAT indexed by the header DIFAT (≤109 FAT sectors), plus
/// the mini stream for entries below the 4096-byte cutoff.
let cfbRead (bytes: byte[]) : CfbStream list =
    if bytes.Length < 512 || bytes.[0..7] <> cfbSig then failwith "not a compound file"
    let sectorShift = u16 bytes 30
    let sectorSize = 1 <<< sectorShift
    let miniShift = u16 bytes 32
    let miniSize = 1 <<< miniShift
    let miniCutoff = u32 bytes 56
    let numFatSectors = u32 bytes 44
    let firstDirSector = u32 bytes 48
    let firstMiniFat = u32 bytes 60
    let sectorOffset s = 512 + s * sectorSize
    // FAT: concatenate the FAT sectors named in the header DIFAT (first 109).
    let fat =
        let entries = System.Collections.Generic.List<int>()
        let mutable di = 0
        while di < 109 && entries.Count / (sectorSize / 4) < numFatSectors do
            let fatSec = u32 bytes (76 + di * 4)
            if fatSec <> FREESECT && fatSec <> ENDOFCHAIN then
                let off = sectorOffset fatSec
                for k in 0 .. sectorSize / 4 - 1 do entries.Add (u32 bytes (off + k * 4))
            di <- di + 1
        entries.ToArray()
    let readChain (start: int) (size: int) : byte[] =
        let out = System.Collections.Generic.List<byte>()
        let mutable s = start
        let mutable guard = 0
        while s <> ENDOFCHAIN && s >= 0 && s < fat.Length && guard < fat.Length + 1 do
            let off = sectorOffset s
            for k in 0 .. sectorSize - 1 do (if off + k < bytes.Length then out.Add bytes.[off + k])
            s <- fat.[s]
            guard <- guard + 1
        let arr = out.ToArray()
        if size > 0 && size <= arr.Length then arr.[0 .. size - 1] else arr
    // Directory entries (each 128 bytes), chained from firstDirSector.
    let dirBytes = readChain firstDirSector 0
    let entryCount = dirBytes.Length / 128
    let entries =
        [ for i in 0 .. entryCount - 1 ->
            let o = i * 128
            let nameLen = u16 dirBytes (o + 64)
            let name =
                if nameLen <= 2 then ""
                else Text.Encoding.Unicode.GetString(dirBytes, o, nameLen - 2)
            {| Name = name
               Type = int dirBytes.[o + 66]
               Start = u32 dirBytes (o + 116)
               Size = u32 dirBytes (o + 120) |} ]
    // Mini stream: the root entry (type 5) start/size locate it in the FAT.
    let root = entries |> List.tryFind (fun e -> e.Type = 5)
    let miniStream = match root with Some r -> readChain r.Start r.Size | None -> [||]
    let miniFat =
        let raw = readChain firstMiniFat 0
        [| for k in 0 .. raw.Length / 4 - 1 -> u32 raw (k * 4) |]
    let readMini (start: int) (size: int) : byte[] =
        let out = System.Collections.Generic.List<byte>()
        let mutable s = start
        let mutable guard = 0
        while s <> ENDOFCHAIN && s >= 0 && s < miniFat.Length && guard < miniFat.Length + 1 do
            let off = s * miniSize
            for k in 0 .. miniSize - 1 do (if off + k < miniStream.Length then out.Add miniStream.[off + k])
            s <- miniFat.[s]
            guard <- guard + 1
        let arr = out.ToArray()
        if size > 0 && size <= arr.Length then arr.[0 .. size - 1] else arr
    [ for e in entries do
        if e.Type = 2 && e.Name <> "" then
            let data = if e.Size < miniCutoff then readMini e.Start e.Size else readChain e.Start e.Size
            yield { Name = e.Name; Data = data } ]

// ---- CFB writer ----

type CfbEntry =
    | CStorage of name: string * children: CfbEntry list
    | CStream of name: string * data: byte[]

type private DirRec =
    { Name: string; Type: int; Data: byte[]
      mutable Left: int; mutable Right: int; mutable Child: int
      mutable Start: int; mutable Size: int }

let private putU16 (buf: System.Collections.Generic.List<byte>) (n: int) =
    buf.Add (byte (n &&& 0xFF)); buf.Add (byte ((n >>> 8) &&& 0xFF))
let private putU32 (buf: System.Collections.Generic.List<byte>) (n: int) =
    buf.Add (byte (n &&& 0xFF)); buf.Add (byte ((n >>> 8) &&& 0xFF))
    buf.Add (byte ((n >>> 16) &&& 0xFF)); buf.Add (byte ((n >>> 24) &&& 0xFF))

/// CFB sibling order: shorter names first, then case-insensitive ordinal.
let private cfbNameCompare (a: string) (b: string) =
    if a.Length <> b.Length then compare a.Length b.Length
    else String.Compare(a, b, StringComparison.OrdinalIgnoreCase)

/// Serialize a compound file from a directory tree (root's children). 512-byte
/// sectors; streams below 4096 bytes go in the mini stream. Produces a
/// structurally-valid CFB (round-trips through cfbRead).
let cfbWrite (rootChildren: CfbEntry list) : byte[] =
    let sectorSize = 512
    let miniSize = 64
    let miniCutoff = 4096
    // 1. Flatten to directory records with stable IDs (DFS).
    let recs = System.Collections.Generic.List<DirRec>()
    let newRec name typ data =
        recs.Add { Name = name; Type = typ; Data = data; Left = NOSTREAM; Right = NOSTREAM; Child = NOSTREAM; Start = ENDOFCHAIN; Size = 0 }
        recs.Count - 1
    // Balanced BST over siblings → returns subtree root id, sets Left/Right.
    let rec buildSiblings (ids: int list) : int =
        match ids with
        | [] -> NOSTREAM
        | _ ->
            let arr = ids |> List.sortWith (fun x y -> cfbNameCompare recs.[x].Name recs.[y].Name) |> List.toArray
            let rec build lo hi =
                if lo > hi then NOSTREAM
                else
                    let mid = (lo + hi) / 2
                    recs.[arr.[mid]].Left <- build lo (mid - 1)
                    recs.[arr.[mid]].Right <- build (mid + 1) hi
                    arr.[mid]
            build 0 (arr.Length - 1)
    let rec emit (e: CfbEntry) : int =
        match e with
        | CStream (name, data) -> newRec name 2 data
        | CStorage (name, children) ->
            let id = newRec name 1 [||]
            let childIds = children |> List.map emit
            recs.[id].Child <- buildSiblings childIds
            id
    let rootId = newRec "Root Entry" 5 [||]
    let topIds = rootChildren |> List.map emit
    recs.[rootId].Child <- buildSiblings topIds
    // 2. Partition stream data into mini (small) vs regular (large).
    let isMini i = recs.[i].Type = 2 && recs.[i].Data.Length < miniCutoff && recs.[i].Data.Length > 0
    // Mini stream container: concat mini-stream data, padded to mini-sector.
    let miniStream = System.Collections.Generic.List<byte>()
    let miniFat = System.Collections.Generic.List<int>()
    for i in 0 .. recs.Count - 1 do
        if isMini i then
            let d = recs.[i].Data
            let startMini = miniStream.Count / miniSize
            recs.[i].Start <- startMini
            recs.[i].Size <- d.Length
            let sectors = (d.Length + miniSize - 1) / miniSize
            for s in 0 .. sectors - 1 do
                for k in 0 .. miniSize - 1 do
                    let idx = s * miniSize + k
                    miniStream.Add (if idx < d.Length then d.[idx] else 0uy)
                miniFat.Add (if s = sectors - 1 then ENDOFCHAIN else startMini + s + 1)
        elif recs.[i].Type = 2 then
            recs.[i].Size <- recs.[i].Data.Length   // big stream; Start set below
    // 3. Lay out regular streams as FAT chains: [big streams…, miniStream, miniFAT, directory].
    //    We assemble the list of "blobs" that occupy regular sectors, in order.
    let pad (data: byte[]) =
        let n = ((data.Length + sectorSize - 1) / sectorSize) * sectorSize
        Array.append data (Array.zeroCreate (n - data.Length))
    let bigStreamIds = [ for i in 0 .. recs.Count - 1 do if recs.[i].Type = 2 && not (isMini i) && recs.[i].Data.Length > 0 then yield i ]
    // Directory bytes (4 entries per sector). Filled after Start sectors known,
    // so we reserve its sector count now using a placeholder length.
    let dirEntryCount = ((recs.Count + 3) / 4) * 4
    let dirByteLen = dirEntryCount * 128
    let miniStreamArr = miniStream.ToArray()
    let miniFatArr =
        let b = System.Collections.Generic.List<byte>()
        for v in miniFat do putU32 b v
        // pad miniFAT sector with FREESECT
        while b.Count % sectorSize <> 0 do putU32 b FREESECT
        b.ToArray()
    // Assign regular sector ranges. Order: big streams, miniStream, miniFAT, directory.
    let mutable sectorCursor = 0
    let allocSectors (len: int) =
        let count = (len + sectorSize - 1) / sectorSize
        let start = sectorCursor
        sectorCursor <- sectorCursor + count
        start, count
    let bigPlacements =
        [ for i in bigStreamIds ->
            let start, count = allocSectors recs.[i].Data.Length
            recs.[i].Start <- start
            i, start, count, pad recs.[i].Data ]
    let miniStreamStart, miniStreamCount = allocSectors miniStreamArr.Length
    let miniFatStart, miniFatCount = if miniFatArr.Length = 0 then ENDOFCHAIN, 0 else allocSectors miniFatArr.Length
    let dirStart, dirCount = allocSectors dirByteLen
    // Root entry points at the mini stream container.
    recs.[rootId].Start <- (if miniStreamArr.Length > 0 then miniStreamStart else ENDOFCHAIN)
    recs.[rootId].Size <- miniStreamArr.Length
    // 4. Build the FAT covering all data sectors + the FAT sectors themselves.
    let dataSectorCount = sectorCursor
    let entriesPerFatSector = sectorSize / 4
    // iterate to a fixed point on FAT sector count
    let mutable numFat = 1
    let mutable changed = true
    while changed do
        let total = dataSectorCount + numFat
        let need = (total + entriesPerFatSector - 1) / entriesPerFatSector
        if need <> numFat then numFat <- need else changed <- false
    let totalSectors = dataSectorCount + numFat
    let fat = Array.create (numFat * entriesPerFatSector) FREESECT
    let chain (start: int) (count: int) =
        for s in 0 .. count - 1 do
            fat.[start + s] <- if s = count - 1 then ENDOFCHAIN else start + s + 1
    for (_, start, count, _) in bigPlacements do chain start count
    if miniStreamCount > 0 then chain miniStreamStart miniStreamCount
    if miniFatCount > 0 then chain miniFatStart miniFatCount
    chain dirStart dirCount
    for s in dataSectorCount .. totalSectors - 1 do fat.[s] <- FATSECT
    // 5. Directory bytes (now that Start/Size are known).
    let dirBuf = System.Collections.Generic.List<byte>()
    for i in 0 .. dirEntryCount - 1 do
        if i < recs.Count then
            let r = recs.[i]
            let nameBytes = Text.Encoding.Unicode.GetBytes r.Name
            for k in 0 .. 63 do dirBuf.Add (if k < nameBytes.Length then nameBytes.[k] else 0uy)
            putU16 dirBuf (nameBytes.Length + 2)        // name length incl null
            dirBuf.Add (byte r.Type)
            dirBuf.Add 1uy                               // color = black
            putU32 dirBuf r.Left
            putU32 dirBuf r.Right
            putU32 dirBuf r.Child
            for _ in 0 .. 15 do dirBuf.Add 0uy           // CLSID
            putU32 dirBuf 0                              // state bits
            for _ in 0 .. 7 do dirBuf.Add 0uy            // creation time
            for _ in 0 .. 7 do dirBuf.Add 0uy            // modified time
            putU32 dirBuf r.Start
            putU32 dirBuf r.Size
            putU32 dirBuf 0                              // size high
        else
            // unused entry
            for _ in 0 .. 63 do dirBuf.Add 0uy
            putU16 dirBuf 0
            dirBuf.Add 0uy; dirBuf.Add 0uy
            putU32 dirBuf NOSTREAM; putU32 dirBuf NOSTREAM; putU32 dirBuf NOSTREAM
            for _ in 0 .. 15 do dirBuf.Add 0uy
            putU32 dirBuf 0
            for _ in 0 .. 7 do dirBuf.Add 0uy
            for _ in 0 .. 7 do dirBuf.Add 0uy
            putU32 dirBuf ENDOFCHAIN
            putU32 dirBuf 0
            putU32 dirBuf 0
    let dirArr = pad (dirBuf.ToArray())
    // 6. Assemble the file: header + data sectors + FAT sectors.
    let header = System.Collections.Generic.List<byte>()
    for b in cfbSig do header.Add b
    for _ in 0 .. 15 do header.Add 0uy                  // CLSID
    putU16 header 0x003E                                // minor version
    putU16 header 0x0003                                // major version (512 sectors)
    putU16 header 0xFFFE                                // byte order
    putU16 header 9                                     // sector shift
    putU16 header 6                                     // mini sector shift
    for _ in 0 .. 5 do header.Add 0uy                   // reserved
    putU32 header 0                                     // num dir sectors (v3)
    putU32 header numFat
    putU32 header dirStart
    putU32 header 0                                     // transaction sig
    putU32 header miniCutoff
    putU32 header (if miniFatCount > 0 then miniFatStart else ENDOFCHAIN)
    putU32 header miniFatCount
    putU32 header ENDOFCHAIN                            // first DIFAT sector
    putU32 header 0                                     // num DIFAT sectors
    // DIFAT: FAT sector numbers (they sit right after the data sectors).
    for k in 0 .. 108 do
        if k < numFat then putU32 header (dataSectorCount + k) else putU32 header FREESECT
    // Body: data sectors in layout order, then FAT sectors.
    let body = System.Collections.Generic.List<byte>()
    for (_, _, _, padded) in bigPlacements do body.AddRange padded
    if miniStreamCount > 0 then body.AddRange (pad miniStreamArr)
    if miniFatCount > 0 then body.AddRange miniFatArr
    body.AddRange dirArr
    let fatBytes = System.Collections.Generic.List<byte>()
    for v in fat do putU32 fatBytes v
    body.AddRange (fatBytes.ToArray())
    let out = System.Collections.Generic.List<byte>()
    out.AddRange (header.ToArray())
    out.AddRange (body.ToArray())
    out.ToArray()

// ===========================================================================
// High-level: VBA module source ↔ vbaProject.bin
// ===========================================================================

let private specialStreams = set [ "dir"; "_VBA_PROJECT"; "PROJECT"; "PROJECTwm" ]
let private isModuleStream (name: string) =
    not (specialStreams.Contains name) && not (name.StartsWith "__SRP_") && not (name.StartsWith "_VBA_")

/// Recover the decompressed source of a module stream: VBA source always
/// begins with `Attribute VB_Name`, so scan for the OVBA signature byte whose
/// decompression yields that. Robust to a leading p-code region.
let private extractModuleSource (data: byte[]) : string option =
    let tryAt off =
        try
            let txt = latin1String (ovbaDecompress data off)
            if txt.StartsWith "Attribute " then Some txt else None
        with _ -> None
    let rec scan off =
        if off >= data.Length then None
        elif data.[off] = 0x01uy then
            match tryAt off with Some s -> Some s | None -> scan (off + 1)
        else scan (off + 1)
    scan 0

/// Extract each VBA module's (name, source) from a vbaProject.bin.
let readModules (vbaProjectBin: byte[]) : (string * string) list =
    cfbRead vbaProjectBin
    |> List.filter (fun s -> isModuleStream s.Name)
    |> List.choose (fun s -> extractModuleSource s.Data |> Option.map (fun src -> s.Name, src))

/// Concatenated module source for `model.Macros` (modules separated by blank
/// lines), stripping the leading `Attribute` lines Excel adds per module.
let readMacroSource (vbaProjectBin: byte[]) : string =
    readModules vbaProjectBin
    |> List.map (fun (_, src) ->
        src.Replace("\r\n", "\n").Split('\n')
        |> Array.filter (fun l -> not (l.TrimStart().StartsWith "Attribute "))
        |> String.concat "\n"
        |> fun s -> s.Trim())
    |> List.filter (fun s -> s <> "")
    |> String.concat "\n\n"

/// The kind of a module of a written project: a standard module, a document module (ThisWorkbook, a
/// sheet's) or a class module.
type ProjectModuleKind =
    | StandardModule
    | DocumentModule
    | ClassModule

/// Build a source-only vbaProject.bin of several modules: each module's stream (OVBA-compressed source
/// at offset 0, no p-code), a dir stream naming them with their types, a stub _VBA_PROJECT (its version
/// matches no host, so Office compiles from source) and PROJECT / PROJECTwm. A module's source without
/// its own `Attribute VB_Name` line gets one. Round-trips through readModules.
let writeProjectModules (modules: (string * ProjectModuleKind * string) list) : byte[] =
    let withName (name: string) (kind: ProjectModuleKind) (source: string) =
        let text = source.Replace("\r\n", "\n")
        if text.TrimStart().StartsWith("Attribute VB_Name", StringComparison.OrdinalIgnoreCase) then text.Replace("\n", "\r\n")
        else
            let header =
                match kind with
                | StandardModule -> sprintf "Attribute VB_Name = \"%s\"\n" name
                | DocumentModule ->
                    sprintf "Attribute VB_Name = \"%s\"\nAttribute VB_Base = \"0{00020820-0000-0000-C000-000000000046}\"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = True\nAttribute VB_TemplateDerived = False\nAttribute VB_Customizable = True\n" name
                | ClassModule ->
                    sprintf "Attribute VB_Name = \"%s\"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = False\nAttribute VB_Exposed = False\n" name
            (header + text).Replace("\n", "\r\n")
    let dirRaw =
        let b = System.Collections.Generic.List<byte>()
        let rec16 id (data: byte[]) =
            putU16 b id; putU32 b data.Length; b.AddRange data
        let mbcs (s: string) = latin1Bytes s
        rec16 0x0004 (mbcs "VBAProject")                                  // PROJECTNAME
        rec16 0x000F [| byte (modules.Length &&& 0xFF); byte (modules.Length >>> 8) |]   // PROJECTMODULES
        putU16 b 0x0013; putU32 b 2; putU16 b 0xFFFF                      // PROJECTCOOKIE
        for name, kind, _ in modules do
            rec16 0x0019 (mbcs name)                                      // MODULENAME
            let streamName = mbcs name                                    // MODULESTREAMNAME (MBCS + reserved + unicode)
            putU16 b 0x001A; putU32 b streamName.Length; b.AddRange streamName
            putU16 b 0x0032; putU32 b (name.Length * 2); b.AddRange (Text.Encoding.Unicode.GetBytes name)
            putU16 b 0x0031; putU32 b 4; putU32 b 0                       // MODULEOFFSET = 0
            putU16 b (if kind = StandardModule then 0x0021 else 0x0022); putU32 b 0   // MODULETYPE procedural / document-class
            putU16 b 0x002B; putU32 b 0                                   // MODULETERMINATOR
        putU16 b 0x0010; putU32 b 0                                       // PROJECT TERMINATOR
        b.ToArray()
    let vbaProjStream = [| 0xCCuy; 0x61uy; 0xFFuy; 0xFFuy; 0x00uy; 0x00uy; 0x00uy |]
    let projectText =
        String.concat "\r\n"
            [ yield "ID=\"{00000000-0000-0000-0000-000000000000}\""
              for name, kind, _ in modules do
                  match kind with
                  | StandardModule -> yield sprintf "Module=%s" name
                  | DocumentModule -> yield sprintf "Document=%s/&H00000000" name
                  | ClassModule -> yield sprintf "Class=%s" name
              yield "Name=\"VBAProject\""
              yield "" ]
    let projectwm =
        Array.concat
            [ for name, _, _ in modules do
                yield latin1Bytes (name + "\x00")
                yield Text.Encoding.Unicode.GetBytes name
                yield [| 0uy; 0uy |]
              yield [| 0uy; 0uy |] ]
    cfbWrite
        [ CStorage("VBA",
            [ yield CStream("_VBA_PROJECT", vbaProjStream)
              yield CStream("dir", ovbaCompress dirRaw)
              for name, kind, source in modules do
                  yield CStream(name, ovbaCompress (latin1Bytes (withName name kind source))) ])
          CStream("PROJECT", latin1Bytes projectText)
          CStream("PROJECTwm", projectwm) ]

/// Build a minimal source-only vbaProject.bin for one standard module (see writeProjectModules).
let writeProject (moduleName: string) (source: string) : byte[] =
    writeProjectModules [ moduleName, StandardModule, source ]

// ===========================================================================
// xlsx/xlsm zip integration (native)
// ===========================================================================

#if !FABLE_COMPILER
open System.IO
open System.IO.Compression

let private readZipEntries (bytes: byte[]) : (string * byte[]) list =
    use ms = new MemoryStream(bytes)
    use za = new ZipArchive(ms, ZipArchiveMode.Read)
    [ for e in za.Entries do
        use s = e.Open()
        use o = new MemoryStream()
        s.CopyTo o
        yield e.FullName.Replace("\\", "/"), o.ToArray() ]

let private writeZip (entries: (string * byte[]) list) : byte[] =
    use ms = new MemoryStream()
    (
        use za = new ZipArchive(ms, ZipArchiveMode.Create, true)
        for (name, data) in entries do
            let e = za.CreateEntry(name, CompressionLevel.Optimal)
            use s = e.Open()
            s.Write(data, 0, data.Length)
    )
    ms.ToArray()

/// Read `xl/vbaProject.bin` from an xlsx/xlsm and return its macro source
/// (None when the workbook has no VBA project).
let extractFromXlsx (xlsxBytes: byte[]) : string option =
    readZipEntries xlsxBytes
    |> List.tryFind (fun (n, _) -> n.ToLowerInvariant() = "xl/vbaproject.bin")
    |> Option.map (fun (_, data) -> readMacroSource data)

/// Add a source-only vbaProject.bin (one module) to an xlsx, returning .xlsm
/// bytes — injects the part, the `bin` content-type, the macro-enabled
/// workbook content-type, and the workbook→vbaProject relationship.
let injectIntoXlsx (xlsxBytes: byte[]) (moduleName: string) (source: string) : byte[] =
    let bin = writeProject moduleName source
    let utf8 = Text.Encoding.UTF8
    let patch name (f: string -> string) (lst: (string * byte[]) list) =
        lst |> List.map (fun (n, data) -> if n = name then n, utf8.GetBytes(f (utf8.GetString data)) else n, data)
    let ct (s: string) =
        // (Its own Override: ".bin" may already be printer settings by default.)
        let s =
            if s.Contains "/xl/vbaProject.bin" then s
            else s.Replace("</Types>", "<Override PartName=\"/xl/vbaProject.bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/></Types>")
        s.Replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                  "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
    let rels (s: string) =
        if s.Contains "vbaProject" then s
        else s.Replace("</Relationships>",
                       "<Relationship Id=\"rIdVBA\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" Target=\"vbaProject.bin\"/></Relationships>")
    let entries =
        readZipEntries xlsxBytes
        |> List.filter (fun (n, _) -> n <> "xl/vbaProject.bin")
        |> patch "[Content_Types].xml" ct
        |> patch "xl/_rels/workbook.xml.rels" rels
    writeZip (("xl/vbaProject.bin", bin) :: entries)
#endif
