/// Encode a Framebuffer to PNG bytes
module Rbt.Widgets.PngEncode

open Rbt.Renderer

let encode (fb: Framebuffer) : byte[] =
    use output = new System.IO.MemoryStream()
    let write (bytes: byte[]) = output.Write(bytes, 0, bytes.Length)

    let crc32table =
        Array.init 256 (fun n ->
            let mutable c = uint32 n
            for _ = 0 to 7 do
                if c &&& 1u <> 0u then c <- 0xEDB88320u ^^^ (c >>> 1) else c <- c >>> 1
            c)
    let crc32 (data: byte[]) (offset: int) (length: int) =
        let mutable c = 0xFFFFFFFFu
        for i = offset to offset + length - 1 do
            c <- crc32table.[int ((c ^^^ uint32 data.[i]) &&& 0xFFu)] ^^^ (c >>> 8)
        c ^^^ 0xFFFFFFFFu

    let writeU32BE (v: uint32) (buf: byte[]) (off: int) =
        buf.[off] <- byte (v >>> 24); buf.[off+1] <- byte (v >>> 16)
        buf.[off+2] <- byte (v >>> 8); buf.[off+3] <- byte v

    let writeChunk (chunkType: string) (data: byte[]) =
        let lenBuf = Array.zeroCreate 4
        writeU32BE (uint32 data.Length) lenBuf 0
        write lenBuf
        let typeBytes = System.Text.Encoding.ASCII.GetBytes(chunkType)
        write typeBytes
        write data
        let crcBuf = Array.zeroCreate (4 + data.Length)
        System.Array.Copy(typeBytes, 0, crcBuf, 0, 4)
        System.Array.Copy(data, 0, crcBuf, 4, data.Length)
        let crcVal = crc32 crcBuf 0 crcBuf.Length
        let crcOut = Array.zeroCreate 4
        writeU32BE crcVal crcOut 0
        write crcOut

    // PNG signature
    write [| 137uy; 80uy; 78uy; 71uy; 13uy; 10uy; 26uy; 10uy |]

    // IHDR
    let ihdr = Array.zeroCreate 13
    writeU32BE (uint32 fb.W) ihdr 0
    writeU32BE (uint32 fb.H) ihdr 4
    ihdr.[8] <- 8uy; ihdr.[9] <- 2uy  // 8-bit RGB
    writeChunk "IHDR" ihdr

    // IDAT - raw pixel data -> zlib compressed
    let rawData = Array.zeroCreate (fb.H * (1 + fb.W * 3))
    let fpx = fb.Px.Span
    for y = 0 to fb.H - 1 do
        let rowOff = y * (1 + fb.W * 3)
        rawData.[rowOff] <- 0uy  // filter byte
        for x = 0 to fb.W - 1 do
            let px = fpx.[y * fb.W + x]
            let off = rowOff + 1 + x * 3
            rawData.[off] <- Pixel.r px
            rawData.[off + 1] <- Pixel.g px
            rawData.[off + 2] <- Pixel.b px
    use ms = new System.IO.MemoryStream()
    ms.WriteByte(0x78uy)
    ms.WriteByte(0x01uy)
    use deflate = new System.IO.Compression.DeflateStream(ms, System.IO.Compression.CompressionLevel.Fastest, true)
    deflate.Write(rawData, 0, rawData.Length)
    deflate.Close()
    let mutable a = 1u
    let mutable b = 0u
    for i = 0 to rawData.Length - 1 do
        a <- (a + uint32 rawData.[i]) % 65521u
        b <- (b + a) % 65521u
    let adler = (b <<< 16) ||| a
    let adlerBuf = Array.zeroCreate 4
    writeU32BE adler adlerBuf 0
    ms.Write(adlerBuf, 0, 4)
    writeChunk "IDAT" (ms.ToArray())

    writeChunk "IEND" [||]
    output.ToArray()
