/// A PNG decoder for the native hosts, which draw pictures from pixels (a
/// browser decodes pictures itself: under Fable this module is empty).
module Rbt.Spreadsheet.Png

#if !FABLE_COMPILER
open System
open System.IO
open System.IO.Compression

/// A decoded picture: `Rgba` holds Width x Height pixels, four bytes each
/// (red, green, blue, alpha - not premultiplied), row by row from the top.
type DecodedImage = {
    Width: int
    Height: int
    Rgba: byte[]
}

let private bigEndian (data: byte[]) (at: int) =
    (int data.[at] <<< 24) ||| (int data.[at + 1] <<< 16) ||| (int data.[at + 2] <<< 8) ||| int data.[at + 3]

/// Decode a PNG: greyscale, RGB, palette, with or without alpha, 8 or 16
/// bits per channel (and palette / greyscale at 1, 2 or 4 bits), not
/// interlaced. None for anything else, and for data that is not a PNG.
let tryDecode (data: byte[]) : DecodedImage option =
    try
        if isNull data || data.Length < 33
           || data.[0] <> 0x89uy || data.[1] <> 0x50uy || data.[2] <> 0x4Euy || data.[3] <> 0x47uy then None
        else
            let mutable width = 0
            let mutable height = 0
            let mutable bitDepth = 0
            let mutable colorType = 0
            let mutable interlace = 0
            let mutable palette : byte[] = [||]
            let mutable transparency : byte[] = [||]
            use compressed = new MemoryStream()
            let mutable at = 8
            let mutable ended = false
            while not ended && at + 8 <= data.Length do
                let length = bigEndian data at
                let kind = Text.Encoding.ASCII.GetString(data, at + 4, 4)
                let body = at + 8
                if length < 0 || body + length > data.Length then ended <- true
                else
                    match kind with
                    | "IHDR" ->
                        width <- bigEndian data body
                        height <- bigEndian data (body + 4)
                        bitDepth <- int data.[body + 8]
                        colorType <- int data.[body + 9]
                        interlace <- int data.[body + 12]
                    | "PLTE" -> palette <- Array.sub data body length
                    | "tRNS" -> transparency <- Array.sub data body length
                    | "IDAT" -> compressed.Write(data, body, length)
                    | "IEND" -> ended <- true
                    | _ -> ()
                    at <- body + length + 4
            let channels =
                match colorType with
                | 0 | 3 -> 1
                | 2 -> 3
                | 4 -> 2
                | 6 -> 4
                | _ -> 0
            if width <= 0 || height <= 0 || channels = 0 || interlace <> 0 || compressed.Length < 3L
               || (int64 width) * (int64 height) > 64000000L
               || not (bitDepth = 8 || bitDepth = 16 || ((colorType = 0 || colorType = 3) && (bitDepth = 1 || bitDepth = 2 || bitDepth = 4))) then None
            else
                // The image data is a zlib stream: two header bytes, then deflate.
                let packed = compressed.ToArray()
                use deflated = new MemoryStream(packed, 2, packed.Length - 2)
                use inflater = new DeflateStream(deflated, CompressionMode.Decompress)
                let bitsPerPixel = channels * bitDepth
                let bytesPerPixel = max 1 (bitsPerPixel / 8)
                let rowBytes = (width * bitsPerPixel + 7) / 8
                let raw = Array.zeroCreate<byte> ((rowBytes + 1) * height)
                let mutable filled = 0
                let mutable more = true
                while more && filled < raw.Length do
                    let read = inflater.Read(raw, filled, raw.Length - filled)
                    if read <= 0 then more <- false else filled <- filled + read
                if filled < raw.Length then None
                else
                    // Undo the row filters in place.
                    let previous = Array.zeroCreate<byte> rowBytes
                    let current = Array.zeroCreate<byte> rowBytes
                    let rgba = Array.zeroCreate<byte> (width * height * 4)
                    for row in 0 .. height - 1 do
                        let start = row * (rowBytes + 1)
                        let filter = raw.[start]
                        for x in 0 .. rowBytes - 1 do
                            let value = raw.[start + 1 + x]
                            let left = if x >= bytesPerPixel then int current.[x - bytesPerPixel] else 0
                            let up = int previous.[x]
                            let upLeft = if x >= bytesPerPixel then int previous.[x - bytesPerPixel] else 0
                            let predicted =
                                match filter with
                                | 1uy -> left
                                | 2uy -> up
                                | 3uy -> (left + up) / 2
                                | 4uy ->
                                    let estimate = left + up - upLeft
                                    let dLeft = abs (estimate - left)
                                    let dUp = abs (estimate - up)
                                    let dUpLeft = abs (estimate - upLeft)
                                    if dLeft <= dUp && dLeft <= dUpLeft then left elif dUp <= dUpLeft then up else upLeft
                                | _ -> 0
                            current.[x] <- byte ((int value + predicted) &&& 0xFF)
                        // The row's samples as 8-bit values.
                        let sample (index: int) : int =
                            if bitDepth = 8 then int current.[index]
                            elif bitDepth = 16 then int current.[index * 2]
                            else
                                let bit = index * bitDepth
                                let shift = 8 - bitDepth - (bit % 8)
                                (int current.[bit / 8] >>> shift) &&& ((1 <<< bitDepth) - 1)
                        let scale (value: int) = if bitDepth >= 8 then value else value * 255 / ((1 <<< bitDepth) - 1)
                        for x in 0 .. width - 1 do
                            let out = (row * width + x) * 4
                            let set (r: int) (g: int) (b: int) (a: int) =
                                rgba.[out] <- byte r
                                rgba.[out + 1] <- byte g
                                rgba.[out + 2] <- byte b
                                rgba.[out + 3] <- byte a
                            match colorType with
                            | 0 ->
                                let raw0 = sample x
                                let grey = scale raw0
                                // (tRNS for greyscale: one sample value that is transparent.)
                                let transparent = transparency.Length >= 2 && bitDepth <= 8 && int transparency.[1] = raw0
                                set grey grey grey (if transparent then 0 else 255)
                            | 2 ->
                                let r, g, b = sample (x * 3), sample (x * 3 + 1), sample (x * 3 + 2)
                                let transparent =
                                    transparency.Length >= 6 && bitDepth = 8
                                    && int transparency.[1] = r && int transparency.[3] = g && int transparency.[5] = b
                                set r g b (if transparent then 0 else 255)
                            | 3 ->
                                let index = sample x
                                if index * 3 + 2 < palette.Length then
                                    set (int palette.[index * 3]) (int palette.[index * 3 + 1]) (int palette.[index * 3 + 2])
                                        (if index < transparency.Length then int transparency.[index] else 255)
                                else set 0 0 0 255
                            | 4 ->
                                let grey = sample (x * 2)
                                set grey grey grey (sample (x * 2 + 1))
                            | _ -> set (sample (x * 4)) (sample (x * 4 + 1)) (sample (x * 4 + 2)) (sample (x * 4 + 3))
                        Array.blit current 0 previous 0 rowBytes
                    Some { Width = width; Height = height; Rgba = rgba }
    with _ -> None

/// The pixels of a decoded picture as 32-bit values, blue in the low byte,
/// then green, red and alpha, with colour premultiplied by alpha: what the
/// renderers' framebuffers hold.
let toPremultipliedBgra (image: DecodedImage) : uint32[] =
    Array.init (image.Width * image.Height) (fun index ->
        let at = index * 4
        let alpha = uint32 image.Rgba.[at + 3]
        let premultiply (value: byte) = (uint32 value * alpha + 127u) / 255u
        premultiply image.Rgba.[at + 2] ||| (premultiply image.Rgba.[at + 1] <<< 8) ||| (premultiply image.Rgba.[at] <<< 16) ||| (alpha <<< 24))
#endif
