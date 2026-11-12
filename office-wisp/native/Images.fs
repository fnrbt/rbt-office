module Rbt.OfficeNative.Images

open System
open System.Runtime.CompilerServices
open Rbt.Renderer
open StbImageNext

let private cache = ConditionalWeakTable<byte[], Framebuffer ref>()

let private decodePixels bytes =
    let image = ImageResult.FromMemory(bytes, Nullable ColorComponents.RedGreenBlueAlpha)
    let source = image.Data
    let pixels =
        Array.init (image.Width * image.Height) (fun index ->
            let offset = index * 4
            let alpha = uint32 source.[offset + 3]
            let red = (uint32 source.[offset] * alpha + 127u) / 255u
            let green = (uint32 source.[offset + 1] * alpha + 127u) / 255u
            let blue = (uint32 source.[offset + 2] * alpha + 127u) / 255u
            blue ||| (green <<< 8) ||| (red <<< 16) ||| (alpha <<< 24))
    ref (Framebuffer.ofArray pixels image.Width image.Height)

let private create = ConditionalWeakTable<byte[], Framebuffer ref>.CreateValueCallback(decodePixels)

/// Decode to premultiplied BGRA. Entries live only as long as their immutable source bytes.
/// Unsupported or malformed image data raises the decoder's exception.
let decode (data: byte[]) =
    if isNull data then nullArg "data"
    (cache.GetValue(data, create)).Value
