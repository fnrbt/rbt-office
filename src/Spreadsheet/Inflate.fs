/// Raw DEFLATE (RFC 1951) decoded in plain F#, the same in .NET and in the browser: what a map chart's geography
/// cache (`cx:geoCache`, base64 of deflated XML) holds.
module Rbt.Spreadsheet.Inflate

exception private Truncated

let private lengthBase = [| 3; 4; 5; 6; 7; 8; 9; 10; 11; 13; 15; 17; 19; 23; 27; 31; 35; 43; 51; 59; 67; 83; 99; 115; 131; 163; 195; 227; 258 |]
let private lengthExtra = [| 0; 0; 0; 0; 0; 0; 0; 0; 1; 1; 1; 1; 2; 2; 2; 2; 3; 3; 3; 3; 4; 4; 4; 4; 5; 5; 5; 5; 0 |]
let private distanceBase =
    [| 1; 2; 3; 4; 5; 7; 9; 13; 17; 25; 33; 49; 65; 97; 129; 193; 257; 385; 513; 769; 1025; 1537; 2049; 3073; 4097; 6145; 8193; 12289; 16385; 24577 |]
let private distanceExtra = [| 0; 0; 0; 0; 1; 1; 2; 2; 3; 3; 4; 4; 5; 5; 6; 6; 7; 7; 8; 8; 9; 9; 10; 10; 11; 11; 12; 12; 13; 13 |]
let private codeOrder = [| 16; 17; 18; 0; 8; 7; 9; 6; 10; 5; 11; 4; 12; 3; 13; 2; 14; 1; 15 |]

/// A canonical Huffman code: how many codes of each length, the symbols in code order.
type private Huffman = { Counts: int[]; Symbols: int[] }

let private huffman (lengths: int[]) : Huffman =
    let counts = Array.zeroCreate 16
    for length in lengths do counts.[length] <- counts.[length] + 1
    counts.[0] <- 0
    let offsets = Array.zeroCreate 16
    for length in 1 .. 14 do offsets.[length + 1] <- offsets.[length] + counts.[length]
    let symbols = Array.zeroCreate lengths.Length
    for symbol in 0 .. lengths.Length - 1 do
        if lengths.[symbol] <> 0 then
            symbols.[offsets.[lengths.[symbol]]] <- symbol
            offsets.[lengths.[symbol]] <- offsets.[lengths.[symbol]] + 1
    { Counts = counts; Symbols = symbols }

/// The bytes a raw deflate stream holds - up to where it ends, or up to where the data is cut off.
let raw (data: byte[]) : byte[] =
    let output = ResizeArray<byte>(data.Length * 2)
    let mutable position = 0
    let mutable buffer = 0
    let mutable count = 0
    let bits (n: int) =
        while count < n do
            if position >= data.Length then raise Truncated
            buffer <- buffer ||| (int data.[position] <<< count)
            position <- position + 1
            count <- count + 8
        let value = buffer &&& ((1 <<< n) - 1)
        buffer <- buffer >>> n
        count <- count - n
        value
    let decode (code: Huffman) =
        let mutable value = 0
        let mutable first = 0
        let mutable index = 0
        let mutable result = -1
        let mutable length = 1
        while result < 0 do
            if length > 15 then raise Truncated
            value <- value ||| bits 1
            let n = code.Counts.[length]
            if value - n < first then result <- code.Symbols.[index + (value - first)]
            else
                index <- index + n
                first <- (first + n) <<< 1
                value <- value <<< 1
                length <- length + 1
        result
    let fixedLiterals = lazy (huffman (Array.init 288 (fun symbol -> if symbol < 144 then 8 elif symbol < 256 then 9 elif symbol < 280 then 7 else 8)))
    let fixedDistances = lazy (huffman (Array.create 30 5))
    let codes (literals: Huffman) (distances: Huffman) =
        let mutable symbol = decode literals
        while symbol <> 256 do
            if symbol < 256 then output.Add (byte symbol)
            else
                let index = symbol - 257
                if index >= 29 then raise Truncated
                let length = lengthBase.[index] + bits lengthExtra.[index]
                let distanceSymbol = decode distances
                if distanceSymbol >= 30 then raise Truncated
                let distance = distanceBase.[distanceSymbol] + bits distanceExtra.[distanceSymbol]
                if distance > output.Count then raise Truncated
                let start = output.Count - distance
                for k in 0 .. length - 1 do output.Add output.[start + k]
            symbol <- decode literals
    try
        let mutable last = false
        while not last do
            last <- bits 1 = 1
            match bits 2 with
            | 0 ->
                // (A stored block: from the next byte, its length and that length's complement, then the bytes.)
                buffer <- 0
                count <- 0
                if position + 4 > data.Length then raise Truncated
                let length = int data.[position] ||| (int data.[position + 1] <<< 8)
                position <- position + 4
                for _ in 1 .. length do
                    if position >= data.Length then raise Truncated
                    output.Add data.[position]
                    position <- position + 1
            | 1 -> codes fixedLiterals.Value fixedDistances.Value
            | 2 ->
                let literalCount = bits 5 + 257
                let distanceCount = bits 5 + 1
                let codeCount = bits 4 + 4
                let codeLengths = Array.zeroCreate 19
                for index in 0 .. codeCount - 1 do codeLengths.[codeOrder.[index]] <- bits 3
                let lengthCode = huffman codeLengths
                let lengths = Array.zeroCreate (literalCount + distanceCount)
                let mutable index = 0
                while index < lengths.Length do
                    let symbol = decode lengthCode
                    if symbol < 16 then
                        lengths.[index] <- symbol
                        index <- index + 1
                    else
                        let value, repeat =
                            match symbol with
                            | 16 -> (if index = 0 then raise Truncated else lengths.[index - 1]), 3 + bits 2
                            | 17 -> 0, 3 + bits 3
                            | _ -> 0, 11 + bits 7
                        for _ in 1 .. repeat do
                            if index < lengths.Length then
                                lengths.[index] <- value
                                index <- index + 1
                codes (huffman lengths.[0 .. literalCount - 1]) (huffman lengths.[literalCount ..])
            | _ -> raise Truncated
    with Truncated -> ()
    output.ToArray()
