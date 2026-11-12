module Rbt.Word.PdfBrowser

#if FABLE_COMPILER
open Fable.Core
open Fable.Core.JsInterop

module Pdf = Rbt.Word.Pdf

[<Emit("fetch(new URL('./ucd.dat', import.meta.url)).then(response => { if (!response.ok) throw new Error('PDF Unicode data could not be loaded'); return response.arrayBuffer(); }).then(buffer => new Uint8Array(buffer))")>]
let private fetchUnicodeData () : JS.Promise<byte[]> = jsNative

let mutable private unicodeLoad: JS.Promise<unit> option = None

let private ensureUnicodeData () =
    match unicodeLoad with
    | Some pending -> pending
    | None ->
        let pending =
            async {
                try
                    let! bytes = fetchUnicodeData () |> Async.AwaitPromise
                    Rbt.Unicode.Ucd.loadData bytes
                with error ->
                    unicodeLoad <- None
                    return raise error
            } |> Async.StartAsPromise
        unicodeLoad <- Some pending
        pending

[<Emit("($0 instanceof Uint8Array ? $0 : ($0 instanceof ArrayBuffer ? new Uint8Array($0) : null))")>]
let private fontBytes (_: obj) : byte[] = jsNative

let private face (value: obj) : Pdf.FontFace =
    let field lower upper =
        let lowerValue: obj = value?(lower)
        if isNullOrUndefined lowerValue then value?(upper) else lowerValue
    let family: string = field "family" "Family" |> unbox
    let data = field "data" "Data" |> fontBytes
    if System.String.IsNullOrWhiteSpace family || isNull data || data.Length = 0 then
        failwith "PDF fonts require a family and nonempty TrueType bytes"
    let flag name upper = let value = field name upper in not (isNullOrUndefined value) && unbox<bool> value
    let bold, italic = flag "bold" "Bold", flag "italic" "Italic"
    let name: string = field "faceName" "FaceName" |> unbox
    { FaceName = if System.String.IsNullOrWhiteSpace name then family.Replace(" ", "") + (if bold then "-Bold" else "") + (if italic then "-Italic" else "") else name
      Family = family; Bold = bold; Italic = italic; Data = data }

// Browser decoding supplies dimensions and converts formats the PDF writer cannot
// decode to JPEG on white. JPEG bytes are retained instead of re-encoding them.
[<Emit("""(async source => {
    const image = new Image();
    image.crossOrigin = 'anonymous';
    image.src = source;
    await image.decode();
    let bytes;
    if (/^data:image\/jpe?g[;,]/i.test(source)) {
        bytes = new Uint8Array(await (await fetch(source)).arrayBuffer());
    } else {
        const canvas = document.createElement('canvas');
        canvas.width = image.naturalWidth; canvas.height = image.naturalHeight;
        const context = canvas.getContext('2d');
        context.fillStyle = 'white'; context.fillRect(0, 0, canvas.width, canvas.height);
        context.drawImage(image, 0, 0);
        const blob = await new Promise((resolve, reject) => canvas.toBlob(value => value ? resolve(value) : reject(new Error('PDF image conversion failed')), 'image/jpeg', 0.95));
        bytes = new Uint8Array(await blob.arrayBuffer());
    }
    return {width:image.naturalWidth, height:image.naturalHeight, bytes};
})($0)""")>]
let private decodeImage (_: string) : JS.Promise<obj> = jsNative

[<Emit("new Blob([$0], {type:'application/pdf'})")>]
let private pdfBlob (_: byte[]) : obj = jsNative

/// A captured immutable whole-document snapshot, not the editor's later state.
let exportPdf (document: Rbt.Word.Document.Document) (fonts: obj array) (allowLinkedImages: bool) : JS.Promise<obj> =
    async {
        if isNull fonts || fonts.Length = 0 then failwith "Configure PDF font bytes before exporting"
        let options: Pdf.RenderOptions = {
            DefaultFontFamily = "Calibri"
            Fonts = fonts |> Array.map face |> Array.toList
            MeasureText = None
            ShapeText = None
            Images = []
        }
        do! ensureUnicodeData () |> Async.AwaitPromise
        let plan = Pdf.collectPaintPlan options document
        let images = ResizeArray<Pdf.ImageResource>()
        for source in plan.ImageSources do
            if not allowLinkedImages && not (source.StartsWith("data:") || source.StartsWith("blob:")) then
                failwith "Linked image PDF export is disabled for this document"
            let! image = decodeImage source |> Async.AwaitPromise
            images.Add {
                Source = source; Width = image?width; Height = image?height
                ColorSpace = "DeviceRGB"; BitsPerComponent = 8
                Data = image?bytes; Filter = Some "DCTDecode"
            }
        return Pdf.toBytesWithPaintPlan { options with Images = List.ofSeq images } plan |> pdfBlob
    } |> Async.StartAsPromise
#endif
