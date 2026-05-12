module RedBat.Spreadsheet.Pdf

open System
open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Layout
open RedBat.Spreadsheet.Paint

module PdfDoc = FsPdf.Doc
module PdfDraw = FsPdf.Draw
module PdfColor = FsPdf.Color
module PdfFont = FsPdf.Font

/// One loadable TTF face the runner passes to us. Family is the abstract
/// family name from the cell ("Calibri", "Arial", ...); Bold/Italic mark the
/// variant.
type FontFace = {
    FaceName: string
    Family: string
    Bold: bool
    Italic: bool
    Data: byte[]
}

type RenderOptions = {
    /// Default family to use when a cell didn't request one.
    DefaultFontFamily: string
    Fonts: FontFace list
    /// Page size in CSS pixels (we convert to points internally).
    PageWidth: float
    PageHeight: float
}

/// Convert "#RRGGBB" / "RRGGBB" to a PDF RGB color in 0..1.
let private parseHexColor (color: string) : PdfColor.Color =
    let s = if color.StartsWith "#" then color.Substring 1 else color
    let clean = if s.Length = 8 then s.Substring 2 else s
    if clean.Length <> 6 then PdfColor.black
    else
        let byteOf (h: string) =
            try float (Convert.ToByte(h, 16)) / 255.0 with _ -> 0.0
        PdfColor.rgb (byteOf (clean.Substring(0, 2))) (byteOf (clean.Substring(2, 2))) (byteOf (clean.Substring(4, 2)))

/// CSS pixels → PDF points. Excel widths flow in CSS px (1 px = 72/96 pt).
let private pxToPt = 72.0 / 96.0

let private px (v: float) = v * pxToPt

/// Pick the loaded TTF for the cell's (family, bold, italic) triple. Falls
/// back to the family's regular weight, then to the default family's regular.
let private pickFace (options: RenderOptions) (familyName: string option) (bold: bool) (italic: bool) : FontFace option =
    let family = familyName |> Option.defaultValue options.DefaultFontFamily
    let exact = options.Fonts |> List.tryFind (fun f -> f.Family = family && f.Bold = bold && f.Italic = italic)
    let weighted = options.Fonts |> List.tryFind (fun f -> f.Family = family && f.Bold = bold && not f.Italic)
    let regular = options.Fonts |> List.tryFind (fun f -> f.Family = family && not f.Bold && not f.Italic)
    let fallback = options.Fonts |> List.tryFind (fun f -> f.Family = options.DefaultFontFamily && not f.Bold && not f.Italic)
    exact
    |> Option.orElseWith (fun () -> weighted)
    |> Option.orElseWith (fun () -> regular)
    |> Option.orElseWith (fun () -> fallback)
    |> Option.orElseWith (fun () -> options.Fonts |> List.tryHead)

/// Convert one Paint item into a list of PDF draw ops. Coordinates flip from
/// our top-left CSS-pixel space into PDF's bottom-left point space.
let private translateItem (options: RenderOptions) (faceByName: Map<string, PdfFont.TrueTypeFont>) (faceResource: Map<string, string>) (pageHeightPt: float) (item: PaintItem) : PdfDraw.DrawOp list =
    let yPdf y = pageHeightPt - px y
    match item with
    | FillRect (rect, color) ->
        [ PdfDraw.SetFillColor (parseHexColor color)
          PdfDraw.Rect (px rect.X, yPdf (rect.Y + rect.Height), px rect.Width, px rect.Height)
          PdfDraw.Fill ]
    | StrokeRect (rect, color, width) ->
        [ PdfDraw.SetStrokeColor (parseHexColor color)
          PdfDraw.SetLineWidth (px width)
          PdfDraw.SetDash ([], 0.0)
          PdfDraw.Rect (px rect.X, yPdf (rect.Y + rect.Height), px rect.Width, px rect.Height)
          PdfDraw.Stroke ]
    | GridLine (x1, y1, x2, y2, color, width) ->
        [ PdfDraw.SetStrokeColor (parseHexColor color)
          PdfDraw.SetLineWidth (px width)
          PdfDraw.SetDash ([], 0.0)
          PdfDraw.MoveTo (px x1, yPdf y1)
          PdfDraw.LineTo (px x2, yPdf y2)
          PdfDraw.Stroke ]
    | DrawText t ->
        match pickFace options t.FontName t.Bold t.Italic with
        | None -> []
        | Some face ->
            match Map.tryFind face.FaceName faceByName with
            | Some ttf ->
                let fontSizePt = px t.FontSize
                // Font units are scaled to font.UnitsPerEm; widths from
                // getTextWidth are in those same units.
                let widthUnits = PdfFont.getTextWidth ttf t.Text
                let measuredPt = float widthUnits / float ttf.UnitsPerEm * fontSizePt
                let rectXPt = px t.Rect.X
                let rectWidthPt = px t.Rect.Width
                let alignedXPt =
                    match t.Align with
                    | AlignLeft -> rectXPt
                    | AlignCenter -> rectXPt + (rectWidthPt - measuredPt) / 2.0
                    | AlignRight -> rectXPt + rectWidthPt - measuredPt
                // Baseline mirrors SpreadsheetGrid.toRendererCommands:
                // baseline = rect.Y + (height + fontSize)/2 - 2.
                let baselineY = t.Rect.Y + (t.Rect.Height + t.FontSize) / 2.0 - 2.0
                let hex = PdfFont.toGlyphHex ttf t.Text
                // addPage remaps SetFont (family, size) into the page's
                // resource name by looking the family up in builder.Fonts;
                // emit the face name, not the embedded resource name, so the
                // remap doesn't turn it into a /BaseFont /F0 garbage stub.
                [ PdfDraw.BeginText
                  PdfDraw.SetFillColor (parseHexColor t.Color)
                  PdfDraw.SetFont (face.FaceName, fontSizePt)
                  PdfDraw.TextMatrix (1.0, 0.0, 0.0, 1.0, alignedXPt, yPdf baselineY)
                  PdfDraw.ShowTextHex hex
                  PdfDraw.EndText ]
            | _ -> []

/// Render a workbook page to PDF bytes. Currently emits a single page sized
/// to `options.PageWidth/PageHeight` containing the whole used range; no
/// pagination yet.
let toBytes (options: RenderOptions) (model: SpreadsheetModel) : byte[] =
    let builder = PdfDoc.create ()

    // Load every TTF face into the document and remember the resource name
    // FsPdf assigned each, plus the parsed TrueTypeFont so we can measure
    // text widths and emit glyph hex for ShowTextHex.
    let faceTtf = System.Collections.Generic.Dictionary<string, PdfFont.TrueTypeFont>()
    let faceResource = System.Collections.Generic.Dictionary<string, string>()
    for face in options.Fonts do
        PdfDoc.loadFontBytes builder face.FaceName face.Data |> ignore
        match PdfDoc.getEmbeddedFont builder face.FaceName with
        | Some ef ->
            faceTtf.[face.FaceName] <- ef.Font
            faceResource.[face.FaceName] <- ef.Name
        | None -> ()

    let pageWidthPt = px options.PageWidth
    let pageHeightPt = px options.PageHeight

    let viewport = {
        ScrollX = 0.0
        ScrollY = 0.0
        Width = options.PageWidth
        Height = options.PageHeight
        HeaderHeight = 22.0
        RowHeaderWidth = 48.0
    }

    let faceTtfMap = faceTtf |> Seq.map (fun kv -> kv.Key, kv.Value) |> Map.ofSeq
    let faceResourceMap = faceResource |> Seq.map (fun kv -> kv.Key, kv.Value) |> Map.ofSeq

    let items = buildPaintItems model viewport
    let ops = items |> List.collect (translateItem options faceTtfMap faceResourceMap pageHeightPt)

    // addPage's first font list controls which entries make it into the
    // page's `/Font` resources dict — without it the SetFont (F#, …) ops we
    // already emitted reference unknown resource names and the reader
    // refuses to draw any text.
    let loadedFaceNames = options.Fonts |> List.map (fun f -> f.FaceName)
    PdfDoc.addPage builder (pageWidthPt, pageHeightPt) loadedFaceNames [] ops |> ignore
    PdfDoc.toBytes builder
