module Rbt.WordWisp.WordRender

// The document render pipeline for wisp-word, lifted from the native WordEditor — but it was always
// Renderer-based (only the default font + the canvas/input were Composit). Here it feeds a Wisp canvas
// instead. Loads nothing toolkit-specific: Rbt.Word does layout + paint (PaintCommand list); this
// converts those to Rbt.Renderer.Cmd. (DrawShape is skipped for now — text/fills/images cover docx bodies.)

open System
open System.Collections.Generic
open System.IO
open Rbt.Renderer
open Rbt.Word.Document

module WordLayout = Rbt.Word.Layout
module WordPaint = Rbt.Word.Paint
module WordTextBuffer = Rbt.Word.TextBuffer

let inline private px (v: float) : float<px> = LanguagePrimitives.FloatWithMeasure v
let private pointsToPixels points = points * 96.0 / 72.0

let mutable private defaultFont : Rbt.Renderer.Text.FontData option = None
let private rendererFont () =
    match defaultFont with
    | Some f -> f
    | None ->
        let f = Rbt.Renderer.Text.loadFontFile "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
        defaultFont <- Some f; f

let private wordFontCache = Dictionary<string, Rbt.Renderer.Text.FontData option>()
let private existingDir path = if Directory.Exists path then Some path else None
let private wordFontDirs () =
    [ Environment.GetEnvironmentVariable "RB_WORD_FONT_DIR"
      Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "artifacts/fonts/windows"))
      "/usr/share/fonts/truetype/liberation" ]
    |> List.choose (fun p -> if String.IsNullOrWhiteSpace p then None else existingDir p)

let private fontFilesFor (family: string option) (bold: bool) (italic: bool) =
    match family |> Option.map (fun v -> v.Trim().ToLowerInvariant()) with
    | Some "calibri" -> if bold then [ "calibrib.ttf"; "LiberationSans-Bold.ttf" ] elif italic then [ "calibrii.ttf"; "LiberationSans-Italic.ttf" ] else [ "calibri.ttf"; "LiberationSans-Regular.ttf" ]
    | Some "arial" -> if bold then [ "arialbd.ttf"; "LiberationSans-Bold.ttf" ] else [ "arial.ttf"; "LiberationSans-Regular.ttf" ]
    | Some "times new roman" -> if bold then [ "timesbd.ttf"; "LiberationSerif-Bold.ttf" ] else [ "times.ttf"; "LiberationSerif-Regular.ttf" ]
    | _ -> []

let private tryLoadWordFont (family: string option) (bold: bool) (italic: bool) =
    let key = sprintf "%A|%b|%b" family bold italic
    match wordFontCache.TryGetValue key with
    | true, cached -> cached
    | _ ->
        let loaded =
            fontFilesFor family bold italic
            |> List.tryPick (fun fileName ->
                wordFontDirs () |> List.tryPick (fun dir ->
                    let path = Path.Combine(dir, fileName)
                    if File.Exists path then (try Some(Rbt.Renderer.Text.loadFontFile path) with _ -> None) else None))
        wordFontCache.[key] <- loaded
        loaded

let private fontForTextStyle (style: WordTextBuffer.TextStyle) =
    tryLoadWordFont style.FontFamily style.Bold style.Italic |> Option.defaultWith rendererFont
let private fontForPaintStyle (style: WordPaint.TextPaintStyle) =
    tryLoadWordFont style.FontFamily style.Bold style.Italic |> Option.defaultWith rendererFont

let measureText (text: string) (style: WordTextBuffer.TextStyle) =
    let font = fontForTextStyle style
    Rbt.Renderer.Text.measureWidth font (pointsToPixels style.FontSize) text + WordTextBuffer.characterSpacingAdvance text style

let private parseHexColor (value: string) =
    let value =
        match value.Trim().ToLowerInvariant() with
        | "black" -> "#000000" | "white" -> "#ffffff" | "red" -> "#ff0000" | "blue" -> "#0000ff"
        | "green" -> "#008000" | "darkgray" -> "#808080" | "lightgray" -> "#c0c0c0" | other -> other
    let parseByte (s: string) =
        match Byte.TryParse(s, Globalization.NumberStyles.HexNumber, Globalization.CultureInfo.InvariantCulture) with
        | true, v -> Some v | _ -> None
    let fromParts r g b = match parseByte r, parseByte g, parseByte b with Some r, Some g, Some b -> Color.rgb r g b | _ -> Color.black
    let s = if value.StartsWith "#" then value.Substring 1 else value
    match s.Length with
    | 3 -> let e (c: char) = string c + string c in fromParts (e s.[0]) (e s.[1]) (e s.[2])
    | 6 -> fromParts s.[0..1] s.[2..3] s.[4..5]
    | 8 -> fromParts s.[0..1] s.[2..3] s.[4..5]
    | _ -> Color.black

let private textFont (style: WordPaint.TextPaintStyle) = fontForPaintStyle style

let layoutSettings (width: float) (doc: Document) : WordLayout.LayoutSettings =
    let sourceWidth = max 1.0 doc.PageSetup.PageWidth
    let scale = width / sourceWidth
    let lineHeight = 24.0
    let hasFooter = not doc.Footers.IsEmpty || not doc.FooterParts.IsEmpty
    let footerReserve = if hasFooter then doc.PageSetup.FooterDistance * scale + lineHeight * 1.5 else doc.PageSetup.MarginBottom * scale
    { Width = width
      PageHeight = max 1.0 (doc.PageSetup.PageHeight * scale)
      FontSize = 16.0
      LineHeight = lineHeight
      PaddingLeft = doc.PageSetup.MarginLeft * scale
      PaddingTop = Rbt.Word.WordLayout.Page.bodyTopWithRunningHeader doc * scale
      MarginRight = doc.PageSetup.MarginRight * scale
      MarginBottom = max (doc.PageSetup.MarginBottom * scale) footerReserve
      InTableCell = false
      TableContinuationHeaderHeight = 0.0
      LongTrailingSpaceFitSafety = 0.25 }

let rec toRendererCommand = function
    | WordPaint.DrawText(text, style, x, baseline, _) ->
        [ Cmd.DrawGlyphs(vec2 x baseline, Rbt.Renderer.Text.layout (textFont style) (pointsToPixels style.FontSize) text, Solid(parseHexColor style.Color)) ]
    | WordPaint.DrawUnderline(color, _, x, y, width) ->
        [ Cmd.Line(vec2 x y, vec2 (x + width) y, parseHexColor color, px 1.0) ]
    | WordPaint.DrawImage(_, x, y, width, height, _, _, _, _) ->
        [ Cmd.StrokeRect(rect x y width height, Color.rgb 180uy 180uy 180uy, px 1.0) ]
    | WordPaint.FillRect(color, x, y, width, height) ->
        [ Cmd.FillRect(rect x y width height, parseHexColor color) ]
    | WordPaint.StrokeRect(color, lineWidth, x, y, width, height) ->
        [ Cmd.StrokeRect(rect x y width height, parseHexColor color, px lineWidth) ]
    | WordPaint.StrokeLine(color, lineWidth, _, _, x1, y1, x2, y2) ->
        [ Cmd.Line(vec2 x1 y1, vec2 x2 y2, parseHexColor color, px lineWidth) ]
    | WordPaint.DrawShape _ -> []
    // (A text box's text and fill, unclipped and upright: this painter draws no shapes.)
    | WordPaint.DrawGroup (_, _, _, _, _, _, inner) -> inner |> List.collect toRendererCommand

/// All paint commands for the document, laid out to `width`, in page-local coordinates.
let documentCommands (width: float) (doc: Document) : Cmd list =
    let settings = layoutSettings width doc
    let _, contentCommands, _ = WordPaint.collectDocumentCommands settings "Calibri" measureText (WordLayout.rangeMeasurerFromTextMeasurer measureText) doc
    let commands = ResizeArray<Cmd>()
    for paintCommand in contentCommands do
        commands.AddRange(toRendererCommand paintCommand)
    List.ofSeq commands

/// Total laid-out height of the document at the given width.
let documentHeight (width: float) (doc: Document) : float =
    let settings = layoutSettings width doc
    match WordLayout.layoutDocument settings measureText doc with
    | [] -> settings.PaddingTop * 2.0 + settings.LineHeight
    | layouts ->
        let mutable bottom = 0.0
        for layout in layouts do
            bottom <- max bottom (layout.Y + layout.Height)
        bottom + settings.MarginBottom
