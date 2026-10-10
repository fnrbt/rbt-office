/// What a document looks like once it is laid out, block by block, in a form
/// two layouts can be compared by: where each block is (page, y, height, line
/// count, line heights), its text, and the formatting it was laid out with —
/// the paragraph's after its styles are applied, the runs' as the lines hold
/// them. Used by the DOCX round-trip guards (tools/docx-roundtrip and
/// tests/word/roundtrip-layout.fsx): a document saved and opened again must
/// give the same fingerprint.
module Rbt.Word.LayoutFingerprint

open Rbt.Word.Document
open Rbt.Word.Layout

type BlockFingerprint = {
    Page: int
    Y: float
    Height: float
    LineCount: int
    /// Height of each line, rounded to hundredths, separated by spaces.
    LineHeights: string
    Text: string
    /// Resolved paragraph formatting.
    Paragraph: string
    /// Resolved run formatting: "count*format" for each stretch of characters with one format.
    Runs: string
}

let private num (value: float) =
    let rounded = System.Math.Round(value * 100.0) / 100.0
    // (Not "-0".)
    if rounded = 0.0 then "0" else string rounded

let private optNum (value: float option) =
    match value with
    | Some v -> num v
    | None -> "-"

let private optStr (value: string option) =
    match value with
    | Some v -> v
    | None -> "-"

let private optBool (value: bool option) =
    match value with
    | Some true -> "1"
    | Some false -> "0"
    | None -> "-"

let private flag (name: string) (value: bool) = if value then name else ""

let private runFormat (style: TextBuffer.TextStyle) =
    String.concat "," [
        optStr style.FontFamily
        num style.FontSize
        flag "b" style.Bold + flag "i" style.Italic + flag "u" style.Underline + flag "caps" style.AllCaps + flag "small" style.SmallCaps
        + flag "strike" style.Strike + flag "dstrike" style.DoubleStrike + flag "hidden" style.Hidden
        (if style.Underline then optStr style.UnderlineStyle + "/" + optStr style.UnderlineColor else "")
        optStr style.Color
        optStr style.Highlight
        optStr style.VertAlign
        num style.Position
        num style.CharacterSpacing
        optStr style.EmphasisMark
        optStr style.RevisionType
        flag "link" (style.Hyperlink.IsSome || style.InactiveHyperlink.IsSome)
        flag "field" style.FieldCode.IsSome
        flag "note" style.NoteId.IsSome
    ]

let private border (value: ParagraphBorder option) =
    match value with
    | Some b -> sprintf "%s/%s/%s/%s" (optStr b.Style) (optNum b.Width) (optStr b.Color) (optNum b.Space)
    | None -> "-"

let private paragraphFormat (layout: BlockLayout) =
    let m = layout.Metadata
    let alignment =
        match m.Alignment with
        | Left -> "left"
        | Center -> "center"
        | Right -> "right"
        | Justify -> "justify"
    let list =
        match m.ListStyle with
        | NoList -> "none"
        | Bullet -> "bullet"
        | Numbered -> "numbered"
    let tabs =
        m.TabStops
        |> List.map (fun tab -> sprintf "%s:%s:%s" (num tab.Position) tab.Alignment (optStr tab.Leader))
        |> String.concat " "
    let borders =
        match m.ParagraphBorders with
        | Some b -> String.concat " " [ border b.Top; border b.Bottom; border b.Left; border b.Right ]
        | None -> "-"
    String.concat ";" [
        "style=" + optStr m.Style
        "align=" + alignment
        "ind=" + optNum m.IndentLeft + "/" + optNum m.IndentRight + "/" + optNum m.FirstLineIndent
        "space=" + optNum m.SpacingBefore + "/" + optNum m.SpacingAfter + "/" + optBool m.ContextualSpacing
        "line=" + optNum m.LineSpacing + "/" + optStr m.LineSpacingRule
        "keep=" + optBool m.KeepNext + optBool m.KeepLines + optBool m.WidowControl
        "break=" + flag "before" m.PageBreakBefore + flag "after" m.PageBreakAfter + optStr m.SectionBreak
        // (The number format and level text show in the prefix; a list without a
        // definition has none and is shown, and saved, as decimal.)
        "list=" + list + "/" + string m.ListLevel
        "prefix=" + layout.Prefix + "/" + num layout.PrefixWidth
        "box=" + num layout.Indent + "/" + num layout.AvailableWidth
        "tabs=" + tabs
        "rtl=" + flag "1" m.Rtl
        "shade=" + optStr m.BackgroundColor
        "borders=" + borders
    ]

/// The fingerprint of one laid-out block. `pageOf` gives the page a y is on.
let ofLayout (pageOf: float -> int) (layout: BlockLayout) : BlockFingerprint =
    let text = System.Text.StringBuilder()
    let heights = System.Text.StringBuilder()
    let runs = ResizeArray<string * int>()
    let add (format: string) (count: int) =
        if count > 0 then
            if runs.Count > 0 && fst runs.[runs.Count - 1] = format then
                runs.[runs.Count - 1] <- (format, snd runs.[runs.Count - 1] + count)
            else
                runs.Add((format, count))
    let mutable lineCount = 0
    for line in layout.Lines do
        lineCount <- lineCount + 1
        if heights.Length > 0 then heights.Append(" ") |> ignore
        heights.Append(num line.Height) |> ignore
        for segment in line.Segments do
            match segment with
            | TextSegment (segmentText, style, _, _, _) ->
                text.Append(segmentText) |> ignore
                add (runFormat style) segmentText.Length
            | ImageSegment (_, _, _, width, height, _, floating, wrap, _, _, _, _, _) ->
                text.Append("￼") |> ignore
                add (sprintf "image %s x %s %s%s" (num width) (num height) (flag "floating " floating) (optStr wrap)) 1
            | MathSegment (_, style, _, width, _, _, _, _) ->
                text.Append("￼") |> ignore
                add ("math " + num width + " " + runFormat style) 1
    let firstY =
        match layout.Lines with
        | line :: _ -> line.Y
        | [] -> layout.Y
    { Page = pageOf firstY
      Y = layout.Y
      Height = layout.Height
      LineCount = lineCount
      LineHeights = heights.ToString()
      Text = text.ToString()
      Paragraph = paragraphFormat layout
      Runs = runs |> Seq.map (fun (format, count) -> string count + "*" + format) |> String.concat " | " }

let ofLayouts (pageOf: float -> int) (layouts: BlockLayout list) : BlockFingerprint list =
    layouts |> List.map (ofLayout pageOf)

/// The first difference between two fingerprints, if any: y and height within
/// `tolerance`, everything else exactly.
let firstDifference (tolerance: float) (before: BlockFingerprint list) (after: BlockFingerprint list) : string option =
    let a = List.toArray before
    let b = List.toArray after
    let mutable found = None
    if a.Length <> b.Length then
        found <- Some (sprintf "%d blocks became %d" a.Length b.Length)
    let mutable i = 0
    let shorten (text: string) = if text.Length > 40 then text.Substring(0, 40) else text
    while found.IsNone && i < min a.Length b.Length do
        let x = a.[i]
        let y = b.[i]
        let differs (what: string) (p: string) (q: string) =
            if found.IsNone && p <> q then
                found <- Some (sprintf "block %d (%s): %s %s -> %s" i (shorten x.Text) what p q)
        differs "text" x.Text y.Text
        differs "page" (string x.Page) (string y.Page)
        if abs (x.Y - y.Y) > tolerance then differs "y" (num x.Y) (num y.Y)
        if abs (x.Height - y.Height) > tolerance then differs "height" (num x.Height) (num y.Height)
        differs "lines" (string x.LineCount) (string y.LineCount)
        differs "line heights" x.LineHeights y.LineHeights
        differs "paragraph format" x.Paragraph y.Paragraph
        differs "run format" x.Runs y.Runs
        i <- i + 1
    found
