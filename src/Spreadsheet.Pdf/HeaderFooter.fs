module RedBat.Spreadsheet.Pdf.HeaderFooter

/// Parse and substitute Excel's header / footer field codes. The PDF
/// emitter calls `resolve` per page to turn the model's raw header/footer
/// strings (Excel's `&L &C &R` section markers + `&P`/`&N`/`&A`/… field
/// codes) into three plain-text sections ready to draw.
///
/// Pure string logic with no PDF-specific dependencies, so it could be
/// lifted into a cross-document `Office.Pdf` shared layer later.

/// Per-page substitution context. The emitter fills this in for each page
/// before resolving the header/footer.
type PageContext = {
    PageNumber: int
    TotalPages: int
    SheetName: string
    /// Bare file name without directory (matches Excel's `&F`).
    FileName: string
    /// Directory the file lives in (matches Excel's `&Z`, sans trailing /).
    FilePath: string
    /// Time at which the document is being rendered. `&D` and `&T` format
    /// from this.
    Now: System.DateTime
}

/// One contiguous header/footer text run after field substitution, carrying
/// the style state active for that run.
type StyledRun = {
    Text: string
    FontName: string option
    FontSize: float option
    Bold: bool
    Italic: bool
    Color: string option
    Underline: bool
    DoubleUnderline: bool
    Strikethrough: bool
    /// Superscript / subscript scale + baseline shift. None = baseline.
    /// `&X` sets Superscript, `&Y` sets Subscript. Mutually exclusive —
    /// toggling one clears the other (matches Excel's HF parser).
    Superscript: bool
    Subscript: bool
}

/// One header / footer section with its plain text plus per-run styling.
/// The section-level FontName / FontSize / Bold / Italic fields mirror the
/// first text run for compatibility with older callers and focused tests.
type StyledSection = {
    Text: string
    FontName: string option
    FontSize: float option
    Bold: bool
    Italic: bool
    Runs: StyledRun list
}

/// One header or footer's three sections, post-substitution.
type Section = {
    Left: StyledSection
    Center: StyledSection
    Right: StyledSection
}

/// Walk a section's raw string, substituting field codes (`&P`, `&A`, …)
/// and turning formatting codes into styled runs. A `&"Font,Style"` carries
/// the family and an optional style tag (`Regular`, `Bold`, `Italic`,
/// `BoldItalic`); `&NN` changes size; `&B` / `&I` toggle weight / slant;
/// `&Kxxxxxx` changes text color.
let private resolveSection (ctx: PageContext) (raw: string) : StyledSection =
    let text = System.Text.StringBuilder(raw.Length)
    let runText = System.Text.StringBuilder(raw.Length)
    let runs = System.Collections.Generic.List<StyledRun>()
    let mutable fontName : string option = None
    let mutable fontSize : float option = None
    let mutable bold = false
    let mutable italic = false
    let mutable color : string option = None
    let mutable underline = false
    let mutable doubleUnderline = false
    let mutable strikethrough = false
    let mutable superscript = false
    let mutable subscript = false
    let mutable i = 0
    let flushRun () =
        if runText.Length > 0 then
            runs.Add {
                Text = runText.ToString()
                FontName = fontName
                FontSize = fontSize
                Bold = bold
                Italic = italic
                Color = color
                Underline = underline
                DoubleUnderline = doubleUnderline
                Strikethrough = strikethrough
                Superscript = superscript
                Subscript = subscript }
            runText.Clear() |> ignore
    let appendText (value: obj) =
        let s = string value
        text.Append s |> ignore
        runText.Append s |> ignore
    let parseHexColor startIndex =
        let stop = min raw.Length (startIndex + 6)
        let mutable j = startIndex
        while j < stop && System.Uri.IsHexDigit raw.[j] do j <- j + 1
        let len = j - startIndex
        let value =
            if len = 6 then
                Some ("#" + raw.Substring(startIndex, len).ToUpperInvariant())
            else None
        value, j
    while i < raw.Length do
        if raw.[i] = '&' && i + 1 < raw.Length then
            let next = raw.[i + 1]
            match next with
            | '&' -> appendText '&'; i <- i + 2
            | 'P' -> appendText ctx.PageNumber; i <- i + 2
            | 'N' -> appendText ctx.TotalPages; i <- i + 2
            | 'A' -> appendText ctx.SheetName; i <- i + 2
            | 'F' -> appendText ctx.FileName; i <- i + 2
            | 'Z' -> appendText ctx.FilePath; i <- i + 2
            | 'D' -> appendText (ctx.Now.ToString("M/d/yyyy", System.Globalization.CultureInfo.InvariantCulture)); i <- i + 2
            | 'T' -> appendText (ctx.Now.ToString("h:mm tt", System.Globalization.CultureInfo.InvariantCulture)); i <- i + 2
            | 'B' ->
                flushRun ()
                bold <- not bold
                i <- i + 2
            | 'I' ->
                flushRun ()
                italic <- not italic
                i <- i + 2
            | 'U' ->
                flushRun ()
                underline <- not underline
                // `&U` and `&E` are mutually exclusive in Excel: toggling
                // single-underline on clears the double-underline state.
                if underline then doubleUnderline <- false
                i <- i + 2
            | 'E' ->
                flushRun ()
                doubleUnderline <- not doubleUnderline
                if doubleUnderline then underline <- false
                i <- i + 2
            | 'S' ->
                flushRun ()
                strikethrough <- not strikethrough
                i <- i + 2
            | 'X' ->
                flushRun ()
                superscript <- not superscript
                if superscript then subscript <- false
                i <- i + 2
            | 'Y' ->
                flushRun ()
                subscript <- not subscript
                if subscript then superscript <- false
                i <- i + 2
            // Outline / shadow codes don't change anything we render.
            // `&G` reserves space for a header/footer image, which we don't
            // emit yet — consume it so the literal "G" doesn't leak into
            // the rendered text.
            | 'O' | 'H' | 'G' -> i <- i + 2
            | 'K' ->
                flushRun ()
                let parsed, nextIndex = parseHexColor (i + 2)
                color <- parsed |> Option.orElse color
                i <- nextIndex
            | '"' ->
                // &"Font,Style" — `Font` is the family, `Style` is one of
                // Regular / Bold / Italic / BoldItalic (case-insensitive).
                flushRun ()
                i <- i + 2
                let startQuote = i
                while i < raw.Length && raw.[i] <> '"' do i <- i + 1
                let inner = raw.Substring(startQuote, i - startQuote)
                let parts = inner.Split(',')
                let family = if parts.Length > 0 then parts.[0] else ""
                let style = if parts.Length > 1 then parts.[1] else ""
                if family.Length > 0 then fontName <- Some family
                if style.Length > 0 then
                    let styleLow = style.ToLowerInvariant()
                    bold <- styleLow.Contains "bold"
                    italic <- styleLow.Contains "italic"
                if i < raw.Length then i <- i + 1
            | c when System.Char.IsDigit c ->
                // &NN font-size code (1- or 3-digit; up to 409 in spec).
                flushRun ()
                i <- i + 1  // consume the `&`; i now points at the first digit
                let startDigits = i
                while i < raw.Length && System.Char.IsDigit raw.[i] do i <- i + 1
                let digits = raw.Substring(startDigits, i - startDigits)
                match System.Double.TryParse digits with
                | true, v -> fontSize <- Some v
                | _ -> ()
            | _ ->
                // Unrecognized code; emit as-is rather than silently drop.
                appendText raw.[i]
                i <- i + 1
        else
            appendText raw.[i]
            i <- i + 1
    flushRun ()
    let runList = List.ofSeq runs
    let firstRun =
        runList
        |> List.tryFind (fun r -> not (System.String.IsNullOrEmpty r.Text))
    { Text = text.ToString()
      FontName = firstRun |> Option.bind (fun r -> r.FontName)
      FontSize = firstRun |> Option.bind (fun r -> r.FontSize)
      Bold = firstRun |> Option.map (fun r -> r.Bold) |> Option.defaultValue false
      Italic = firstRun |> Option.map (fun r -> r.Italic) |> Option.defaultValue false
      Runs = runList }

/// Split an Excel header/footer string on the `&L`/`&C`/`&R` section
/// markers. Text before any marker is treated as the center section,
/// matching Excel's behavior when no marker is present. Field codes
/// (`&P`, `&N`, …) pass through unchanged — `substituteFields` resolves
/// them at the per-page step.
let private parseSections (raw: string) : string * string * string =
    let left = System.Text.StringBuilder()
    let center = System.Text.StringBuilder()
    let right = System.Text.StringBuilder()
    let mutable current = center
    let mutable i = 0
    while i < raw.Length do
        if raw.[i] = '&' && i + 1 < raw.Length then
            match raw.[i + 1] with
            | 'L' -> current <- left; i <- i + 2
            | 'C' -> current <- center; i <- i + 2
            | 'R' -> current <- right; i <- i + 2
            | _ ->
                // Pass other &-codes through to the current section;
                // substituteFields handles them later.
                current.Append raw.[i] |> ignore
                current.Append raw.[i + 1] |> ignore
                i <- i + 2
        else
            current.Append raw.[i] |> ignore
            i <- i + 1
    left.ToString(), center.ToString(), right.ToString()

/// Resolve an Excel header/footer template to its three sections with
/// every supported field code substituted for the given page.
let resolve (ctx: PageContext) (raw: string) : Section =
    let l, c, r = parseSections raw
    { Left = resolveSection ctx l
      Center = resolveSection ctx c
      Right = resolveSection ctx r }
