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

/// One header or footer's three sections, post-substitution.
type Section = {
    Left: string
    Center: string
    Right: string
}

let private substituteFields (ctx: PageContext) (raw: string) : string =
    let sb = System.Text.StringBuilder(raw.Length)
    let mutable i = 0
    while i < raw.Length do
        if raw.[i] = '&' && i + 1 < raw.Length then
            let next = raw.[i + 1]
            match next with
            | '&' -> sb.Append '&' |> ignore; i <- i + 2
            | 'P' -> sb.Append ctx.PageNumber |> ignore; i <- i + 2
            | 'N' -> sb.Append ctx.TotalPages |> ignore; i <- i + 2
            | 'A' -> sb.Append ctx.SheetName |> ignore; i <- i + 2
            | 'F' -> sb.Append ctx.FileName |> ignore; i <- i + 2
            | 'Z' -> sb.Append ctx.FilePath |> ignore; i <- i + 2
            | 'D' -> sb.Append (ctx.Now.ToString "yyyy-MM-dd") |> ignore; i <- i + 2
            | 'T' -> sb.Append (ctx.Now.ToString "HH:mm") |> ignore; i <- i + 2
            // Style-only codes we don't render in the plain-text output:
            // bold / italic / underline / strikethrough / sub / super /
            // outline / shadow / extended-formatting "&K{hex}" / etc.
            // Consume them silently so the surrounding text comes through
            // clean.
            | 'B' | 'I' | 'U' | 'E' | 'S' | 'X' | 'Y' | 'O' | 'H' -> i <- i + 2
            | 'K' ->
                // &Kxxxxxx — six-hex color code; just skip past it.
                i <- i + 2
                let stop = min raw.Length (i + 6)
                while i < stop && System.Uri.IsHexDigit raw.[i] do i <- i + 1
            | '"' ->
                // &"Font,Style" — skip to next ".
                i <- i + 2
                while i < raw.Length && raw.[i] <> '"' do i <- i + 1
                if i < raw.Length then i <- i + 1
            | c when System.Char.IsDigit c ->
                // &NN font-size code — skip digits.
                i <- i + 1
                while i < raw.Length && System.Char.IsDigit raw.[i] do i <- i + 1
            | _ ->
                // Unrecognized code; emit as-is rather than silently drop.
                sb.Append raw.[i] |> ignore
                i <- i + 1
        else
            sb.Append raw.[i] |> ignore
            i <- i + 1
    sb.ToString()

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
    { Left = substituteFields ctx l
      Center = substituteFields ctx c
      Right = substituteFields ctx r }
