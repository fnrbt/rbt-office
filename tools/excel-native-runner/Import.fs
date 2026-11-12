module ExcelNativeRunner

open System
open Rbt.Spreadsheet.Model

module Xlsx = XlsxReader

/// Apply the native oracle's font metrics and pixel rounding after the shared XLSX import.
let private nativeGeometry (sheet: Xlsx.Worksheet) (workbook: Xlsx.Workbook) (model: SpreadsheetModel) =
    let cells = model.Cells
    let richRunsMap = model.RichRuns
    // Excel col.Width is in character units. The chars-to-pixels ratio is
    // the "Maximum Digit Width" (MDW) of the workbook's default font —
    // technically the rendered pixel width of the widest digit 0-9 in that
    // font at the workbook's default size, as Excel computes it. Use the
    // real loaded font instead of a Calibri-anchored linear fit, since
    // fonts vary enough at the same point size to skew column widths
    // (Verdana 10 needs 8 px/char, not the 7 a Calibri-only estimate gives).
    let defaultFontSize =
        if workbook.Styles.Length > 0 then
            workbook.Styles.[0].FontSize |> Option.defaultValue 11.0
        else 11.0
    let defaultFontName =
        if workbook.Styles.Length > 0 then
            workbook.Styles.[0].FontName
        else None
    let mdw =
        try
            let font = Rbt.OfficeNative.Fonts.resolve defaultFontName false false
            let pxSize = defaultFontSize * 96.0 / 72.0
            // Excel's MDW table is the integer pixel advance Excel uses for
            // chars-to-pixels column-width math at 96 DPI. Empirically:
            //   Calibri 11 = 7  (measured 7.43)
            //   Calibri 12 = 8  (measured 8.11)
            //   Calibri 14 = 10 (measured 9.46) — synthetic-vertalign
            //   Calibri 18 = 14 (measured 12.16, Microsoft table value)
            //   Arial   10 = 7  (measured 7.42)
            //   Verdana 10 = 8  (measured 8.48)
            // Floor alone matches Calibri 11/12 (where measured truncates
            // cleanly) but falls short at 14+ because Excel's GDI hinting
            // bumps the glyph advance past the raw font-table value. The
            // observation `Excel_MDW = pt - 4` lines up across all known
            // Calibri sizes ≥ 11; `max(floor(measured), pt-4)` therefore
            // covers both the small-size (floor wins, e.g. Verdana 10 = 8)
            // and large-size (pt-4 wins, e.g. Calibri 14 = 10) regimes.
            let zero = Rbt.Renderer.Text.measureWidth font pxSize "0"
            let measured = System.Math.Floor zero
            let ptBased = defaultFontSize - 4.0
            max 7.0 (max measured ptBased)
        with _ ->
            // Renderer unavailable (font cache miss, etc.) — fall back to
            // the original Calibri-anchored linear estimate.
            max 7.0 (System.Math.Round(defaultFontSize * 7.0 / 11.0))
    // OOXML §18.3.1.13 width formula:
    //   width_stored = TRUNC((chars * mdw + 5) / mdw * 256) / 256
    // So when we have a STORED <col width="N"> attribute, N already
    // encodes the 5-px padding adjustment: pixels = N * mdw with no
    // additional padding. When the column has NO stored width (uses the
    // sheet default chars value), the padding must be added explicitly:
    // pixels = chars * mdw + 5. Verified against phpsp-26template's col B
    // (stored 86.285 -> 604 px = 453 pt) and default 8.43 chars columns
    // (-> 64 px = 48 pt).
    let charsToPixelsDefault (chars: float) = chars * mdw + 5.0
    // Excel reports `cell.Width` in integer pixels (it rounds away the
    // 1/256 chars granularity from the stored width attr). Verified against
    // pandas-multiindex_no_index_names: stored width 5.41 × MDW 7 = 37.87
    // px which Excel rounds to 38 px (= 28.5 pt). Without rounding our
    // column x-positions drift by ~0.1 px per column, accumulating over a
    // wide sheet (the row 1 cells in col F sat 0.6 pt to the left of where
    // Excel placed them, 5 narrow columns of error).
    let storedWidthToPixels (storedWidth: float) =
        System.Math.Round(storedWidth * mdw)

    // sheetFormatPr defaults. Excel uses these for every column without an
    // explicit <col width="..."> entry. Without seeding the defaults across
    // used columns, our renderer falls back to 100 px per column and drifts
    // every column past A. Excel's implicit default when sheetFormatPr is
    // absent is 8.43 character units (the same value Excel writes when you
    // explicitly request "AutoFit Column Width" on an unsized column).
    let defaultColWidthPx =
        match sheet.Format.DefaultColWidth with
        // defaultColWidth attribute is stored using OOXML §18.3.1.13's
        // encoded form (same as <col width>) — pixels = stored*MDW. Verified
        // against pandas-dimension_large (defaultColWidth=11.55, MDW=7 → 81
        // px = 60.75 pt, matches Excel COM cell.Width).
        | Some w -> storedWidthToPixels w
        | None ->
            match sheet.Format.BaseColWidth with
            // OOXML §18.3.1.13 spec: when only baseColWidth is given,
            // defaultColWidth = baseColWidth + (cellPaddingPx + 1 grid) / MDW
            // with cellPadding = 2 px (Excel) for a total of 3 px. So the
            // effective char count is baseColWidth + 3/MDW. Combined with the
            // chars→pixels formula (chars*MDW + 5) this collapses to
            // baseColWidth*MDW + 8. Verified against pandas-test1
            // (baseColWidth=10, MDW=8 → 88 px = 66 pt, matches Excel COM
            // cell.Width = 66) and synthetic-print (implicit baseColWidth=8,
            // MDW=7 → 8.43 chars → 64 px = 48 pt).
            | Some baseW -> charsToPixelsDefault (baseW + 3.0 / mdw)
            | None -> charsToPixelsDefault 8.43

    let defaultRowHeightPx =
        sheet.Format.DefaultRowHeight |> Option.map (fun h -> h * (96.0 / 72.0))

    // Excel's font-derived auto row height. Derived from the font's
    // OS/2 typographic metrics:
    //
    //   pixels = ceil(typoAscender_px) + ceil(|typoDescender|_px)
    //          + ceil(typoLineGap_px) + 1
    //
    // The trick is that Excel ceils each metric SEPARATELY before summing,
    // not the total winLineHeight. Each metric snaps to its own integer-
    // pixel slot in the line box, then a single 1-px leading sits below.
    //
    // Empirical verification (pixels @ 96 dpi, all match Excel exactly):
    //   - Calibri 11: ceil(11.0)+ceil(3.67)+ceil(3.23)+1  = 11+4+4+1 = 20
    //   - Calibri 12: ceil(12.0)+ceil(4.0) +ceil(3.53)+1  = 12+4+4+1 = 21
    //   - Calibri 14: ceil(14.0)+ceil(4.67)+ceil(4.12)+1  = 14+5+5+1 = 25
    //   - Calibri 18: ceil(18.0)+ceil(6.0) +ceil(5.30)+1  = 18+6+6+1 = 31
    //   - Arial 11:   ceil(10.68)+ceil(3.09)+ceil(2.20)+1 = 11+4+3+1 = 19
    let autoRowHeightPt (familyName: string option) (fontPt: float) : float =
        // Fonts we don't ship a TTF for: hard-code the auto-height ratio.
        // Arial Black has heavier metrics than Arial Regular (typoAscender
        // ~2050 vs Arial's 1854). We don't have ariblk.ttf so we can't
        // derive from the font; the 1.7× ratio matches Excel's reported
        // 18.75 pt at 11 pt and 24 pt at 14 pt.
        let unshippedFontRatio =
            match familyName with
            | Some n when n.Trim().Equals("Arial Black", System.StringComparison.OrdinalIgnoreCase) ->
                Some 1.7
            | _ -> None
        match unshippedFontRatio with
        | Some ratio ->
            // Round to the 0.75-pt grid Excel snaps row heights to.
            let raw = fontPt * ratio
            System.Math.Round(raw * 4.0 / 3.0) * 3.0 / 4.0
        | None ->
        try
            let font = Rbt.OfficeNative.Fonts.resolve familyName false false
            let em = float font.OtFont.Head.UnitsPerEm
            let ptToPx = 96.0 / 72.0
            let unit2px (units: float) = units / em * fontPt * ptToPx
            match font.OtFont.OS2 with
            | Some os2 ->
                let asc  = unit2px (float os2.TypoAscender)
                let desc = unit2px (float -os2.TypoDescender)
                let lg   = unit2px (float os2.TypoLineGap)
                let pixels =
                    System.Math.Ceiling asc
                    + System.Math.Ceiling desc
                    + System.Math.Ceiling lg
                    + 1.0
                pixels * 72.0 / 96.0
            | None ->
                // Fonts without OS/2 fall back to hhea metrics.
                let asc  = unit2px (float font.OtFont.Hhea.Ascender)
                let desc = unit2px (float -font.OtFont.Hhea.Descender)
                let lg   = unit2px (float font.OtFont.Hhea.LineGap)
                let pixels =
                    System.Math.Ceiling asc
                    + System.Math.Ceiling desc
                    + System.Math.Ceiling lg
                    + 1.0
                pixels * 72.0 / 96.0
        with _ ->
            // Last-resort: ~Calibri 11 ratio rounded to 0.75-pt grid.
            System.Math.Round(fontPt * 1.36 * 4.0) / 4.0

    // <col min=M max=N width=W hidden=1/> applies W to every column in M..N
    // (inclusive). Hidden columns collapse to width=0 regardless of the
    // declared width — Excel hides them by zero-width, not by skipping the
    // entry, so a hidden col with width=11 still renders as 0 px. poi-51585
    // has 748 hidden cols in its template; without this they'd push the
    // real data way off-screen.
    let explicitColumns =
        sheet.Columns
        |> List.collect (fun col ->
            let widthPx =
                if col.Hidden then Some 0.0
                else col.Width |> Option.map storedWidthToPixels
            match widthPx with
            | None -> []
            | Some px -> [ for idx in col.Index .. col.MaxIndex -> idx - 1, px ])
        |> Map.ofList

    let explicitRows =
        // Hidden rows collapse to 0 px (matches Excel's screen + print
        // behavior). Otherwise honor an explicit height attribute, snapped
        // DOWN to the nearest pixel — Excel rounds stored `ht` values to
        // the 1-px (0.75-pt at 96 DPI) grid for display, and reports the
        // floored value as `cell.Height` / `cell.RowHeight`. Without the
        // floor pandas-df_empty's ht=19.4 rendered at 19.4 pt while Excel
        // showed 18.75 (= 25 px floor), pulling every subsequent row above
        // Excel's grid.
        sheet.Rows
        |> List.choose (fun row ->
            if row.Hidden then Some (row.Index - 1, 0.0)
            else row.Height |> Option.map (fun h ->
                let px = System.Math.Floor(h * (96.0 / 72.0))
                row.Index - 1, px))
        |> Map.ofList

    // Excel's grid is a fixed 1,048,576 × 16,384 (2^20 × 2^14) — not the
    // used range. The used range (maxRow/maxCol, the <dimension> element)
    // is just metadata; the displayed grid is always the full size, which
    // is why Excel paints gridlines across the whole viewport even past
    // the data. Layout.visibleRange virtualizes to the viewport, so a
    // full-size model costs nothing at paint time.
    let excelRowCount = 1048576
    let excelColCount = 16384

    // Sheet defaults cover unlisted axes; only explicit widths need entries.
    let columnWidths = explicitColumns

    // Rows fall into four height-resolution buckets:
    //   1. IMPLICIT  — no <row> XML entry. Use max(defaultRowHeight, fontAuto).
    //      pandas-dimension_large row 2 is the canonical case.
    //   2. EXPLICIT customHeight=true — landed in explicitRows above, the
    //      stored `ht` wins literally (the user manually shrank/grew the row,
    //      so even sub-font-height is respected).
    //   3. EXPLICIT `ht` present, customHeight absent OR false — Excel uses
    //      max(stored ht, fontAuto). The stored `ht` is a hint that gets
    //      raised to fit the font. phpsp-26template row 1 (`ht=27`, TNR 22
    //      → 27) and phpsp-27template row 2 (`ht=17`, Arial Black 11 →
    //      18.75) are the two sides of this rule.
    //   4. EXPLICIT customHeight=true but no ht (rare) — falls through to
    //      max(defaultRowHeight, fontAuto) like bucket 1.
    // Rows with a stored `ht` attribute but `customHeight` not explicitly
    // true. Excel raises these to fit the font (max(ht, fontAuto)) instead
    // of either honoring the stored value literally (the customHeight=true
    // case in `explicitRows`) or ignoring it entirely.
    let rawHeightOnly =
        sheet.Rows
        |> List.choose (fun row ->
            let isCustom = row.CustomHeight = Some true
            if row.Hidden || isCustom then None
            else row.RawHeight |> Option.map (fun h ->
                let px = System.Math.Floor (h * (96.0 / 72.0))
                row.Index - 1, px))
        |> Map.ofList
    let defaultNativeRowHeight =
        let fontHeight = autoRowHeightPt defaultFontName defaultFontSize * 96.0 / 72.0
        if sheet.Format.CustomDefaultHeight then
            defaultRowHeightPx |> Option.defaultValue fontHeight
        else
            match defaultRowHeightPx with
            | Some height -> max (System.Math.Floor height) fontHeight
            | None -> fontHeight


    let rowHeights =
        // When `<sheetFormatPr customHeight="1"/>` is set, Excel treats
        // `defaultRowHeight` as the explicit row height for every row
        // without its own `ht` attribute — the font-derived auto-height
        // does NOT override it. pandas-ints_spelled_with_decimals does
        // this to lock every row at 15.75 pt even though the cells use
        // Arial 10 (which would otherwise auto-height to 12.75).
        if sheet.Format.CustomDefaultHeight then
            explicitRows
        else
        // Build a per-row max-font map from cells we've collected. Any row
        // present in `explicitRows` (customHeight="true") keeps its stored
        // value; everything else gets Excel's font-derived auto-height,
        // which is the MAX of:
        //   - the workbook-default font's auto-height (acts as a floor —
        //     pandas-test1 has a Calibri 12 workbook default that floors
        //     every row at 15.75 pt even when individual cells are 11 pt)
        //   - the row's own max cell font auto-height (so a Title row with
        //     Calibri 18 pushes the row up to 23.25 pt).
        let maxFontByRow =
            let acc = System.Collections.Generic.Dictionary<int, string option * float>()
            let consider row name pt =
                match acc.TryGetValue row with
                | true, (_, cur) when cur >= pt -> ()
                | _ -> acc.[row] <- (name, pt)
            for KeyValue(addr, cd) in cells do
                let pt = cd.Format.FontSize |> Option.map float |> Option.defaultValue defaultFontSize
                let name = cd.Format.FontName |> Option.orElse defaultFontName
                consider addr.Row name pt
            // Rich-text runs can override the cell's base font size for
            // individual segments. Excel auto-grows the row to fit the
            // tallest run, not just the base font. synthetic-inline-rich's
            // A3 has a <sz val="14"/> run inside an otherwise Calibri 11
            // cell — Excel reports row 3 at 18.75 pt (Calibri 14 auto)
            // while we used to report 15 pt (Calibri 11 auto).
            for KeyValue(addr, runs) in richRunsMap do
                for (_, ovr) in runs do
                    match ovr.FontSize with
                    | Some pt ->
                        let name =
                            ovr.FontName
                            |> Option.orElse (
                                Map.tryFind addr cells
                                |> Option.bind (fun cd -> cd.Format.FontName)
                                |> Option.orElse defaultFontName)
                        consider addr.Row name pt
                    | None -> ()
            acc
        // Track rows that contain a superscript or subscript run. Excel
        // bumps row height to fit the raised/lowered run's ascent above
        // (or descent below) the normal baseline. Empirical bump from
        // synthetic-vertalign: super → +2.25 pt, sub → +1.5 pt for
        // Calibri 14. Scale linearly with the row's max-font pt size.
        let rowVertAlign =
            // 0 = none, 1 = sub only, 2 = super (overrides sub when both).
            // Earlier versions wrote `Some Superscript` / `Some Subscript`
            // which bound the names as fresh variables (FS0049) instead of
            // matching the cases — every run ended up as kind=2, so sub-only
            // rows got the larger super bump.
            let acc = System.Collections.Generic.Dictionary<int, int>()
            for KeyValue(addr, runs) in richRunsMap do
                for (_, ovr) in runs do
                    let kind =
                        match ovr.VertAlign with
                        | Some VertAlignSuperscript -> 2
                        | Some VertAlignSubscript -> 1
                        | _ -> 0
                    if kind > 0 then
                        let cur = match acc.TryGetValue addr.Row with true, v -> v | _ -> 0
                        acc.[addr.Row] <- max cur kind
            acc
        let vertAlignBumpPt (kind: int) (fontPt: float) =
            // Excel bumps the row to fit a super/subscript run's raised /
            // lowered glyph. Empirical bumps from synthetic-vertalign at
            // Calibri 14: super → +2.25 pt (21.0 = 18.75 base + 2.25),
            // sub → +1.5 pt (20.25 = 18.75 base + 1.5). Scale linearly
            // with the row's max-font pt size.
            match kind with
            | 2 -> fontPt * 2.25 / 14.0  // super
            | 1 -> fontPt * 1.5 / 14.0   // sub
            | _ -> 0.0
        // The workbook-default-font auto-height that floors every row
        // (per-row max font lands here when smaller).
        let defAutoPt = autoRowHeightPt defaultFontName defaultFontSize
        let defAutoPx = defAutoPt * 96.0 / 72.0
        // sheetFormatPr defaultRowHeight, floored to integer pixels.
        let defaultRowPx =
            sheet.Format.DefaultRowHeight
            |> Option.map (fun h -> System.Math.Floor (h * (96.0 / 72.0)))
        Seq.append maxFontByRow.Keys (rawHeightOnly |> Map.toSeq |> Seq.map fst)
        |> Seq.distinct
        |> Seq.fold (fun acc r ->
            if Map.containsKey r explicitRows then acc
            else
                // Take the larger of defaultRowHeight (floored to pixel
                // grid) and the row's font-derived auto-height. Then, if a
                // raw `ht` was stored without customHeight=true (bucket 3
                // above), raise to at least the stored value — Excel
                // honors any explicit `ht` as a floor unless the font is
                // even taller.
                let fontPx =
                    match maxFontByRow.TryGetValue r with
                    | true, (name, pt) ->
                        let basePt = autoRowHeightPt name pt
                        let vaKind = match rowVertAlign.TryGetValue r with true, v -> v | _ -> 0
                        let bump = vertAlignBumpPt vaKind pt
                        max defAutoPx ((basePt + bump) * 96.0 / 72.0)
                    | _ -> defAutoPx
                let baseAuto =
                    match defaultRowPx with
                    | Some d -> max d fontPx
                    | None -> fontPx
                let px =
                    match Map.tryFind r rawHeightOnly with
                    | Some stored -> max baseAuto stored
                    | None -> baseAuto
                if px = defaultNativeRowHeight then acc else Map.add r px acc) explicitRows

    // Auto-grow rows for wrapText cells. Excel auto-sizes the row to
    // accommodate wrapped text whenever a row has no explicit `ht`. We
    // measure the cell's text against its column width (real font metrics
    // via Rbt.Renderer.Text.measureWidth, with `mdw × len` as the fallback),
    // count wrap lines through Paint.wrapLines (which respects literal
    // `\n` and breaks on word boundaries), and bump rowHeights[r] when
    // the wrapped block exceeds the current height. Skipped when the
    // row carries an explicit height — Excel honors the user-set value
    // even if the text doesn't fit.
    let rowHeights =
        let defaultPxHeight = defaultNativeRowHeight
        // Cache one renderer-font per (familyName, bold, italic) tuple.
        // Without caching every wrap-cell would re-resolve the font, which
        // hits the Composit font cache repeatedly.
        let fontCache = System.Collections.Generic.Dictionary<string option * bool * bool, _>()
        let getFont (familyName: string option) (bold: bool) (italic: bool) =
            let key = familyName, bold, italic
            match fontCache.TryGetValue key with
            | true, f -> f
            | _ ->
                let f = Rbt.OfficeNative.Fonts.resolve familyName bold italic
                fontCache.[key] <- f
                f
        let measureCell (familyName: string option) (bold: bool) (italic: bool) (fontPx: float) (text: string) : float =
            try
                Rbt.Renderer.Text.measureWidth (getFont familyName bold italic) fontPx text
            with _ -> float text.Length * mdw
        let mutable result = rowHeights
        for KeyValue(addr, cd) in cells do
            let text =
                match cd.Value with
                | String s -> s
                | _ -> ""
            // Excel auto-grows the row when EITHER wrapText is set OR the
            // cell text has literal `\n` line breaks (Alt+Enter authored).
            // For the wrap case we count word-wrapped lines; for the
            // hard-break case we just count `\n`+1 lines (no word wrap).
            let colWidth =
                Map.tryFind addr.Col columnWidths
                |> Option.defaultValue defaultColWidthPx
            // Excel only honors wrapping / manual line breaks (and grows the row)
            // when wrapText is ON. A cell with an embedded `\n` but wrapText OFF
            // renders on a single line (Excel ignores the break), so it must NOT
            // grow the row. A hidden column (0 width) never grows either.
            let needsGrow = cd.Format.WrapText && text.Length > 0 && colWidth > 0.0
            let alreadyExplicit = Map.containsKey addr.Row explicitRows
            if needsGrow && not alreadyExplicit then
                let fontPt = cd.Format.FontSize |> Option.map float |> Option.defaultValue 11.0
                let fontPx = fontPt * 96.0 / 72.0
                let lineCount =
                    if cd.Format.WrapText then
                        let measure = fun (s: string) -> measureCell cd.Format.FontName cd.Format.Bold cd.Format.Italic fontPx s
                        let lines = Rbt.Spreadsheet.CellText.wrapLines (Some measure) colWidth fontPt text
                        lines.Length
                    else
                        // Hard-break only: split on \n without word-wrap.
                        text.Split('\n').Length
                // Excel's wrapped-line height matches the font's auto row
                // height (OS/2 typo-metric-derived). Calibri 11 → 15 pt =
                // 20 px / line, NOT fontPx*1.2 = 17.6 px which under-counted
                // synthetic-wrap-autoheight A2's 8-line block at 123 px vs
                // Excel's 160 px (= 8 * 20).
                let lineHeightPx = autoRowHeightPt cd.Format.FontName fontPt * 96.0 / 72.0
                let needed = float lineCount * lineHeightPx
                let current =
                    Map.tryFind addr.Row result
                    |> Option.defaultValue defaultPxHeight
                if needed > current then
                    result <- Map.add addr.Row needed result
        result

    // Preserve hidden rows introduced by the shared filter policy as well as
    // hidden axes from the source, without replacing the bridge's axis metadata.
    let rowHeights =
        model.RowHeights |> Map.fold (fun sizes row height ->
            if height = 0.0 then Map.add row 0.0 sizes else sizes) rowHeights
    let columnWidths =
        model.ColumnWidths |> Map.fold (fun sizes col width ->
            if width = 0.0 then Map.add col 0.0 sizes else sizes) columnWidths
    { model with
        RowCount = excelRowCount
        ColCount = excelColCount
        Selection = SingleCell(-1, -1)
        RowHeights = rowHeights
        ColumnWidths = columnWidths
        DefaultColumnWidth = defaultColWidthPx
        DefaultRowHeight = defaultNativeRowHeight }

/// Import every sheet together so formulas, cached values, names and tables share context.
let buildWorkbookContext (workbook: Xlsx.Workbook) =
    let imported = Rbt.Spreadsheet.XlsxBridge.Import.importWorkbook workbook
    let sheets =
        workbook.Sheets
        |> List.mapi (fun index sheet ->
            let id = sprintf "sheet%d" (index + 1)
            let current = Map.find id imported.Context.Sheets
            id, { current with Model = nativeGeometry sheet workbook current.Model })
        |> Map.ofList
    { imported.Context with Sheets = sheets }
