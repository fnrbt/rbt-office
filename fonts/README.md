# Default fonts

Free faces with the character widths of the Office fonts. The PDF export in the browser sets text in
them when the host has not supplied a face of its own (`SetPdfFonts` / `WordShell.setPdfFonts`), so a
PDF has all its text without any host setup. A face the host supplies is always used first. The build
copies this folder to `dist/fonts/`. The PDF modules fetch a file only when an export needs that face,
and the PDF embeds only the glyphs it uses. Which Office family maps to which face is
`src/Office.Formats/src/DefaultFonts.fs`.

| Files | Stands in for | Version | Licence |
|---|---|---|---|
| `Carlito-*.ttf` | Calibri | 20230309 (Debian `fonts-crosextra-carlito`) | SIL OFL 1.1, `OFL-Carlito.txt` |
| `Caladea-*.ttf` | Cambria | 20200211 (Debian `fonts-crosextra-caladea`) | SIL OFL 1.1, `OFL-Caladea.txt` |
| `LiberationSans-*.ttf` | Arial, Helvetica, and any family without a closer match | 2.1.5 (Debian `fonts-liberation`) | SIL OFL 1.1, `OFL-Liberation.txt` |
| `LiberationSerif-*.ttf` | Times New Roman and other serif families | 2.1.5 | SIL OFL 1.1, `OFL-Liberation.txt` |
| `LiberationMono-*.ttf` | Courier New, Consolas and other monospaced families | 2.1.5 | SIL OFL 1.1, `OFL-Liberation.txt` |

Each family has `-Regular`, `-Bold`, `-Italic` and `-BoldItalic`. The files are unmodified. The OFL
requires the licence to travel with them, which is why the `OFL-*.txt` files are here and in `dist/fonts/`.
