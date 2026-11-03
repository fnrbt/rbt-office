/// Table styles as Excel defines and draws them. A table style is a list of elements - the whole
/// table, the header and total rows, the first and last columns, row and column stripes, the first
/// and last cells of the header and total rows - each with a differential format (fill, font colour,
/// bold, italic, lines). The built-in styles (TableStyleLight1-21, TableStyleMedium1-28,
/// TableStyleDark1-11) are Excel's own definitions, the `presetTableStyles.xml` of ECMA-376 Part 1
/// (tools/excel-parity/make-table-style-presets.py writes the table below from it), in theme colours that
/// are resolved against the workbook's theme; a custom style comes from `<tableStyles>` of styles.xml.
/// `cellLook` lays the elements that cover a cell over each other in Excel's order, each element's
/// lines on the edges of the block it covers.
module Rbt.Spreadsheet.TableStyles

open System

/// A line of a table style: its kind ("thin", "medium", "thick", "double", ...) and colour (#RRGGBB).
type TableLine = { Style: string; Color: string }

/// What an element of a table style gives the cells it covers. `Top` .. `Right` are the edges of the
/// block the element covers, `Horizontal` and `Vertical` the lines between its cells.
type TableStyleFormat = {
    Fill: string option
    Text: string option
    Bold: bool option
    Italic: bool option
    Top: TableLine option
    Bottom: TableLine option
    Left: TableLine option
    Right: TableLine option
    Horizontal: TableLine option
    Vertical: TableLine option
}

/// An element of a table style: "wholeTable", "headerRow", "totalRow", "firstColumn", "lastColumn",
/// "firstRowStripe", "secondRowStripe", "firstColumnStripe", "secondColumnStripe", "firstHeaderCell",
/// "lastHeaderCell", "firstTotalCell", "lastTotalCell"; `Size` is the rows (columns) a stripe spans.
type TableStyleElement = { Kind: string; Size: int; Format: TableStyleFormat }

let emptyFormat : TableStyleFormat =
    { Fill = None; Text = None; Bold = None; Italic = None
      Top = None; Bottom = None; Left = None; Right = None; Horizontal = None; Vertical = None }

/// What a table's style gives one of its cells.
type TableCellLook = {
    Fill: string option
    Text: string option
    Bold: bool
    Italic: bool
    Top: TableLine option
    Bottom: TableLine option
    Left: TableLine option
    Right: TableLine option
}

/// The colours of the Office theme (2013-2022) in the order of `<color theme="N">`: lt1, dk1, lt2,
/// dk2, accent1-6, hlink, folHlink. The colours of a style made where no workbook theme is known.
let officeTheme : string array =
    [| "FFFFFF"; "000000"; "E7E6E6"; "44546A"; "4472C4"; "ED7D31"; "A5A5A5"; "FFC000"; "5B9BD5"; "70AD47"; "0563C1"; "954F72" |]

// Each style as `element[*size]=format|...`; a format is `;`-separated: `F<colour>` fill, `C<colour>`
// font colour, `B` bold, `I` italic, `<edge><kind>:<colour>` a line (edges t b l r h v; kinds 1 thin,
// 2 medium, 3 thick, d double); a colour is a theme index with an optional tint (`4,0.399976`).
let private presets : (string * string) array =
    [|
      "TableStyleLight1", "wholeTable=C1;t1:1;b1:1|headerRow=B;C1;b1:1|totalRow=B;C1;t1:1|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F0,-0.149998|firstColumnStripe=F0,-0.149998"
      "TableStyleLight2", "wholeTable=C4,-0.249977;t1:4;b1:4|headerRow=B;C4,-0.249977;b1:4|totalRow=B;C4,-0.249977;t1:4|firstColumn=B;C4,-0.249977|lastColumn=B;C4,-0.249977|firstRowStripe=F4,0.799982|firstColumnStripe=F4,0.799982"
      "TableStyleLight3", "wholeTable=C5,-0.249977;t1:5;b1:5|headerRow=B;C5,-0.249977;b1:5|totalRow=B;C5,-0.249977;t1:5|firstColumn=B;C5,-0.249977|lastColumn=B;C5,-0.249977|firstRowStripe=F5,0.799982|firstColumnStripe=F5,0.799982"
      "TableStyleLight4", "wholeTable=C6,-0.249977;t1:6;b1:6|headerRow=B;C6,-0.249977;b1:6|totalRow=B;C6,-0.249977;t1:6|firstColumn=B;C6,-0.249977|lastColumn=B;C6,-0.249977|firstRowStripe=F6,0.799982|firstColumnStripe=F6,0.799982"
      "TableStyleLight5", "wholeTable=C7,-0.249977;t1:7;b1:7|headerRow=B;C7,-0.249977;b1:7|totalRow=B;C7,-0.249977;t1:7|firstColumn=B;C7,-0.249977|lastColumn=B;C7,-0.249977|firstRowStripe=F7,0.799982|firstColumnStripe=F7,0.799982"
      "TableStyleLight6", "wholeTable=C8,-0.249977;t1:8;b1:8|headerRow=B;C8,-0.249977;b1:8|totalRow=B;C8,-0.249977;t1:8|firstColumn=B;C8,-0.249977|lastColumn=B;C8,-0.249977|firstRowStripe=F8,0.799982|firstColumnStripe=F8,0.799982"
      "TableStyleLight7", "wholeTable=C9,-0.249977;t1:9;b1:9|headerRow=B;C9,-0.249977;b1:9|totalRow=B;C9,-0.249977;t1:9|firstColumn=B;C9,-0.249977|lastColumn=B;C9,-0.249977|firstRowStripe=F9,0.799982|firstColumnStripe=F9,0.799982"
      "TableStyleLight8", "wholeTable=C1;l1:1;r1:1;t1:1;b1:1|headerRow=F1;B;C0|totalRow=B;C1;td:1|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=t1:1|secondRowStripe=t1:1|firstColumnStripe=l1:1|secondColumnStripe=l1:1"
      "TableStyleLight9", "wholeTable=C1;l1:4;r1:4;t1:4;b1:4|headerRow=F4;B;C0|totalRow=B;C1;td:4|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=t1:4|secondRowStripe=t1:4|firstColumnStripe=l1:4|secondColumnStripe=l1:4"
      "TableStyleLight10", "wholeTable=C1;l1:5;r1:5;t1:5;b1:5|headerRow=F5;B;C0|totalRow=B;C1;td:5|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=t1:5|secondRowStripe=t1:5|firstColumnStripe=l1:5|secondColumnStripe=l1:5"
      "TableStyleLight11", "wholeTable=C1;l1:6;r1:6;t1:6;b1:6|headerRow=F6;B;C0|totalRow=B;C1;td:6|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=t1:6|secondRowStripe=t1:6|firstColumnStripe=l1:6|secondColumnStripe=l1:6"
      "TableStyleLight12", "wholeTable=C1;l1:7;r1:7;t1:7;b1:7|headerRow=F7;B;C0|totalRow=B;C1;td:7|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=t1:7|secondRowStripe=t1:7|firstColumnStripe=l1:7|secondColumnStripe=l1:7"
      "TableStyleLight13", "wholeTable=C1;l1:8;r1:8;t1:8;b1:8|headerRow=F8;B;C0|totalRow=B;C1;td:8|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=t1:8|secondRowStripe=t1:8|firstColumnStripe=l1:8|secondColumnStripe=l1:8"
      "TableStyleLight14", "wholeTable=C1;l1:9;r1:9;t1:9;b1:9|headerRow=F9;B;C0|totalRow=B;C1;td:9|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=t1:9|secondRowStripe=t1:9|firstColumnStripe=l1:9|secondColumnStripe=l1:9"
      "TableStyleLight15", "wholeTable=C1;l1:1;r1:1;t1:1;b1:1;v1:1;h1:1|headerRow=B;C1;b2:1|totalRow=B;C1;td:1|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F0,-0.149998|firstColumnStripe=F0,-0.149998"
      "TableStyleLight16", "wholeTable=C1;l1:4;r1:4;t1:4;b1:4;v1:4;h1:4|headerRow=B;C1;b2:4|totalRow=B;C1;td:4|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F4,0.799982|firstColumnStripe=F4,0.799982"
      "TableStyleLight17", "wholeTable=C1;l1:5;r1:5;t1:5;b1:5;v1:5;h1:5|headerRow=B;C1;b2:5|totalRow=B;C1;td:5|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F5,0.799982|firstColumnStripe=F5,0.799982"
      "TableStyleLight18", "wholeTable=C1;l1:6;r1:6;t1:6;b1:6;v1:6;h1:6|headerRow=B;C1;b2:6|totalRow=B;C1;td:6|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F6,0.799982|firstColumnStripe=F6,0.799982"
      "TableStyleLight19", "wholeTable=C1;l1:7;r1:7;t1:7;b1:7;v1:7;h1:7|headerRow=B;C1;b2:7|totalRow=B;C1;td:7|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F7,0.799982|firstColumnStripe=F7,0.799982"
      "TableStyleLight20", "wholeTable=C1;l1:8;r1:8;t1:8;b1:8;v1:8;h1:8|headerRow=B;C1;b2:8|totalRow=B;C1;td:8|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F8,0.799982|firstColumnStripe=F8,0.799982"
      "TableStyleLight21", "wholeTable=C1;l1:9;r1:9;t1:9;b1:9;v1:9;h1:9|headerRow=B;C1;b2:9|totalRow=B;C1;td:9|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F9,0.799982|firstColumnStripe=F9,0.799982"
      "TableStyleMedium1", "wholeTable=C1;l1:1;r1:1;t1:1;b1:1;h1:1|headerRow=F1;B;C0|totalRow=B;C1;td:1|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F0,-0.149998|firstColumnStripe=F0,-0.149998"
      "TableStyleMedium2", "wholeTable=C1;l1:4,0.399976;r1:4,0.399976;t1:4,0.399976;b1:4,0.399976;h1:4,0.399976|headerRow=F4;B;C0|totalRow=B;C1;td:4|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F4,0.799982|firstColumnStripe=F4,0.799982"
      "TableStyleMedium3", "wholeTable=C1;l1:5,0.399976;r1:5,0.399976;t1:5,0.399976;b1:5,0.399976;h1:5,0.399976|headerRow=F5;B;C0|totalRow=B;C1;td:5|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F5,0.799982|firstColumnStripe=F5,0.799982"
      "TableStyleMedium4", "wholeTable=C1;l1:6,0.399976;r1:6,0.399976;t1:6,0.399976;b1:6,0.399976;h1:6,0.399976|headerRow=F6;B;C0|totalRow=B;C1;td:6|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F6,0.799982|firstColumnStripe=F6,0.799982"
      "TableStyleMedium5", "wholeTable=C1;l1:7,0.399976;r1:7,0.399976;t1:7,0.399976;b1:7,0.399976;h1:7,0.399976|headerRow=F7;B;C0|totalRow=B;C1;td:7|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F7,0.799982|firstColumnStripe=F7,0.799982"
      "TableStyleMedium6", "wholeTable=C1;l1:8,0.399976;r1:8,0.399976;t1:8,0.399976;b1:8,0.399976;h1:8,0.399976|headerRow=F8;B;C0|totalRow=B;C1;td:8|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F8,0.799982|firstColumnStripe=F8,0.799982"
      "TableStyleMedium7", "wholeTable=C1;l1:9,0.399976;r1:9,0.399976;t1:9,0.399976;b1:9,0.399976;h1:9,0.399976|headerRow=F9;B;C0|totalRow=B;C1;td:9|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F9,0.799982|firstColumnStripe=F9,0.799982"
      "TableStyleMedium8", "wholeTable=F0,-0.149998;C1;v1:0;h1:0|headerRow=F1;B;C0;b3:0|totalRow=F1;B;C0;t3:0|firstColumn=F1;B;C0|lastColumn=F1;B;C0|firstRowStripe=F0,-0.349986|firstColumnStripe=F0,-0.349986"
      "TableStyleMedium9", "wholeTable=F4,0.799982;C1;v1:0;h1:0|headerRow=F4;B;C0;b3:0|totalRow=F4;B;C0;t3:0|firstColumn=F4;B;C0|lastColumn=F4;B;C0|firstRowStripe=F4,0.599994|firstColumnStripe=F4,0.599994"
      "TableStyleMedium10", "wholeTable=F5,0.799982;C1;v1:0;h1:0|headerRow=F5;B;C0;b3:0|totalRow=F5;B;C0;t3:0|firstColumn=F5;B;C0|lastColumn=F5;B;C0|firstRowStripe=F5,0.599994|firstColumnStripe=F5,0.599994"
      "TableStyleMedium11", "wholeTable=F6,0.799982;C1;v1:0;h1:0|headerRow=F6;B;C0;b3:0|totalRow=F6;B;C0;t3:0|firstColumn=F6;B;C0|lastColumn=F6;B;C0|firstRowStripe=F6,0.599994|firstColumnStripe=F6,0.599994"
      "TableStyleMedium12", "wholeTable=F7,0.799982;C1;v1:0;h1:0|headerRow=F7;B;C0;b3:0|totalRow=F7;B;C0;t3:0|firstColumn=F7;B;C0|lastColumn=F7;B;C0|firstRowStripe=F7,0.599994|firstColumnStripe=F7,0.599994"
      "TableStyleMedium13", "wholeTable=F8,0.799982;C1;v1:0;h1:0|headerRow=F8;B;C0;b3:0|totalRow=F8;B;C0;t3:0|firstColumn=F8;B;C0|lastColumn=F8;B;C0|firstRowStripe=F8,0.599994|firstColumnStripe=F8,0.599994"
      "TableStyleMedium14", "wholeTable=F9,0.799982;C1;v1:0;h1:0|headerRow=F9;B;C0;b3:0|totalRow=F9;B;C0;t3:0|firstColumn=F9;B;C0|lastColumn=F9;B;C0|firstRowStripe=F9,0.599994|firstColumnStripe=F9,0.599994"
      "TableStyleMedium15", "wholeTable=C1;l1:1;r1:1;t2:1;b2:1;v1:1;h1:1|headerRow=F1;B;C0;b2:1|totalRow=td:1|firstColumn=F1;B;C0|lastColumn=F1;B;C0|firstRowStripe=F0,-0.149998|firstColumnStripe=F0,-0.149998"
      "TableStyleMedium16", "wholeTable=C1;t2:1;b2:1|headerRow=F4;B;C0;b2:1|totalRow=td:1|firstColumn=F4;B;C0|lastColumn=F4;B;C0|firstRowStripe=F0,-0.149998|firstColumnStripe=F0,-0.149998"
      "TableStyleMedium17", "wholeTable=C1;t2:1;b2:1|headerRow=F5;B;C0;b2:1|totalRow=td:1|firstColumn=F5;B;C0|lastColumn=F5;B;C0|firstRowStripe=F0,-0.149998|firstColumnStripe=F0,-0.149998"
      "TableStyleMedium18", "wholeTable=C1;t2:1;b2:1|headerRow=F6;B;C0;b2:1|totalRow=td:1|firstColumn=F6;B;C0|lastColumn=F6;B;C0|firstRowStripe=F0,-0.149998|firstColumnStripe=F0,-0.149998"
      "TableStyleMedium19", "wholeTable=C1;t2:1;b2:1|headerRow=F7;B;C0;b2:1|totalRow=td:1|firstColumn=F7;B;C0|lastColumn=F7;B;C0|firstRowStripe=F0,-0.149998|firstColumnStripe=F0,-0.149998"
      "TableStyleMedium20", "wholeTable=C1;t2:1;b2:1|headerRow=F8;B;C0;b2:1|totalRow=td:1|firstColumn=F8;B;C0|lastColumn=F8;B;C0|firstRowStripe=F0,-0.149998|firstColumnStripe=F0,-0.149998"
      "TableStyleMedium21", "wholeTable=C1;t2:1;b2:1|headerRow=F9;B;C0;b2:1|totalRow=td:1|firstColumn=F9;B;C0|lastColumn=F9;B;C0|firstRowStripe=F0,-0.149998|firstColumnStripe=F0,-0.149998"
      "TableStyleMedium22", "wholeTable=F0,-0.149998;C1;l1:1;r1:1;t1:1;b1:1;v1:1;h1:1|headerRow=B;C1|totalRow=B;C1;t2:1|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F0,-0.349986|firstColumnStripe=F0,-0.349986"
      "TableStyleMedium23", "wholeTable=F4,0.799982;C1;l1:4,0.399976;r1:4,0.399976;t1:4,0.399976;b1:4,0.399976;v1:4,0.399976;h1:4,0.399976|headerRow=B;C1|totalRow=B;C1;t2:4|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F4,0.599994|firstColumnStripe=F4,0.599994"
      "TableStyleMedium24", "wholeTable=F5,0.799982;C1;l1:5,0.399976;r1:5,0.399976;t1:5,0.399976;b1:5,0.399976;v1:5,0.399976;h1:5,0.399976|headerRow=B;C1|totalRow=B;C1;t2:5|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F5,0.599994|firstColumnStripe=F5,0.599994"
      "TableStyleMedium25", "wholeTable=F6,0.799982;C1;l1:6,0.399976;r1:6,0.399976;t1:6,0.399976;b1:6,0.399976;v1:6,0.399976;h1:6,0.399976|headerRow=B;C1|totalRow=B;C1;t2:6|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F6,0.599994|firstColumnStripe=F6,0.599994"
      "TableStyleMedium26", "wholeTable=F7,0.799982;C1;l1:7,0.399976;r1:7,0.399976;t1:7,0.399976;b1:7,0.399976;v1:7,0.399976;h1:7,0.399976|headerRow=B;C1|totalRow=B;C1;t2:7|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F7,0.599994|firstColumnStripe=F7,0.599994"
      "TableStyleMedium27", "wholeTable=F8,0.799982;C1;l1:8,0.399976;r1:8,0.399976;t1:8,0.399976;b1:8,0.399976;v1:8,0.399976;h1:8,0.399976|headerRow=B;C1|totalRow=B;C1;t2:8|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F8,0.599994|firstColumnStripe=F8,0.599994"
      "TableStyleMedium28", "wholeTable=F9,0.799982;C1;l1:9,0.399976;r1:9,0.399976;t1:9,0.399976;b1:9,0.399976;v1:9,0.399976;h1:9,0.399976|headerRow=B;C1|totalRow=B;C1;t2:9|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F9,0.599994|firstColumnStripe=F9,0.599994"
      "TableStyleDark1", "wholeTable=F1,0.449995;C0|headerRow=F1;B;C0;b2:0|totalRow=F1,0.149998;B;C0;t2:0|firstColumn=F1,0.249977;B;C0;r2:0|lastColumn=F1,0.249977;B;C0;l2:0|firstRowStripe=F1,0.249977|firstColumnStripe=F1,0.249977"
      "TableStyleDark2", "wholeTable=F4;C0|headerRow=F1;B;C0;b2:0|totalRow=F4,-0.499985;B;C0;t2:0|firstColumn=F4,-0.249977;B;C0;r2:0|lastColumn=F4,-0.249977;B;C0;l2:0|firstRowStripe=F4,-0.249977|firstColumnStripe=F4,-0.249977"
      "TableStyleDark3", "wholeTable=F5;C0|headerRow=F1;B;C0;b2:0|totalRow=F5,-0.499985;B;C0;t2:0|firstColumn=F5,-0.249977;B;C0;r2:0|lastColumn=F5,-0.249977;B;C0;l2:0|firstRowStripe=F5,-0.249977|firstColumnStripe=F5,-0.249977"
      "TableStyleDark4", "wholeTable=F6;C0|headerRow=F1;B;C0;b2:0|totalRow=F6,-0.499985;B;C0;t2:0|firstColumn=F6,-0.249977;B;C0;r2:0|lastColumn=F6,-0.249977;B;C0;l2:0|firstRowStripe=F6,-0.249977|firstColumnStripe=F6,-0.249977"
      "TableStyleDark5", "wholeTable=F7;C0|headerRow=F1;B;C0;b2:0|totalRow=F7,-0.499985;B;C0;t2:0|firstColumn=F7,-0.249977;B;C0;r2:0|lastColumn=F7,-0.249977;B;C0;l2:0|firstRowStripe=F7,-0.249977|firstColumnStripe=F7,-0.249977"
      "TableStyleDark6", "wholeTable=F8;C0|headerRow=F1;B;C0;b2:0|totalRow=F8,-0.499985;B;C0;t2:0|firstColumn=F8,-0.249977;B;C0;r2:0|lastColumn=F8,-0.249977;B;C0;l2:0|firstRowStripe=F8,-0.249977|firstColumnStripe=F8,-0.249977"
      "TableStyleDark7", "wholeTable=F9;C0|headerRow=F1;B;C0;b2:0|totalRow=F9,-0.499985;B;C0;t2:0|firstColumn=F9,-0.249977;B;C0;r2:0|lastColumn=F9,-0.249977;B;C0;l2:0|firstRowStripe=F9,-0.249977|firstColumnStripe=F9,-0.249977"
      "TableStyleDark8", "wholeTable=F0,-0.149998|headerRow=F1;C0|totalRow=B;C1;td:1|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F0,-0.349986|firstColumnStripe=F0,-0.349986"
      "TableStyleDark9", "wholeTable=F4,0.799982|headerRow=F5;C0|totalRow=B;C1;td:1|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F4,0.599994|firstColumnStripe=F4,0.599994"
      "TableStyleDark10", "wholeTable=F6,0.799982|headerRow=F7;C0|totalRow=B;C1;td:1|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F6,0.599994|firstColumnStripe=F6,0.599994"
      "TableStyleDark11", "wholeTable=F8,0.799982|headerRow=F9;C0|totalRow=B;C1;td:1|firstColumn=B;C1|lastColumn=B;C1|firstRowStripe=F8,0.599994|firstColumnStripe=F8,0.599994"
    |]

/// A colour of a style as `#RRGGBB`: `#RRGGBB` as it is, or a theme index of `theme` with a tint.
let private colorOf (theme: string array) (text: string) : string option =
    if text.StartsWith "#" then Some (text.ToUpperInvariant())
    else
        let parts = text.Split ','
        match Int32.TryParse(parts.[0], Globalization.NumberStyles.Integer, Globalization.CultureInfo.InvariantCulture) with
        | true, index when index >= 0 && index < theme.Length ->
            let tint =
                if parts.Length > 1 then
                    match Double.TryParse(parts.[1], Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
                    | true, value -> value
                    | _ -> 0.0
                else 0.0
            let hex = XlsxReader.Xlsx.applyColorTint theme.[index] tint
            Some ("#" + hex.TrimStart('#').ToUpperInvariant())
        | _ -> None

let private lineKind (code: char) =
    match code with
    | '1' -> "thin" | '2' -> "medium" | '3' -> "thick" | 'd' -> "double" | 'h' -> "hair" | 'o' -> "dotted" | 'a' -> "dashed"
    | _ -> "thin"

let private formatOf (theme: string array) (text: string) : TableStyleFormat =
    text.Split([| ';' |], StringSplitOptions.RemoveEmptyEntries)
    |> Array.fold (fun (format: TableStyleFormat) (token: string) ->
        match token.[0] with
        | 'F' -> { format with Fill = colorOf theme (token.Substring 1) }
        | 'C' -> { format with Text = colorOf theme (token.Substring 1) }
        | 'B' -> { format with Bold = Some true }
        | 'I' -> { format with Italic = Some true }
        | edge when token.Length > 3 && token.[2] = ':' ->
            let line = colorOf theme (token.Substring 3) |> Option.map (fun color -> { Style = lineKind token.[1]; Color = color })
            match edge with
            | 't' -> { format with Top = line }
            | 'b' -> { format with Bottom = line }
            | 'l' -> { format with Left = line }
            | 'r' -> { format with Right = line }
            | 'h' -> { format with Horizontal = line }
            | 'v' -> { format with Vertical = line }
            | _ -> format
        | _ -> format) emptyFormat

/// Whether Excel has a built-in table style of this name ("TableStyleMedium2").
let isBuiltIn (name: string) : bool = presets |> Array.exists (fun (known, _) -> known = name)

/// The elements of a built-in table style in the colours of `theme` (12 colours `RRGGBB` in the
/// order of `<color theme="N">`); None for a name that is not one.
let builtIn (theme: string array) (name: string) : TableStyleElement list option =
    presets
    |> Array.tryFind (fun (known, _) -> known = name)
    |> Option.map (fun (_, text) ->
        text.Split '|'
        |> Array.toList
        |> List.map (fun element ->
            let cut = element.IndexOf '='
            let head, body = element.Substring(0, cut), element.Substring(cut + 1)
            let kind, size =
                match head.IndexOf '*' with
                | -1 -> head, 1
                | star -> head.Substring(0, star), int (head.Substring(star + 1))
            { Kind = kind; Size = size; Format = formatOf theme body }))

/// The order Excel lays the elements of a style over each other in: a later one wins.
let private order =
    [| "wholeTable"; "firstColumnStripe"; "secondColumnStripe"; "firstRowStripe"; "secondRowStripe"; "lastColumn"; "firstColumn"
       "headerRow"; "totalRow"; "firstHeaderCell"; "lastHeaderCell"; "firstTotalCell"; "lastTotalCell" |]

/// What the elements of a style give the 0-based cell (row, col) of a table over rows r0..r1 and
/// columns c0..c1 with `headerRows` header rows and `totalRows` total rows, the style's options as
/// the table turns them on. None outside the table.
let cellLook (elements: TableStyleElement list) (r0: int, c0: int, r1: int, c1: int) (headerRows: int) (totalRows: int)
             (showFirstColumn: bool, showLastColumn: bool, showRowStripes: bool, showColumnStripes: bool)
             (row: int) (col: int) : TableCellLook option =
    if row < r0 || row > r1 || col < c0 || col > c1 || elements.IsEmpty then None
    else
        let headerEnd = r0 + headerRows - 1
        let totalStart = r1 - totalRows + 1
        let isHeader = headerRows > 0 && row <= headerEnd
        let isTotal = totalRows > 0 && row >= totalStart
        let dataTop, dataBottom = r0 + headerRows, r1 - totalRows
        let sizeOf (kind: string) =
            elements |> List.tryFind (fun element -> element.Kind = kind) |> Option.map (fun element -> max 1 element.Size) |> Option.defaultValue 1
        /// The stripe a position of the body is in (first or second) and the block of that stripe.
        let stripe (first: string) (second: string) (start: int) (finish: int) (at: int) =
            let a, b = sizeOf first, sizeOf second
            let offset = (at - start) % (a + b)
            let blockStart = at - offset
            if offset < a then first, blockStart, min finish (blockStart + a - 1)
            else second, blockStart + a, min finish (blockStart + a + b - 1)
        // The block (top, left, bottom, right) each element covers at this cell, if it covers it.
        let covers (kind: string) : (int * int * int * int) option =
            match kind with
            | "wholeTable" -> Some (r0, c0, r1, c1)
            | "headerRow" when isHeader -> Some (r0, c0, headerEnd, c1)
            | "totalRow" when isTotal -> Some (totalStart, c0, r1, c1)
            | "firstColumn" when showFirstColumn && col = c0 -> Some (r0, c0, r1, c0)
            | "lastColumn" when showLastColumn && col = c1 -> Some (r0, c1, r1, c1)
            | ("firstRowStripe" | "secondRowStripe") when showRowStripes && not isHeader && not isTotal ->
                let which, top, bottom = stripe "firstRowStripe" "secondRowStripe" dataTop dataBottom row
                if which = kind then Some (top, c0, bottom, c1) else None
            | ("firstColumnStripe" | "secondColumnStripe") when showColumnStripes && not isHeader && not isTotal ->
                let which, left, right = stripe "firstColumnStripe" "secondColumnStripe" c0 c1 col
                if which = kind then Some (dataTop, left, dataBottom, right) else None
            | "firstHeaderCell" when isHeader && showFirstColumn && col = c0 -> Some (r0, c0, headerEnd, c0)
            | "lastHeaderCell" when isHeader && showLastColumn && col = c1 -> Some (r0, c1, headerEnd, c1)
            | "firstTotalCell" when isTotal && showFirstColumn && col = c0 -> Some (totalStart, c0, r1, c0)
            | "lastTotalCell" when isTotal && showLastColumn && col = c1 -> Some (totalStart, c1, r1, c1)
            | _ -> None
        let mutable look = { Fill = None; Text = None; Bold = false; Italic = false; Top = None; Bottom = None; Left = None; Right = None }
        for kind in order do
            for element in elements do
                if element.Kind = kind then
                    match covers kind with
                    | Some (top, left, bottom, right) ->
                        let format = element.Format
                        let pick (outer: TableLine option) (inner: TableLine option) (atEdge: bool) (current: TableLine option) =
                            (if atEdge then outer else inner) |> Option.orElse current
                        look <-
                            { Fill = format.Fill |> Option.orElse look.Fill
                              Text = format.Text |> Option.orElse look.Text
                              Bold = defaultArg format.Bold look.Bold
                              Italic = defaultArg format.Italic look.Italic
                              Top = pick format.Top format.Horizontal (row = top) look.Top
                              Bottom = pick format.Bottom format.Horizontal (row = bottom) look.Bottom
                              Left = pick format.Left format.Vertical (col = left) look.Left
                              Right = pick format.Right format.Vertical (col = right) look.Right }
                    | None -> ()
        Some look
