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

// The built-in pivot styles (PivotStyleLight1-28, PivotStyleMedium1-28, PivotStyleDark1-28) in the same form, with
// the elements only a pivot table has: subtotal rows and columns, row and column subheadings, blank rows, page fields
// (`make-table-style-presets.py presetTableStyles.xml --pivot`).
let private pivotPresets : (string * string) array =
    [|
      "PivotStyleLight1", "wholeTable=C1;h1:0,-0.149998|headerRow=B;C1;t1:1,0.499985;b1:1,0.499985|totalRow=F0;B;C1;t1:1,0.499985;b1:1,0.499985|firstRowStripe=F0,-0.149998;t1:0,-0.349986;b1:0,-0.349986|firstColumnStripe=F0,-0.149998;l1:0,-0.349986;r1:0,-0.349986;t1:0,-0.349986;b1:0,-0.349986;v1:0,-0.349986;h1:0,-0.349986|firstSubtotalRow=B;C1|secondSubtotalRow=B;C1,0.499985|firstRowSubheading=B;C1|secondRowSubheading=B;C1,0.499985|pageFieldLabels=t1:1,0.499985;b1:1,0.499985|pageFieldValues=t1:1,0.499985;b1:1,0.499985"
      "PivotStyleLight2", "wholeTable=C1;h1:4,0.799982|headerRow=B;C1;t1:4;b1:4|totalRow=F0;B;C1;t1:4;b1:4|firstRowStripe=F4,0.799982;t1:4,0.599994;b1:4,0.599994|firstColumnStripe=F4,0.799982;l1:4,0.599994;r1:4,0.599994;t1:4,0.599994;b1:4,0.599994;v1:4,0.599994;h1:4,0.599994|firstSubtotalRow=B;C1|secondSubtotalRow=C4|firstRowSubheading=B;C1|secondRowSubheading=C4|pageFieldLabels=t1:4;b1:4|pageFieldValues=t1:4;b1:4"
      "PivotStyleLight3", "wholeTable=C1;h1:5,0.799982|headerRow=B;C1;t1:5;b1:5|totalRow=F0;B;C1;t1:5;b1:5|firstRowStripe=F5,0.799982;t1:5,0.599994;b1:5,0.599994|firstColumnStripe=F5,0.799982;l1:5,0.599994;r1:5,0.599994;t1:5,0.599994;b1:5,0.599994;v1:5,0.599994;h1:5,0.599994|firstSubtotalRow=B;C1|secondSubtotalRow=B;C5|firstRowSubheading=B;C1|secondRowSubheading=B;C5|pageFieldLabels=t1:5;b1:5|pageFieldValues=t1:5;b1:5"
      "PivotStyleLight4", "wholeTable=C1;h1:6,0.799982|headerRow=B;C1;t1:6;b1:6|totalRow=F0;B;C1;t1:6;b1:6|firstRowStripe=F6,0.799982;t1:6,0.599994;b1:6,0.599994|firstColumnStripe=F6,0.799982;l1:6,0.599994;r1:6,0.599994;t1:6,0.599994;b1:6,0.599994;v1:6,0.599994;h1:6,0.599994|firstSubtotalRow=B;C1|secondSubtotalRow=B;C6|firstRowSubheading=B;C1|secondRowSubheading=B;C6|pageFieldLabels=t1:6;b1:6|pageFieldValues=t1:6;b1:6"
      "PivotStyleLight5", "wholeTable=C1;h1:7,0.799982|headerRow=B;C1;t1:7;b1:7|totalRow=F0;B;C1;t1:7;b1:7|firstRowStripe=F7,0.799982;t1:7,0.599994;b1:7,0.599994|firstColumnStripe=F7,0.799982;l1:7,0.599994;r1:7,0.599994;t1:7,0.599994;b1:7,0.599994;v1:7,0.599994;h1:7,0.599994|firstSubtotalRow=B;C1|secondSubtotalRow=B;C7|firstRowSubheading=B;C1|secondRowSubheading=B;C7|pageFieldLabels=t1:7;b1:7|pageFieldValues=t1:7;b1:7"
      "PivotStyleLight6", "wholeTable=C1;h1:8,0.799982|headerRow=B;C1;t1:8;b1:8|totalRow=F0;B;C1;t1:8;b1:8|firstRowStripe=F8,0.799982;t1:8,0.599994;b1:8,0.599994|firstColumnStripe=F8,0.799982;l1:8,0.599994;r1:8,0.599994;t1:8,0.599994;b1:8,0.599994;v1:8,0.599994;h1:8,0.599994|firstSubtotalRow=B;C1|secondSubtotalRow=B;C8|firstRowSubheading=B;C1|secondRowSubheading=B;C8|pageFieldLabels=t1:8;b1:8|pageFieldValues=t1:8;b1:8"
      "PivotStyleLight7", "wholeTable=C1;h1:9,0.799982|headerRow=B;C1;t1:9;b1:9|totalRow=F0;B;C1;t1:9;b1:9|firstRowStripe=F9,0.799982;t1:9,0.599994;b1:9,0.599994|firstColumnStripe=F9,0.799982;l1:9,0.599994;r1:9,0.599994;t1:9,0.599994;b1:9,0.599994;v1:9,0.599994;h1:9,0.599994|firstSubtotalRow=B;C1|secondSubtotalRow=B;C9|firstRowSubheading=B;C1|secondRowSubheading=B;C9|pageFieldLabels=t1:9;b1:9|pageFieldValues=t1:9;b1:9"
      "PivotStyleLight8", "wholeTable=C1;h1:0,-0.149998|headerRow=B;C1;l2:1,0.499985;r2:1,0.499985;t2:1,0.499985;b2:1,0.499985;h1:0|totalRow=B;C1;l2:1,0.499985;r2:1,0.499985;t2:1,0.499985;b2:1,0.499985|firstColumn=r1:1,0.499985|firstRowStripe=l1:0,-0.349986;r1:0,-0.349986;t1:0,-0.349986;b1:0,-0.349986|firstColumnStripe=l1:0,-0.449995;r1:0,-0.449995;t1:0,-0.449995;b1:0,-0.449995|firstSubtotalRow=F0,-0.149998;B;C1;t1:0,-0.349986;b1:0,-0.349986|secondSubtotalRow=B;C1|secondColumnSubheading=t1:0,-0.349986|firstRowSubheading=F0,-0.149998;B;C1;b1:0|secondRowSubheading=B;C1"
      "PivotStyleLight9", "wholeTable=C4,-0.249977;h1:4,0.799982|headerRow=B;C1;l2:4;r2:4;t2:4;b2:4;h1:0|totalRow=B;C1;l2:4;r2:4;t2:4;b2:4|firstColumn=r1:4|firstRowStripe=l1:4,0.599994;r1:4,0.599994;t1:4,0.599994;b1:4,0.599994|firstColumnStripe=l1:4;r1:4;t1:4;b1:4|firstSubtotalRow=F4,0.799982;B;C1;t1:4,0.599994;b1:4,0.599994|secondSubtotalRow=B;C1|secondColumnSubheading=t1:4,0.799982|thirdColumnSubheading=t1:4,0.799982|firstRowSubheading=F4,0.799982;B;C1;b1:0|secondRowSubheading=B;C1"
      "PivotStyleLight10", "wholeTable=C5,-0.249977;h1:5,0.799982|headerRow=B;C1;l2:5;r2:5;t2:5;b2:5;h1:0|totalRow=B;C1;l2:5;r2:5;t2:5;b2:5|firstColumn=r1:5|firstRowStripe=l1:5,0.599994;r1:5,0.599994;t1:5,0.599994;b1:5,0.599994|firstColumnStripe=l1:5;r1:5;t1:5;b1:5|firstSubtotalRow=F5,0.799982;B;C1;t1:5,0.599994;b1:5,0.599994|secondSubtotalRow=B;C1|secondColumnSubheading=t1:5,0.799982|thirdColumnSubheading=t1:5,0.799982|firstRowSubheading=F5,0.799982;B;C1;b1:0|secondRowSubheading=B;C1"
      "PivotStyleLight11", "wholeTable=C6,-0.249977;h1:6,0.799982|headerRow=B;C1;l2:6;r2:6;t2:6;b2:6;h1:0|totalRow=B;C1;l2:6;r2:6;t2:6;b2:6|firstColumn=r1:6|firstRowStripe=l1:6,0.599994;r1:6,0.599994;t1:6,0.599994;b1:6,0.599994|firstColumnStripe=l1:6;r1:6;t1:6;b1:6|firstSubtotalRow=F6,0.799982;B;C1;t1:6,0.599994;b1:6,0.599994|secondSubtotalRow=B;C1|secondColumnSubheading=t1:6,0.799982|thirdColumnSubheading=t1:6,0.799982|firstRowSubheading=F6,0.799982;B;C1;b1:0|secondRowSubheading=B;C1"
      "PivotStyleLight12", "wholeTable=C7,-0.249977;h1:7,0.799982|headerRow=B;C1;l2:7;r2:7;t2:7;b2:7;h1:0|totalRow=B;C1;l2:7;r2:7;t2:7;b2:7|firstColumn=r1:7|firstRowStripe=l1:7,0.599994;r1:7,0.599994;t1:7,0.599994;b1:7,0.599994|firstColumnStripe=l1:7;r1:7;t1:7;b1:7|firstSubtotalRow=F7,0.799982;B;C1;t1:7,0.599994;b1:7,0.599994|secondSubtotalRow=B;C1|secondColumnSubheading=t1:7,0.799982|thirdColumnSubheading=t1:7,0.799982|firstRowSubheading=F7,0.799982;B;C1;b1:0|secondRowSubheading=B;C1"
      "PivotStyleLight13", "wholeTable=C8,-0.249977;h1:8,0.799982|headerRow=B;C1;l2:8;r2:8;t2:8;b2:8;h1:0|totalRow=B;C1;l2:8;r2:8;t2:8;b2:8|firstColumn=r1:8|firstRowStripe=l1:8,0.599994;r1:8,0.599994;t1:8,0.599994;b1:8,0.599994|firstColumnStripe=l1:8;r1:8;t1:8;b1:8|firstSubtotalRow=F8,0.799982;B;C1;t1:8,0.599994;b1:8,0.599994|secondSubtotalRow=B;C1|secondColumnSubheading=t1:8,0.799982|thirdColumnSubheading=t1:8,0.799982|firstRowSubheading=F8,0.799982;B;C1;b1:0|secondRowSubheading=B;C1"
      "PivotStyleLight14", "wholeTable=C9,-0.249977;h1:9,0.799982|headerRow=B;C1;l2:9;r2:9;t2:9;b2:9;h1:0|totalRow=B;C1;l2:9;r2:9;t2:9;b2:9|firstColumn=r1:9|firstRowStripe=l1:9,0.599994;r1:9,0.599994;t1:9,0.599994;b1:9,0.599994|firstColumnStripe=l1:9;r1:9;t1:9;b1:9|firstSubtotalRow=F9,0.799982;B;C1;t1:9,0.599994;b1:9,0.599994|secondSubtotalRow=B;C1|secondColumnSubheading=t1:9,0.799982|thirdColumnSubheading=t1:9,0.799982|firstRowSubheading=F9,0.799982;B;C1;b1:0|secondRowSubheading=B;C1"
      "PivotStyleLight15", "headerRow=F0,-0.149998;B;C1;b1:0,-0.349986|totalRow=F0,-0.149998;B;C1;t1:0,-0.349986|firstRowStripe=F0,-0.149998|firstColumnStripe=F0,-0.149998;l1:0,-0.249977;r1:0,-0.249977|firstSubtotalColumn=F0,-0.149998|firstSubtotalRow=B;C1;t1:1,0.499985;b1:1,0.499985|secondSubtotalRow=B;C1|firstRowSubheading=B;C1;b1:0,-0.349986|secondRowSubheading=B;C1|pageFieldLabels=F0,-0.149998;b1:0,-0.349986|pageFieldValues=F0,-0.149998;b1:0,-0.349986"
      "PivotStyleLight16", "headerRow=F4,0.799982;B;C1;b1:4,0.399976|totalRow=F4,0.799982;B;C1;t1:4,0.399976|firstRowStripe=F0,-0.149998|firstColumnStripe=F0,-0.149998;l1:0,-0.249977;r1:0,-0.249977|firstSubtotalColumn=F0,-0.149998|firstSubtotalRow=B;C1;t1:4;b1:4|secondSubtotalRow=B;C1|firstRowSubheading=B;C1;b1:4,0.399976|secondRowSubheading=B;C1|pageFieldLabels=F4,0.799982;b1:4,0.399976|pageFieldValues=F4,0.799982;b1:4,0.399976"
      "PivotStyleLight17", "headerRow=F5,0.799982;B;C1;b1:5,0.399976|totalRow=F5,0.799982;B;C1;t1:5,0.399976|firstRowStripe=F0,-0.149998|firstColumnStripe=F0,-0.149998;l1:0,-0.249977;r1:0,-0.249977|firstSubtotalColumn=F0,-0.149998|firstSubtotalRow=B;C1;t1:5;b1:5|secondSubtotalRow=B;C1|firstRowSubheading=B;C1;b1:5,0.399976|secondRowSubheading=B;C1|pageFieldLabels=F5,0.799982;b1:5,0.399976|pageFieldValues=F5,0.799982;b1:5,0.399976"
      "PivotStyleLight18", "headerRow=F6,0.799982;B;C1;b1:6,0.399976|totalRow=F6,0.799982;B;C1;t1:6,0.399976|firstRowStripe=F0,-0.149998|firstColumnStripe=F0,-0.149998;l1:0,-0.249977;r1:0,-0.249977|firstSubtotalColumn=F0,-0.149998|firstSubtotalRow=B;C1;t1:6;b1:6|secondSubtotalRow=B;C1|firstRowSubheading=B;C1;b1:6,0.399976|secondRowSubheading=B;C1|pageFieldLabels=F6,0.799982;b1:6,0.399976|pageFieldValues=F6,0.799982;b1:6,0.399976"
      "PivotStyleLight19", "headerRow=F7,0.799982;B;C1;b1:7,0.399976|totalRow=F7,0.799982;B;C1;t1:7,0.399976|firstRowStripe=F0,-0.149998|firstColumnStripe=F0,-0.149998;l1:0,-0.249977;r1:0,-0.249977|firstSubtotalColumn=F0,-0.149998|firstSubtotalRow=B;C1;t1:7;b1:7|secondSubtotalRow=B;C1|firstRowSubheading=B;C1;b1:7,0.399976|secondRowSubheading=B;C1|pageFieldLabels=F7,0.799982;b1:7,0.399976|pageFieldValues=F7,0.799982;b1:7,0.399976"
      "PivotStyleLight20", "headerRow=F8,0.799982;B;C1;b1:8,0.399976|totalRow=F8,0.799982;B;C1;t1:8,0.399976|firstRowStripe=F0,-0.149998|firstColumnStripe=F0,-0.149998;l1:0,-0.249977;r1:0,-0.249977|firstSubtotalColumn=F0,-0.149998|firstSubtotalRow=B;C1;t1:8;b1:8|secondSubtotalRow=B;C1|firstRowSubheading=B;C1;b1:8,0.399976|secondRowSubheading=B;C1|pageFieldLabels=F8,0.799982;b1:8,0.399976|pageFieldValues=F8,0.799982;b1:8,0.399976"
      "PivotStyleLight21", "headerRow=F9,0.799982;B;C1;b1:9,0.399976|totalRow=F9,0.799982;B;C1;t1:9,0.399976|firstRowStripe=F0,-0.149998|firstColumnStripe=F0,-0.149998;l1:0,-0.249977;r1:0,-0.249977|firstSubtotalColumn=F0,-0.149998|firstSubtotalRow=B;C1;t1:9;b1:9|secondSubtotalRow=B;C1|firstRowSubheading=B;C1;b1:9,0.399976|secondRowSubheading=B;C1|pageFieldLabels=F9,0.799982;b1:9,0.399976|pageFieldValues=F9,0.799982;b1:9,0.399976"
      "PivotStyleLight22", "wholeTable=C1;l1:1,0.499985;r1:1,0.499985;t1:1,0.499985;b1:1,0.499985;v1:1,0.499985|headerRow=B;C1;b1:1,0.499985|totalRow=B;C1;t1:1,0.499985|firstColumn=B;C1|firstRowStripe=F0,-0.149998;t1:1,0.499985;b1:1,0.499985|firstColumnStripe=F0,-0.149998|firstSubtotalColumn=B;C1;l1:0,-0.349986;r1:0,-0.349986|firstSubtotalRow=B;C1|secondSubtotalRow=B;C1|pageFieldLabels=B;C1"
      "PivotStyleLight23", "wholeTable=C4,-0.249977;l1:4;r1:4;t1:4;b1:4;v1:4|headerRow=C4,-0.249977;b1:4|totalRow=C4,-0.249977;t1:4|firstColumn=C4,-0.249977|firstRowStripe=F4,0.799982;t1:1,0.499985;b1:1,0.499985|firstColumnStripe=F4,0.799982|firstSubtotalColumn=C4,-0.249977|firstSubtotalRow=C4,-0.249977|secondSubtotalRow=C4,-0.249977|pageFieldLabels=C4,-0.249977"
      "PivotStyleLight24", "wholeTable=C5,-0.249977;l1:5;r1:5;t1:5;b1:5;v1:5|headerRow=C5,-0.249977;b1:5|totalRow=C5,-0.249977;t1:5|firstColumn=C5,-0.249977|firstRowStripe=F5,0.799982;t1:1,0.499985;b1:1,0.499985|firstColumnStripe=F5,0.799982|firstSubtotalColumn=C5,-0.249977|firstSubtotalRow=C5,-0.249977|secondSubtotalRow=C5,-0.249977|pageFieldLabels=C5,-0.249977"
      "PivotStyleLight25", "wholeTable=C6,-0.249977;l1:6;r1:6;t1:6;b1:6;v1:6|headerRow=C6,-0.249977;b1:6|totalRow=C6,-0.249977;t1:6|firstColumn=C6,-0.249977|firstRowStripe=F6,0.799982;t1:1,0.499985;b1:1,0.499985|firstColumnStripe=F6,0.799982|firstSubtotalColumn=C6,-0.249977|firstSubtotalRow=C6,-0.249977|secondSubtotalRow=C6,-0.249977|pageFieldLabels=C6,-0.249977"
      "PivotStyleLight26", "wholeTable=C7,-0.249977;l1:7;r1:7;t1:7;b1:7;v1:7|headerRow=C7,-0.249977;b1:7|totalRow=C7,-0.249977;t1:7|firstColumn=C7,-0.249977|firstRowStripe=F7,0.799982;t1:1,0.499985;b1:1,0.499985|firstColumnStripe=F7,0.799982|firstSubtotalColumn=C7,-0.249977|firstSubtotalRow=C7,-0.249977|secondSubtotalRow=C7,-0.249977|pageFieldLabels=C7,-0.249977"
      "PivotStyleLight27", "wholeTable=C8,-0.249977;l1:8;r1:8;t1:8;b1:8;v1:8|headerRow=C8,-0.249977;b1:8|totalRow=C8,-0.249977;t1:8|firstColumn=C8,-0.249977|firstRowStripe=F8,0.799982;t1:1,0.499985;b1:1,0.499985|firstColumnStripe=F8,0.799982|firstSubtotalColumn=C8,-0.249977|firstSubtotalRow=C8,-0.249977|secondSubtotalRow=C8,-0.249977|pageFieldLabels=C8,-0.249977"
      "PivotStyleLight28", "wholeTable=C9,-0.249977;l1:9;r1:9;t1:9;b1:9;v1:9|headerRow=C9,-0.249977;b1:9|totalRow=C9,-0.249977;t1:9|firstColumn=C9,-0.249977|firstRowStripe=F9,0.799982;t1:1,0.499985;b1:1,0.499985|firstColumnStripe=F9,0.799982|firstSubtotalColumn=C9,-0.249977|firstSubtotalRow=C9,-0.249977|secondSubtotalRow=C9,-0.249977|pageFieldLabels=C9,-0.249977"
      "PivotStyleMedium1", "wholeTable=C1;h1:0,-0.149998|headerRow=F1,0.499985;C0;h1:1,0.499985|totalRow=B;C1;td:1,0.499985|firstRowStripe=t1:1,0.499985;b1:1,0.499985;h1:1,0.499985|firstColumnStripe=l1:1,0.499985;r1:1,0.499985|firstHeaderCell=B;C0|firstSubtotalRow=F0,-0.349986;B;C0|secondSubtotalRow=F0,-0.149998;B;C1|firstColumnSubheading=b1:0,-0.149998|firstRowSubheading=F0,-0.349986;C0;b1:0,-0.149998;h1:0,-0.149998|secondRowSubheading=F0,-0.149998;b1:0,-0.149998|pageFieldLabels=t1:0,-0.149998;b1:0,-0.149998|pageFieldValues=t1:0,-0.149998;b1:0,-0.149998"
      "PivotStyleMedium2", "wholeTable=C1;h1:4,0.799982|headerRow=F4,-0.249977;C0;h1:4,-0.249977|totalRow=B;C1;td:4,-0.249977|firstRowStripe=t1:4,-0.249977;b1:4,-0.249977;h1:4,-0.249977|firstColumnStripe=l1:4,-0.249977;r1:4,-0.249977|firstHeaderCell=B;C0|firstSubtotalRow=F4,0.399976;B;C0|secondSubtotalRow=F0,-0.149998;B;C1|firstColumnSubheading=b1:4,0.599994|firstRowSubheading=F4,0.399976;C0;b1:4,0.799982;h1:4,0.399976|secondRowSubheading=F4,0.799982;b1:4|pageFieldLabels=t1:4,0.799982;b1:4,0.799982|pageFieldValues=t1:4,0.799982;b1:4,0.799982"
      "PivotStyleMedium3", "wholeTable=C1;h1:5,0.799982|headerRow=F5,-0.249977;C0;h1:5,-0.249977|totalRow=B;C1;td:5,-0.249977|firstRowStripe=t1:5,-0.249977;b1:5,-0.249977;h1:5,-0.249977|firstColumnStripe=l1:5,-0.249977;r1:5,-0.249977|firstHeaderCell=B;C0|firstSubtotalRow=F5,0.399976;B;C0|secondSubtotalRow=F0,-0.149998;B;C1|firstColumnSubheading=b1:5,0.599994|firstRowSubheading=F5,0.399976;C0;b1:5,0.799982;h1:5,0.399976|secondRowSubheading=F5,0.799982;b1:5|pageFieldLabels=t1:5,0.799982;b1:5,0.799982|pageFieldValues=t1:5,0.799982;b1:5,0.799982"
      "PivotStyleMedium4", "wholeTable=C1;h1:6,0.799982|headerRow=F6,-0.249977;C0;h1:6,-0.249977|totalRow=B;C1;td:6,-0.249977|firstRowStripe=t1:6,-0.249977;b1:6,-0.249977;h1:6,-0.249977|firstColumnStripe=l1:6,-0.249977;r1:6,-0.249977|firstHeaderCell=B;C0|firstSubtotalRow=F6,0.399976;B;C0|secondSubtotalRow=F0,-0.149998;B;C1|firstColumnSubheading=b1:6,0.599994|firstRowSubheading=F6,0.399976;C0;b1:6,0.799982;h1:6,0.399976|secondRowSubheading=F6,0.799982;b1:6|pageFieldLabels=t1:6,0.799982;b1:6,0.799982|pageFieldValues=t1:6,0.799982;b1:6,0.799982"
      "PivotStyleMedium5", "wholeTable=C1;h1:7,0.799982|headerRow=F7,-0.249977;C0;h1:7,-0.249977|totalRow=B;C1;td:7,-0.249977|firstRowStripe=t1:7,-0.249977;b1:7,-0.249977;h1:7,-0.249977|firstColumnStripe=l1:7,-0.249977;r1:7,-0.249977|firstHeaderCell=B;C0|firstSubtotalRow=F7,0.399976;B;C0|secondSubtotalRow=F0,-0.149998;B;C1|firstColumnSubheading=b1:7,0.599994|firstRowSubheading=F7,0.399976;C0;b1:7,0.799982;h1:7,0.399976|secondRowSubheading=F7,0.799982;b1:7|pageFieldLabels=t1:7,0.799982;b1:7,0.799982|pageFieldValues=t1:7,0.799982;b1:7,0.799982"
      "PivotStyleMedium6", "wholeTable=C1;h1:8,0.799982|headerRow=F8,-0.249977;C0;h1:8,-0.249977|totalRow=B;C1;td:8,-0.249977|firstRowStripe=t1:8,-0.249977;b1:8,-0.249977;h1:8,-0.249977|firstColumnStripe=l1:8,-0.249977;r1:8,-0.249977|firstHeaderCell=B;C0|firstSubtotalRow=F8,0.399976;B;C0|secondSubtotalRow=F0,-0.149998;B;C1|firstColumnSubheading=b1:8,0.599994|firstRowSubheading=F8,0.399976;C0;b1:8,0.799982;h1:8,0.399976|secondRowSubheading=F8,0.799982;b1:8|pageFieldLabels=t1:8,0.799982;b1:8,0.799982|pageFieldValues=t1:8,0.799982;b1:8,0.799982"
      "PivotStyleMedium7", "wholeTable=C1;h1:9,0.799982|headerRow=F9,-0.249977;C0;h1:9,-0.249977|totalRow=B;C1;td:9,-0.249977|firstRowStripe=t1:9,-0.249977;b1:9,-0.249977;h1:9,-0.249977|firstColumnStripe=l1:9,-0.249977;r1:9,-0.249977|firstHeaderCell=B;C0|firstSubtotalRow=F9,0.399976;B;C0|secondSubtotalRow=F0,-0.149998;B;C1|firstColumnSubheading=b1:9,0.599994|firstRowSubheading=F9,0.399976;C0;b1:9,0.799982;h1:9,0.399976|secondRowSubheading=F9,0.799982;b1:9|pageFieldLabels=t1:9,0.799982;b1:9,0.799982|pageFieldValues=t1:9,0.799982;b1:9,0.799982"
      "PivotStyleMedium8", "wholeTable=C1|headerRow=F1,0.499985;B;C0;t2:1,0.499985|totalRow=B;C1;t1:1;b2:1,0.499985|firstRowStripe=t1:0,-0.249977;b1:0,-0.249977;h1:0,-0.249977|firstColumnStripe=l2:0,-0.149998;r2:0,-0.149998|firstSubtotalColumn=B;C1;l2:0,-0.249977;r2:0,-0.249977;t2:0,-0.249977;b2:0,-0.249977|firstSubtotalRow=F0,-0.249977;B;C1|secondSubtotalRow=B;C1|firstRowSubheading=F0,-0.149998;B;C1|secondRowSubheading=B;C1|pageFieldLabels=F0,-0.149998|pageFieldValues=F0,-0.149998"
      "PivotStyleMedium9", "wholeTable=C1|headerRow=F4;B;C0;t2:4,-0.249977|totalRow=B;C1;t1:4,-0.249977;b2:4,-0.249977|firstRowStripe=t1:4,0.399976;b1:4,0.399976;h1:4,0.399976|firstColumnStripe=l1:4,0.399976;r1:4,0.399976|firstSubtotalColumn=B;C1;l2:4,0.599994;r2:4,0.599994;t2:4,0.599994;b2:4,0.599994|firstSubtotalRow=F4,0.599994;B;C1|secondSubtotalRow=B;C1|firstRowSubheading=F4,0.799982;B;C1|secondRowSubheading=B;C1|pageFieldLabels=F4,0.799982|pageFieldValues=F4,0.799982"
      "PivotStyleMedium10", "wholeTable=C1|headerRow=F5;B;C0;t2:5,-0.249977|totalRow=B;C1;t1:5,-0.249977;b2:5,-0.249977|firstRowStripe=t1:5,0.399976;b1:5,0.399976;h1:5,0.399976|firstColumnStripe=l1:5,0.399976;r1:5,0.399976|firstSubtotalColumn=B;C1;l2:5,0.599994;r2:5,0.599994;t2:5,0.599994;b2:5,0.599994|firstSubtotalRow=F5,0.599994;B;C1|secondSubtotalRow=B;C1|firstRowSubheading=F5,0.799982;B;C1|secondRowSubheading=B;C1|pageFieldLabels=F5,0.799982|pageFieldValues=F5,0.799982"
      "PivotStyleMedium11", "wholeTable=C1|headerRow=F6;B;C0;t2:6,-0.249977|totalRow=B;C1;t1:6,-0.249977;b2:6,-0.249977|firstRowStripe=t1:6,0.399976;b1:6,0.399976;h1:6,0.399976|firstColumnStripe=l1:6,0.399976;r1:6,0.399976|firstSubtotalColumn=B;C1;l2:6,0.599994;r2:6,0.599994;t2:6,0.599994;b2:6,0.599994|firstSubtotalRow=F6,0.599994;B;C1|secondSubtotalRow=B;C1|firstRowSubheading=F6,0.799982;B;C1|secondRowSubheading=B;C1|pageFieldLabels=F6,0.799982|pageFieldValues=F6,0.799982"
      "PivotStyleMedium12", "wholeTable=C1|headerRow=F7;B;C0;t2:7,-0.249977|totalRow=B;C1;t1:7,-0.249977;b2:7,-0.249977|firstRowStripe=t1:7,0.399976;b1:7,0.399976;h1:7,0.399976|firstColumnStripe=l1:7,0.399976;r1:7,0.399976|firstSubtotalColumn=B;C1;l2:7,0.599994;r2:7,0.599994;t2:7,0.599994;b2:7,0.599994|firstSubtotalRow=F7,0.599994;B;C1|secondSubtotalRow=B;C1|firstRowSubheading=F7,0.799982;B;C1|secondRowSubheading=B;C1|pageFieldLabels=F7,0.799982|pageFieldValues=F7,0.799982"
      "PivotStyleMedium13", "wholeTable=C1|headerRow=F8;B;C0;t2:8,-0.249977|totalRow=B;C1;t1:8,-0.249977;b2:8,-0.249977|firstRowStripe=t1:8,0.399976;b1:8,0.399976;h1:8,0.399976|firstColumnStripe=l1:8,0.399976;r1:8,0.399976|firstSubtotalColumn=B;C1;l2:8,0.599994;r2:8,0.599994;t2:8,0.599994;b2:8,0.599994|firstSubtotalRow=F8,0.599994;B;C1|secondSubtotalRow=B;C1|firstRowSubheading=F8,0.799982;B;C1|secondRowSubheading=B;C1|pageFieldLabels=F8,0.799982|pageFieldValues=F8,0.799982"
      "PivotStyleMedium14", "wholeTable=C1|headerRow=F9;B;C0;t2:9,-0.249977|totalRow=B;C1;t1:9,-0.249977;b2:9,-0.249977|firstRowStripe=t1:9,0.399976;b1:9,0.399976;h1:9,0.399976|firstColumnStripe=l1:9,0.399976;r1:9,0.399976|firstSubtotalColumn=B;C1;l2:9,0.599994;r2:9,0.599994;t2:9,0.599994;b2:9,0.599994|firstSubtotalRow=F9,0.599994;B;C1|secondSubtotalRow=B;C1|firstRowSubheading=F9,0.799982;B;C1|secondRowSubheading=B;C1|pageFieldLabels=F9,0.799982|pageFieldValues=F9,0.799982"
      "PivotStyleMedium15", "wholeTable=F0,-0.049989;C1;l1:0,-0.149998;r1:0,-0.149998;v1:0,-0.149998|headerRow=F1;C0|totalRow=F1;C0|firstRowStripe=F0,-0.149998;t1:0,-0.249977;b1:0,-0.249977|firstColumnStripe=F0,-0.149998;l1:0,-0.249977;r1:0,-0.249977|firstSubtotalColumn=F0,-0.249977;C1;l1:0,-0.349986;r1:0,-0.349986;t1:0,-0.349986|firstSubtotalRow=B;C1;b1:0,-0.249977|secondSubtotalRow=B;C1,0.499985|firstRowSubheading=B;C1|secondRowSubheading=B;C1,0.499985|thirdRowSubheading=B;C1|pageFieldLabels=t1:1;b1:1|pageFieldValues=t1:1;b1:1"
      "PivotStyleMedium16", "wholeTable=F4,0.799982;C1;l1:4,0.599994;r1:4,0.599994;v1:4,0.599994|headerRow=F1;C0|totalRow=F1;C0|firstRowStripe=F4,0.799982;t1:4,0.599994;b1:4,0.599994|firstColumnStripe=F4,0.799982;l1:4,0.599994;r1:4,0.599994|firstSubtotalColumn=F4,0.599994;C1;l1:4,0.399976;r1:4,0.399976;t1:4,0.399976|firstSubtotalRow=B;C1;b1:4,0.599994|secondSubtotalRow=B;C1,0.499985|firstRowSubheading=B;C1|secondRowSubheading=B;C1,0.499985|thirdRowSubheading=B;C1|pageFieldLabels=t1:1;b1:1|pageFieldValues=t1:1;b1:1"
      "PivotStyleMedium17", "wholeTable=F5,0.799982;C1;l1:5,0.599994;r1:5,0.599994;v1:5,0.599994|headerRow=F1;C0|totalRow=F1;C0|firstRowStripe=F5,0.799982;t1:5,0.599994;b1:5,0.599994|firstColumnStripe=F5,0.799982;l1:5,0.599994;r1:5,0.599994|firstSubtotalColumn=F5,0.599994;C1;l1:5,0.399976;r1:5,0.399976;t1:5,0.399976|firstSubtotalRow=B;C1;b1:5,0.599994|secondSubtotalRow=B;C1,0.499985|firstRowSubheading=B;C1|secondRowSubheading=B;C1,0.499985|thirdRowSubheading=B;C1|pageFieldLabels=t1:1;b1:1|pageFieldValues=t1:1;b1:1"
      "PivotStyleMedium18", "wholeTable=F6,0.799982;C1;l1:6,0.599994;r1:6,0.599994;v1:6,0.599994|headerRow=F1;C0|totalRow=F1;C0|firstRowStripe=F6,0.799982;t1:6,0.599994;b1:6,0.599994|firstColumnStripe=F6,0.799982;l1:6,0.599994;r1:6,0.599994|firstSubtotalColumn=F6,0.599994;C1;l1:6,0.399976;r1:6,0.399976;t1:6,0.399976|firstSubtotalRow=B;C1;b1:6,0.599994|secondSubtotalRow=B;C1,0.499985|firstRowSubheading=B;C1|secondRowSubheading=B;C1,0.499985|thirdRowSubheading=B;C1|pageFieldLabels=t1:1;b1:1|pageFieldValues=t1:1;b1:1"
      "PivotStyleMedium19", "wholeTable=F7,0.799982;C1;l1:7,0.599994;r1:7,0.599994;v1:7,0.599994|headerRow=F1;C0|totalRow=F1;C0|firstRowStripe=F7,0.799982;t1:7,0.599994;b1:7,0.599994|firstColumnStripe=F7,0.799982;l1:7,0.599994;r1:7,0.599994|firstSubtotalColumn=F7,0.599994;C1;l1:7,0.399976;r1:7,0.399976;t1:7,0.399976|firstSubtotalRow=B;C1;b1:7,0.599994|secondSubtotalRow=B;C1,0.499985|firstRowSubheading=B;C1|secondRowSubheading=B;C1,0.499985|thirdRowSubheading=B;C1|pageFieldLabels=t1:1;b1:1|pageFieldValues=t1:1;b1:1"
      "PivotStyleMedium20", "wholeTable=F8,0.799982;C1;l1:8,0.599994;r1:8,0.599994;v1:8,0.599994|headerRow=F1;C0|totalRow=F1;C0|firstRowStripe=F8,0.799982;t1:8,0.599994;b1:8,0.599994|firstColumnStripe=F8,0.799982;l1:8,0.599994;r1:8,0.599994|firstSubtotalColumn=F8,0.599994;C1;l1:8,0.399976;r1:8,0.399976;t1:8,0.399976|firstSubtotalRow=B;C1;b1:8,0.599994|secondSubtotalRow=B;C1,0.499985|firstRowSubheading=B;C1|secondRowSubheading=B;C1,0.499985|thirdRowSubheading=B;C1|pageFieldLabels=t1:1;b1:1|pageFieldValues=t1:1;b1:1"
      "PivotStyleMedium21", "wholeTable=F9,0.799982;C1;l1:9,0.599994;r1:9,0.599994;v1:9,0.599994|headerRow=F1;C0|totalRow=F1;C0|firstRowStripe=F9,0.799982;t1:9,0.599994;b1:9,0.599994|firstColumnStripe=F9,0.799982;l1:9,0.599994;r1:9,0.599994|firstSubtotalColumn=F9,0.599994;C1;l1:9,0.399976;r1:9,0.399976;t1:9,0.399976|firstSubtotalRow=B;C1;b1:9,0.599994|secondSubtotalRow=B;C1,0.499985|firstRowSubheading=B;C1|secondRowSubheading=B;C1,0.499985|thirdRowSubheading=B;C1|pageFieldLabels=t1:1;b1:1|pageFieldValues=t1:1;b1:1"
      "PivotStyleMedium22", "wholeTable=F0,-0.149998;C1,0.499985;v1:0|headerRow=B;C1,0.499985;b1:0|totalRow=B;C1,0.499985|firstColumn=F0,-0.249977;B;C1,0.499985|secondRowStripe=F0,-0.249977|secondColumnStripe=F0,-0.249977|firstSubtotalRow=B;C1|secondSubtotalRow=B;C1|firstRowSubheading=B;C1|pageFieldLabels=F0,-0.249977;B;C1,0.499985"
      "PivotStyleMedium23", "wholeTable=F4,0.799982;C4,-0.249977;v1:0|headerRow=B;C4,-0.249977;b1:0|totalRow=B;C4,-0.249977|firstColumn=F4,0.599994;B;C4,-0.249977|secondRowStripe=F4,0.599994|secondColumnStripe=F4,0.599994|firstSubtotalRow=B;C1|secondSubtotalRow=B;C1|firstRowSubheading=B;C1|pageFieldLabels=F4,0.599994;B;C4,-0.249977"
      "PivotStyleMedium24", "wholeTable=F5,0.799982;C5,-0.249977;v1:0|headerRow=B;C5,-0.249977;b1:0|totalRow=B;C5,-0.249977|firstColumn=F5,0.599994;B;C5,-0.249977|secondRowStripe=F5,0.599994|secondColumnStripe=F5,0.599994|firstSubtotalRow=B;C1|secondSubtotalRow=B;C1|firstRowSubheading=B;C1|pageFieldLabels=F5,0.599994;B;C5,-0.249977"
      "PivotStyleMedium25", "wholeTable=F6,0.799982;C6,-0.249977;v1:0|headerRow=B;C6,-0.249977;b1:0|totalRow=B;C6,-0.249977|firstColumn=F6,0.599994;B;C6,-0.249977|secondRowStripe=F6,0.599994|secondColumnStripe=F6,0.599994|firstSubtotalRow=B;C1|secondSubtotalRow=B;C1|firstRowSubheading=B;C1|pageFieldLabels=F6,0.599994;B;C6,-0.249977"
      "PivotStyleMedium26", "wholeTable=F7,0.799982;C7,-0.249977;v1:0|headerRow=B;C7,-0.249977;b1:0|totalRow=B;C7,-0.249977|firstColumn=F7,0.599994;B;C7,-0.249977|secondRowStripe=F7,0.599994|secondColumnStripe=F7,0.599994|firstSubtotalRow=B;C1|secondSubtotalRow=B;C1|firstRowSubheading=B;C1|pageFieldLabels=F7,0.599994;B;C7,-0.249977"
      "PivotStyleMedium27", "wholeTable=F8,0.799982;C8,-0.249977;v1:0|headerRow=B;C8,-0.249977;b1:0|totalRow=B;C8,-0.249977|firstColumn=F8,0.599994;B;C8,-0.249977|secondRowStripe=F8,0.599994|secondColumnStripe=F8,0.599994|firstSubtotalRow=B;C1|secondSubtotalRow=B;C1|firstRowSubheading=B;C1|pageFieldLabels=F8,0.599994;B;C8,-0.249977"
      "PivotStyleMedium28", "wholeTable=F9,0.799982;C9,-0.249977;v1:0|headerRow=B;C9,-0.249977;b1:0|totalRow=B;C9,-0.249977|firstColumn=F9,0.599994;B;C9,-0.249977|secondRowStripe=F9,0.599994|secondColumnStripe=F9,0.599994|firstSubtotalRow=B;C1|secondSubtotalRow=B;C1|firstRowSubheading=B;C1|pageFieldLabels=F9,0.599994;B;C9,-0.249977"
      "PivotStyleDark1", "wholeTable=F0,-0.249977;C1;h1:0,-0.149998|headerRow=F1,0.499985;B;C0;b1:0;h1:1,0.499985|totalRow=F1,0.499985;B;C0|secondRowStripe=F0,-0.349986|secondColumnStripe=l1:0,-0.149998;r1:0,-0.149998|firstSubtotalRow=B;C1;b1:0,-0.349986|firstRowSubheading=F0,-0.149998;B;C1;t1:0,-0.349986;b1:0,-0.349986|secondRowSubheading=B;C1|pageFieldLabels=F1,0.499985;B;C0;h1:1,0.499985|pageFieldValues=F1,0.499985;C0;h1:1,0.499985"
      "PivotStyleDark2", "wholeTable=F4,0.599994;C1;h1:4,0.799982|headerRow=F4,-0.499985;B;C0;b1:4;h1:4,-0.499985|totalRow=F4,-0.499985;B;C0|secondRowStripe=F4,0.399976|secondColumnStripe=l1:4,0.799982;r1:4,0.799982|firstSubtotalRow=B;C1;b1:4,0.799982|firstRowSubheading=F4,0.799982;B;C1;t1:4,0.399976;b1:4,0.399976|secondRowSubheading=B;C1|pageFieldLabels=F4,-0.499985;B;C0;h1:4,-0.499985|pageFieldValues=F4,-0.499985;C0;h1:4,-0.499985"
      "PivotStyleDark3", "wholeTable=F5,0.599994;C1;h1:5,0.799982|headerRow=F5,-0.499985;B;C0;b1:5;h1:5,-0.499985|totalRow=F5,-0.499985;B;C0|secondRowStripe=F5,0.399976|secondColumnStripe=l1:5,0.799982;r1:5,0.799982|firstSubtotalRow=B;C1;b1:5,0.799982|firstRowSubheading=F5,0.799982;B;C1;t1:5,0.399976;b1:5,0.399976|secondRowSubheading=B;C1|pageFieldLabels=F5,-0.499985;B;C0;h1:5,-0.499985|pageFieldValues=F5,-0.499985;C0;h1:5,-0.499985"
      "PivotStyleDark4", "wholeTable=F6,0.599994;C1;h1:6,0.799982|headerRow=F6,-0.499985;B;C0;b1:6;h1:6,-0.499985|totalRow=F6,-0.499985;B;C0|secondRowStripe=F6,0.399976|secondColumnStripe=l1:6,0.799982;r1:6,0.799982|firstSubtotalRow=B;C1;b1:6,0.799982|firstRowSubheading=F6,0.799982;B;C1;t1:6,0.399976;b1:6,0.399976|secondRowSubheading=B;C1|pageFieldLabels=F6,-0.499985;B;C0;h1:6,-0.499985|pageFieldValues=F6,-0.499985;C0;h1:6,-0.499985"
      "PivotStyleDark5", "wholeTable=F7,0.599994;C1;h1:7,0.799982|headerRow=F7,-0.499985;B;C0;b1:7;h1:7,-0.499985|totalRow=F7,-0.499985;B;C0|secondRowStripe=F7,0.399976|secondColumnStripe=l1:7,0.799982;r1:7,0.799982|firstSubtotalRow=B;C1;b1:7,0.799982|firstRowSubheading=F7,0.799982;B;C1;t1:7,0.399976;b1:7,0.399976|secondRowSubheading=B;C1|pageFieldLabels=F7,-0.499985;B;C0;h1:7,-0.499985|pageFieldValues=F7,-0.499985;C0;h1:7,-0.499985"
      "PivotStyleDark6", "wholeTable=F8,0.599994;C1;h1:8,0.799982|headerRow=F8,-0.499985;B;C0;b1:8;h1:8,-0.499985|totalRow=F8,-0.499985;B;C0|secondRowStripe=F8,0.399976|secondColumnStripe=l1:8,0.799982;r1:8,0.799982|firstSubtotalRow=B;C1;b1:8,0.799982|firstRowSubheading=F8,0.799982;B;C1;t1:8,0.399976;b1:8,0.399976|secondRowSubheading=B;C1|pageFieldLabels=F8,-0.499985;B;C0;h1:8,-0.499985|pageFieldValues=F8,-0.499985;C0;h1:8,-0.499985"
      "PivotStyleDark7", "wholeTable=F9,0.599994;C1;h1:9,0.799982|headerRow=F9,-0.499985;B;C0;b1:9;h1:9,-0.499985|totalRow=F9,-0.499985;B;C0|secondRowStripe=F9,0.399976|secondColumnStripe=l1:9,0.799982;r1:9,0.799982|firstSubtotalRow=B;C1;b1:9,0.799982|firstRowSubheading=F9,0.799982;B;C1;t1:9,0.399976;b1:9,0.399976|secondRowSubheading=B;C1|pageFieldLabels=F9,-0.499985;B;C0;h1:9,-0.499985|pageFieldValues=F9,-0.499985;C0;h1:9,-0.499985"
      "PivotStyleDark8", "wholeTable=F0,-0.149998;C1;l2:1,0.499985;r2:1,0.499985;t2:1,0.499985;b2:1,0.499985|headerRow=F1,0.249977;B;C0|totalRow=F1,0.249977;B;C0|secondRowStripe=t1:0,-0.249977;b1:0,-0.249977|firstColumnStripe=l1:0,-0.449995;r1:0,-0.449995|secondColumnStripe=l1:0,-0.449995;r1:0,-0.449995|firstSubtotalRow=F0,-0.249977;B;C1|secondColumnSubheading=t2:0,-0.249977|thirdColumnSubheading=t2:0,-0.249977|firstRowSubheading=F0,-0.249977;B;C1;b2:0,-0.149998|secondRowSubheading=B;C1|pageFieldLabels=l2:1,0.499985;t2:1,0.499985;b2:1,0.499985|pageFieldValues=r2:1,0.499985;t2:1,0.499985;b2:1,0.499985"
      "PivotStyleDark9", "wholeTable=F4,0.799982;C1;l2:1,0.499985;r2:1,0.499985;t2:1,0.499985;b2:1,0.499985|headerRow=F1,0.249977;B;C0|totalRow=F1,0.249977;B;C0|secondRowStripe=t1:4,0.599994;b1:4,0.599994|firstColumnStripe=l1:4,0.399976;r1:4,0.399976|secondColumnStripe=l1:4,0.399976;r1:4,0.399976|firstSubtotalRow=F4,0.599994;B;C1|secondColumnSubheading=t2:4,0.799982|thirdColumnSubheading=t2:4,0.799982|firstRowSubheading=F4,0.599994;B;C1;b2:4,0.799982|secondRowSubheading=B;C1|pageFieldLabels=l2:1,0.499985;t2:1,0.499985;b2:1,0.499985|pageFieldValues=r2:1,0.499985;t2:1,0.499985;b2:1,0.499985"
      "PivotStyleDark10", "wholeTable=F5,0.799982;C1;l2:1,0.499985;r2:1,0.499985;t2:1,0.499985;b2:1,0.499985|headerRow=F1,0.249977;B;C0|totalRow=F1,0.249977;B;C0|secondRowStripe=t1:5,0.599994;b1:5,0.599994|firstColumnStripe=l1:5,0.399976;r1:5,0.399976|secondColumnStripe=l1:5,0.399976;r1:5,0.399976|firstSubtotalRow=F5,0.599994;B;C1|secondColumnSubheading=t2:5,0.799982|thirdColumnSubheading=t2:5,0.799982|firstRowSubheading=F5,0.599994;B;C1;b2:5,0.799982|secondRowSubheading=B;C1|pageFieldLabels=l2:1,0.499985;t2:1,0.499985;b2:1,0.499985|pageFieldValues=r2:1,0.499985;t2:1,0.499985;b2:1,0.499985"
      "PivotStyleDark11", "wholeTable=F6,0.799982;C1;l2:1,0.499985;r2:1,0.499985;t2:1,0.499985;b2:1,0.499985|headerRow=F1,0.249977;B;C0|totalRow=F1,0.249977;B;C0|secondRowStripe=t1:6,0.599994;b1:6,0.599994|firstColumnStripe=l1:6,0.399976;r1:6,0.399976|secondColumnStripe=l1:6,0.399976;r1:6,0.399976|firstSubtotalRow=F6,0.599994;B;C1|secondColumnSubheading=t2:6,0.799982|thirdColumnSubheading=t2:6,0.799982|firstRowSubheading=F6,0.599994;B;C1;b2:6,0.799982|secondRowSubheading=B;C1|pageFieldLabels=l2:1,0.499985;t2:1,0.499985;b2:1,0.499985|pageFieldValues=r2:1,0.499985;t2:1,0.499985;b2:1,0.499985"
      "PivotStyleDark12", "wholeTable=F7,0.799982;C1;l2:1,0.499985;r2:1,0.499985;t2:1,0.499985;b2:1,0.499985|headerRow=F1,0.249977;B;C0|totalRow=F1,0.249977;B;C0|secondRowStripe=t1:7,0.599994;b1:7,0.599994|firstColumnStripe=l1:7,0.399976;r1:7,0.399976|secondColumnStripe=l1:7,0.399976;r1:7,0.399976|firstSubtotalRow=F7,0.599994;B;C1|secondColumnSubheading=t2:7,0.799982|thirdColumnSubheading=t2:7,0.799982|firstRowSubheading=F7,0.599994;B;C1;b2:7,0.799982|secondRowSubheading=B;C1|pageFieldLabels=l2:1,0.499985;t2:1,0.499985;b2:1,0.499985|pageFieldValues=r2:1,0.499985;t2:1,0.499985;b2:1,0.499985"
      "PivotStyleDark13", "wholeTable=F8,0.799982;C1;l2:1,0.499985;r2:1,0.499985;t2:1,0.499985;b2:1,0.499985|headerRow=F1,0.249977;B;C0|totalRow=F1,0.249977;B;C0|secondRowStripe=t1:8,0.599994;b1:8,0.599994|firstColumnStripe=l1:8,0.399976;r1:8,0.399976|secondColumnStripe=l1:8,0.399976;r1:8,0.399976|firstSubtotalRow=F8,0.599994;B;C1|secondColumnSubheading=t2:8,0.799982|thirdColumnSubheading=t2:8,0.799982|firstRowSubheading=F8,0.599994;B;C1;b2:8,0.799982|secondRowSubheading=B;C1|pageFieldLabels=l2:1,0.499985;t2:1,0.499985;b2:1,0.499985|pageFieldValues=r2:1,0.499985;t2:1,0.499985;b2:1,0.499985"
      "PivotStyleDark14", "wholeTable=F9,0.799982;C1;l2:1,0.499985;r2:1,0.499985;t2:1,0.499985;b2:1,0.499985|headerRow=F1,0.249977;B;C0|totalRow=F1,0.249977;B;C0|secondRowStripe=t1:9,0.599994;b1:9,0.599994|firstColumnStripe=l1:9,0.399976;r1:9,0.399976|secondColumnStripe=l1:9,0.399976;r1:9,0.399976|firstSubtotalRow=F9,0.599994;B;C1|secondColumnSubheading=t2:9,0.799982|thirdColumnSubheading=t2:9,0.799982|firstRowSubheading=F9,0.599994;B;C1;b2:9,0.799982|secondRowSubheading=B;C1|pageFieldLabels=l2:1,0.499985;t2:1,0.499985;b2:1,0.499985|pageFieldValues=r2:1,0.499985;t2:1,0.499985;b2:1,0.499985"
      "PivotStyleDark15", "wholeTable=F0,-0.449995;C0,-0.149998|headerRow=F1;B;C0|totalRow=F1;B;C0|firstRowStripe=t1:0,-0.349986;b1:0,-0.349986|firstColumnStripe=l1:0,-0.249977;r1:0,-0.249977|firstSubtotalColumn=F1,0.499985;l1:0,-0.349986;r1:0,-0.349986;t1:0,-0.349986;b1:0,-0.349986|firstSubtotalRow=B;C0|firstColumnSubheading=b1:0,-0.349986|firstRowSubheading=F1,0.499985;B;C0;b1:4,0.799982|secondRowSubheading=B;C0|pageFieldLabels=B;C0"
      "PivotStyleDark16", "wholeTable=F4;C4,0.799982|headerRow=F1;B;C0|totalRow=F1;B;C0|firstRowStripe=t1:4,0.399976;b1:4,0.399976|firstColumnStripe=l1:4,0.399976;r1:4,0.399976|firstSubtotalColumn=F4,-0.249977;l1:4,0.399976;r1:4,0.399976;t1:4,0.399976;b1:4,0.399976|firstSubtotalRow=B;C0|firstColumnSubheading=b1:4,0.399976|firstRowSubheading=F4,-0.249977;B;C0|secondRowSubheading=B;C0|pageFieldLabels=B;C0"
      "PivotStyleDark17", "wholeTable=F5;C5,0.799982|headerRow=F1;B;C0|totalRow=F1;B;C0|firstRowStripe=t1:5,0.399976;b1:5,0.399976|firstColumnStripe=l1:5,0.399976;r1:5,0.399976|firstSubtotalColumn=F5,-0.249977;l1:5,0.399976;r1:5,0.399976;t1:5,0.399976;b1:5,0.399976|firstSubtotalRow=B;C0|firstColumnSubheading=b1:5,0.399976|firstRowSubheading=F5,-0.249977;B;C0|secondRowSubheading=B;C0|pageFieldLabels=B;C0"
      "PivotStyleDark18", "wholeTable=F6;C6,0.799982|headerRow=F1;B;C0|totalRow=F1;B;C0|firstRowStripe=t1:6,0.399976;b1:6,0.399976|firstColumnStripe=l1:6,0.399976;r1:6,0.399976|firstSubtotalColumn=F6,-0.249977;l1:6,0.399976;r1:6,0.399976;t1:6,0.399976;b1:6,0.399976|firstSubtotalRow=B;C0|firstColumnSubheading=b1:6,0.399976|firstRowSubheading=F6,-0.249977;B;C0|secondRowSubheading=B;C0|pageFieldLabels=B;C0"
      "PivotStyleDark19", "wholeTable=F7;C7,0.799982|headerRow=F1;B;C0|totalRow=F1;B;C0|firstRowStripe=t1:7,0.399976;b1:7,0.399976|firstColumnStripe=l1:7,0.399976;r1:7,0.399976|firstSubtotalColumn=F7,-0.249977;l1:7,0.399976;r1:7,0.399976;t1:7,0.399976;b1:7,0.399976|firstSubtotalRow=B;C0|firstColumnSubheading=b1:7,0.399976|firstRowSubheading=F7,-0.249977;B;C0|secondRowSubheading=B;C0|pageFieldLabels=B;C0"
      "PivotStyleDark20", "wholeTable=F8;C8,0.799982|headerRow=F1;B;C0|totalRow=F1;B;C0|firstRowStripe=t1:8,0.399976;b1:8,0.399976|firstColumnStripe=l1:8,0.399976;r1:8,0.399976|firstSubtotalColumn=F8,-0.249977;l1:8,0.399976;r1:8,0.399976;t1:8,0.399976;b1:8,0.399976|firstSubtotalRow=B;C0|firstColumnSubheading=b1:8,0.399976|firstRowSubheading=F8,-0.249977;B;C0|secondRowSubheading=B;C0|pageFieldLabels=B;C0"
      "PivotStyleDark21", "wholeTable=F9;C9,0.799982|headerRow=F1;B;C0|totalRow=F1;B;C0|firstRowStripe=t1:9,0.399976;b1:9,0.399976|firstColumnStripe=l1:9,0.399976;r1:9,0.399976|firstSubtotalColumn=F9,-0.249977;l1:9,0.399976;r1:9,0.399976;t1:9,0.399976;b1:9,0.399976|firstSubtotalRow=B;C0|firstColumnSubheading=b1:9,0.399976|firstRowSubheading=F9,-0.249977;B;C0|secondRowSubheading=B;C0|pageFieldLabels=B;C0"
      "PivotStyleDark22", "wholeTable=F0,-0.499985;C0,-0.149998;v1:0|headerRow=F1,0.499985;b2:0|totalRow=B;C0;t2:0|firstColumn=F1,0.499985|secondRowStripe=F0,-0.449995|secondColumnStripe=F0,-0.449995|firstHeaderCell=B;C0|firstSubtotalRow=B;C0,-0.149998|secondSubtotalRow=B;C1|firstRowSubheading=B;C0|secondRowSubheading=B;C0|pageFieldLabels=F1,0.499985"
      "PivotStyleDark23", "wholeTable=F4;C4,0.799982;v1:0|headerRow=F4,-0.249977;b2:0|totalRow=B;C0;t2:0|firstColumn=F4,-0.249977|secondRowStripe=F4,0.399976|secondColumnStripe=F4,0.399976|firstHeaderCell=B;C0|firstSubtotalRow=C4,0.799982|secondSubtotalRow=B;C0|firstRowSubheading=B;C0|secondRowSubheading=B;C0|pageFieldLabels=F4,-0.249977"
      "PivotStyleDark24", "wholeTable=F5;C5,0.799982;v1:0|headerRow=F5,-0.249977;b2:0|totalRow=B;C0;t2:0|firstColumn=F5,-0.249977|secondRowStripe=F5,0.399976|secondColumnStripe=F5,0.399976|firstHeaderCell=B;C0|firstSubtotalRow=C5,0.799982|secondSubtotalRow=B;C0|firstRowSubheading=B;C0|secondRowSubheading=B;C0|pageFieldLabels=F5,-0.249977"
      "PivotStyleDark25", "wholeTable=F6;C6,0.799982;v1:0|headerRow=F6,-0.249977;b2:0|totalRow=B;C0;t2:0|firstColumn=F6,-0.249977|secondRowStripe=F6,0.399976|secondColumnStripe=F6,0.399976|firstHeaderCell=B;C0|firstSubtotalRow=C6,0.799982|secondSubtotalRow=B;C0|firstRowSubheading=B;C0|secondRowSubheading=B;C0|pageFieldLabels=F6,-0.249977"
      "PivotStyleDark26", "wholeTable=F7;C7,0.799982;v1:0|headerRow=F7,-0.249977;b2:0|totalRow=B;C0;t2:0|firstColumn=F7,-0.249977|secondRowStripe=F7,0.399976|secondColumnStripe=F7,0.399976|firstHeaderCell=B;C0|firstSubtotalRow=C7,0.799982|secondSubtotalRow=B;C0|firstRowSubheading=B;C0|secondRowSubheading=B;C0|pageFieldLabels=F7,-0.249977"
      "PivotStyleDark27", "wholeTable=F8;C8,0.799982;v1:0|headerRow=F8,-0.249977;b2:0|totalRow=B;C0;t2:0|firstColumn=F8,-0.249977|secondRowStripe=F8,0.399976|secondColumnStripe=F8,0.399976|firstHeaderCell=B;C0|firstSubtotalRow=C8,0.799982|secondSubtotalRow=B;C0|firstRowSubheading=B;C0|secondRowSubheading=B;C0|pageFieldLabels=F8,-0.249977"
      "PivotStyleDark28", "wholeTable=F9;C9,0.799982;v1:0|headerRow=F9,-0.249977;b2:0|totalRow=B;C0;t2:0|firstColumn=F9,-0.249977|secondRowStripe=F9,0.399976|secondColumnStripe=F9,0.399976|firstHeaderCell=B;C0|firstSubtotalRow=C9,0.799982|secondSubtotalRow=B;C0|firstRowSubheading=B;C0|secondRowSubheading=B;C0|pageFieldLabels=F9,-0.249977"
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
let isBuiltIn (name: string) : bool =
    presets |> Array.exists (fun (known, _) -> known = name) || pivotPresets |> Array.exists (fun (known, _) -> known = name)

/// The elements of a built-in table or pivot style in the colours of `theme` (12 colours `RRGGBB` in the
/// order of `<color theme="N">`); None for a name that is not one.
let builtIn (theme: string array) (name: string) : TableStyleElement list option =
    (if name.StartsWith "PivotStyle" then pivotPresets else presets)
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

/// A pivot table as its style sees it (0-based rows and columns): where it stands (`location`), where its
/// header ends and its data begins in it, its lines of rows and columns (`rowItems` / `colItems`: the kind -
/// "data", "default" and the other subtotals, "grand", "blank" -, how many leading fields it repeats from the
/// line before, how many items it names), its row and column fields, the page fields above it, which row fields
/// are laid out in tabular form, and the style's options as it turns them on.
type PivotShape = {
    Top: int
    Left: int
    Bottom: int
    Right: int
    FirstHeaderRow: int
    FirstDataRow: int
    FirstDataCol: int
    RowLines: (string * int * int) array
    ColumnLines: (string * int * int) array
    RowFields: int
    ColumnFields: int
    PageFields: int
    RowFieldTabular: bool array
    ShowRowHeaders: bool
    ShowColumnHeaders: bool
    ShowRowStripes: bool
    ShowColumnStripes: bool
    ShowLastColumn: bool
}

/// The order Excel lays the elements of a pivot style over each other in: a later one wins a cell's fill and font.
/// Found in Excel 2019 (tools/excel-parity/probe-pivot-style-order.vbs: a style of an own colour an element, every
/// two elements that cover one cell alone - which one shows - on compact, outline and tabular pivot tables).
let private pivotOrder =
    [| "wholeTable"; "firstColumnStripe"; "secondColumnStripe"; "pageFieldLabels"; "pageFieldValues"; "firstRowStripe"
       "secondRowStripe"; "firstColumn"; "headerRow"; "firstHeaderCell"; "firstSubtotalColumn"; "secondSubtotalColumn"
       "thirdSubtotalColumn"; "firstColumnSubheading"; "secondColumnSubheading"; "thirdColumnSubheading"; "blankRow"
       "firstSubtotalRow"; "secondSubtotalRow"; "thirdSubtotalRow"; "firstRowSubheading"; "secondRowSubheading"
       "thirdRowSubheading"; "lastColumn"; "totalRow" |]

// (Their lines are laid in the same order. Excel's lines do not follow one order: probe-pivot-style-order.vbs finds
// a row stripe's line over a row subheading's where the stripe's row is below, the subheading's where it is
// below, the header row's over both - not matched: 50 of the 1,779 cells of synthetic-pivot-styles.xlsx, an edge.)

let private levelName = [| "first"; "second"; "third" |]

/// The blocks (top, left, bottom, right) each element of a pivot style covers in a pivot table.
let pivotBlocks (shape: PivotShape) : (string * (int * int * int * int)) list =
    let r0, c0, r1, c1 = shape.Top, shape.Left, shape.Bottom, shape.Right
    let top = r0 + shape.FirstDataRow
    let fdc = shape.FirstDataCol
    let blocks = ResizeArray<string * (int * int * int * int)>()
    let add kind block = blocks.Add((kind, block))
    add "wholeTable" (r0, c0, r1, c1)
    for p in 0 .. shape.PageFields - 1 do
        let row = r0 - 1 - shape.PageFields + p
        add "wholeTable" (row, c0, row, c0 + 1)
        add "pageFieldLabels" (row, c0, row, c0)
        add "pageFieldValues" (row, c0 + 1, row, c0 + 1)
    if shape.ShowColumnHeaders && top > r0 then add "headerRow" (r0, c0, top - 1, c1)
    if shape.ShowRowHeaders && fdc > 0 then add "firstColumn" (r0, c0, r1, c0 + fdc - 1)
    if top - 2 >= r0 && fdc > 0 then add "firstHeaderCell" (r0, c0, top - 2, c0 + fdc - 1)
    let rows = shape.RowLines
    let mutable stripe = 0
    for k in 0 .. rows.Length - 1 do
        let kind, repeat, count = rows.[k]
        let row = top + k
        match kind with
        | "grand" -> add "totalRow" (row, c0, row, c1)
        | "blank" -> add "blankRow" (row, c0, row, c1)
        | "data" ->
            let level = repeat + count - 1
            for lv in repeat .. repeat + count - 1 do
                if lv < shape.RowFields - 1 && lv < 3 then
                    let tabular = lv < shape.RowFieldTabular.Length && shape.RowFieldTabular.[lv]
                    if tabular then
                        // (In tabular form an item's label stands in its field's column beside the lines of its
                        // items: the subheading is that column down those lines.)
                        let mutable last = k
                        while last + 1 < rows.Length && (let kind, repeat, _ = rows.[last + 1] in kind = "data" && repeat > lv) do
                            last <- last + 1
                        add (levelName.[lv] + "RowSubheading") (row, c0 + lv, top + last, c0 + lv)
                    elif lv = level then
                        add (levelName.[lv] + "RowSubheading") (row, c0, row, c1)
        | _ ->
            let level = repeat + count - 1
            if level >= 0 && level < 3 then add (levelName.[level] + "SubtotalRow") (row, c0, row, c1)
        if shape.ShowRowStripes then
            add (if stripe % 2 = 0 then "firstRowStripe" else "secondRowStripe") (row, c0, row, c1)
        stripe <- stripe + 1
    let columns = shape.ColumnLines
    let headerTop = r0 + shape.FirstHeaderRow
    for j in 0 .. columns.Length - 1 do
        let kind, repeat, count = columns.[j]
        let col = c0 + fdc + j
        match kind with
        | "grand" -> if shape.ShowLastColumn then add "lastColumn" (r0, col, r1, col)
        | "data" ->
            for lv in repeat .. min 2 (min (shape.ColumnFields - 1) (repeat + count - 1)) do
                add (levelName.[lv] + "ColumnSubheading") (headerTop + lv, col, headerTop + lv, col)
        | _ ->
            let level = repeat + count - 1
            if level >= 0 && level < 3 then add (levelName.[level] + "SubtotalColumn") (headerTop, col, r1, col)
        if shape.ShowColumnStripes then
            add (if j % 2 = 0 then "firstColumnStripe" else "secondColumnStripe") (top, col, r1, col)
    List.ofSeq blocks

/// What a pivot style (its elements) gives each cell of a pivot table, from the rows of its page fields to its
/// last row: the elements laid over each other in Excel's order (`pivotOrder`), each painting its fill and font
/// on the cells of its blocks and its lines on the edges of its blocks (`Top` .. `Right`) and between their
/// cells (`Horizontal`, `Vertical`) - a later element's line replaces the one on that edge. A function of a
/// cell, None outside.
let pivotLooks (elements: TableStyleElement list) (shape: PivotShape) : int -> int -> TableCellLook option =
    let firstRow = shape.Top - (if shape.PageFields > 0 then shape.PageFields + 1 else 0)
    let r0, c0, r1, c1 = firstRow, shape.Left, shape.Bottom, max shape.Right (shape.Left + 1)
    let height, width = r1 - r0 + 1, c1 - c0 + 1
    if elements.IsEmpty || height <= 0 || width <= 0 then (fun _ _ -> None)
    else
        let fills : string option array = Array.create (height * width) None
        let texts : string option array = Array.create (height * width) None
        let bolds = Array.create (height * width) false
        let italics = Array.create (height * width) false
        let covered = Array.create (height * width) false
        // (Edges: across the top of each row and one below the last; down the left of each column and one past the last.)
        // An edge between two rows gets a line from three sides: the top of a block below it, the bottom of a block above
        // it, a block across it (its lines between rows). Excel 2019 shows the top of the block below, unless the block
        // above is the header (its bottom line stays) or its bottom is the heavier line; either beats a line between a
        // block's rows
        // (probe-pivot-style-order.vbs; synthetic-pivot-styles.xlsx: a row stripe's top over the subheading above, the
        // header's bottom over the subheading below, a column subheading's bottom over the whole table's inner line).
        let lowerTops : TableLine option array = Array.create ((height + 1) * width) None
        let upperBottoms : TableLine option array = Array.create ((height + 1) * width) None
        let headerBottoms = Array.create ((height + 1) * width) false
        let insides : TableLine option array = Array.create ((height + 1) * width) None
        let down : TableLine option array = Array.create (height * (width + 1)) None
        let blocks = pivotBlocks shape
        let paint (order: string array) (withFill: bool) (withLines: bool) =
          for kind in order do
            match elements |> List.tryFind (fun element -> element.Kind = kind) with
            | None -> ()
            | Some element ->
                let format = element.Format
                for (k, (top, left, bottom, right)) in blocks do
                    if k = kind then
                        let top, bottom = max top r0, min bottom r1
                        let left, right = max left c0, min right c1
                        for row in top .. bottom do
                            for col in left .. right do
                              let at = (row - r0) * width + (col - c0)
                              if withFill then
                                covered.[at] <- true
                                if format.Fill.IsSome then fills.[at] <- format.Fill
                                if format.Text.IsSome then texts.[at] <- format.Text
                                match format.Bold with Some b -> bolds.[at] <- b | None -> ()
                                match format.Italic with Some i -> italics.[at] <- i | None -> ()
                              if withLines then
                                let lineAcross (edgeRow: int) (line: TableLine option) =
                                    if line.IsSome then
                                        let at = (edgeRow - r0) * width + (col - c0)
                                        if edgeRow = top then lowerTops.[at] <- line
                                        elif edgeRow = bottom + 1 then
                                            upperBottoms.[at] <- line
                                            headerBottoms.[at] <- (kind = "headerRow" || kind = "firstHeaderCell")
                                        else insides.[at] <- line
                                let lineDown (edgeCol: int) (line: TableLine option) =
                                    if line.IsSome then down.[(row - r0) * (width + 1) + (edgeCol - c0)] <- line
                                lineAcross row (if row = top then format.Top else format.Horizontal)
                                if row = bottom then lineAcross (row + 1) format.Bottom
                                lineDown col (if col = left then format.Left else format.Vertical)
                                if col = right then lineDown (col + 1) format.Right
        paint pivotOrder true true
        fun row col ->
            if row < r0 || row > r1 || col < c0 || col > c1 then None
            else
                let at = (row - r0) * width + (col - c0)
                let weight (line: TableLine option) =
                    match line with
                    | Some l when l.Style = "medium" -> 2
                    | Some l when l.Style = "thick" || l.Style = "double" -> 3
                    | Some _ -> 1
                    | None -> 0
                let across (edge: int) =
                    let lower, upper = lowerTops.[edge], upperBottoms.[edge]
                    if headerBottoms.[edge] && upper.IsSome then upper
                    elif weight upper > weight lower then upper
                    else lower |> Option.orElse upper |> Option.orElse insides.[edge]
                let top = across at
                let bottom = across (at + width)
                let left = down.[(row - r0) * (width + 1) + (col - c0)]
                let right = down.[(row - r0) * (width + 1) + (col - c0) + 1]
                if not covered.[at] && top.IsNone && bottom.IsNone && left.IsNone && right.IsNone then None
                else
                    Some { Fill = fills.[at]; Text = texts.[at]; Bold = bolds.[at]; Italic = italics.[at]
                           Top = top; Bottom = bottom; Left = left; Right = right }
