# Host API reference (@rbt/office spreadsheet component)

## Conventions

- Rows, columns and sheet indices are 0-based numbers. Ranges are A1 strings: `"B2"`, `"B2:D5"`, `"A:C"` (whole columns), `"3:5"` (whole rows); `$` signs are accepted and ignored.
- Objects returned by the members below use camelCase keys; an absent value is `null`. (The older event payloads - `OnChange`, `OnSelectionChange`, `GetActiveCell` - keep their PascalCase keys.)
- Invalid input throws an `Error` whose message starts with the member's name, e.g. `GetCell: row -1 is outside the sheet (0 to 999)`.
- Members act on the active sheet unless they take a sheet index or a sheet-qualified range.

### The `format` object

Used wherever a cell format is read or written (GetCell, GetSelectionFormat, and the formatting / conditional-format members):

```ts
type Border = { style: BorderStyleName; color: string /* "#RRGGBB" */ } | null;
type BorderStyleName = "thin" | "medium" | "thick" | "dashed" | "dotted" | "double" | "hair" | "mediumDashed"
                     | "dashDot" | "mediumDashDot" | "dashDotDot" | "mediumDashDotDot" | "slantDashDot";
type Format = {
  bold: boolean; italic: boolean;
  underline: boolean;          // true for a single or a double underline
  doubleUnderline: boolean;
  strikethrough: boolean;
  fontFamily: string | null;   // null: the default font
  fontSize: number | null;     // points; null: the default (11)
  textColor: string | null;    // "#RRGGBB"; null: automatic
  fill: string | null;         // "#RRGGBB"; null: no fill
  horizontalAlign: "left" | "center" | "right" | "justify" | "fill" | "centerContinuous" | "distributed" | null; // null: general
  verticalAlign: "top" | "middle" | "bottom" | "justify" | "distributed" | null;                             // null: bottom
  wrap: boolean; indent: number /* 0..250 */; shrinkToFit: boolean;
  textRotation: number;        // degrees: 1..90 counter-clockwise, -1..-90 clockwise, 0 none, 255 stacked
  numberFormat: string | null; // an Excel format code ("0.00", "#,##0", "yyyy-mm-dd"); null: General
  borders: { top: Border; bottom: Border; left: Border; right: Border; diagonal: Border; diagonalUp: boolean; diagonalDown: boolean };
};
```

Returned formats have every key. Where a member takes a format, it is a patch: only the keys present are applied, `null` resets a nullable property, colours may be `"#RGB"` / `"RRGGBB"`, `horizontalAlign: "general"` equals `null`, a border may be `{ style: "none" }`. Unknown keys and values of the wrong type throw.

## Reading cells

### `GetCell(row: number, col: number): Cell`

```ts
type Cell = {
  row: number; col: number; address: string;          // "B2"
  value: number | string | boolean | null;            // of a formula: its result; an error is its text ("#DIV/0!"); dates are serial numbers
  valueType: "empty" | "number" | "string" | "boolean" | "error";
  text: string;                                       // what the cell shows (value through its number format)
  formula: string | null;                             // with the leading "="
  format: Format;                                     // the cell's own format, else its column's default style, else the default; conditional formats are NOT applied
  isMerged: boolean; mergeRange: string | null;       // "B2:C3": the merged cell it is part of (content lives in the first cell)
  comment: { text: string; author: string | null; runs: CommentRun[] | null; thread: object | null } | null;   // see Comments
  hyperlink: { target: string; isInternal: boolean; range: string } | null;
  validation: {
    type: "list" | "whole" | "decimal" | "date" | "time" | "textLength" | "custom";
    operator: "between" | "notBetween" | "equal" | "notEqual" | "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual" | null;
    formula1: string | null; formula2: string | null;  // operands (numbers or formulas without "="); of a list, its source
    values: string[] | null;                           // entries of a list written out in the rule
    allowBlank: boolean; showError: boolean; errorStyle: "stop" | "warning" | "information";
    errorTitle: string | null; errorMessage: string | null; inputTitle: string | null; inputMessage: string | null;
  } | null;
};
```

Throws when `row` / `col` are not whole numbers inside the sheet. Read-only; no undo step, no change event.

### `GetSpillRange(row: number, col: number): { anchor: string, range: string } | null`
The dynamic array a cell of the active sheet is part of: `anchor` is the formula cell (`"B2"`), `range` what the array fills, the anchor
included (`"B2:D5"`). The same object is `GetCell(...).spill` (and each cell of `GetRange`): for the anchor **and** for every cell the array
filled (those have `formula: null`). `null` for any other cell, for a formula that gives a single value, and for an array that is blocked
(`#SPILL!`). Throws when `row` / `col` are not whole numbers inside the sheet. In the grid the array of the active cell has a thin blue
outline (`.spreadsheet-spill-outline[data-range]`, one piece per frozen pane it reaches), as in Excel.

### Editing a spilled cell: `GetCurrentCellFormula(): string`, `GetGhostFormula(): string | null`
A cell a dynamic array spilled into holds nothing of its own. As in Excel, an entry in it **starts empty** (F2, a double click,
`StartEdit()`), and `GetCurrentCellFormula()` - "what an entry in the active cell starts with" - is `""` there (it was the spilled value).
`GetGhostFormula()` is the anchor's formula for such a cell (what Excel's formula bar shows greyed) and `null` for every other cell: a
formula bar shows `GetCurrentCellFormula()`, or `GetGhostFormula()` in grey and not editable when that is not `null`. Committing an entry
in a spilled cell blocks the array (`#SPILL!` on the anchor), as before.

### `GetRange(a1: string): Cell[][]`

Rows of `Cell` for a range of the active sheet, or of another sheet when qualified (`"Data!A1:B2"`, `"'My sheet'!A:A"`; a lazily loaded sheet is converted, which is not a document change). Whole columns / rows are clipped to the sheet's used range; a block is clipped to the sheet's last row / column (a block entirely outside gives `[]`). Throws when `a1` is not a string or not a range, when the sheet does not exist, or when the range holds more than 1 000 000 cells.

### `GetSelectionFormat(): SelectionFormat`

The `Format` of the selection with the string `"mixed"` in place of every property in which the selected cells differ. `borders` is `{ top, bottom, left, right, diagonal }`, each `Border`, `null` or `"mixed"` (no diagonalUp/Down). Empty cells count with their column style / the default format. Cheap: blocks up to 20 000 cells are read cell by cell (1 ms for 1 000 cells); larger selections cost one pass over the existing cells (84 ms for a sheet of 846 000 cells); the result is cached until the cells or the selection change (the same object is returned: do not mutate it). Safe to call from an `OnSelectionChange` listener.


## Formatting

All of these act on the selection (unless they take a range), are ONE undo step, bump the revision and fire `OnChange`; a call that changes nothing adds no step. They do not recalculate (except the Clear members that remove values). A selection of up to 20 000 cells gives every cell the format (empty ones too); a larger one (whole rows / columns of a big sheet) changes only the cells that exist. All formats survive `ExportXlsx` → `ImportXlsx` (a re-imported cell without a font of its own reports the workbook's default font, e.g. `fontFamily: "Calibri", fontSize: 11`, instead of `null`).

### `ApplyFormat(command: string): void`
- toggles (by the active cell: on for all if off there, else off for all): `"bold"`, `"italic"`, `"underline"`, `"doubleUnderline"`, `"strikethrough"`, `"wrap"`, `"shrinkToFit"`
- steps: `"fontIncrease"`, `"fontDecrease"` (2 pt, 8..409), `"indentIncrease"`, `"indentDecrease"`, `"decimalIncrease"`, `"decimalDecrease"`
- alignment: `"alignLeft"`, `"alignCenter"`, `"alignRight"`, `"alignJustify"`, `"alignGeneral"`, `"alignTop"`, `"alignMiddle"`, `"alignBottom"`
- Throws `ApplyFormat: unknown command "..."` for anything else (it used to ignore unknown commands).

### Setters
| Member | Argument |
|---|---|
| `SetFontFamily(family)` | string, or `null` for the default font. Throws on an empty name. |
| `SetFontSize(points)` | number 1..409 (rounded to whole points), or `null` for the default (11). |
| `SetStrikethrough(on)` | boolean (required). |
| `SetHorizontalAlign(align)` | `"left" \| "center" \| "right" \| "justify" \| "fill" \| "centerContinuous" \| "distributed"`, or `"general"` / `null`. |
| `SetVerticalAlign(align)` | `"top" \| "middle" \| "bottom" \| "justify" \| "distributed"`, or `null`. |
| `SetWrapText(on)` | boolean. |
| `SetIndent(indent)` | whole number 0..250. `IncreaseIndent()` / `DecreaseIndent()` step by one. |
| `SetTextRotation(degrees)` | whole number -90..90 (positive = counter-clockwise), `0` none, `255` stacked. |
| `SetShrinkToFit(on)` | boolean. |
| `SetNumberFormat(code)` | an Excel format code, or `null` / `"General"`. Throws when the code is empty, over 255 characters, has an open quote / bracket, a dangling `\`, `_`, `*`, or more than four sections. (`ApplyNumberFormat(code)` remains and does not validate.) |
| `ApplyTextColor(color)`, `ApplyBgColor(color)`, `ApplyBorder(side, style, color)` | unchanged. |

Each throws `<Member>: ...` on a value of the wrong type or out of range.

### `ApplyFormatPatch(format: Partial<Format>): void`
Sets the keys present in `format` on every cell of the selection and leaves the cells' other properties alone (`null` resets a property). Throws on unknown keys / wrong values.

### `IncreaseDecimals(): void`, `DecreaseDecimals(): void`
Per cell holding a number: its number format gains / loses a decimal place in every section (`"0.0"` → `"0.00"`, `"0%"` → `"0.0%"`, `"#,##0.00;[Red]-#,##0.00"` → both sections). A General cell starts from the decimals its value shows (12.5 → `"0.00"` / `"0"`). Date / time formats, fractions, text and empty cells are left alone; up to 30 decimals. (Excel applies the active cell's new format to the whole selection; here every cell steps from its own.)

### Format painter: `GetFormatOf(a1: string): Format`, `ApplyFormatTo(a1: string, format: Format): void`
`GetFormatOf` returns the `Format` of the first cell of `a1` (same as `GetCell(...).format`). `ApplyFormatTo` REPLACES the format of every cell of the range with `format` (keys left out are defaults, unlike `ApplyFormatPatch`); values, comments and merges are untouched; the selection does not move. Ranges are of the active sheet. Throws when `a1` is not a range inside the sheet or `format` is not a format object.

### `ClearContents(a1?: string)`, `ClearFormats(a1?: string)`, `ClearAll(a1?: string)`
Act on the selection, or on the range `a1`.
- `ClearContents`: values and formulas (and rich-text runs) go; formats and comments stay; dependants are recalculated; clearing a spilling formula takes its spilled array back.
- `ClearFormats`: the cells' formats go back to the default (font, colours, alignment, number format, borders). Merged cells and conditional formats stay.
- `ClearAll`: contents, formats and cell comments (notes). Merged cells, conditional formats, validations, hyperlinks and threaded comments stay (they have their own members).
Throw when `a1` is not a range. Clearing cells that hold nothing is not a change.

### Cell styles
Not available: the model has no named cell styles (a file's named styles are carried through export untouched for cells whose format was not edited, but cannot be listed or applied).

### What the grid draws
Font family, size, bold / italic, single / double underline, strikethrough, colours, horizontal alignment, vertical alignment (top / middle / bottom; a cell without one is drawn centred in its row), wrapped text (several lines, bottom-aligned unless the format says otherwise), indent, shrink to fit, and text rotation (drawn through the centre of the cell — an approximation of Excel's placement; stacked text as upright letters).

## Conditional formats

All members act on the active sheet, except `GetConditionalFormats(sheetIndex)` which can read any sheet. Every mutating member is one undo step (no recalculation), bumps the revision and fires `OnChange`; a call that changes nothing is not a change. Rules follow inserted / deleted rows and columns (ranges and the formulas inside rules), and round-trip through `ExportXlsx` → `ImportXlsx`.

### `GetConditionalFormats(sheetIndex?: number): ConditionalFormat[]`
The sheet's rules in order of precedence:
```
ConditionalFormat = {
  id: number,          // names the rule on its sheet while the workbook is open (see "Ids")
  range: string,       // "A1:A9", "A:A", or several areas separated by spaces: "A1:A5 C1:C5"
  priority: number,    // 1 = first. Where two matching rules set the same property, the first wins
  stopIfTrue: boolean, // a match hides the rules after it
  rule: Rule
}
```
Throws on a `sheetIndex` that is not a sheet.

### The `rule` object (the same in both directions; `type` says which keys it has)
```
{ type: "cellIs", operator: "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual" | "equal" | "notEqual" | "between" | "notBetween",
  formula1: string | number, formula2: string | number | null   /* between / notBetween only */, format }
{ type: "expression", formula: string, format }                               // formatted where the formula is true
{ type: "containsText" | "notContainsText" | "beginsWith" | "endsWith", text: string, format }     // case-insensitive
{ type: "top", rank: number, percent: boolean, bottom: boolean, format }      // top / bottom N values, or N percent of the range
{ type: "aboveAverage", above: boolean, orEqual: boolean, stdDev: 0 | 1 | 2 | 3, format }          // above: false = below average
{ type: "duplicateValues" | "uniqueValues" | "blanks" | "noBlanks" | "errors" | "noErrors", format }
{ type: "dateOccurring", period: "today" | "yesterday" | "tomorrow" | "last7Days" | "thisWeek" | "lastWeek" | "nextWeek"
                                 | "thisMonth" | "lastMonth" | "nextMonth", format }               // weeks start on Sunday
{ type: "colorScale", stops: [{ type, value, color: "#RRGGBB" }, ...] }       // 2 or 3 stops, lowest first; format is null
{ type: "dataBar", min: Threshold, max: Threshold, color: "#RRGGBB" }         // format is null
{ type: "iconSet", iconSet: string, thresholds: Threshold[], reverse: boolean }   // one threshold per icon, the first is the floor; format is null

Threshold = { type: "min" | "max" | "number" | "percent" | "percentile" | "formula", value: number | string | null }
            // value: a number (0..100 for percent / percentile), the formula text for "formula", null for min / max.
            // icon-set thresholds cannot be "min" / "max".
format    = { textColor: "#RRGGBB" | null, fill: "#RRGGBB" | null, bold, italic, underline, strikethrough: boolean }
            // the part of the `format` object a conditional format can apply; other format keys are refused.
iconSet   = "3Arrows" | "3ArrowsGray" | "3Flags" | "3TrafficLights1" | "3TrafficLights2" | "3Signs" | "3Symbols" | "3Symbols2"
          | "4Arrows" | "4ArrowsGray" | "4RedToBlack" | "4Rating" | "4TrafficLights" | "5Arrows" | "5ArrowsGray" | "5Rating" | "5Quarters"
```
- Formulas (`formula`, `formula1`, `formula2`, a threshold's formula) are returned WITHOUT the leading `=`; on input a leading `=` is accepted and dropped. They are formula text: a number (`"10"`, or the number `10`), a quoted text (`"\"north\""`), a reference (`"$C$1"`) or any formula, written for the top-left cell of the rule's first range (relative references move with each cell, `$` ones stay).
- Numbers read back as strings in `formula1` / `formula2` (`"25"`); threshold values read back as numbers.
- On input, keys with a default may be left out: `percent`, `bottom`, `orEqual`, `reverse`: false; `above`: true; `stdDev`: 0; `rank`: 10; a data bar's `min` (min), `max` (max), `color` (`#638EC6`); an icon set's `iconSet` (`3TrafficLights1`) and `thresholds` (percent 0/33/67, 0/25/50/75, 0/20/40/60/80); `format`: no formatting. On output every key of the rule's type is present.
- Rules the old built-in dialog made (greater than / less than / between / equal with a literal number, text contains) read back as the `cellIs` / `containsText` rules they are equal to.
- **Excel 2010+ forms (the x14 extension of a sheet).** They are rules like the others: listed, editable, drawn, undoable, following inserted / deleted
  rows and columns, written back in the x14 form Excel writes.
  - A data bar with x14 options has, besides `min` / `max` / `color`, these keys (a plain 2007 data bar has none of them - their presence is what says
    "this is an Excel 2010+ data bar"):
    ```
    { type: "dataBar", min, max, color,
      gradient: boolean,                       // false: a solid fill
      borderColor: "#RRGGBB" | null,           // null: no border
      negativeColor: "#RRGGBB" | null,         // fill of the bars of negative values; null: the same as `color`
      negativeBorderColor: "#RRGGBB" | null,   // null: the same as `borderColor`
      axisPosition: "automatic" | "middle" | "none", axisColor: "#RRGGBB" | null,
      direction: "context" | "leftToRight" | "rightToLeft",
      showValue: boolean,                      // false: the cells show the bar only
      minLength: number, maxLength: number }   // shortest / longest bar in percent of the cell (0..100)
    ```
    Its `min` / `max` may also be `{ type: "autoMin" }` / `{ type: "autoMax" }` (Excel's "Automatic": the lowest value or 0 / the highest value or 0).
    On input, giving ANY of these keys makes the rule an Excel 2010+ data bar; the keys left out take Excel's defaults for a new data bar (`gradient:
    true`, `borderColor: null`, `negativeColor: "#FF0000"`, `negativeBorderColor: null`, `axisPosition: "automatic"`, `axisColor: "#000000"`,
    `direction: "context"`, `showValue: true`, `minLength: 0`, `maxLength: 100`); a colour key given as `null` means "none / same as positive".
    Drawn (DOM grid and painter): with `axisPosition: "automatic"` and negative values the axis is where zero falls between the two ends, negative
    bars grow leftwards from it in `negativeColor`, positive ones rightwards; `"middle"` puts the axis at the middle of the cell; `"none"` draws
    every bar from the cell's edge between `minLength` and `maxLength`; `"rightToLeft"` mirrors; a gradient fades to white away from the axis; the
    border is a 1 px line. A plain data bar is drawn as before (a solid bar from the left edge).
  - An icon set may have `showValue: boolean` and `icons: [{ iconSet, index }, ...] | null` (custom icons, one per threshold from the lowest step up;
    `index` 0 is the icon of the family's lowest step; `{ iconSet: "NoIcons", index: 0 }` shows no icon for that step). These two keys are present on
    reading only when the rule uses them or names one of the icon sets of Excel 2010+: `"3Stars"`, `"3Triangles"`, `"5Boxes"` (accepted by `iconSet`
    now). `icons` and `reverse: true` do not go together. The DOM grid draws an icon as a coloured character of its kind (arrows ↓ → ↑ ↘ ↗, triangles,
    flags, check / exclamation / cross marks, stars, part-filled circles for ratings, quarters and boxes, discs for traffic lights), with
    `data-cf-icon` (index) and `data-cf-icon-set` on its node; the painter draws arrows and triangles as shapes and the others as a square in the
    icon's colour.
  - A rule whose formula reads another sheet (`{ type: "expression", formula: "Lists!$B$1>$A1" }`, `cellIs` operands, a threshold formula) is an
    ordinary rule here; it is evaluated against the other sheet and follows edits of it. In the file it is written in the x14 list (`<x14:cfRule>`
    with `<xm:f>` and its format in place as `<x14:dxf>`), as Excel does, and read from there.
  - Still not understood: an x14 rule of an unknown type and pivot-table conditional formats (`pivot="1"`). When a sheet's x14 list has one, the whole
    list is carried through as it is (also when the sheet's other rules are edited; new cross-sheet rules are added to that list) and its rules are
    not listed; its 2007 twins (plain data bars) are. As for sparklines, a carried list is dropped when rows / columns of the sheet are inserted or
    deleted. Not kept either: the `gte="0"` ("greater than" instead of "greater than or equal") of a threshold.

### `AddConditionalFormat(range: string, rule: Rule, options?: { stopIfTrue?: boolean }): number`
Adds a rule to the cells of `range` and returns its id. The new rule takes precedence over the sheet's other rules (as in Excel); use `ReorderConditionalFormats` / `UpdateConditionalFormat(id, { priority })` to change that.
Throws `AddConditionalFormat: ...` on: a range that is not one; a missing / non-object rule; an unknown `type`; a key the type does not have; a missing operand (`formula1`, `formula2` for between, `formula`, `text`, `period`); an unknown operator, period or icon set; `rank` outside 1..1000 (1..100 for percent); `stdDev` outside 0..3; a colour scale without 2 or 3 stops; a threshold of an unknown type, without its value, with a value that is not a number or a percentage outside 0..100; an icon set with the wrong number of thresholds; a colour that is not `#RRGGBB` / `#RGB`; a `format` on a colour scale / data bar / icon set; a `format` key a rule cannot apply; an unknown `options` key.
Legacy form, still supported: `AddConditionalFormat(row, col, ruleType: "greaterThan" | "lessThan" | "equal", threshold: number, bgColor)` → id (one cell, a fill). It is an undo step now, and throws on an unknown `ruleType` (it used to fall back to "greaterThan").

### `UpdateConditionalFormat(id: number, changes: { range?, priority?, stopIfTrue?, rule?, format? }): void`
- `range`: new range(s). `stopIfTrue`: boolean.
- `priority`: the rule's new place in the order of precedence (1..number of rules); the other rules keep their relative order.
- `rule`: a whole `Rule`; it replaces the rule. Without a `format` key the rule keeps the format it has.
- `format`: a patch of the rule's format (only the keys given change). Refused for colour scales, data bars and icon sets.
Throws on an unknown id, an unknown key of `changes`, and everything `AddConditionalFormat` refuses.

### `RemoveConditionalFormat(id: number): void`
Throws when the sheet has no rule with that id.

### `ReorderConditionalFormats(ids: number[]): void`
`ids` lists EVERY rule of the active sheet exactly once, the one that takes precedence first. Throws on an unknown id, a repeated id, or a rule left out (the message lists the sheet's ids).

### `ClearConditionalFormats(range?: string): number`
Without an argument: removes every rule of the active sheet. With a range (several areas allowed): takes the rules off those cells: a rule covering more cells keeps the rest of its range (as up to four areas per hole), a rule left with no cell is removed. Returns how many rules were removed or cut (0: nothing changed).

### `RemoveConditionalFormats(row: number, col: number): void` (existing)
The same for one cell. Now an undo step.

### `ShowConditionalFormattingDialog(): void` (existing, unchanged)
The built-in dialog offers five kinds (greater than, less than, between, equal to, text contains) on the active cell only. It was left as it is: hosts should build their own rule manager on the members above. Rules made in the dialog are undo steps now and show up in `GetConditionalFormats`.

### Ids
An id is unique on its sheet and stays with its rule through edits of the rule and of other rules, undo / redo, and inserted / deleted rows and columns; the component does not hand the id of a removed rule to a new one. A file stores no ids: after `ImportXlsx` each sheet's rules are numbered 1, 2, 3 ... in order of precedence (the same file gives the same ids), so a host re-reads `GetConditionalFormats` after opening a file. Ids of different sheets may coincide.

## Data validation

Validations belong to a sheet and are stored by range: a whole column (`"A:A"`) is one validation, no cell is made for it, and it is written to
XLSX as `A1:A1048576`. A cell is covered by at most one validation. All mutating members act on the active sheet, are one undo step each (no
recalculation), bump the revision and fire `OnChange`. Validations follow inserted / deleted rows and columns (ranges move, the formulas in
the rules are re-pointed) and round-trip through `ExportXlsx` / `ImportXlsx`.
A validation whose formula or list source reads another sheet (`{ type: "list", source: "Lists!$A$1:$A$9" }`, `formula1: "Limits!$B$2"`) is read from
and written to the sheet's x14 list (`<x14:dataValidations>` with `<xm:f>` / `<xm:sqref>`), the way Excel writes it; in this API it is a validation
like any other. Because that list comes after the sheet's other validations in the file, the ids of a workbook opened again are numbered in that
order (the validations that read another sheet last); ids are never stable across a save and a reopen anyway.

### `GetValidations(sheetIndex?: number): { id: number, range: string, rule: ValidationRule }[]`
The validations of the active sheet (or of the sheet at `sheetIndex`; a lazily loaded sheet is converted), in the order they were made.
`range`: the areas it covers, separated by spaces (`"B2:B9 D2:D9"`, `"A:A"`). Throws `GetValidations: there is no sheet N ...`.

### The `ValidationRule` object (the same in both directions; `type` says which keys count, the others read back as `null`)
```
{ type: "any" }                                   every value is allowed (carries an input message only)
{ type: "list", values: string[], dropdown? }      the allowed values written out (no commas inside a value)
{ type: "list", source: string, dropdown? }        read from cells when needed: "$A$1:$A$9", "Lists!$B$2:$B$20", "$H:$H" or a defined name
{ type: "whole" | "decimal" | "date" | "time" | "textLength",
  operator: "between" | "notBetween"                                   (formula1 and formula2)
          | "equal" | "notEqual" | "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual",   (formula1)
  formula1, formula2 }
{ type: "custom", formula }                        valid where the formula is true
```
plus, for every type (defaults in brackets):
`allowBlank` [true], `showInput` [true], `inputTitle`, `inputMessage` (shown beside the cell while it is selected),
`showError` [true] (false: the rule is advisory, nothing is refused), `errorStyle`: `"stop"` [default; the entry is refused] | `"warning"` |
`"information"` (the entry is accepted and the user told), `errorTitle`, `errorMessage`; for lists `dropdown` [true] (the in-cell dropdown button).

- `formula1` / `formula2` / `formula` / `source` are texts without the leading `=` (one is accepted and dropped): a number (`"10"`), a
  reference (`"$K$1"`) or any formula, written for the top-left cell of the validation's first range; relative references move with the cell.
- On input an operand may be a number; a `"date"` operand may be `"2024-03-31"`, `"2024-03-31T14:30"` or a JS `Date`, a `"time"` operand `"14:30"` /
  `"14:30:15"`. They read back as serial numbers (`"45292"`, `"0.5"`). `value1` / `value2` are accepted as aliases of `formula1` / `formula2`.
- `between` with two plain numbers, `equal`, `greaterThanOrEqual` and `lessThanOrEqual` with a plain number read back exactly as given.
- List matching is case-insensitive; an empty entry is always accepted; a list whose `source` cannot be read refuses nothing.

### `GetValidationAt(row: number, col: number): (ValidationRule & { id: number, range: string }) | null`
The validation on a cell of the active sheet. The object can be handed to `SetValidation` for other cells (`id` and `range` are ignored there).
`GetCell(row, col).validation` is this same object. Throws when row / col are outside the sheet.

### `SetValidation(range: string, rule: ValidationRule): number`
Puts a validation on the cells of `range` (`"B2:B20"`, `"A:A"`, several areas `"A1:A5 C1:C5"`), replacing whatever validation those cells had (a
validation that covered more cells keeps the rest). Returns the new validation's id. Existing values are not checked (see `CircleInvalid`).
Throws (message starts `SetValidation:`) for a missing / invalid range, a missing rule, an unknown `type`, `operator`, `errorStyle` or key, a key of
the wrong type, a missing operand (`formula1`, `formula2` for between / notBetween), a custom rule without `formula`, a list without `values` or
`source` (or with both), and a list value containing a comma.

### `ClearValidation(range: string): number`
Takes the validations off the cells of `range`; a validation that covers more keeps the rest, one left with no cell is removed. Returns how many
validations were removed or cut (0: nothing changed, no undo step). Throws for a missing / invalid range.

### `GetValidationList(row: number, col: number): string[] | null`
The values the list validation on a cell offers, in order (a `source` range is read now, from another sheet if need be); `null` when the cell has no
list validation or the source cannot be read. For hosts that draw their own dropdown.

The `source` may be a range (`$J$1:$J$4`, `Lists!$A$1:$A$9`), a defined name, or **a formula that gives cells or an array**: `OFFSET(...)`,
`Lists!$A$1:INDEX(Lists!$A:$A,COUNTA(Lists!$A:$A))`, `INDIRECT(...)`, or a defined name whose refers-to is one of these (a dynamic named
range) or an array constant. A range and a name that is a range are read off the cells; anything else is evaluated by the formula engine
with the workbook in reach (every sheet is converted for it) at the cell asked for - relative references move with the cell as in any
rule - and the filled cells / items it gives are the list, blanks and errors left out, numbers as their plain text. The grid's own dropdown,
entry checking (`SetCell`, typing) and `CircleInvalid` use the same list. The rule is written back as it was read (`<formula1>Regions</formula1>`);
nothing is expanded into the file. (The engine now also reads a range operator after a sheet-qualified cell: `Sheet!A1:INDEX(...)`,
`Sheet!A1:Sheet!B2`.)

### `ShowValidationList(): boolean`
Opens the built-in dropdown of the active cell's list validation (what the button beside the cell and Alt+Down do); picking a value enters it
like typed input (validated, one undo step). `false` when the cell has no list or the grid is read-only. The popup is `#validation-list-dropdown`
with `.validation-list-item` children; the button is `.spreadsheet-list-button` (shown beside the active cell when its list has `dropdown: true`).

### `CircleInvalid(draw?: boolean): { row, col, address, value, text, errorStyle, errorTitle, errorMessage }[]`
The filled cells of the active sheet whose current value breaks the validation on them (Excel's "Circle Invalid Data"), in row order. Rules with
`showError: false` count too; empty cells never do; only existing cells are looked at (a whole-column rule costs what the column holds). Unless
`draw` is `false`, the grid draws a red ellipse (`.spreadsheet-invalid-circle`, up to 2000) around each until `ClearInvalidCircles()`, the next
`CircleInvalid()`, a change of the sheet's validations through the API, or a new workbook. The circles are view state: no undo step, no revision, not saved, and they
do not follow later edits of the cells.

### `ClearInvalidCircles(): void`
Removes the circles (on every sheet).

### Entering values
`EnterText`, `CommitEdit`, `SetCurrentCellFromFormula` and typing validate (`false` when a "stop" rule refused the entry; `OnValidationNotice(style,
title, message)` and `GetValidationNotice()` report it). `SetCell`, paste, fill and `ImportCSV` do not validate.

### Ids
A validation's `id` is stable for the life of the workbook in the component (through undo / redo and edits of other validations; ids are not
reused). They are not stored in the file: at each import a sheet's validations are numbered 1, 2, ... in file order.

## Defined names

Names live in the workbook, not in a sheet. A change of them is a document change: one undo step (on the active sheet's undo stack), a new
revision, and `OnChange` fires with the new key `WorkbookChanged: true` (its `Sheets` lists the sheets whose values changed with it).
Print areas and print titles (`_xlnm.Print_Area` / `_xlnm.Print_Titles`) are not names here: they belong to the page setup.

### `GetDefinedNames(): DefinedName[]`
`{ name: string, refersTo: string, scope: number | null, scopeName: string | null, comment: string | null, isBuiltIn: boolean, hidden: boolean }`
- `name` as it was written (lookups ignore case); `refersTo` the formula without a leading `=` (`"Data!$A$1:$A$3"`, `"0.19"`).
- `scope`: 0-based index of the sheet a sheet-level name belongs to, `null` for a workbook-level name; `scopeName` that sheet's name.
- `isBuiltIn`: an `_xlnm.` name; `hidden`: the file hides it from the name manager (kept through a save).
Workbook-level names first, then sheet-level ones by sheet, each group by name (upper-cased order).

### `SetDefinedName(name: string, refersTo: string, scopeSheetIndex?: number | null, comment?: string | null): void`
Defines a name or replaces the one of the same scope (names differ by more than case; the new spelling and comment replace the old ones).
`refersTo` is a reference or formula with or without `=`; a reference without a sheet is read on the sheet of the formula that uses the name.
Every formula is recalculated. One undo step. Defining a name exactly as it is does nothing (no step).
Throws: name empty / longer than 255 / not starting with a letter, `_` or `\` / holding anything but letters, digits, `_ . \ ?` / reading as a
cell reference (`A1`, `R1C1`, `R`, `C`) or `TRUE` / `FALSE`; a built-in `_xlnm.` name; the name of a table; `refersTo` empty or not a formula;
a sheet index that does not exist; arguments of the wrong type.

### `DeleteDefinedName(name: string, scope?: number | null): void`
Deletes the workbook-level name, or with `scope` that sheet's. Formulas using it show `#NAME?` (or fall back to the workbook-level name when a
sheet-level one is deleted). One undo step. Throws when there is no such name in that scope, for a built-in name, or a sheet index that does not exist.

### `GoTo(nameOrRef: string): { sheetIndex: number, range: string }`
The name box: a reference (`"B3"`, `"$B$2:$D$5"`, `"A:A"`, `"Sheet2!B3:C9"`, `"'My sheet'!A1"`), a defined name that refers to a range (the
active sheet's own name before a workbook-level one; `"Sheet2!Name"` for another sheet's; of several areas the first), or a table name.
Switches the sheet if needed, selects the range (clipped to the sheet's rows and columns), scrolls its first cell into view; an entry in
progress is committed. Not a document change. Throws when the text is not a reference, name or table, names a sheet that does not exist, lies
outside the sheet, or is a name that does not refer to a range (a constant or formula).

## Freeze panes and view

All of these act on the active sheet; each sheet has its own panes, zoom and view options, saved with it (XLSX `sheetView`).

### `FreezeAt(row: number, col: number): void`
Freezes the rows above `row` and the columns left of `col` (0-based), like Excel's "Freeze Panes" with that cell selected:
`FreezeAt(1, 0)` = first row, `FreezeAt(0, 1)` = first column, `FreezeAt(2, 3)` = two rows and three columns; `FreezeAt(0, 0)` unfreezes.
Panes are always counted from row 0 / column 0 (not from where the grid is scrolled to). One undo step. Throws when the cell is not on the sheet.

### `Unfreeze(): void`
Removes the panes. One undo step (nothing when none are frozen).

### `GetFrozen(): { rows: number, cols: number }`

### `SetZoom(factor: number): void` / `GetZoom(): number`
`factor` 0.1 .. 4 (1 = 100 %), rounded to a whole percent. Cells, headers, selection, pictures, the edit box and the validation button scale;
hit-testing, `ScrollToCell`, resize drags and the scroll extent follow; the cell in the top-left corner stays there. Text is laid out at the
zoomed size (CSS `zoom` on the grid's content, not a bitmap scale). **Zooming is not an edit** (as in Excel): no undo step, no new revision,
no `OnChange`, `IsModified()` unchanged; the zoom is nevertheless written by `ExportXlsx` (`zoomScale`) and read by `ImportXlsx`.
A host zoom control should re-read `GetZoom()` on `OnSelectionChange` when `SheetIndex` changes (each sheet has its own zoom).
Throws when `factor` is not a number in the range.

### `SetShowGridLines(show: boolean)`, `SetShowHeaders(show: boolean)`, `SetShowFormulas(show: boolean)`, `SetShowZeros(show: boolean)`
Each is one undo step, a new revision and an `OnChange`; each throws when `show` is not a boolean.
- grid lines: the light cell grid (borders stay).
- headers: row numbers and column letters; without them cells start at the grid's top-left corner (and there are no header resize handles,
  filter buttons or header clicks).
- formulas: a formula cell shows its text (`=A1*2`) instead of its result (Excel's Ctrl+`). `GetCell().text` / clipboard text still give results.
- zeros: with `false` a cell holding the number 0 (typed or calculated) shows nothing.
### `GetShowGridLines()`, `GetShowHeaders()`, `GetShowFormulas()`, `GetShowZeros()`: boolean
### `GetViewOptions(): { zoom, showGridLines, showHeaders, showFormulas, showZeros, frozen: { rows, cols } }`

## Sheets (group 7)

Sheets are addressed by 0-based index in tab order. Every mutating member below is ONE undo step, gives a new revision, and is
reported by `OnChange` (with `StructureChanged: true`) and `OnSheetsChange`. All throw `Error("Member: reason")` on invalid input.

**Where the undo step lives.** Undo histories are per sheet (`Undo()` / `CanUndo()` act on the sheet on display). A sheet-list step is put
on the history of the sheet that is on display once the operation is done: the new sheet after `AddSheet` / `DuplicateSheet`, the sheet that
took a deleted (or hidden) sheet's place, otherwise the sheet that was on display. So `Undo()` right after any of these calls takes it back.
When undoing / redoing changes which sheet is on display (undoing `AddSheet` removes the sheet on display), the step moves to the history of
the sheet then on display, so `Redo()` finds it there. After the user switches to another sheet, `Undo()` acts on that sheet's own history.

| Member | Arguments | Returns | Notes |
|---|---|---|---|
| `GetSheetInfos()` | | `[{ index, id, name, visibility: "visible"\|"hidden"\|"veryHidden", tabColor: "#RRGGBB"\|null, isLoaded, isActive, kind }]` | `id` stays with the sheet through renames, moves and other sheets being added / deleted (new sheets get an id never used before). `isLoaded`: see `IsSheetLoaded`. `kind`: see `GetSheetKind`. A tab colour a file gives as a theme colour is reported as the Office theme's colour without its tint; an indexed colour as null (both are written back as the file had them). |
| `GetSheetKind(index?)` | default: the sheet on display | `"worksheet"` \| `"chartsheet"` \| `"dialogsheet"` | What the opened file has the sheet as (every other sheet is a worksheet). A chart sheet or dialog sheet has no cells: it is shown as an empty sheet (a chart sheet with the frame of its chart) and written back as the part it was read from. **It takes no entry**: while it is on display the grid behaves as in viewer mode (no editing, filling, pasting, resizing; `IsReadOnly()` still reports only what the host set), `EnterText` returns false and `SetCell` changes nothing there. Before, an entry was accepted, shown, and lost at the save. |
| `OnSheetsChange(cb)` | `cb(infos)` gets the `GetSheetInfos()` array | unsubscribe function | Fires after the list changed: a sheet added, deleted, renamed, moved, copied, hidden / shown, tab colour; by a call, by undo / redo, or by `InstallWorkbook` / `ImportXlsx`. Not for a mere switch of the sheet on display (`OnSelectionChange` / `OnSheetChanged`). Fires while the grid is mounted. |
| `AddSheet(name?, atIndex?)` | `name` string (default: first free "Sheet1", "Sheet2", ...); `atIndex` 0..count (default: at the end) | index of the new sheet | The new sheet (1000 x 26) is shown. Throws on an invalid / duplicate name or a position outside 0..count. Formulas that named a sheet of that name before it existed are recalculated. |
| `RenameSheet(index, name)` | | | Formulas of every sheet (cells, conditional formats, validations incl. list sources), defined names and internal hyperlinks follow, quoted where needed (`'My Sheet'!A1`, `'It''s'!A1`). Throws: empty, > 31 chars, any of `: \ / ? * [ ]`, leading / trailing apostrophe, "History", a duplicate (case-insensitive; changing only the case of the sheet's own name is allowed), no such sheet. Values are not recalculated (nothing changes) - see limits. |
| `DeleteSheet(index)` | | | References to the sheet in other sheets' formulas, rules and defined names become `#REF!`; names scoped to it and its tables go with it; what the file kept for it (pictures, charts, pivot tables, comments) is dropped from the exported package; internal hyperlinks of other sheets into it get the target `#REF!`. If it was on display, the nearest visible sheet is shown. Undo restores the sheet, the names, the tables and the references (a formula edited in between keeps the edit). Throws when it is the only visible sheet, or no such sheet. |
| `MoveSheet(from, to)` | `to` = index the sheet has afterwards (0..count-1) | | The sheet on display stays on display. |
| `DuplicateSheet(index)` | | index of the copy | The copy is placed right after the original, named "Name (2)" ("(3)", ...; cut to 31 chars) and shown. Copied: cells, formats, sizes, merges, comments, validations, conditional formats, pictures, page setup, sheet-scoped names, tables (as "Table_2"; the copy's structured references follow). The copy's formulas that named the original sheet name the copy. Also copied (since 0.1.9): the drawing objects the editor only keeps from the file - charts, shapes; `GetDrawings()` lists them on the copy, and `ExportXlsx` writes each copied chart as a chart part of its own (new part name, its own relationships and content type) whose series read the copy where the original's read the original. NOT copied: **pivot tables** (the copy has their cells as plain values; the pivot table, its cache and slicers stay with the original) and named cell styles. |
| `SetSheetVisibility(index, visibility)` | `"visible"` \| `"hidden"` \| `"veryHidden"` | | Hiding the sheet on display shows the nearest visible one. Throws when it would leave no visible sheet. `SheetVisibility()` (bool per sheet) keeps working. |
| `SetSheetTabColor(index, color)` | `"#RRGGBB"` or null (none) | | The component draws no tab strip; the colour is workbook content (GetSheetInfos, the saved file's `<tabColor>`). |

Existing members that keep working: `GetSheetNames()`, `SheetNames()`, `SheetVisibility()`, `GetActiveSheetIndex()`, `SwitchSheet(index)`,
`OnSheetChanged(cb(index))` (now also fires when a sheet-list operation or its undo changes the sheet on display), `IsSheetLoaded(index)`.

Lazy sheets: `MoveSheet`, `SetSheetVisibility`, `SetSheetTabColor` convert no sheet. `RenameSheet` / `DeleteSheet` / `AddSheet` convert the sheets
that have formulas or internal hyperlinks (as any first edit does); plain-data sheets stay unconverted. `DuplicateSheet` converts the sheet it copies.

Limits: a rename does not recalculate (a formula that builds the sheet's name as text - `INDIRECT("Old!A1")` - keeps its value until the
next recalculation and then no longer finds the sheet, as in Excel).

Kept parts follow the sheets (since 0.1.9; applied when the workbook is written, `ExportXlsx`): after a rename the series of the file's charts
(`<c:f>`, `<cx:f>`: `Sales!$B$2:$B$9` -> `'Net sales'!$B$2:$B$9`, also where they are qualified defined names), the sheet of a pivot cache's
`<worksheetSource sheet>` and the formulas of kept extensions (`<xm:f>` of sparklines) name the sheet by its new name; references in chart
series to a sheet that was deleted become `#REF!`. Internal hyperlinks follow a rename also on sheets that were not converted yet, and a link
into a deleted sheet has the target `#REF!` (`GetHyperlink(...).target`, written as `location="#REF!"`; undo restores it). Parts nothing was
renamed in go back byte for byte. Not followed: slicer caches and external-link parts (they do not name sheets of the workbook), and text
inside chart titles.

## Rows and columns (group 8)

Sizes are **sheet pixels at 100 % zoom** (a new sheet: columns 100, rows 25; an opened file: what the file says). Where a member takes
"`target`" it takes a 0-based row / column **number** or an **A1 string** whose rows / columns are meant: `"2:5"`, `"B:D"`, or any range
(`"B2:D9"` = rows 2..9 or columns B..D); ranges are clipped to the sheet. All members act on the active sheet. Every mutating member is one
undo step, a new revision and one `OnChange`; a call that changes nothing is none of these. All of it follows inserted / deleted rows and
columns and is written to / read from XLSX (`<row ht hidden outlineLevel collapsed>`, `<col width hidden outlineLevel collapsed>`).

| Member | Arguments | Returns | Notes |
|---|---|---|---|
| `SetColumnWidth(target, px)` | `px` number 0..1785 | | 0 hides (as `HideColumns`); any other width also shows a hidden column. Throws on a column off the sheet, a bad range, a width outside the limits. |
| `SetRowHeight(target, px)` | `px` number 0..546 | | 0 hides. Throws likewise. |
| `GetColumnWidth(col)` | 0-based | number | 0 when hidden. Throws when the column is not on the sheet. |
| `GetRowHeight(row)` | 0-based | number | 0 when hidden (by `HideRows`, a collapsed group, or a filter). |
| `AutoFitColumns(target?)` | default: the columns of the selection | | Each column gets the width of its widest cell: the text the cell displays (number format applied), measured in the cell's font (family, size, bold, italic), plus indent and the cell's 10 px padding; at most 1785. A wrapped cell counts with its longest line (as broken by line breaks in the text); text turned 90 degrees counts a line high; cells merged across columns do not count; hidden rows are left out; hidden columns stay hidden; a column without text keeps its width. |
| `AutoFitRows(target?)` | default: the rows of the selection | | Each row gets the height of its tallest cell: lines x line height (round(font px x 1.3)) + 6; wrapped text takes the lines it needs at the width of its column (or of its merged columns); at least 20 (one line of the default font), at most 546; a row without text gets 20. Cells merged across rows do not count; hidden rows stay hidden. |
| `HideRows(target, count?)` | number (+ `count` rows from it, default 1) or A1 range | | A hidden row keeps its height for `UnhideRows`. |
| `UnhideRows(target, count?)` | | | Shows the hidden rows among those named at the height they had. Rows hidden by a filter are the filter's (not shown by this). |
| `HideColumns(target, count?)` / `UnhideColumns(target, count?)` | | | Likewise. |
| `GetHiddenRows()` | | `string[]`: `["3:5", "9:9"]` | Runs of hidden rows in order; rows hidden by a filter are not listed. |
| `GetHiddenColumns()` | | `string[]`: `["B:D"]` | |
| `GroupRows(target, count?)` / `GroupColumns(target, count?)` | | | Excel's Data > Group: the outline level of each row / column named goes up by one (max 7). |
| `UngroupRows(target, count?)` / `UngroupColumns(target, count?)` | | | The level goes down by one (rows in no group are left alone). |
| `SetGroupCollapsed(axis, index, collapsed)` | `axis` `"rows"` \| `"columns"`; `index` 0-based row / column inside the group; `collapsed` boolean | | Acts on the group the row / column is in at its own level (its neighbours at that level or deeper). Collapsing hides the group's rows; expanding shows them, except those of groups inside it that are collapsed themselves. The row / column after the group is its summary and carries the collapsed mark (Excel's default "summary below / right"). Throws when the row / column is in no group, or `axis` is not one of the two. |
| `ShowOutlineLevel(axis, level)` | `axis` `"rows"` \| `"columns"`; `level` 1..8 | | Excel's level buttons: shows the rows / columns of the levels below `level`, collapses every group from `level` on (1 = only what is in no group, deepest + 1 = everything). Rows hidden by hand inside the groups are shown too. |
| `SetOutlineSummary(below, right)` | two booleans | | Where a group's summary row / column is (the sheet's `<outlinePr summaryBelow summaryRight>`): below / right (`true`, Excel's default) or above / left. The +/- button and the `collapsed` mark of a group are on its summary row / column. Collapsed groups stay collapsed. |
| `SetShowOutline(show)` / `GetShowOutline()` | boolean | boolean | Whether the grid draws the outline strips (default `true`). A setting of the component for every sheet: not saved, no undo step, no revision. |
| `SetShowOutlineSymbols(show)` / `GetShowOutlineSymbols()` | boolean | boolean | The active **sheet's** own switch for its outline symbols (the sheet view's `showOutlineSymbols`; Excel's "Show outline symbols if an outline is applied"): read from and written to the file, one undo step, a new revision. `false` hides the strips of that sheet; its groups, hidden rows and `GetOutline()` stay. |
| `GetOutline()` | | `{ rows: [group], columns: [group], rowLevels, columnLevels, summaryBelow, summaryRight, showSymbols }` with `group = { first, last, range: "3:7" \| "B:D", level: 1..7, collapsed }` | Outer groups come before the groups inside them; `rowLevels` / `columnLevels` are the deepest level (0 = no groups). |

In the grid: dragging a header's resize handle is now one undo step (it was none); a **double click on a resize handle** fits that column /
row (`AutoFitColumns` / `AutoFitRows`), and when whole columns / rows are selected and the handle is one of theirs, all of them. Arrow keys,
selection, drawing anchors, the portable painter and the PDF path treat hidden rows / columns as before (size 0).

**The outline strips.** A sheet with grouped rows gets a strip left of the row numbers, one with grouped columns a strip above the column
letters, each `(deepest level + 1) x 15 + 3` px wide (48 px for two levels): a bracket per group (`.spreadsheet-outline-line`), a +/- button
on the group's summary row / column (`.spreadsheet-outline-button` with `data-axis="rows|columns"`, `data-first`, `data-last`, `data-level`,
`data-collapsed`), and the level buttons 1, 2, 3 ... in the corner (`.spreadsheet-outline-level` with `data-axis`, `data-level`). A click on
+/- is `SetGroupCollapsed`, on a level button `ShowOutlineLevel` (one undo step each); it selects nothing. **The cells and headers move right /
down by the strip**: the header area is 50 x 30 px only on a sheet without groups (read it from `.spreadsheet-corner`'s size, or use
`GetCellRect`-style members rather than fixed offsets). The strips scale with the zoom, stay in place while the sheet scrolls, and go with
`SetShowOutline(false)` (the component, every sheet) or `SetShowOutlineSymbols(false)` (the sheet, saved). **A sheet without headers
(`SetShowHeaders(false)`) keeps its strips**: the row numbers and column letters go, and a band of 18 px stays beside a strip for the level
buttons of the other axis (so the header area is e.g. 66 x 51 px with two row levels and one column level, 48 x 18 px with row groups only).
**In a viewer (`SetReadOnly(true)`) the buttons work**: +/- and the level buttons collapse and expand as in the editor, but as view state -
no undo step, no new revision, no `OnChange` call, and an export still gives the file as it was opened. (The API members `SetGroupCollapsed` /
`ShowOutlineLevel` are edits as before.) A group whose summary row / column is hidden or off the sheet (a group at row 1 with the summary above) has no button:
use the level buttons or `SetGroupCollapsed`. At most 1500 strip elements are drawn for the rows / columns in view.

`<sheetPr><outlinePr showOutlineSymbols>` (a second place a file can say it) is read too: the symbols are shown only when neither it nor the sheet view's flag says otherwise; a save writes the sheet view's flag. Limits: `applyStyles` is not kept.
Auto-fit looks only at the cells that exist; on a very large sheet a whole column costs a
pass over the sheet's cells (dashboard "Data", 846 000 cells: one column 0.46 s, eleven columns 3.4 s, all 76 931 rows 0.9 s).
Auto-fit does not account for conditional-format icons / data bars, filter buttons, or rotation other than 90 degrees.
A row that was hidden in an opened file without a height of its own is written with the file's default height once the sheet's rows change.

## Sort and filter

All act on the active sheet (except `GetFilters(sheetIndex)`); each mutating member is one undo step, a new revision and one `OnChange`.
A **column** argument is a 0-based sheet column number or its letters (`2` or `"C"`).

| Member | Arguments | Returns | Notes |
|---|---|---|---|
| `SortRange(range, keys, options?)` | `range` A1 (`"A1:D20"`; `"A:D"` is clipped to the used rows). `keys`: `[{ column, descending?: boolean, by?: "value" \| "cellColor" \| "fontColor", color?: "#RRGGBB" \| null }]` (one key object or a bare column is accepted too). `options`: `{ hasHeader?: boolean, caseSensitive?: boolean }` (both false by default) | | See "Sort semantics". Throws on a bad range / key / option key, a key column outside the range, merged cells among the rows to sort, a dynamic array (spilling formula or spilled cells) among them. |
| `SortCurrentColumnAsc()` / `SortCurrentColumnDesc()` | | | Quick sort: sorts the block of data around the **active cell** (Excel's current region, see "Quick sort") by the active cell's column, as one undo step. Only the cells of that block move. Nothing happens when there is no data around the cell. Throws like `SortRange` for merged cells or a dynamic array in the block. |
| `SetAutoFilter(range)` | A1 range; its first row is the header row | | Gives the sheet its AutoFilter (one per sheet), replacing the one it has; column filters inside the new range stay, others go. Throws on a bad range or one outside the sheet. |
| `ClearAutoFilter()` | | | Removes the AutoFilter and its column filters; every row it hid is shown. |
| `SetFilter(column, criteria)` | see "Filter criteria" | | Replaces the column's filter; rows failing it are hidden (rows failing any column's filter are). Without an AutoFilter the sheet gets an implicit one over the block of data around the column (it goes away with its last column filter). Throws on a column outside the AutoFilter's range and on criteria that are not understood. |
| `ClearFilter(column?)` | | | Takes the filter off one column, or off all; the AutoFilter (range) stays. |
| `ReapplyFilter()` | | | Works the hidden rows out again from the cells as they are now. An undo step only when the hidden rows change. |
| `GetFilters(sheetIndex?)` | | `null` or `{ range: "A1:D20", columns: [{ column, letter, criteria }], hiddenRows: number[], hiddenRowCount, visibleRowCount }` | `hiddenRows` are 0-based rows the filter hides; `visibleRowCount` counts the shown rows below the header (down to the last used row of the range). Throws on a sheet index that does not exist. |
| `GetFilterValues(column)` | | `{ values: string[], blanks: boolean }` | The distinct texts the column's cells show in the filter's rows (number format applied; numbers first, in order, then text A-Z) and whether it has blank cells there: what a checklist offers. Without an AutoFilter: the block of data around the column, below its first row. |

### Filter buttons: `OnFilterButton(cb): () => void`, `GetFilterButtonRects(): FilterButton[]`
`FilterButton = { column, row, sheetIndex, clientRect: { left, top, width, height }, active: boolean, sorted: "asc" | "desc" | null }`: `column` / `row` 0-based (the header cell), `clientRect` the button in viewport (client) pixels as `getBoundingClientRect` gives them (zoom included: 16 x 16 at 100 %, 24 x 24 at 150 %), `active` = the column has a filter, `sorted` = the column is the sheet's sorted column.
- `OnFilterButton(cb)`: `cb(button)` is called when a button is pressed with the primary mouse button, and when **Alt+Down** is typed with the active cell on a header cell that has a button. Returns the unsubscribe function. **While at least one subscriber is registered** the grid's own checklist is not opened (the host shows its menu, anchored at `clientRect`, and calls `SetFilter` / `ClearFilter` / `SortRange`), the buttons beside the column letters are hidden, and a viewer draws the buttons too. The press does not select the cell under the button. A callback that throws is reported through `OnError` (source `"OnFilterButton"`).
- `GetFilterButtonRects()`: the buttons that can be seen now, in the order header cells left to right, the AutoFilter's before the tables': not those scrolled out of the grid or under a frozen pane, not those of hidden columns. Call it after a change (it reads the DOM the grid has drawn; the grid draws synchronously with every change and on scroll). Empty when the grid is not mounted, and in a viewer without a subscriber.
- A host that draws buttons of its own over the grid should hide the built-in ones (`.spreadsheet-filter-button { visibility: hidden }` keeps the rectangles; `display: none` makes them empty), or better drop its own and subscribe.
- A table's own filter is still the sheet's (one filter per sheet): a table button is `active` when the sheet's AutoFilter covers the table's rows and filters that column. The built-in checklist on a table button without a sheet AutoFilter filters the block of data around the column.

### Filtering in a viewer (`SetReadOnly(true)`)
A viewer draws **no filter buttons of its own** (its own filter menu edits). When a host subscribes with `OnFilterButton` the buttons are
drawn and pressed as in the editor (`GetFilterButtonRects()` lists them), and the host answers a press with `SetFilter` / `ClearFilter` /
`ReapplyFilter`. In a viewer those calls are **view state**: the rows are hidden or shown and `SUBTOTAL` / `AGGREGATE` follow, but there is
no undo step, no new revision and no `OnChange` call, and `ExportXlsx()` still gives the file as it was opened (the filter is not saved).
In the editor the same calls are edits, as before. `SortRange` and the quick sorts change cells and stay refused / edits.

### Filter criteria (given to `SetFilter`, returned by `GetFilters`)
One of:
- `{ values: string[], blanks?: boolean }`: the cells showing one of the texts (compared without case; the cell's displayed text, or its unformatted value), plus the blank cells when `blanks` (an `""` among the values means the same).
- `{ conditions: [{ operator, value }], and?: boolean }`: one or two comparisons, joined by AND (default) or OR. `operator`: `"equals"`, `"notEquals"`, `"greaterThan"`, `"greaterOrEqual"`, `"lessThan"`, `"lessOrEqual"`, `"contains"`, `"notContains"`, `"beginsWith"`, `"endsWith"`. `value`: a number, a text, a `Date`, or a date written `"2024-03-31"` (optionally with a time; it is compared as the date's serial number). With `equals` / `notEquals` a text may hold the wildcards `*` and `?` (`~*`, `~?`, `~~` for the characters themselves), as in Excel's custom filter; the four text operators take their text literally. `equals ""` is the blank cells. Numbers compare with number cells, text with text cells (case-insensitive). "Between" is two conditions.
- `{ top: { count, percent?: boolean, bottom?: boolean } }`: the `count` largest (or smallest) numbers of the column, or that percentage of its numbers (at least one); ties with the last one are let through.
- `{ color: "#RRGGBB" | null, by?: "cellColor" | "fontColor" }`: cells with that fill (default) or text colour; `null` = no fill / the automatic colour. (The cell's own format: conditional-format colours do not count.)
- `{ dynamic: "aboveAverage" | "belowAverage" }`.
- read only: `{ kept: true }`: a filter of an opened file that the editor does not evaluate (date groups, icon filters, the other dynamic filters). It is written back as it was; the rows the file had hidden stay hidden until the column gets another filter.

**Reading criteria back.** `GetFilters()` returns, for a column filtered with `SetFilter`, a copy of **what was given**: the same keys and the same types (a condition set with the number `5` reads back with the number `5`, one set with `"5"` with the text, a `Date` as a `Date`), and no key that was not given (no `and`, `blanks`, `percent`, `bottom` or `by` is added). This holds for as long as the column's filter is that one, through undo and redo.

For a filter that did not come through `SetFilter` (an opened file - XLSX stores every operand as text - or the grid's own filter buttons) the criteria are derived from the model in a canonical shape: keys at their default are left out (`blanks: false`; `and` unless it is `false` with two conditions; `percent: false`, `bottom: false`; `by: "cellColor"`), the `value` of `equals` / `notEquals` / `greaterThan` / `greaterOrEqual` / `lessThan` / `lessOrEqual` is a number when it is written as one (else a text), dates are their serial numbers, and `contains` etc. come back as such (text patterns that are more than that as `equals` with the pattern).

### Filter semantics
- Filtered rows are **hidden, not removed**: `GetCell` still reads them, `GetRowHeight` gives 0, the grid does not draw them and keeps the row numbers of the others. They are separate from rows hidden with `HideRows` (those are not listed in `hiddenRows`, are not shown by `ClearFilter`, and `GetHiddenRows` does not list filtered rows).
- As in Excel, the rows are worked out when a filter is set (or `ReapplyFilter()` is called), **not** when a cell changes: an edited row stays visible / hidden until then. Inserted and deleted rows and columns move the range, the column filters and the hidden rows; a new row is visible.
- On a selection that spans filtered rows, these act on the **visible rows only**: `GetSelectionText` / `GetSelectionHtml` / `Copy` (the copy holds the visible rows one after the other and pastes as one block), `Cut` and the Delete key, `ClearContents` / `ClearFormats` / `ClearAll`, every format command (`ApplyFormat`, `ApplyFormatPatch`, `ApplyBorder`, ...), `FillDown` / `FillRight` / `AutoFill`, and `SortRange` (hidden rows stay where they are). A **paste** into a filtered range is not filtered: it fills the rows one after the other, hidden ones included (Excel does the same). `SetCell` / `EnterText` write to any cell.
- `SUBTOTAL` / `AGGREGATE` leave out the rows a filter hides; setting a filter recalculates.
- XLSX: `<autoFilter ref>` with a `<filterColumn>` per filtered column (`<filters>`, `<customFilters>`, `<top10>`, `<colorFilter>` with its `<dxf>`, `<dynamicFilter>`), the hidden rows as `<row hidden="1">` (with their height), and the hidden sheet-scoped name `_xlnm._FilterDatabase`. An opened file with a filtered sheet opens filtered, with the rows the file hides inside the filter's range as the filter's rows; a sheet whose filter was not touched is written back as the file states it. (Before this the sheet-level `<autoFilter>` of a file was dropped on export.) The built-in name is derived from the filter and is not listed by `GetDefinedNames`.
- **Filter buttons.** The grid draws a button (`.spreadsheet-filter-button`, with `data-row`, `data-col`, `data-active="1|0"`, `data-sorted="asc|desc|"`) at the bottom right of every header cell of the AutoFilter range and of every table header row whose table has a filter (`hasAutoFilter`), hidden columns left out. It sits in the cell's pane: a frozen header row keeps its buttons, a scrolled one takes them along (and under the frozen panes). A filtered column's button is tinted and marked, a sorted one (the `SortState` column) shows an arrow. See `OnFilterButton` / `GetFilterButtonRects` below. Without a subscriber a press opens the grid's checklist of `GetFilterValues`, which sets a `{ values }` filter (an undo step). The older button beside each column letter (`.col-filter-btn`, it filters the block of data around the column) is still drawn when no subscriber is registered. A viewer (`SetReadOnly(true)`) has neither, unless a host subscribes.
- One AutoFilter per sheet. A table's own filter (Excel allows one per table besides the sheet's) is not modelled separately: filter a table by `SetAutoFilter(<table range>)`.

### Quick sort (`SortCurrentColumnAsc` / `SortCurrentColumnDesc`, a click on a column header's letter)
- **The block** is Excel's current region around the active cell: the smallest rectangle holding it that empty rows and empty columns (or the sheet's edges) bound on every side; cells touching at a corner join it. Data elsewhere on the sheet - above, below or beside, even in the same rows - is not touched. A table's totals row is left out.
- **The header row** is guessed as Excel guesses it: the first row of the block stays in place when it is a table's header row, when a label stands over a value of another kind in some column (text over a number, a date or a boolean - the first filled cell below decides), or when it is formatted unlike the row under it (bold, italic, underline, font, size, text colour, fill). A block of text formatted alike all the way down has no header: its first row is sorted with the rest (Excel does the same). A number over text in some column means the first row is data.
- **Blanks are last** in both directions; otherwise the order is `SortRange`'s. The sort is `SortRange(block, [{ column, descending }], { hasHeader })`: formulas that move are re-pointed, hidden filter rows stay, one undo step.
- A click on a column header's letter sorts the block around the active cell when that cell is in the column, else the block around the column's first entry. (Before 0.1.9 these sorted every row of the sheet's longest run of filled rows, across all columns, and took any all-text first row for a header.)

### Sort semantics
- Whole rows **of the range** move; cells beside the range stay. The first key decides, later keys break ties, rows equal on every key keep their order (stable).
- Value order ascending: numbers, text (case-insensitive; with `caseSensitive` lower case before upper case among otherwise equal texts), FALSE, TRUE, errors. `descending` reverses that. **Blank cells are last in either direction.** A formula sorts by its value.
- Colour keys: rows whose key cell has `color` as fill (`cellColor`) / text colour (`fontColor`) come first (last with `descending`); the others keep their order.
- What moves with a cell: value / formula, format, cell comment, rich-text runs, threaded comment, a hyperlink on that single cell. What stays in place: row heights, validations, conditional formats, merged cells (none are allowed in the rows), hyperlinks over several cells.
- Formulas: a formula that moves is re-pointed as Excel does on a sort, i.e. like a copy of it to the new row: relative row references shift by the rows it moved (`=B2*2` in row 2 becomes `=B7*2` in row 7; a reference pushed off the sheet becomes `#REF!`), `$`-locked rows stay. Formulas elsewhere that refer to the sorted cells are **not** re-pointed (they see the new order). The workbook is recalculated.
- Rows hidden by the AutoFilter are not sorted (they keep their place); rows hidden with `HideRows` are sorted like any other.
- The first key sets the header's sort arrow (`SortState`).

## Comments (group 10a)
All on the active sheet unless a sheet index is taken. Every mutating member is one undo step (no recalculation), bumps the revision and fires `OnChange`.

| Member | Arguments | Returns | Notes |
|---|---|---|---|
| `GetComment(row, col)` | 0-based | comment object \| `null` | Also what `GetCell(...).comment` is. |
| `GetComments(sheetIndex?)` | | `[{ row, col, address, text, author, runs, thread }]` | Row by row. Throws for a sheet that is not there. |
| `SetComment(row, col, text, author?)` | | | A cell without a comment, or with a note, gets a note. On a cell with a thread the first comment's text (and author, when given) changes; replies stay. Throws on an empty text. (Was not an undo step before.) |
| `SetCommentRuns(row, col, runs, author?)` | `runs: CommentRun[]` (only `text` is required) | | A note whose text is formatted in runs; the note's text is the runs' texts joined. Runs without text are dropped; with no formatting in any run the note is a plain one (`runs: null`). On a cell with a thread the first comment gets the joined text (threads hold plain text). Throws when the runs hold no text, on a `fontSize` outside 1..409 or a colour that is not `#RGB` / `#RRGGBB`. One undo step. |
| `DeleteComment(row, col)` | | `boolean` (there was one) | Removes the note or the whole thread. `RemoveComment(row, col)` is the older name (returns nothing). |
| `AddCommentReply(row, col, text, author?)` | | the reply's id (string) | A note becomes the first comment of a thread (Excel's "convert to comments"). Dated now (UTC). Throws when the cell has no comment or the text is empty. |
| `DeleteCommentReply(row, col, replyId)` | | | Throws when the cell has no thread / no such reply. |
| `ResolveComment(row, col, resolved = true)` | | | Marks a thread resolved / open. Throws on a note. |
| `GetCommentThread(row, col)` | | `{ author, date, text, replies }` \| `null` | Older member, kept; prefer `GetComment`. |

Comment object: `{ text, author: string | null, runs, thread }`; `runs` is `null` for a plain note (and for a thread), else `CommentRun[]`: `{ text: string, bold: boolean, italic: boolean, underline: boolean, strikethrough: boolean, color: "#RRGGBB" | null, fontName: string | null, fontSize: number | null }` (points; `null` = the note's usual look); `thread` is `null` for a note, else `{ id, author, date: "2026-01-02T03:04:05.00" | null, text, resolved: boolean, replies: [{ id, author, date, text }] }` (then `text` / `author` are the first comment's).

- The cell shows a corner marker (red for a note, violet for a thread). While the pointer is over the cell (no button down) the comment is shown in a box beside it: `.spreadsheet-comment-box` with `data-row`, `data-col` and `data-kind="note" | "thread"`; a note in its runs (one `<span>` per run with its font, size, weight, style, decoration and colour), a thread as one `.spreadsheet-comment-entry` per comment (author, date, text; replies indented) and "Resolved" when it is. The box takes no pointer events. The cell element itself carries the comment's text as `data-comment` (it no longer has a `title` for it: the browser tooltip would cover the box).
- Comments move with inserted / deleted rows and columns and with `SortRange`; `ClearAll` removes notes and threads; `ClearContents` / `ClearFormats` keep them.
- XLSX: notes go to the comments part, threads to the threaded-comments part with their persons and the `done` mark, plus the note Excel keeps beside each thread for old versions. Ids of thread comments survive a save; a note has no id.

## Hyperlinks (group 10b)
On the active sheet unless a sheet index is taken. `SetHyperlink` / `RemoveHyperlink` are one undo step each, bump the revision and fire `OnChange`.

| Member | Arguments | Returns | Notes |
|---|---|---|---|
| `GetHyperlink(row, col)` | 0-based | hyperlink object \| `null` | Also `GetCell(...).hyperlink`. |
| `GetHyperlinks(sheetIndex?)` | | hyperlink objects, in the order defined | |
| `SetHyperlink(range, target, display?, tooltip?)` | `range`: `"B2"`, `"B2:C3"`, several areas `"A1:A3 C1"` | | Replaces the links of those cells (a link that covered more keeps the rest). `target`: an address (`"https://…"`, `"mailto:…"`; any text is stored) or a place in the workbook: `"Sheet2!A1"`, `"'My sheet'!B2:C3"`, `"A1"`, a defined name, or any of these after `#`. `display` becomes the text of the first cell; without it an empty first cell shows the target. `tooltip` is the screen tip. Cells get the link look (text colour `#0563C1`, underline) unless they have a text colour of their own. Throws on a missing / invalid range, an empty target, a non-text display / tooltip. |
| `RemoveHyperlink(range)` | | number of links removed or cut | Also removes the link look (Excel's "Remove Hyperlink"); the texts stay. |
| `ActivateHyperlink(row, col)` | | `{ target, row, col, sheetIndex, isInternal, safe, tooltip, followed }` \| `null` (no link) | What Ctrl+click does: an internal link is followed (`followed: true`; false when the place no longer exists), an external one is handed to `OnHyperlinkActivate` subscribers. |
| `OnHyperlinkActivate(cb)` | `cb({ target, row, col, sheetIndex, isInternal: false, safe, tooltip, followed: false })` | unsubscribe function | Only for links that lead out of the workbook. |

Hyperlink object: `{ range, target, isInternal: boolean, tooltip: string | null, display: string | null, safe: boolean }`.

- **The grid never navigates.** No `window.open`, no `location` change, no anchor element: an external link only raises `OnHyperlinkActivate`. `safe` is true only for `http:`, `https:`, `mailto:` and `tel:` addresses; a host should open nothing else (`javascript:`, `file:`, `data:` … are reported with `safe: false`).
- Activation: Ctrl+click (Cmd+click) on a linked cell; a plain click in read-only mode (cursor is a pointer there). Internal links go through `GoTo` (sheet switch, selection, scroll).
- A linked cell has `title` = tooltip (else the target) unless it has a comment, and `data-link="internal" | "external"`.
- Links move with rows / columns, follow `RenameSheet`, and (single-cell links) `SortRange`.
- XLSX: `<hyperlink ref location | r:id tooltip display>`; addresses are external relationships of the sheet. Tooltip and display are new in the model and round trip.

## Pictures (group 10c)
On the active sheet. Every mutating member is one undo step (no recalculation), bumps the revision and fires `OnChange`.

| Member | Arguments | Returns | Notes |
|---|---|---|---|
| `GetPictures()` | | `[{ id, kind: "picture", name, description, contentType, row, col, x, y, width, height }]` back to front | `x`, `y`, `width`, `height` in sheet pixels at 100% from the sheet's top-left cell; `row`, `col` the anchor cell. |
| `GetDrawings()` | | the same objects for every drawing object (`kind`: `"picture"`, `"chart"`, `"shape"`, `"other"`; `contentType` null for non-pictures) | Now with `id` and `contentType`; `name` / `description` are `null` when absent. |
| `InsertPicture(source, anchorCell, size?, name?)` | `source`: `Uint8Array`, `ArrayBuffer` or a data URL; `anchorCell`: `"B3"` or `{ row, col }`; `size`: `{ width?, height? }` px | the new id (number) | PNG, JPEG, GIF only, recognised from the bytes (an SVG / BMP / WebP / anything else throws, whatever the data URL's type says). Without `size` the picture's own pixel size; with one side the other keeps the proportions. Put in front; one-cell anchor (moves with the cell, keeps its size). Default name `"Picture <id>"`. |
| `MovePicture(id, place)` | `place`: `"B3"` / `{ row, col }` (top-left corner to the cell's) or `{ x, y }` sheet pixels | | Size stays. Throws for an unknown id, a chart / kept object, a bad place. |
| `ResizePicture(id, { width?, height? })` | | | One side alone keeps the present proportions. Top-left corner stays. |
| `DeletePicture(id)` | | | |
| `SetPictureAltText(id, description, name?)` | | | `""` removes the description. |
| `GetSelectedPicture()` | | id \| `null` | |
| `SelectPicture(id \| null)` | | | Shows the frame with 8 handles; `null` deselects. |

- Ids: a number per sheet, stable through undo / redo, moves and re-stacking; never reused while a higher one exists. When a file is opened its drawing objects (pictures and charts alike) are numbered 1..n back to front, so ids are stable across save / reopen as long as the order of the objects is.
- In the grid (not in read-only mode, where pictures let clicks through to the cells): click selects a picture (frame + 8 handles: `.spreadsheet-picture-frame`, `.spreadsheet-picture-handle[data-handle]`; the image is `.spreadsheet-picture[data-drawing-id]`), drag moves it, a handle sizes it (corners keep the proportions, Shift frees them; sides change one dimension; minimum 8 px), Delete / Backspace removes it, Escape or a click on a cell deselects. Each drag is one undo step. Correct under zoom and when scrolled.
- `OnSelectionChange` payload has `Picture: id | null` and fires when the picture selection changes.
- **Frozen panes.** A drawing (picture, chart frame, kept object) that reaches over frozen rows / columns is cut along the frozen edges, as in Excel: the part over the frozen rows / columns stays in place, the part beside them scrolls and goes under the frozen cells. (Before, a picture anchored over frozen cells was hidden under them.) In the DOM such a drawing is one `.spreadsheet-drawing-part[data-drawing-id][data-pane="0..3"]` box per pane it reaches (0 scrolled, 1 frozen rows, 2 frozen columns, 3 the frozen corner), each holding a copy of the `.spreadsheet-picture` / `.spreadsheet-object` node: **a selector for the image may match up to four nodes**. Without frozen panes, and for a drawing wholly in the scrolled part, it is the one node as before. The selection frame sits in the pane of the picture's top-left corner.
- Charts and other kept objects are not movable through this API (the members throw); they keep following their cells.
- XLSX: pictures are written into the sheet's drawing part beside the kept charts (which stay byte for byte); identical image data is one media part; deleting the last drawing object of a sheet drops its drawing part and media.

## Find and replace (group 11)
A search reads the sheet on display, or every sheet that has a tab. Replacing is an undo step; searching and stepping are not changes.

| Member | Arguments | Returns | Notes |
|---|---|---|---|
| `Find(text, options?)` | `options`: `{ matchCase?: boolean (false), wholeCell?: boolean (false), regex?: boolean (false), lookIn?: "values" \| "formulas" ("values"), scope?: "sheet" \| "workbook" ("sheet"), range?: "B2:D9" }` | `[{ sheetIndex, row, col, address, text }]` | Matches row by row, sheets in tab order starting at the sheet on display. Selects nothing; the search is remembered for the three members below. `text` of a match is the text that was searched (shown, or entered). Throws for an empty text, a non-string, an unreadable regular expression, unknown option keys / values, `range` together with `scope: "workbook"`. |
| `FindNext()` | | match \| `null` (after `Find`); `true` / `false` (after the old `StartSearch`) | The first match after the active cell (then the following sheets for a workbook search; after the last match the first again). Shows its sheet, selects and scrolls to the cell. |
| `FindPrevious()` | | match \| `null` | The match before the active cell (before the first: the last). Throws when `Find` was not called. |
| `ReplaceMatch(replacement)` | string | `{ replaced: boolean, next: match \| null }` | Replaces in the active cell when it matches, then steps to the next match. One undo step, recalculates. Not on a match: nothing replaced, the next match is selected. Throws when `Find` was not called. |
| `ReplaceAll(text, replacement, options)` | `options` as for `Find` (an object; `{}` for the defaults) | number of cells changed | ONE undo step, also with `scope: "workbook"` (one `Undo()` on the sheet it was called from restores every sheet; `OnChange` names every changed sheet). 0 matches: no step. The old call `ReplaceAll(text, replacement, caseSensitive: boolean)` still works (active sheet, text contained in value or formula text) and now also returns the count. |
| `StartSearch(query, caseSensitive)`, `Replace(find, replace, caseSensitive)`, `GetSearchResultsCount()`, `GetSearchIndex()`, `ShowFindReplaceDialog()` | | | The earlier single-sheet search, unchanged (the built-in dialog uses it). `StartSearch` ends a search started by `Find`. |

- `lookIn: "values"` searches what cells show (number formats applied: `1,234.50`); a formula cell matches by its result. `lookIn: "formulas"` searches what was entered: a formula with its `=`, a number / date as it would be typed, text as it is.
- `wholeCell`: the whole text has to match (for a regular expression: from the first to the last character).
- `regex`: in the browser the pattern is a JavaScript regular expression (`new RegExp(text, "g" + (matchCase ? "" : "i"))`); `$1`, `$2` in a replacement are its groups. Without `regex` the text and the replacement are literal (`$` is a dollar sign).
- Replacing: every occurrence in the cell's text is replaced and the result is read the way typed input is (a number, a date, a boolean, a formula when it starts with `=`); the cell keeps its format and note; an empty result empties the cell. It is NOT checked against data validation. Looking in values, formula cells are skipped and the replacement is made in the text shown (`1,234.50` with `1,234` -> `2,345` gives the number 2345.5); looking in formulas it is made in the formula text. The cells a dynamic array spilled into are never changed.
- Hidden / very hidden sheets are not searched (unless one is on display); the cells of hidden rows and filtered-out rows are.
- A workbook search converts the sheets of an opened file that were not converted yet (`IsSheetLoaded`); a sheet search converts none.
- Matches are found again for a sheet whose cells changed since (`FindNext` after an edit sees the edit); stepping is always relative to the active cell.
- Cost: 846 000 cells are searched in about 1 s (values) / under 2 s (formulas); a step afterwards costs nothing until cells change.

## Tables (group 12)
A table has a name (unique in the workbook, also against defined names), a range on one sheet (header row, data rows, optional total row), named columns and a style. Structured references (`Sales[Amount]`, `Sales[@Amount]`, `Sales[[#Totals],[Amount]]`, `Sales[#Headers]`, `Sales[#All]`, `Sales[[A]:[B]]`) read it from any sheet. Every mutating member is ONE undo step of the workbook (it sits on the history of the sheet on display, like sheet operations), bumps the revision, fires `OnChange`, recalculates every formula and is written to / read from XLSX (`xl/tables/tableN.xml`).

| Member | Arguments | Returns | Notes |
|---|---|---|---|
| `GetTables()` | | `[{ name, sheetIndex, range, hasHeaders, hasTotals, columns: string[], style: { name, showRowStripes, showColumnStripes, showFirstColumn, showLastColumn } \| null, hasAutoFilter }]` | Sheet by sheet in tab order, each sheet's from the top left. `range` includes header and total row. `hasAutoFilter`: the header cells carry filter buttons in the file. |
| `DefineTable(range, name?, hasHeaders?, style?)` | `range` `"A1:D9"` on the active sheet; `name` default the first free `"TableN"`; `hasHeaders` default `true`; `style`: style name, `{ name, showRowStripes?, showColumnStripes?, showFirstColumn?, showLastColumn? }` or `"none"` (default `"TableStyleMedium2"`) | the table's name | With headers: blank header cells become `ColumnN`, a repeated name gets a number appended (`Amount`, `Amount2`), and header cells become text (written into the cells). Without headers: columns are `Column1..`, every row is data, NO header row is inserted (Excel inserts one). Throws: invalid / taken name (rules of defined names), range outside the sheet, whole rows / columns, overlap with another table, merged cells in the range, a header row with no row below, unknown style name. |
| `DefineTable(name, startRow, startCol, endRow, endCol, columns, hasHeaders)` | the earlier call (0-based, `columns` a string array) | the name | Still works: registers the table for structured references only (no style, no header changes, not written as a table part unless a style is set later). Now an undo step. |
| `ResizeTable(name, range)` | | | The first row stays in its row; the new range overlaps the old one; with a total row the range includes it (it moves to the new last row, which must be free). New columns are named by their header cells. Throws for an unknown table or a bad range. |
| `RenameTable(name, newName)` | | | Structured references in every formula of the workbook (cells, conditional formats, validations, defined names) follow. Throws for an unknown table, an invalid or taken name. |
| `DeleteTable(name, keepData?)` | `keepData` default `true` | | `true`: convert to range - cells stay, every structured reference becomes the plain absolute reference it stood for (`=SUM(Sales[Amount])` -> `=SUM($B$2:$B$9)`, with the sheet name on other sheets; a this-row reference becomes the cell of its row; one that selects nothing becomes `#REF!`). `false`: the table's cells are emptied too (formats stay) and references to the table show `#REF!` (their text is left). |
| `SetTableTotalRow(name, show)` | boolean | | Shown: the row below the data gets `Total` in the first column and, in the last, `=SUBTOTAL(109,Table[Col])` when that column holds numbers, else `=SUBTOTAL(103,...)` (a count). The row below must be empty (nothing is moved down: throws otherwise). Removed: the row's cells are emptied. |
| `SetTableStyle(name, style)` | style name, style object, or `null` / `"none"` | | |
| `GetTableNames()` | | the names (an F# list, kept for older hosts; use `GetTables`) | |

- Filtering a table: `SetAutoFilter(table.range without the total row)` + `SetFilter` (group 9; one filter per sheet). `SUBTOTAL` totals leave filtered rows out.
- Rows / columns inserted or deleted inside a table grow / shrink it (a new column is `ColumnN`); deleting its header row removes the table.
- **Typing in and around a table** (an entry the user types in the grid, `EnterText`, `CommitEdit`, the formula bar - NOT `SetCell`, paste or fill,
  which write cells only) does what Excel does, as ONE undo step that holds the entry and everything that follows from it (a workbook-level step: it
  bumps the revision once, fires `OnChange` with `WorkbookChanged: true`, and the workbook is recalculated):
  - *A header cell*: the column takes the text as its name. A cleared header becomes `ColumnN` (N = the column's place), a name another column of the
    table has gets the next free number appended (`Region` -> `Region2`; names compare case-insensitively), a number or a formula result is stored as
    its text. The header cell holds exactly that name, and structured references to the column follow in every formula of the workbook - cells,
    conditional formats, validations, defined names; `Sales[Amount]`, `Sales[@Amount]`, `Sales[[Amount]:[Qty]]` anywhere and the unqualified
    `[@Amount]` in the table's own cells. `GetTables()[i].columns` shows the new name at once.
  - *Directly below the last row* (a table without a total row; any of its columns) or *directly right of the last column* (any of its rows): the
    table grows by that row / column. It does not when the entry is empty, when another cell of the new row / column (within the table's width /
    height) already holds something, or when the table would run into another table or a merged cell - the entry is then an ordinary one. A new row
    takes the table's calculated columns (a column whose data cells all hold one formula); a new column is named by what is in the header row
    (`ColumnN` when nothing is, written into the header cell). The AutoFilter range, the bands and the structured references cover the new cells.
  - *A formula in a data cell of a column* whose other data cells are empty, or all hold the one formula of a calculated column: the formula fills
    the whole column (relative references move row by row; `[@Col]` stays). A column that holds values or differing formulas keeps them.
  Hosts that enter cells themselves on behalf of the user and want this should call `EnterText`; a host that must not get it uses `SetCell`.
- **Table styles are drawn from the style's name**, in the DOM grid and by the portable painter (PNG / PDF / native), for the built-in families
  `TableStyleLight1-21`, `TableStyleMedium1-28`, `TableStyleDark1-11`. The colour is the theme accent the number stands for (every seven: black, accent
  1 ... accent 6; taken from the theme of an opened file, the Office theme otherwise). Drawn: the header row (filled with white bold text in Medium /
  Dark and Light8-14, bold between lines in Light1-7 / 15-21), the total row (bold; a double line above it, or filled in Medium8-14 / 22-28 and
  Dark1-7), banded rows (`showRowStripes`: tint bands; two tints in Medium8-14 / 22-28 and the Dark styles; grey in Medium15-21) and banded columns
  (`showColumnStripes`), first / last column emphasis (`showFirstColumn` / `showLastColumn`: bold, filled where the style fills), the lines of the
  Light styles (between rows in Light8-14, a full grid in Light15-21), dark fills with white text in Dark1-7. A cell's own fill, text colour and
  borders win over the table's. Approximations: Dark8-11 use one accent (Excel pairs two), the inner lines of Medium1-7 are not drawn, custom
  (workbook-defined) table styles draw as `TableStyleMedium2`-like header and bands in black.
- XLSX: columns, `headerRowCount`, `totalsRowCount`, per-column `totalsRowLabel` / `totalsRowFunction` / `totalsRowFormula`, `autoFilter` (header + data), `tableStyleInfo`. Tables of an opened file that were not touched are written as the file had them; a deleted table's part is not written.
- `DuplicateSheet` copies a sheet's tables under new names (`Sales_2`), `DeleteSheet` removes them, `RenameSheet` leaves them.

## Events and state (group 13)

| Member | Arguments | Returns | Notes |
|---|---|---|---|
| `OnContextMenu(cb)` | `cb({ target, row, col, sheetIndex, clientX, clientY, pictureId, inSelection })` | unsubscribe function | `target`: `"cell"` \| `"rowHeader"` \| `"colHeader"` \| `"corner"` \| `"picture"` \| `"sheetTab"`. `row` / `col` 0-based, `-1` where the target has none (a column header has no row, the corner neither; a picture gives its anchor cell; a tab gives `-1, -1`). `clientX` / `clientY`: where to open the menu (the pointer; for the keyboard the middle of the active cell / picture). `pictureId`: number \| null. `inSelection`: the target was part of the selection already. |
| `RaiseSheetTabContextMenu(sheetIndex, clientX, clientY)` | | `true` when a subscriber was told | For a tab strip (the component draws none): raises `{ target: "sheetTab", row: -1, col: -1, sheetIndex, clientX, clientY, pictureId: null, inSelection: sheet is on display }`. Does not switch sheets. Throws for a sheet that is not there. |
| `OnError(cb)` | `cb({ source, message, error })` | unsubscribe function | Errors the component catches instead of letting them reach the host. |
| `GetStatus()` | | `{ sum, average, count, numericCount, min, max }` | Of the selection of the active sheet. |
| `GetUsedRange(sheetIndex?)` | | `{ range: "B2:F40", firstRow, firstCol, lastRow, lastCol }` \| `null` (empty sheet) | |

- Context menu: raised by a right click (`contextmenu`) in the grid and by the menu key / Shift+F10 (active cell, or the selected picture). As in Excel a right click outside the selection first selects what it is on (the cell; the whole row / column of a header; the picture) and one inside the selection keeps it; `OnSelectionChange` fires before the callback. A right click on the corner selects nothing. While at least one subscriber is registered the browser's own menu is prevented in the grid, except in the edit box of a cell being edited (spelling / paste menu stays). The component has no built-in menu; without a subscriber nothing changes (the browser menu shows, nothing is selected). Works in read-only mode too (the host decides what to offer). Correct under zoom, frozen panes and hidden headers (the cell is the one under the pointer).
- `OnError` sources: the name of the subscription whose callback threw (`"OnChange"`, `"OnSelectionChange"`, `"OnSheetsChange"`, `"OnSheetChanged"`, `"OnEditStateChange"`, `"OnHyperlinkActivate"`, `"OnContextMenu"`, `"OnFilterButton"`; the other subscribers are still told); `"event:<dom event>"` (`event:keydown`, `event:mousedown`, `event:paste`, `event:scroll` ...) for an error while the grid handles a DOM event: it is reported and then thrown on, so the console shows it as before; `"picture"` (a picture drag that could not be applied); `"ImportXlsx"` (also `ImportXlsxFromBuffer`): the file could not be read; the returned promise still rejects. NOT reported: what a member called by the host throws (it is thrown to the caller); a clipboard the browser refuses (Copy / Cut / Paste resolve to `false`); an OnValidationNotice callback that throws (it propagates to the caller of the entry). An error thrown inside an OnError callback is dropped.
- `GetStatus`: `count` = cells that hold something (empty text and format-only cells do not count), `numericCount` = numbers among them (dates are numbers; a formula counts by its result; text, booleans and errors are not numbers), `sum` / `average` / `min` / `max` over the numbers, `null` when there is none. Rows hidden by the AutoFilter are left out (Excel's status bar); rows hidden by hand are counted. Cached: the same object is returned until the sheet's cells, the filter's hidden rows or the selection change. Cost: a selection up to 20 000 cells is walked cell by cell (well under 1 ms for typical selections); larger ones (whole rows / columns / sheet) cost one pass over the cells that exist: about 100 ms for the whole 846 000-cell dashboard "Data" sheet, once.
- `OnSelectionChange` payload gained `Status`: the `GetStatus()` object of the new selection, worked out when the property is first read (a subscriber that does not read it pays nothing: a selection change on the 846k-cell sheet stays under 1 ms). The selection event does not fire when the cells of an unchanged selection change: call `GetStatus()` from `OnChange` to keep a status bar current.
- `GetUsedRange`: the smallest range holding every cell with a value, a format or a note and every merged cell, as Excel's `UsedRange` (it need not start at A1; it does not shrink for cells that were emptied but keep a format). A sheet of an opened file that was not converted yet (`IsSheetLoaded`) is converted by the call.

## Print / page setup (group 14)
Page setup is per sheet. Every setter is ONE undo step (on the sheet on display; for another sheet a step of the workbook that is undone from the sheet on display), bumps the revision, fires `OnChange` naming the sheet, changes no cell, and is written to / read from XLSX (`pageSetup`, `pageMargins`, `printOptions`, `headerFooter`, `sheetPr/pageSetUpPr`, `rowBreaks` / `colBreaks`, and the sheet's `_xlnm.Print_Area` / `_xlnm.Print_Titles` names). `sheetIndex` is optional everywhere (default: the active sheet).

| Member | Arguments | Returns | Notes |
|---|---|---|---|
| `GetPageSetup(sheetIndex?)` | | the `pageSetup` object below | |
| `SetPageSetup(changes, sheetIndex?)` | a `pageSetup` object with only the keys to change | | A patch. `SetPageSetup(GetPageSetup())` changes nothing (and makes no undo step). Throws for an unknown key or a bad value; nothing is changed then. |
| `GetPageBreaks(sheetIndex?)` | | `{ rows: number[], columns: number[] }` | Manual breaks: the 0-based rows / columns BEFORE which a break sits (`rows: [20]` = row 21 starts a page), ascending. |
| `SetPageBreaks({ rows?, columns? }, sheetIndex?)` | | | Replaces the lists given; one not given stays. Duplicates are dropped. Throws for a break before row / column 0 or outside the sheet. |
| `AddPageBreak(row \| null, col?, sheetIndex?)` | 0-based | `true` when a break was added | Throws when neither is given. |
| `RemovePageBreak(row \| null, col?, sheetIndex?)` | | `true` when one was removed | |
| `ClearPageBreaks(sheetIndex?)` | | how many there were | |
| `GetPrintPages(sheetIndex?, fileName?)` | `fileName`: what `&F` prints (default `""`) | `{ sheetIndex, sheetName, setup, scale, printAreas, printTitles, breaks, pageCount, pages }` | The pages the sheet prints on, **as the PDF export lays them out** (one pagination, `Pagination.paginatePrinted`: rows and columns at their print sizes, manual breaks, the scale or the fit to pages, print titles and printed headings taken off every page, a print area of several areas paginated area by area, `pageOrder`). `setup` is the object of `GetPageSetup` (orientation, paper, margins, grid lines / headings / centring flags ...); `scale` the percent the sheet is printed at (the one that fits it when it is fitted to pages); `printAreas: ["A1:F40", ...]` what is printed (the print area's areas in order, or the used range); `printTitles: { rows: { first, last } \| null, columns: { first, last } \| null }` (0-based); `breaks` as `GetPageBreaks`. Each of `pages`: `{ index, number, area, range, firstRow, lastRow, firstCol, lastCol, header, footer }` - `number` is what `&P` prints (`firstPageNumber` counts), `area` an index into `printAreas`, the rows / columns (0-based) are the page's body (print titles ride along beside it), `header` / `footer` are `{ left, center, right }` with the codes resolved for that page (`&P &N &D &T &A &F ...`; the first-page / even-page texts where the setup has them) or null. Read only; a sheet that was not converted yet is converted. A page-break view draws its lines at `lastRow` / `lastCol` of the pages. |

The `pageSetup` object:
```
{ orientation: "portrait" | "landscape",
  paperSize: "letter" | "tabloid" | "legal" | "executive" | "a3" | "a4" | "a5" | "b5" | { code: number },   // { code }: any other OOXML paper code (kept as is)
  pageWidth, pageHeight: number,          // points, of the paper in portrait; read only (ignored by SetPageSetup)
  margins: { top, bottom, left, right, header, footer: number },   // INCHES (as the file and Excel's dialog); 0..49
  scale: number | null,                   // percent 10..400; null while the sheet is fitted to pages
  fitToWidth, fitToHeight: number | null, // pages, 0 = as many as it takes; null while a scale applies
  printArea: "A1:F40" | "A1:C10,E1:G10" | null,   // one area, or several separated by commas
  printAreas: ["A1:C10", "E1:G10"],       // the same, one entry per area ([] when none); both given: printAreas counts (what GetPageSetup returns can be given back)
  printTitles: { rows: "1:2" | null, columns: "A:B" | null },
  header, footer, firstHeader, firstFooter, evenHeader, evenFooter: { left, center, right, raw: string },
  differentFirst, differentOddEven, scaleHeaderFooter: boolean,
  gridlines, headings, centerHorizontally, centerVertically: boolean,
  pageOrder: "downThenOver" | "overThenDown", blackAndWhite: boolean,
  firstPageNumber: number | null,         // null: automatic
  comments: "none" | "asDisplayed" | "atEnd" }
```
- **A print area of several areas** (`Print_Area` = `Sheet1!$A$1:$C$10,Sheet1!$E$1:$G$10`) is read, kept and written (each area with its
  sheet). Each area is **its own run of pages**, in the order given: the paginator and the PDF output lay out the first area's pages, then
  the next area's (page order, manual breaks and print titles apply within each; cells between the areas are on no page). Fit-to-page
  fits each area into the pages given, all at one scale - the smallest any area needs. Areas follow inserted / deleted rows and columns;
  an area that is deleted whole goes. `PageSetup.PrintArea` is the first area, `PageSetup.ExtraPrintAreas` the others (`Model.printAreasOf`).
- Setting: `scale` (a number) switches fitting off; `fitToWidth` and / or `fitToHeight` switch it on (the side not given keeps its number, initially 1); giving a scale and a fit in one call throws; `null` for any of the three is ignored (so the object of `GetPageSetup` can be handed back). `paperSize`: a name, a code number or `{ code }`. `margins` may name some sides. `printArea`: an A1 range (`$` allowed) or `null`. `printTitles`: `{ rows?, columns? }` (whole rows `"1:2"` / whole columns `"A:B"`, `null` clears that one) or `null` (clears both). `firstPageNumber: null` = automatic.
- Headers and footers are text in Excel's code syntax: `&P` page number, `&N` number of pages, `&D` date, `&T` time, `&F` file name, `&A` sheet name, `&&` a literal ampersand, `&"Font,Style"`, `&12` font size, `&K RRGGBB` colour, `&B` `&I` `&U` toggles. `left` / `center` / `right` are the sections of `raw` (`&L...&C...&R...`; text before the first marker belongs to the centre). To set one: a string (the raw text), `null` (none), `{ raw }`, or `{ left?, center?, right? }` (sections not given stay as they are). At most 255 characters with the codes. `firstHeader` / `firstFooter` apply when `differentFirst`, `evenHeader` / `evenFooter` when `differentOddEven` (the flags are separate keys and are not set implicitly).
- The print area, the print titles and the breaks follow inserted and deleted rows / columns (and their undo); the print names follow a sheet rename (they are written from the page setup, not stored as defined names: `GetDefinedNames` does not list them and `SetDefinedName` refuses `_xlnm.` names).
- XLSX: a sheet of an opened file whose setup was not changed is written exactly as the file had it (also a multi-area print area and printer settings the editor does not model). Once the page setup of a sheet is changed, the `<pageSetup>` element is written from the model: `horizontalDpi` / `verticalDpi`, the printer-settings relationship (`r:id`), `draft`, `errors`, `copies` and `autoPageBreaks` are not written for that sheet.
- Page count and page rectangles: `GetPrintPages` (above). The pagination (`Pagination.paginatePrinted`) is part of the browser bundle and is the one the PDF exporter (`Rbt.Spreadsheet.Pdf`, .NET only) lays its pages out with; only the PDF output itself is .NET only.

## Calculation (groups 15 / 16)
Nothing to call for the recalculation itself: every edit recalculates what it has to. What a host can read and set:

- `GetRecalcStats()` -> `{ Kind: "incremental" | "full" | "none", Reason, ChangedCells, ReachedFormulas, EvaluatedFormulas, Evaluations, Formulas }` (PascalCase keys). `EvaluatedFormulas` counts each formula once (never more than `Formulas`); `Evaluations` is how many evaluations the pass took: the formulas of a circular reference that is iterated are evaluated once per sweep, so it can be several times `EvaluatedFormulas`. (Before, `EvaluatedFormulas` was that total: "23006 of 13346" after a paste that closed a cycle.)
  The kind is "incremental" after an edit of cells - also in workbooks with dynamic arrays, OFFSET / INDIRECT and 3-D references, and when the edit
  makes or reaches a circular reference - and (since 0.1.9) after edits that are more than cells; only the formulas that can tell are evaluated:
  - a **circular reference** made or reached by the edit (a paste, a typed formula): its formulas show the cycle error, or - with iterative
    calculation - are calculated in sweeps, inside the incremental pass. With iterative calculation a circular reference that did not settle is
    calculated again at every edit (as in Excel); one that settled is left alone until something it reads changes. (The circular references
    of a file that was just opened are taken as settled - the file came calculated - until an edit reaches one.)
  - a **filter** set, changed or removed, rows hidden or shown: SUBTOTAL / AGGREGATE formulas and what reads them.
  - a **defined name** defined, changed or deleted: the formulas that use that name (directly or through another name) and what reads them.
  - a **table** defined, resized, renamed, removed, its total row: the formulas with structured references and what reads them.
  - **rows / columns inserted or deleted**: formulas that only follow their cells (`=B5` becoming `=B6`) are re-pointed and keep their value
    without being evaluated; evaluated are the formulas the edit really changes (a range it took cells from, `#REF!`), those that read the sheet
    at a place that stayed while cells moved under it (a whole column, `INDEX(B:B,3)`), ROW() / COLUMN() and the like, SUBTOTAL / AGGREGATE, and
    the arrays of the sheet. The undo (and redo) of such an edit is calculated the same way.
  - **sheets** added, deleted, moved, copied back by an undo: the formulas that read a deleted / restored sheet, formulas waiting for a sheet of
    that name, 3-D references and SHEET() / SHEETS(). (A rename and a copy change no value: the graph is built again at the next edit.)
  `EvaluatedFormulas` counts evaluations (the sweeps of an iterated circular reference count each time).
  It is "full" (every formula of the workbook, in one dependency order across the sheets) with one of these reasons:
  - "the setting of iterative calculation changed" / "the workbook was calculated from scratch" (`SetIterativeCalculation`)
  - "the dynamic arrays do not settle (a circular reference through spilled cells)"
  - "formulas that choose their precedents (OFFSET, INDIRECT) do not settle (a circular reference through them)"
  - "rows or columns changed on a sheet that is not converted" (does not happen through the component: it converts the sheets first)
  - "no calculated state to start from" / "incremental recalculation is switched off"
- `SetIncrementalRecalculation(enabled: boolean)` (existing): off = every edit is a full recalculation.
- Circular references. The formulas ON a circular reference - within a sheet, across sheets, through a defined name, a 3-D reference or the cells of a
  dynamic array - have the engine's cycle error as their value (GetCell: `{ value: "#CYCLE", valueType: "error" }`; it is not one of Excel's error values - Excel shows 0 and a warning); a formula that only reads one has that error as any formula
  reading an error has (IFERROR catches it). This is the same after every edit and whatever the order of the sheets.
- `GetIterativeCalculation()` -> `{ enabled: boolean, maxIterations: number, maxChange: number }`. Excel's "Enable iterative calculation" (File > Options >
  Formulas), a setting of the workbook that is read from and written to the file (`<calcPr iterate iterateCount iterateDelta>`). When off: `{ enabled: false,
  maxIterations: 100, maxChange: 0.001 }`. When on, the formulas of a circular reference are not errors: they are calculated again and again, from the
  values they have, until none changes by `maxChange` or more, at most `maxIterations` times per recalculation (as in Excel a cycle that does not converge
  moves on by that many steps at every recalculation that reaches it). Financial models with a deliberate circular reference (interest on a loan that
  finances the interest) need it: four of the six real-world workbooks in the test corpus have it on.
- `SetIterativeCalculation(enabled: boolean, maxIterations?: number = 100, maxChange?: number = 0.001)` -> void. Recalculates the workbook. ONE undo step
  (on the active sheet's history, like the other workbook-level steps), bumps the revision, fires OnChange with `WorkbookChanged: true`. Throws
  `SetIterativeCalculation: ...` when `enabled` is not a boolean, `maxIterations` is not 1..32767 or `maxChange` is negative.
- 3-D references (`=SUM(Jan:Mar!B2)`): read the range on every sheet between the two named sheets in tab order, whatever the order they are named in.
  `MoveSheet` / `AddSheet` / `DeleteSheet` of a sheet inside the span change what is read, not the text. `DeleteSheet` of a sheet a span ends at re-points
  the span to that sheet's neighbour inside it (`Jan:Mar` without Mar is `'Jan:Feb'`, a span of two is left as a plain reference to the one sheet, a span
  of that one sheet is `#REF!`), as Excel does; undo restores the text. `RenameSheet` re-points either end. Inserting or deleting rows / columns on ONE
  sheet of the span does not re-point a 3-D reference (Excel does not either: it is the same cells on every sheet of the span); it reads whatever is at
  those cells afterwards.

## XLSX: what is written how

### What is read when
`ImportXlsx` shows **the sheet the file was saved on** (`activeTab` of `xl/workbook.xml`; the first visible sheet when that one is hidden
or the file names none) - `GetActiveSheetIndex()` says which. `ImportXlsxFromBuffer(buffer, name)` and `InstallImportedWorkbook(wb, -1)`
show the sheet asked for, else the first visible one, as before. Only the sheet on display is converted (`IsSheetLoaded`); the **rows of
the other sheets are not even read** from their parts until something asks for them (the sheet is shown or read by a formula or a host
call, or - for a sheet with formulas - in a task of its own right after the workbook is shown, so that the first edit does not wait for
it; not in a viewer). The formulas the dependency graph needs and the row heights a save needs are found in the part without its other
cells. A sheet nobody looks at costs the inflating of its part at open and nothing at a save.
`XlsxReader.Worksheet.Rows` (office-formats) is now a member that reads the rows on demand (`RowSource: SheetRows`, with `IsLoaded`,
`FormulaCells`, `Headers`); `XlsxWriter.XlsxSheetOut.Cells` likewise (`CellSource`; build one with `XlsxWriter.sheetCells cells`). A sheet
with a what-if data table is read at once.

### An unchanged workbook is exported as the file that was opened
`ExportXlsx()` of a workbook that `ImportXlsx` opened and in which **nothing was changed** resolves to a copy of the bytes that were opened:
every part byte-identical, nothing regenerated. "Nothing changed" is: `GetRevision()` is the revision the workbook opened with (also after
undoing every change back to it), the sheet on display is the one shown after opening, every converted sheet is shown as its file says
(zoom, grid lines, headers, formulas, zeros), and the package declares its workbook part's content type (a rough package is written anew, as
before, which makes it whole). Selections and scrolling are not changes. Anything else - an edit, a setting switched and switched back by
hand (a new revision), another sheet on display - is written by the writer as before: sheets, styles and shared strings from the model
(untouched cells, tables, validations and conditional formats as the file states them), everything else carried through (next section).
Only `ImportXlsx` keeps the bytes (one copy of the file, compressed as it came); `ImportXlsxFromBuffer`, `InstallImportedWorkbook` and the
.NET hosts always write.

### A sheet nothing was changed in is written as the part it was read from
When the workbook did change, `ExportXlsx()` still writes every **worksheet nothing was changed in** as the file's own part: the same
bytes under the same name, with everything the part's relationships lead to (comments and their boxes, threaded comments, drawings, tables,
pivot tables, printer settings) as it was. Only the sheets that changed are written from the model. So that the untouched parts keep their
meaning, `xl/styles.xml` and `xl/sharedStrings.xml` are then **the file's parts, added to**: formats, fonts, fills, borders, number
formats, dxfs and strings the edited sheets need and the file does not have are put behind the file's own, and no index of the file
changes (named styles, the indexed palette, table styles and extensions of the styles part stay as they were, too).
- **The styles and shared strings are added to in every save of an opened workbook**, also when every sheet changed: a cell that keeps
  its format is written with the very `cellXfs` entry the file gave it (`XlsxCellFormat.SourceXf`), so `<protection locked="0"/>`,
  `quotePrefix`, `pivotButton`, theme colours and anything else of that entry stay; custom `<tableStyles>`, `<colors>` (indexed palette,
  `mruColors`), the styles' `<extLst>` (slicer styles, timeline styles, x14 dxfs), named styles and dxfs stay where they are. A cell
  whose format was changed gets a new entry **made from the file's entry** (`XlsxCellFormat.BasedOnXf`): what the change touched - the
  number format, the font, the fill, the border, the alignment, each only when it is no longer what the file's entry said - is replaced,
  everything else of the entry stays (`xfId` / the named style, `<protection>`, `quotePrefix`, `pivotButton`, unmodelled attributes of
  `<xf>` and `<alignment>`, `<extLst>`, and the fonts / fills / borders the change left alone, with their theme colours). A cell the
  file did not have gets an entry made from the model.
- **`xl/workbook.xml` is the file's part** with the sheet list, the defined names, `activeTab` and the calculation flags (`iterate...`,
  `fullCalcOnLoad` when a formula has no result) made anew: `fileVersion`, `workbookPr` (`hidePivotFieldList`, `codeName` ...; not
  `date1904`, see below), `workbookProtection`, the window of `bookViews`, `calcPr`'s other attributes, `mc:AlternateContent` and the
  extension list stay as they were.
- **A sheet's comments part** (`xl/commentsN.xml`: `shapeId`, authors, phonetic runs) goes back as it is while the sheet's comments are
  the file's, also when the sheet itself is written from the model. When the comments changed the part is written from the model:
  every note keeps its runs (font, size, bold, italic, underline, strikethrough, colour, super/subscript); a run's `charset` / `family`
  and phonetic text are not kept.
- In a browser a sheet that goes back as it was read, and every other part the reader kept that is written as the file's bytes (chart
  parts, pivot caches and their records, media, the theme ...), is put into the new package as **the compressed entry it was** (not
  deflated again).
- **A sheet with little changed is the file's part with that changed in it** (not written from the model). "Little" is: cells
  (values, formulas, formats, rich-text runs - typed, pasted, filled, cleared, or formulas with another result after a recalculation),
  column widths and column styles, merged cells, conditional formats, validations, the view (zoom, grid lines / headers / formulas /
  zeros, frozen panes) and the tab colour. Then only this is replaced in the part's text: the `<c>` elements of the changed cells (a
  new cell is put into its row, a new row among the rows, a removed cell is taken out; a formula that only has another result keeps
  its `<f>` - shared, array - and gets the new `<v>`), `<dimension>` when a cell lies outside it, and the elements whose content
  changed: `<cols>`, `<mergeCells>`, `<conditionalFormatting>`, `<dataValidations>` made from the model; `<sheetView>` with only the
  changed switches / `zoomScale` set (selection, `topLeftCell`, `zoomScaleNormal` stay; other panes: the view is made anew);
  `<sheetPr>` with only `<tabColor>` replaced. Everything else is the file's text, character for character: `sheetPr` (`codeName`,
  `pageSetUpPr`), the views, `sheetFormatPr`, every `<row>` with its attributes (`s`, `customFormat`, `ht`, `spans` unless a cell was
  added to the row), untouched cells, shared formulas, `ignoredErrors`, page setup and margins, `extLst`, and every relationship id
  (the sheet's `.rels` and what it leads to are the file's). Nothing the file did not have is added (no empty `<headerFooter>`).
  The sheet is **written from its model** instead when rows / columns were inserted, deleted, hidden or grouped, a row height
  changed, a range was sorted or filtered, comments / threads / tables / drawings / hyperlinks / the page setup / data tables changed, a
  changed cell is the master of a shared formula other cells still use, rules or validations of the x14 list are involved, the
  workbook is in the 1904 date system, the part names a renamed or deleted sheet, or the part is not laid out as expected (a prefix on
  its elements, rows or cells without `r`, comments between elements). (`SheetExtras.Patch`; `XlsxBridge` decides, `XlsxWriter` patches.)
- **A formula's text result is written in its cell** (`t="str"`), as Excel does; it is never added to `xl/sharedStrings.xml` (which
  stays byte-identical unless a text was typed that the file does not have).
- **`[Content_Types].xml`, `_rels/.rels` and `xl/_rels/workbook.xml.rels` are the file's text** while they say what is written: same
  entries, same order, same (relative) targets, same ids. When a part is added or goes, only its entry is added (at the end) or taken
  out. A sheet of the file keeps the relationship id the file's workbook names it by; styles, shared strings and theme likewise.
  In `xl/workbook.xml` the `<sheets>` and `<definedNames>` elements stay the file's text while they say the same (defined names in
  any order, a sheet name quoted or not); `activeTab` is touched only when another sheet is on display; an apostrophe in a defined
  name's text is written as it is (`'My Sheet'!$A$1`, not `&apos;`). `xl/styles.xml` is byte-identical when no format was added (a
  `<dxfs count="0"/>` is opened only to add a dxf).
- "Nothing changed in a sheet" is: its cells (values, formulas, their cached results, formats, comments), sizes, outline, merges, rules,
  validations, tables, drawings, threads, links, page setup, panes, filter, zoom, grid lines / headers / formulas / zeros and tab colour are
  what they were when it was opened. A formula result that changed because another sheet was edited is a change (the cell's `<v>` is
  replaced in the sheet's part, see above). Renaming, hiding or moving the sheet is not (that is in `xl/workbook.xml`).
- The one thing rewritten in such a part is `tabSelected` of its view, when the sheet on display is no longer the one the file had on
  display (two sheets that say so are a group to Excel); the part is otherwise the file's text.
- **Every sheet is written from the model** (the styles, strings and workbook part are still the file's, added to) in a workbook in the
  1904 date system (its serials are converted on import, and `date1904` is not written). In a workbook in which **a sheet of the file was
  renamed or deleted**, a sheet whose part names one of those sheets anywhere in its text (a formula, a rule, a link) is written from the
  model, where the references have followed; the sheets that do not name any still go back as the parts they were read from (also when
  such a name has an apostrophe, a quotation mark, `&`, `<`, `>` or letters outside ASCII: the name is looked for as it is, with its
  apostrophes doubled as in a quoted reference, and with those characters as XML writes them). **Everything is written from the model as before** with a
  styles or shared-strings part the writer cannot add to (no `cellXfs` / `fonts` / `fills` / `borders`, element names with a prefix, or entries the reader counted
  differently); a workbook that was not opened with `ImportXlsx` in a browser (`ImportXlsxFromBuffer`, `InstallImportedWorkbook`: the
  editor does not have the file's bytes). A copy of a sheet (`DuplicateSheet`) is always written from the model. The copy has the source's charts (parts of their own, reading the copy), pictures and shapes; the source's **pivot tables, slicers and timelines stay with the source** - the copy holds the pivot tables' cells as plain values and no slicer / timeline frame (a frame without its slicer makes Excel repair the file).
- .NET: `XlsxBridge.Export.toWorkbookOut` does the same for a workbook read with `XlsxReader.loadBytesNative` (the reader keeps the
  bytes in `Preserved.SourceParts`). In a browser the reader does not keep the large parts; `Export.verbatimPartsWanted input` names the
  parts to take from the file and `Export.toWorkbookOutWith parts input` writes with them (this is what `ExportXlsx` does with the bytes
  `ImportXlsx` kept). `ToXlsxWorkbook()` is the workbook without those parts: every sheet from the model.
- `tests/office-formats/xlsx-parts.fsx` edits one cell of the first sheet of every test workbook and asserts that every other worksheet
  part and what it reaches is byte-identical, and that every cell of every sheet reads back with the value and format it had; the browser
  check `export-keeps-parts` asserts the byte identity through `ImportXlsx` / `ExportXlsx`.

### Parts the editor does not model
Nothing of an opened package is dropped on a save. Besides the parts made from the model (workbook, styles, shared strings, the worksheets,
comments, threaded comments and persons) every part goes back **as it was read, under its own name, with its relationship and content type**:
- of the workbook: pivot and slicer caches, external links, connections, cell metadata (`xl/metadata`) and rich values, the Power Pivot data
  model (`xl/model/item.data`), custom XML, XML maps, and any other relationship type that is not a sheet, the styles, the shared strings,
  the theme (kept), the calculation chain, the persons or the macros;
- of the package: document properties, the thumbnail, sensitivity labels (`docMetadata`);
- of a sheet: pivot tables, slicers, timelines, printer settings, and what the elements `customProperties`, `legacyDrawing`,
  `legacyDrawingHF`, `drawingHF`, `picture`, `oleObjects`, `controls`, `webPublishItems` name (form controls and camera pictures with their
  VML drawing and images, embedded objects, header pictures); those elements are written back as they were;
- **a sheet's drawing part** (`xl/drawings/drawingN.xml`, with its relationships, charts and media) while the sheet's pictures and objects are
  the file's: byte for byte, under its name. Once a picture is added, moved or removed, or rows / columns move an anchor, the drawing is
  written from the model as `xl/drawings/sheetDrawingN.xml`;
- **a table part** (`xl/tables/tableN.xml`, with its query table) while the table's range, columns, header / total rows and style are what
  the part says. A table of the file that changed (resized, a header retyped, renamed, totals or filter switched, another style) is
  written as **the file's part with what changed replaced**: its id and revision ids, `headerRowDxfId` / `dataDxfId` and each column's
  own id, `dataDxfId`, `calculatedColumnFormula` and `queryTableFieldId`, the filter's criteria, `tableType="queryTable"` and the part's
  relationship to its query table stay. Columns are matched by position while their number is the file's, by name otherwise; a new
  column gets a free id; a `sortState` goes when the range changed. A renamed table is found as the one table of the file it overlaps.
  Only a table that is new to the workbook is written from the model alone. New table parts take numbers and ids no kept part uses.
- A VML drawing that holds nothing but the boxes of the sheet's comments goes back while the comments are the file's and is made anew when
  they change; one that also holds form controls or camera pictures keeps every such shape as it was, and the boxes of the comments in it
  follow the comments (a new comment gets a box in that drawing, the box of a removed comment is taken out; untouched it stays the file's).
- **Chart sheets and dialog sheets** are sheets of the workbook the editor has no model of. They are listed like any sheet (a chart sheet
  shows as an empty grid with its chart on it) and go back as the part they were read from, byte for byte, at their place in the sheet
  order, as a chart sheet (relationship type, content type, the drawing and chart behind it). Only `tabSelected` of the part's view is
  rewritten when the sheet on display changed. Cells typed into such a sheet are not saved. A copy of one (`DuplicateSheet`) is a worksheet.
- A worksheet of the file is written under the name its part had (`xl/worksheets/sheet3.xml` stays that), new sheets under free numbers.
- Not carried: `xl/calcChain.xml` once a formula cell was added, removed or moved (Excel rebuilds it).
`tests/office-formats/xlsx-parts.fsx` asserts this for every workbook under `tests/real-world` and `tests/test-documents-excel`.

### Dynamic arrays and legacy array formulas

- A formula that spills (`=SEQUENCE(3)`, `=SORT(A1:A9)`, `=A1:A3*2`) is written as Excel writes a dynamic array: `<c r="A1" cm="1"><f t="array"
  ref="A1:A3">`, with `xl/metadata.xml` holding the XLDAPR future metadata (`xda:dynamicArrayProperties fDynamic="1"`) that the `cm` names, the part's
  content type and its workbook relationship. Excel opens it as a spilling formula (no `{…}`, no implicit `@`). The `ref` is the range the formula
  fills at the time of the save.
- A file's dynamic array formulas are read back as spilling formulas, including one that fills a single cell (`ref="H1"` with `cm`): it is written as a
  dynamic array again for as long as its cell holds that formula.
- A legacy (Ctrl+Shift+Enter) array formula - `<f t="array">` on a cell without dynamic-array metadata - is still calculated as an array formula over
  its cells, and is written as a legacy array formula over the range the file gave it, again for as long as the cell holds the formula it was read
  with (rows / columns inserted or deleted and a sort move the formula and keep it). Typing another formula into the cell makes it an ordinary (or
  spilling) formula. There is no host call that makes a formula a legacy array formula.
- A workbook that already has a metadata part with other entries (rich values: pictures in cells, data types) keeps it: the part, its relationship, the
  rich-value parts it goes with and their content types are carried through; when a dynamic array has to be written and the part has no XLDAPR record
  yet, the type, the future-metadata block and a `cellMetadata` record are ADDED to it (the existing types, blocks and `valueMetadata` keep their
  positions and numbers). Not carried: the `vm` attribute of a cell that names a rich value (the cell keeps its fallback value, as before).
- A workbook without dynamic arrays gets no metadata part.

### Conditional formats and validations of Excel 2010+ (x14)

- A data bar with x14 options is written twice, as Excel does: the 2007 `<cfRule type="dataBar">` (with `min` / `max` ends where the x14 ones are
  automatic) carrying `<extLst><ext uri="{B025F937-…}"><x14:id>{…}</x14:id></ext></extLst>`, and its `<x14:cfRule type="dataBar" id="{…}">` twin in
  the sheet's `<ext uri="{78C0D931-…}"><x14:conditionalFormattings>` list with `minLength`, `maxLength`, `showValue`, `border`, `gradient`,
  `direction`, `negativeBarColorSameAsPositive`, `negativeBarBorderColorSameAsPositive`, `axisPosition`, the two `x14:cfvo` and `x14:fillColor` /
  `borderColor` / `negativeFillColor` / `negativeBorderColor` / `axisColor`. Colours are written as RGB (a theme colour of the file is resolved).
- Icon sets with custom icons or of the families `3Stars` / `3Triangles` / `5Boxes`, and rules whose formulas read another sheet, are written in that
  list only (`<x14:iconSet custom="1">` with `<x14:cfIcon>`; `<xm:f>` and `<x14:dxf>`); validations that read another sheet in `<ext
  uri="{CCE6A557-…}"><x14:dataValidations>`. Priorities are one numbering across both lists.
- Sparkline groups, slicer lists and anything else in a sheet's extension list are carried through untouched, as before.
