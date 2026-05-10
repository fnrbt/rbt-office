# Side-by-Side Comparison: Grid vs Word

## Text Styling

### RedBat.Grid (Spreadsheet)
```fsharp
// Cell-level formatting (one format per cell)
type CellFormat = {
    Bold: bool
    Italic: bool
    Underline: bool
    TextColor: string option
    BgColor: string option
    FontSize: int option          // ← int
    TextAlign: string option      // ← Grid-specific
    NumberFormat: string option   // ← Grid-specific
}

// One format per cell - no runs within cell
type CellData = {
    Value: CellValue
    Format: CellFormat
}
```

### RedBat.Word (Word Processor)
```fsharp
// Character-level formatting (runs within paragraph)
type TextStyle = {
    Bold: bool
    Italic: bool
    Underline: bool
    FontSize: float               // ← float (points)
    FontFamily: string option     // ← Word-specific
    Color: string option
}

// Multiple runs with different styles in one paragraph
type TextRun = {
    Start: int        // Character offset
    Length: int       // Number of characters
    Style: TextStyle  // Style for this range
}

type TextBuffer = {
    Rope: Rope
    Runs: TextRun list  // Non-overlapping styled runs
}
```

## Similarities (Can Share!)

| Feature | Grid | Word | Can Share? |
|---------|------|------|------------|
| Bold | ✓ | ✓ | ✅ YES |
| Italic | ✓ | ✓ | ✅ YES |
| Underline | ✓ | ✓ | ✅ YES |
| Color | ✓ (TextColor) | ✓ (Color) | ✅ YES |
| Font Size | ✓ (int) | ✓ (float) | ✅ YES (use float) |
| Background | ✓ (BgColor) | ✓ (implicit) | ✅ YES |

## Differences (Keep Separate!)

| Feature | Grid | Word | Should Share? |
|---------|------|------|---------------|
| Font Family | ✗ | ✓ | ⚠️ FUTURE (add to Grid) |
| Text Align | ✓ | ✗ | ❌ NO (Grid-specific) |
| Number Format | ✓ | ✗ | ❌ NO (Grid-specific) |
| Text Runs | ✗ | ✓ | ❌ NO (Word-specific) |
| Block Types | ✗ | ✓ (paragraphs/headings) | ❌ NO (Word-specific) |

## Rendering Approach

### Grid
```
DOM-based:
- Virtual scrolling
- Absolutely positioned divs
- One div per cell
- Apply CSS styles directly

┌─────────┬─────────┬─────────┐
│  Cell   │  Cell   │  Cell   │
│ (div)   │ (div)   │ (div)   │
├─────────┼─────────┼─────────┤
│  Cell   │  Cell   │  Cell   │
│ (div)   │ (div)   │ (div)   │
└─────────┴─────────┴─────────┘
```

### Word
```
Canvas-based:
- Text layout engine
- Line wrapping
- Measure text runs
- Draw styled segments

┌────────────────────────────┐
│ This is **bold** and       │
│ _italic_ text that wraps   │
│ across multiple lines.     │
└────────────────────────────┘
```

**Verdict:** ❌ Too different to share

## Selection Model

### Grid
```fsharp
// Cell-based selection
type Selection =
    | SingleCell of row: int * col: int
    | Range of startRow: int * startCol: int * endRow: int * endCol: int

// Select cells: A1, B2:D5
```

### Word
```fsharp
// Character-based selection
type Selection = {
    Anchor: Position  // Start of selection
    Focus: Position   // End/cursor
}

type Position = {
    BlockId: BlockId
    Offset: int  // Character offset
}

// Select characters: "Hello w|orld" (cursor at position 7)
```

**Verdict:** ❌ Too different to share (different granularity)

## Format Application

### Grid
```fsharp
// Apply to whole cell at once
let applyFormat formatType cellAddr =
    let cellData = getCell cellAddr model
    let newFormat =
        match formatType with
        | "bold" -> { cellData.Format with Bold = not cellData.Format.Bold }
        | ...
    setCell cellAddr { cellData with Format = newFormat } model
```

### Word
```fsharp
// Apply to character range
let applyStyle start len styleTransform buffer =
    // Split runs at boundaries
    // Apply transform to overlapping runs
    // Merge adjacent runs with same style
    { buffer with Runs = transformedRuns }

// Example: Bold characters 10-20
buffer |> applyStyle 10 10 (fun s -> { s with Bold = true })
```

**Difference:**
- Grid: Entire cell styled at once
- Word: Character ranges within text

## Toolbar Integration

### Grid (Current)
```javascript
const toolbar = createFormatToolbar((action) => {
    if (action === 'bold') {
        spreadsheet.ApplyFormat('bold');  // Applies to selected cell
    }
});
```

### Word (Proposed)
```javascript
const toolbar = createFormatToolbar((action) => {
    if (action === 'bold') {
        wordEditor.ApplyFormat('bold');  // Applies to selected characters
    }
});
```

**Verdict:** ✅ Same toolbar can work with both! Just different handlers.

## Unified TextStyle Proposal

```fsharp
module RedBat.Style

/// Shared text style (use in both Grid and Word)
type TextStyle = {
    // Common formatting
    Bold: bool
    Italic: bool
    Underline: bool
    Strikethrough: bool

    // Font properties
    FontSize: float        // in points (11.0 = 11pt)
    FontFamily: string option

    // Colors
    ForegroundColor: string option  // text color
    BackgroundColor: string option  // highlight/cell background
}

/// Grid extends with cell-specific properties
type CellFormat = {
    TextStyle: TextStyle          // ← Reuse shared style
    TextAlign: string option      // Grid-specific
    NumberFormat: string option   // Grid-specific
    VerticalAlign: string option  // Grid-specific
}

/// Word uses directly in runs
type TextRun = {
    Start: int
    Length: int
    Style: TextStyle  // ← Same type as Grid uses
}
```

## Usage Example: Copy Style Between Components

```fsharp
// Copy formatting from Grid cell to Word document
let cellStyle: RedBat.Style.TextStyle =
    cellData.Format.TextStyle

let wordRun: TextRun = {
    Start = 0
    Length = 10
    Style = cellStyle  // ✓ Type-safe, same structure
}

// Or vice versa
let runStyle: RedBat.Style.TextStyle =
    textRun.Style

let cellFormat: CellFormat = {
    TextStyle = runStyle  // ✓ Works!
    TextAlign = Some "left"
    NumberFormat = None
    VerticalAlign = None
}
```

## Shared Toolbar Example

```javascript
// Same toolbar component works with both!

// In Grid
const toolbar = createFormattingToolbar((action) => {
    match action with
    | "bold" -> grid.ApplyBold()
    | "italic" -> grid.ApplyItalic()
    | "textColor" -> grid.ApplyTextColor(color)
});

// In Word
const toolbar = createFormattingToolbar((action) => {
    match action with
    | "bold" -> word.ApplyBold()
    | "italic" -> word.ApplyItalic()
    | "textColor" -> word.ApplyTextColor(color)
});
```

Same toolbar UI, different targets. Perfect separation!

## Summary Table

| Component | Grid | Word | Share? | Notes |
|-----------|------|------|--------|-------|
| **TextStyle Type** | ✓ | ✓ | ✅ YES | Unify into RedBat.Style |
| **Toolbar UI** | ✓ | ✗ | ✅ YES | Grid has it, Word can reuse |
| **Color Picker** | ✗ | ✗ | ✅ YES | Neither has it, build shared |
| **Font Picker** | ✗ | ✗ | ✅ YES | Neither has it, build shared |
| **Rendering** | DOM | Canvas | ❌ NO | Too different |
| **Selection** | Cell | Char | ❌ NO | Different granularity |
| **Data Model** | Cells | Runs | ❌ NO | Different structure |

## Recommendation

**Create shared components:**
1. `RedBat.Core` - Shared types (TextStyle)
2. `RedBat.UI` - Shared UI (Toolbar, ColorPicker, FontPicker)

**Keep separate:**
- Rendering engines (DOM vs Canvas)
- Selection models (cell vs character)
- Data structures (sparse cells vs text runs)

**Result:** Consistent formatting UI across Grid and Word, while preserving their unique strengths.
