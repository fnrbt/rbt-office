# Shared Components Analysis: Grid vs Word Processor

Analysis of what can be shared between RedBat.Grid (spreadsheet) and RedBat.Word (word processor).

## Current State

### RedBat.Word (/home/dave/Dev/redbatnet-feature-word/components/RedBat.Word)

**Text Styling Model:**
```fsharp
type TextStyle = {
    Bold: bool
    Italic: bool
    Underline: bool
    FontSize: float              // in points (11.0 default)
    FontFamily: string option
    Color: string option
}
```

**Architecture:**
- Rope data structure (efficient text buffer)
- TextBuffer with styled runs (non-overlapping ranges)
- Canvas-based renderer
- Selection model (anchor/focus)

### RedBat.Grid (Current)

**Cell Formatting Model:**
```fsharp
type CellFormat = {
    Bold: bool
    Italic: bool
    Underline: bool
    TextColor: string option
    BgColor: string option
    FontSize: int option
    TextAlign: string option
    NumberFormat: string option
}
```

**Architecture:**
- Sparse cell storage (Map)
- Virtual scrolling renderer
- DOM-based rendering
- Single cell = one format (no runs within cell)

## What Can Be Shared

### 1. Text Style Type (★★★ High Priority)

**Opportunity:** Unify text styling model across both components.

**Proposed Shared Module:** `RedBat.Style` or `RedBat.Formatting`

```fsharp
module RedBat.Style

/// Core text styling (used by both Grid and Word)
type TextStyle = {
    Bold: bool
    Italic: bool
    Underline: bool
    Strikethrough: bool
    FontSize: float  // in points
    FontFamily: string option
    ForegroundColor: string option  // text color
    BackgroundColor: string option  // highlight/background
}

let defaultStyle = {
    Bold = false
    Italic = false
    Underline = false
    Strikethrough = false
    FontSize = 11.0
    FontFamily = None
    ForegroundColor = None
    BackgroundColor = None
}

/// Apply style transform
let applyBold style = { style with Bold = not style.Bold }
let applyItalic style = { style with Italic = not style.Italic }
let applyUnderline style = { style with Underline = not style.Underline }
let setFontSize size style = { style with FontSize = size }
let setForegroundColor color style = { style with ForegroundColor = Some color }
let setBackgroundColor color style = { style with BackgroundColor = Some color }
```

**Benefits:**
- Consistent styling across Grid and Word
- Shared toolbar components work with both
- Single source of truth for text formatting
- Type safety when passing styles between components

**Usage:**

*In RedBat.Grid:*
```fsharp
type CellFormat = {
    TextStyle: RedBat.Style.TextStyle  // Core text styling
    TextAlign: string option           // Grid-specific
    NumberFormat: string option        // Grid-specific
}
```

*In RedBat.Word:*
```fsharp
type TextRun = {
    Start: int
    Length: int
    Style: RedBat.Style.TextStyle  // Same type
}
```

### 2. Toolbar Components (★★★ High Priority)

**Opportunity:** Format toolbar already exists in Grid. Word processor can reuse it.

**Current Grid Toolbar:**
```fsharp
// Spreadsheet/Toolbar.fs
type FormatToolbar(onAction: string -> unit)
    // Bold, Italic, Underline, Font size, Colors
```

**Make Generic:**
```fsharp
module RedBat.UI.Toolbar

/// Generic formatting toolbar (works with Grid and Word)
type FormattingToolbar(onAction: string -> unit) =
    // Bold, Italic, Underline, Font size, Colors
    // Doesn't know about cells or documents
    // Just fires actions: "bold", "italic", etc.
```

**Usage:**

*In Grid:*
```javascript
const toolbar = createFormattingToolbar((action) => {
    spreadsheet.ApplyFormat(action);
});
```

*In Word:*
```javascript
const toolbar = createFormattingToolbar((action) => {
    wordEditor.ApplyFormat(action);
});
```

**Benefits:**
- One toolbar implementation
- Consistent UI across apps
- Maintain once, use everywhere

### 3. Color Picker (★★ Medium Priority)

**Opportunity:** Both need color selection UI.

**Current State:**
- Grid: Hardcoded colors in toolbar action handlers
- Word: Not yet implemented

**Proposed:**
```fsharp
module RedBat.UI.ColorPicker

type ColorPicker(onColor: string -> unit) =
    // Render color palette
    // Trigger callback on selection
```

**Benefits:**
- Consistent color UI
- Shared color palettes
- Single implementation

### 4. Font Picker (★★ Medium Priority)

**Opportunity:** Both need font family selection.

**Current State:**
- Grid: Not implemented
- Word: Supported in TextStyle

**Proposed:**
```fsharp
module RedBat.UI.FontPicker

type FontPicker(onFont: string -> unit) =
    // List of system fonts
    // Font preview
    // Search/filter
```

### 5. Selection Model (★ Lower Priority)

**Current State:**
- Grid: Single cell or range (row/col based)
- Word: Anchor/Focus model (character-based)

**Difference:** Too different to unify directly.

Grid selection is cell-based, Word is character-based. Keep separate.

### 6. Rendering (✗ Not Shareable)

**Grid:** DOM-based, virtual scrolling, absolutely positioned divs
**Word:** Canvas-based, text layout, line wrapping

These are fundamentally different and should stay separate.

## Recommended Shared Structure

```
components/
├── RedBat.Core/           (NEW - shared foundation)
│   ├── Style.fs          ← TextStyle type, style transforms
│   └── RedBat.Core.fsproj
│
├── RedBat.UI/             (NEW - shared UI components)
│   ├── Toolbar.fs        ← Generic formatting toolbar
│   ├── ColorPicker.fs    ← Color selection
│   ├── FontPicker.fs     ← Font selection
│   └── RedBat.UI.fsproj
│
├── RedBat.Grid/
│   ├── (existing files)
│   └── RedBat.Grid.fsproj ← Reference RedBat.Core, RedBat.UI
│
└── RedBat.Word/
    ├── (existing files)
    └── RedBat.Word.fsproj ← Reference RedBat.Core, RedBat.UI
```

## Migration Plan

### Phase 1: Extract TextStyle (Immediate)

1. Create `RedBat.Core` project
2. Define shared `TextStyle` type
3. Update Grid's `CellFormat` to use it
4. Update Word's `TextStyle` to use it

### Phase 2: Extract Toolbar (Near-term)

1. Create `RedBat.UI` project
2. Move `Toolbar.fs` from Grid to UI
3. Make it generic (no spreadsheet-specific logic)
4. Update Grid to reference shared toolbar
5. Use same toolbar in Word

### Phase 3: Add Pickers (Later)

1. Add `ColorPicker` to RedBat.UI
2. Add `FontPicker` to RedBat.UI
3. Integrate into both Grid and Word

## Specific Differences to Preserve

### Grid-Specific

- `NumberFormat` - Excel-style number formatting
- `TextAlign` - Cell alignment (left, center, right)
- Virtual scrolling
- Cell-based selection
- Formula bar integration

### Word-Specific

- Text runs (styled ranges within paragraph)
- Line wrapping and reflow
- Block structure (paragraphs, headings)
- Character-based selection
- Canvas rendering

## Benefits of Sharing

### Consistency

Users see same formatting UI in both Grid and Word:
- Same toolbar layout
- Same keyboard shortcuts
- Same color palettes
- Same font lists

### Maintenance

Fix a bug in toolbar? Both Grid and Word benefit.
Add a new color preset? Appears in both.

### Code Reuse

Formatting logic written once, tested once, used everywhere.

### Type Safety

Shared types prevent mismatches:
```fsharp
// Can safely pass styles between components
let cellStyle: RedBat.Style.TextStyle = cell.Format.TextStyle
let textRun: TextRun = { Start = 0; Length = 10; Style = cellStyle }
```

## Implementation Notes

### Dependencies

```
RedBat.Core (no deps)
    ↓
RedBat.UI (depends on Core)
    ↓
RedBat.Grid, RedBat.Word (depend on Core + UI)
```

### Fable Compatibility

All shared components must:
- Target `netstandard2.0`
- Use Fable-compatible F#
- Avoid .NET-specific APIs

### Versioning

Shared components should version independently:
- `RedBat.Core` v1.0.0
- `RedBat.UI` v1.0.0
- Grid and Word reference specific versions

## Current Compatibility

### What Works Today

✓ Both use Fable
✓ Both target netstandard2.0
✓ Similar styling models (easy to unify)
✓ Similar toolbar needs

### What Needs Work

- Extract shared types from both codebases
- Create new shared projects
- Update project references
- Test interoperability

## Questions for Architecture Decision

1. **Where to place shared projects?**
   - Option A: `components/RedBat.Core/`
   - Option B: `shared/RedBat.Core/`
   - Recommendation: `components/` (consistent with existing structure)

2. **Versioning strategy?**
   - Option A: All components share one version
   - Option B: Independent versioning
   - Recommendation: Independent (allows Grid/Word to evolve separately)

3. **Breaking changes?**
   - Option A: Migrate both at once
   - Option B: Support old and new styles temporarily
   - Recommendation: Migrate Grid first (smaller), then Word

## Conclusion

**High-value shares:**
1. ✅ TextStyle type (foundational)
2. ✅ Formatting toolbar (UI consistency)
3. ✅ Color/font pickers (when implemented)

**Keep separate:**
1. ✗ Rendering (too different)
2. ✗ Selection model (different granularity)
3. ✗ Data structures (cells vs text runs)

**Recommendation:** Create `RedBat.Core` and `RedBat.UI` shared projects, starting with TextStyle extraction.
