# TODO: Unify Components with RedBat.Word

Status: **PLANNING** - To be implemented later

## Summary

RedBat.Grid (spreadsheet) and RedBat.Word (word processor) have similar text formatting needs. This document outlines the unification plan.

## What's Already Analyzed

✓ `SHARED-COMPONENTS.md` - Detailed sharing analysis
✓ `COMPARISON-GRID-VS-WORD.md` - Side-by-side feature comparison

## Quick Reference

### Word Processor Location
```
/home/dave/Dev/redbatnet-feature-word/components/RedBat.Word/
```

### Current Similar Types

**Grid (CellFormat):**
```fsharp
type CellFormat = {
    Bold: bool
    Italic: bool
    Underline: bool
    TextColor: string option
    BgColor: string option
    FontSize: int option
    TextAlign: string option      // Grid-specific
    NumberFormat: string option   // Grid-specific
}
```

**Word (TextStyle):**
```fsharp
type TextStyle = {
    Bold: bool
    Italic: bool
    Underline: bool
    FontSize: float
    FontFamily: string option
    Color: string option
}
```

## Unification Checklist

### Phase 1: Create Shared Projects

- [ ] Create `components/RedBat.Core/`
  - [ ] Create `RedBat.Core.fsproj`
  - [ ] Add `Style.fs` with unified `TextStyle` type
  - [ ] Target `netstandard2.0`
  - [ ] Add `Fable.Core` dependency

- [ ] Create `components/RedBat.UI/`
  - [ ] Create `RedBat.UI.fsproj`
  - [ ] Reference `RedBat.Core`
  - [ ] Target `netstandard2.0`
  - [ ] Add Fable dependencies

### Phase 2: Extract Shared TextStyle

- [ ] Define unified `RedBat.Core.Style.TextStyle`:
  ```fsharp
  type TextStyle = {
      Bold: bool
      Italic: bool
      Underline: bool
      Strikethrough: bool
      FontSize: float  // in points
      FontFamily: string option
      ForegroundColor: string option
      BackgroundColor: string option
  }
  ```

### Phase 3: Extract Toolbar to RedBat.UI

- [ ] Move `Spreadsheet/Toolbar.fs` to `RedBat.UI/Toolbar.fs`
- [ ] Make it generic (remove spreadsheet-specific logic)
- [ ] Update Grid to reference `RedBat.UI.Toolbar`
- [ ] Update Word to reference `RedBat.UI.Toolbar`

### Phase 4: Update Grid

- [ ] Add project reference to `RedBat.Core`
- [ ] Add project reference to `RedBat.UI`
- [ ] Update `CellFormat` to use `RedBat.Core.Style.TextStyle`:
  ```fsharp
  type CellFormat = {
      TextStyle: RedBat.Core.Style.TextStyle  // Shared
      TextAlign: string option                 // Grid-specific
      NumberFormat: string option              // Grid-specific
  }
  ```
- [ ] Update formatting methods to work with shared type
- [ ] Update toolbar integration
- [ ] Test all formatting features

### Phase 5: Update Word (Coordinate with Word Team)

- [ ] Add project reference to `RedBat.Core`
- [ ] Add project reference to `RedBat.UI`
- [ ] Update `TextRun` to use `RedBat.Core.Style.TextStyle`
- [ ] Update formatting methods
- [ ] Integrate shared toolbar
- [ ] Test all formatting features

### Phase 6: Add Shared Pickers

- [ ] Create `RedBat.UI/ColorPicker.fs`
  - [ ] Color palette component
  - [ ] Custom color input
  - [ ] Recent colors

- [ ] Create `RedBat.UI/FontPicker.fs`
  - [ ] System font list
  - [ ] Font preview
  - [ ] Search/filter

- [ ] Integrate pickers into both Grid and Word

## Project Structure (After Unification)

```
components/
├── RedBat.Core/
│   ├── Style.fs              ← Shared TextStyle
│   └── RedBat.Core.fsproj
│
├── RedBat.UI/
│   ├── Toolbar.fs            ← Shared formatting toolbar
│   ├── ColorPicker.fs        ← Shared color picker
│   ├── FontPicker.fs         ← Shared font picker
│   └── RedBat.UI.fsproj
│
├── RedBat.Grid/
│   ├── Spreadsheet/Model.fs  ← Uses RedBat.Core.Style
│   └── RedBat.Grid.fsproj    ← References Core + UI
│
└── RedBat.Word/              (other repo)
    ├── TextBuffer.fs         ← Uses RedBat.Core.Style
    └── RedBat.Word.fsproj    ← References Core + UI
```

## Benefits Summary

**For Users:**
- Consistent formatting UI across spreadsheet and word processor
- Same keyboard shortcuts
- Same color palettes
- Learn once, use everywhere

**For Developers:**
- Single source of truth for text styling
- Fix bugs once
- Add features once
- Type-safe style passing between components

**For Maintenance:**
- Less code duplication
- Easier to test
- Single toolbar implementation
- Shared UI components

## What Stays Separate

✗ Rendering engines (DOM vs Canvas)
✗ Selection models (cell-based vs character-based)
✗ Data structures (sparse cells vs text runs)
✗ Grid-specific features (NumberFormat, TextAlign)
✗ Word-specific features (Paragraphs, Headings, TextRuns)

## Notes

- Word processor is in separate repo: `/home/dave/Dev/redbatnet-feature-word/`
- Coordinate with Word team before Phase 5
- Test thoroughly - formatting is core functionality
- Keep backward compatibility during migration
- Document breaking changes if any

## Resources

- `SHARED-COMPONENTS.md` - Full analysis
- `COMPARISON-GRID-VS-WORD.md` - Feature comparison
- `ARCHITECTURE.md` - Grid architecture
- Word processor code: `/home/dave/Dev/redbatnet-feature-word/components/RedBat.Word/`

## Status

**Current:** Planning phase, notes saved
**Next:** Create RedBat.Core and RedBat.UI projects
**Timeline:** TBD - implement when ready to unify

---

*Last Updated: 2026-01-27*
*Note: Do not modify Word processor files - coordinate with Word team*
