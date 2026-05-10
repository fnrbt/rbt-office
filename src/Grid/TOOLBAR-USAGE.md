# Toolbar Usage Guide

The spreadsheet component provides toolbar components that apps can place, configure, and connect as needed.

## Overview

Toolbars are provided by the Fable module but apps control:
- **Placement**: Where toolbars appear in the UI
- **Configuration**: Which actions are enabled
- **Connection**: How toolbar actions affect the spreadsheet

## Available Toolbars

### FormatToolbar
Text formatting controls:
- Bold (B)
- Italic (I)
- Underline (U)
- Font increase (A+)
- Font decrease (A-)
- Text color (🎨)
- Background color (⬛)

### AlignmentToolbar
Text alignment controls:
- Align left (⬅)
- Align center (↔)
- Align right (➡)

### NumberFormatToolbar
Number formatting controls:
- Number format (1,234.56)
- Currency ($)
- Percent (%)
- Increase decimals (.0+)
- Decrease decimals (.0-)

## Basic Usage

```javascript
import createSpreadsheet from './Spreadsheet.js';
import { createFormatToolbar, createAlignmentToolbar, createNumberFormatToolbar } from './Toolbar.js';

// Create spreadsheet
const spreadsheet = createSpreadsheet(container);
spreadsheet.Mount();

// Create toolbar with action handler
const formatToolbar = createFormatToolbar((action) => {
    console.log('Format action:', action);

    // Handle color actions specially
    if (action === 'textColor') {
        spreadsheet.ApplyTextColor('#ff0000');
    } else if (action === 'bgColor') {
        spreadsheet.ApplyBgColor('#ffff00');
    } else {
        // Other actions pass through to ApplyFormat
        spreadsheet.ApplyFormat(action);
    }
});

// Mount toolbar to container
formatToolbar.Mount(toolbarContainer);
```

## Spreadsheet Format Methods

### ApplyFormat(formatType: string)
Apply text formatting to selected cell:
- `"bold"` - Toggle bold
- `"italic"` - Toggle italic
- `"underline"` - Toggle underline
- `"fontIncrease"` - Increase font size by 2px
- `"fontDecrease"` - Decrease font size by 2px
- `"alignLeft"` - Align left
- `"alignCenter"` - Align center
- `"alignRight"` - Align right

### ApplyTextColor(color: string)
Set text color (CSS color value):
```javascript
spreadsheet.ApplyTextColor('#ff0000');  // Red
spreadsheet.ApplyTextColor('rgb(255, 0, 0)');  // Also red
```

### ApplyBgColor(color: string)
Set background color (CSS color value):
```javascript
spreadsheet.ApplyBgColor('#ffff00');  // Yellow
```

### ApplyNumberFormat(format: string)
Set number format:
- `"number"` - Number format (1,234.56)
- `"currency"` - Currency format ($1,234.56)
- `"percent"` - Percent format (50%)

## Placement Examples

### Horizontal Toolbar (Default)
```javascript
const toolbar = createFormatToolbar(handleAction);
toolbar.Mount(document.getElementById('top-toolbar'));
```

### Sidebar (Vertical)
```javascript
// Override toolbar styles for vertical layout
const style = document.createElement('style');
style.textContent = `
    .sidebar .toolbar {
        flex-direction: column;
        align-items: stretch;
    }
    .sidebar .toolbar-button {
        margin: 4px 0;
    }
`;
document.head.appendChild(style);

const toolbar = createFormatToolbar(handleAction);
toolbar.Mount(document.getElementById('sidebar'));
```

### Multiple Toolbars
```javascript
// Place different toolbars in different locations
const formatToolbar = createFormatToolbar(handleFormat);
formatToolbar.Mount(sidebarElement);

const numberToolbar = createNumberFormatToolbar(handleNumber);
numberToolbar.Mount(topBarElement);

const alignmentToolbar = createAlignmentToolbar(handleAlign);
alignmentToolbar.Mount(bottomBarElement);
```

## Custom Action Handlers

Apps can customize how toolbar actions work:

```javascript
const formatToolbar = createFormatToolbar((action) => {
    // Custom color picker
    if (action === 'textColor') {
        const color = showColorPicker();  // Your custom UI
        spreadsheet.ApplyTextColor(color);
        return;
    }

    // Custom confirmation
    if (action === 'bold') {
        if (confirm('Make text bold?')) {
            spreadsheet.ApplyFormat('bold');
        }
        return;
    }

    // Default handling
    spreadsheet.ApplyFormat(action);
});
```

## Test Examples

See the test files for complete examples:

- **test-spreadsheet.html**: Simple layout with all three toolbars in one row
- **test-custom-toolbar.html**: Custom layout with toolbars in sidebar and top bar

## Notes

- Toolbars operate on the **currently selected cell**
- If no cell is selected, format actions have no effect
- Format state is stored per-cell in the spreadsheet model
- Toolbars do not show active state (apps can implement if needed)
- Button styles can be overridden with CSS
