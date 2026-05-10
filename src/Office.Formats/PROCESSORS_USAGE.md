# DocumentProcessors Usage Guide

## Overview

The `DocumentProcessors` module transforms parsed Word documents into application-specific formats with fully resolved styles. It provides an editor-friendly format that's ready to use in web applications.

## Installation & Build

```bash
npm run build
```

This compiles the F# code to JavaScript in the `dist/` directory.

## API

### `processors.forEditor(doc)`

Transforms a WordDocument into an array of EditorParagraph objects with fully resolved styles.

**Input:** `WordDocument` (from `word.load()`)

**Output:** Array of `EditorParagraph` objects:

```typescript
type EditorParagraph = {
    Runs: EditorRun[],
    Alignment: "left" | "center" | "right" | "justify"
}

type EditorRun = {
    Text: string,
    Bold: boolean,
    Italic: boolean,
    Underline: boolean,
    Strike: boolean,
    FontSize: number,        // In points (e.g., 11.0, 12.0)
    FontFamily: string,      // E.g., "Calibri", "Arial"
    Color: string | null     // CSS format: "#FF0000" or null
}
```

## Browser Usage Example

```html
<!DOCTYPE html>
<html>
<head>
    <title>Word Document Editor</title>
</head>
<body>
    <input type="file" id="fileInput" accept=".docx">
    <div id="editor"></div>

    <script type="module">
        import officeFormat from './dist/src/Main.js';

        const api = officeFormat();

        document.getElementById('fileInput').addEventListener('change', async (e) => {
            const file = e.target.files[0];
            if (!file) return;

            // Load the document
            const doc = await api.word.load(file);

            if (doc) {
                // Process for editor
                const paragraphs = api.processors.forEditor(doc);

                // Render in the editor
                const editor = document.getElementById('editor');
                editor.innerHTML = '';

                paragraphs.forEach(para => {
                    const paraEl = document.createElement('div');
                    paraEl.style.textAlign = para.Alignment;
                    paraEl.style.marginBottom = '1em';

                    para.Runs.forEach(run => {
                        const span = document.createElement('span');
                        span.textContent = run.Text;

                        // Apply formatting
                        span.style.fontSize = `${run.FontSize}pt`;
                        span.style.fontFamily = run.FontFamily;

                        if (run.Bold) {
                            span.style.fontWeight = 'bold';
                        }
                        if (run.Italic) {
                            span.style.fontStyle = 'italic';
                        }
                        if (run.Underline) {
                            span.style.textDecoration = 'underline';
                        }
                        if (run.Strike) {
                            span.style.textDecoration = 'line-through';
                        }
                        if (run.Color) {
                            span.style.color = run.Color;
                        }

                        paraEl.appendChild(span);
                    });

                    editor.appendChild(paraEl);
                });
            }
        });
    </script>
</body>
</html>
```

## Key Features

### 1. **Style Resolution**

The processor resolves Word's complex style inheritance hierarchy:

```
Document Defaults (11pt Calibri)
    ↓ override
Base Style (via BasedOn chain)
    ↓ override
Paragraph Style
    ↓ override
Direct Formatting
    ↓ result
Fully Resolved Properties
```

### 2. **Type Conversion**

- **Font sizes**: Converted from Word half-points to actual points
  - Word: "24" → Output: 12.0
- **Colors**: Normalized to CSS format
  - Word: "FF0000" → Output: "#FF0000"
  - Word: "auto" → Output: null
- **Alignment**: Normalized to web standards
  - Word: "both" → Output: "justify"
- **Font families**: Theme fonts resolved
  - Word: "+majorFont" → Output: "Calibri"

### 3. **Content Flattening**

- Tables are flattened into a linear sequence of paragraphs
- All content (body, table cells) is processed uniformly
- Empty runs are automatically filtered out

### 4. **Safety Features**

- Handles circular style references (A → B → A)
- Gracefully defaults when styles are missing
- Converts malformed values to sensible defaults

## Architecture

```
Parsing Layer              Processing Layer           Application
┌────────────┐            ┌────────────────┐         ┌──────────┐
│ DocxReader │────────────>│ Processors     │────────>│ RedBat   │
│            │            │                │         │ Word     │
│ Faithful   │            │ Style          │         │ Editor   │
│ OOXML      │            │ Resolution     │         └──────────┘
│ Extraction │            │ Type           │
└────────────┘            │ Conversion     │
                          └────────────────┘
```

## Implementation Details

### TypeConversion Module

Pure functions for converting OOXML strings to typed values:
- `fontSizeToPoints`: Half-points → points
- `normalizeColor`: Word color → CSS color
- `normalizeAlignment`: Word alignment → web alignment
- `normalizeFontFamily`: Theme fonts → actual fonts

### StyleResolution Module

Implements the style cascade:
- `buildStyleMap`: Pre-computes all style inheritance
- `resolveRunProps`: Applies cascade for text properties
- `resolveParagraphProps`: Applies cascade for paragraph properties
- `mergeRunProps`: Combines base + override properties

### EditorProcessor Module

Transforms to editor format:
- `extractRunText`: Extracts text from various run content types
- `toEditorRun`: Converts Run → EditorRun
- `toEditorParagraph`: Converts Paragraph → EditorParagraph
- `extractForEditor`: Main entry point

## Testing in Browser

Since the library is designed for browser use, testing requires a browser environment. To test:

1. Build the project: `npm run build`
2. Create an HTML file with the example above
3. Serve it with a local server (e.g., `npx serve .`)
4. Open in browser and load a .docx file

## Future Enhancements

The architecture allows for easy addition of new processors:

- **MarkdownProcessor**: Convert documents to Markdown
- **SearchProcessor**: Build search indexes
- **PlainTextProcessor**: Extract plain text with metadata
- **DiffProcessor**: Compare document versions

Each processor can be added as a new module without modifying the parsing layer.
