# Office Format Reader (F# + Fable)

A comprehensive MS Word (.docx) and Excel (.xlsx) file format reader written in F# that compiles to JavaScript via Fable.

## Features

### DOCX Reading
- ✅ **Paragraphs & Runs** with full formatting
- ✅ **Text Formatting**: Bold, Italic, Underline, Strike, Font size, Font family
- ✅ **Colors**: Font color and highlight color
- ✅ **Hyperlinks**: Extract links with targets
- ✅ **Images**: Extract image metadata with relationships
- ✅ **Tables**: Full table structure with nested content
- ✅ **Lists**: Numbered and bulleted list support
- ✅ **Styles**: Style definitions with cascading
- ✅ **Headers/Footers**: Document headers and footers
- ✅ **Document Properties**: Title, author, dates, description
- ✅ **Paragraph Properties**: Alignment, spacing, indentation

### XLSX Reading
- ✅ **Cell Values**: Numbers, strings, booleans, dates, formulas
- ✅ **Shared Strings**: Automatic resolution
- ✅ **Cell Styles**: Font colors, fill colors, bold, italic, font sizes
- ✅ **Number Formats**: Cell formatting codes
- ✅ **Merged Cells**: Extract merge regions
- ✅ **Column Info**: Column widths and visibility
- ✅ **Row Heights**: Row height information
- ✅ **Multiple Sheets**: Full workbook support

### F# Abstractions & DSLs
- **Active Patterns**: XML matching (`ElementNamed`, `HasAttr`, `TextValue`)
- **Computation Expressions**: Monadic navigation (`xml {}`, `promise {}`, `document {}`)
- **Lenses**: Composable accessors with `>->` operator
- **Point-Free Combinators**: `|>>`, `=>>`, `?>` for concise queries
- **Discriminated Unions**: Type-safe domain models
- **Higher-Order Functions**: Functional pipelines throughout

## Installation

```bash
npm install
```

## Usage

### Read Word Document

```javascript
import api from './dist/src/Main.js'

// From file input
const file = document.querySelector('input[type=file]').files[0]
const buffer = await file.arrayBuffer()

const doc = await api().word.load(buffer)
if (doc) {
  // Extract all text
  const text = api().word.extractText(doc)
  console.log(text)

  // Get structured content
  const paragraphs = api().word.paragraphs(doc)
  const tables = api().word.tables(doc)

  // Get formatting and metadata
  const styles = api().word.styles(doc)
  const properties = api().word.properties(doc)

  // Extract hyperlinks
  const links = api().word.hyperlinks(doc)
  links.forEach(link => {
    console.log(`${link.Text} -> ${link.Target}`)
  })

  // Extract images
  const images = api().word.images(doc)
  images.forEach(img => {
    console.log(`Image: ${img.Description}, ${img.Width}x${img.Height}`)
  })

  // Get headers and footers
  const headers = api().word.headers(doc)
  const footers = api().word.footers(doc)
}
```

### Access Formatting

```javascript
paragraphs.forEach(para => {
  console.log(`Alignment: ${para.Properties.Alignment}`)
  console.log(`Style: ${para.Properties.Style}`)

  para.Runs.forEach(run => {
    const props = run.Properties
    console.log({
      text: run.Content.tag === 'Text' ? run.Content.fields[0] : '',
      bold: props.Bold,
      italic: props.Italic,
      color: props.Color,
      fontSize: props.FontSize,
      fontFamily: props.FontFamily,
      highlight: props.Highlight
    })
  })
})
```

### Read Excel Workbook

```javascript
const file = document.querySelector('input[type=file]').files[0]
const buffer = await file.arrayBuffer()

const workbook = await api().excel.load(buffer)
if (workbook) {
  // Get sheet names
  const sheets = api().excel.sheetNames(workbook)
  console.log('Sheets:', sheets)

  // Get specific sheet
  const sheet = api().excel.sheet('Sheet1', workbook)

  // Convert to 2D array
  if (sheet) {
    const data = api().excel.toArray(workbook, sheet)
    console.log(data)

    // Get specific cell with value
    const cellValue = api().excel.cellAt('A1', sheet)
    console.log('A1:', cellValue)

    // Get cell with full style information
    const cell = api().excel.cellWithStyleAt('A1', sheet)
    if (cell && cell.Style) {
      console.log({
        value: cell.Value,
        fontColor: cell.Style.FontColor,
        fillColor: cell.Style.FillColor,
        bold: cell.Style.Bold,
        italic: cell.Style.Italic,
        fontSize: cell.Style.FontSize,
        numberFormat: cell.Style.NumberFormat
      })
    }

    // Get merged cells
    const merged = api().excel.mergedCells(sheet)
    merged.forEach(m => {
      console.log(`Merged: ${m.StartRef.Column}${m.StartRef.Row} to ${m.EndRef.Column}${m.EndRef.Row}`)
    })

    // Get column information
    const columns = api().excel.columns(sheet)
    columns.forEach(col => {
      console.log(`Column ${col.Index}: width=${col.Width}, hidden=${col.Hidden}`)
    })
  }
}
```

## Build

```bash
# Build F# project
dotnet build

# Compile to JavaScript with Fable
npm run build

# Watch mode
npm run watch
```

## Architecture

### Core Modules

1. **XmlDsl.fs** (140 lines) - XML navigation DSL
   - Active patterns for XML matching
   - Computation expression for optional chaining
   - Point-free combinators
   - Namespace-aware helpers

2. **OfficeDsl.fs** (80 lines) - Office document abstractions
   - Relationship parsing
   - Document builder computation expression
   - Lens pattern for nested access
   - Common text extraction patterns

3. **ZipInterop.fs** (75 lines) - Fable-compatible ZIP handling
   - JSZip interop
   - Promise monad
   - Archive loading and file extraction

4. **DocxReader.fs** (420 lines) - Word document reader
   - Full paragraph and run extraction
   - Style resolution from styles.xml
   - Hyperlink and image extraction
   - List numbering support
   - Header/footer parsing
   - Document properties

5. **XlsxReader.fs** (395 lines) - Excel workbook reader
   - Cell value extraction (all types)
   - Style parsing from styles.xml
   - Merged cell detection
   - Column and row metadata
   - Shared string resolution

6. **Main.fs** (170 lines) - Public API
   - Clean functional API
   - Example usage
   - Browser integration helpers

**Total: ~1,280 lines of F#**

## DSL Examples

### XML Navigation

```fsharp
// Active patterns
match node with
| ElementNamed "paragraph" (attrs, children) -> // ...
| _ -> ()

// Computation expression
xml {
    let! name = Xml.attr "name" node
    let! value = Xml.child "value" node |> Xml.text
    return (name, value)
}

// Combinators
node ?> "child" |>> Xml.text
nodes =>> (Xml.child "item")
```

### Document Building

```fsharp
let content = document {
    for para in paragraphs do
        yield Para para
    for table in tables do
        yield Tbl table
}
```

### Promise Monad

```fsharp
let parseWorkbook archive = promise {
    let! wbXml = getFileText archive "xl/workbook.xml"
    let! stylesXml = getFileText archive "xl/styles.xml"
    let! sheets = loadAllSheets wbXml
    return { Sheets = sheets; Styles = parseStyles stylesXml }
}
```

## Completeness

### DOCX Support: ~90%
- ✅ Text, formatting, colors, styles
- ✅ Hyperlinks, images
- ✅ Tables, lists
- ✅ Headers, footers
- ✅ Document properties
- ⚠️ Advanced features: comments, track changes, SmartArt, charts (not implemented)

### XLSX Support: ~85%
- ✅ Cell values, formulas
- ✅ Styles, colors, fonts
- ✅ Merged cells
- ✅ Column/row metadata
- ⚠️ Advanced features: charts, pivot tables, data validation, conditional formatting (not implemented)

## Key F# Features Used

1. **Active Patterns** - Pattern matching on XML structures
2. **Computation Expressions** - 3 custom: xml, promise, document
3. **Discriminated Unions** - Type-safe domain models
4. **Pattern Matching** - Exhaustive case analysis
5. **Lenses** - Composable nested accessors
6. **Higher-Order Functions** - map, filter, fold abstractions
7. **Option Type** - Explicit handling of missing values
8. **Records** - Immutable data structures
9. **Inline Functions** - Zero-cost abstractions
10. **Partial Application** - Point-free style

## DRY Principles

- Shared XML DSL for all readers
- Reusable namespace helpers
- Generic property extraction patterns
- Composable lens accessors
- Single promise monad for all async
- Common relationship parsing

## License

MIT

## Contributing

This is a demonstration of F# functional programming with DSLs. Contributions welcome!
