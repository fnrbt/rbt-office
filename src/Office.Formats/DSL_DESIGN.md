# DSL Design & F# Abstraction Features

This document explains the F# abstractions and DSLs used throughout the Office format reader implementation.

## 1. XML Navigation DSL

### Active Patterns for Pattern Matching
```fsharp
// Match elements by name
let (|ElementNamed|_|) name = function
    | Element(n, attrs, children) when n = name -> Some(attrs, children)
    | _ -> None

// Usage in pattern matching
match node with
| ElementNamed "paragraph" (attrs, children) -> processContent children
| _ -> None
```

### Computation Expression for Optional Chaining
```fsharp
type XmlBuilder() =
    member _.Bind(node, f) = match node with Some n -> f n | None -> None
    member _.Return(x) = Some x

let xml = XmlBuilder()

// Usage - monadic navigation with automatic None handling
xml {
    let! name = Xml.attr "name" node
    let! value = Xml.child "value" node |> Xml.text
    return (name, value)
}
```

### Point-Free Combinators
```fsharp
// Custom operators for concise navigation
let (|>>) node f = Option.bind f (Some node)    // Single element pipeline
let (=>>) nodes f = List.choose f nodes         // Multiple elements filter
let (?>) node name = child name node            // Child accessor

// Usage
node ?> "body" |>> Xml.text                     // Get body's text
paragraphs =>> (Xml.attr "style")               // Get all paragraph styles
```

## 2. Office Document DSL

### Lens Pattern for Nested Access
```fsharp
type Lens<'a, 'b> = {
    Get: 'a -> 'b option
    GetAll: 'a -> 'b list
}

let (>->) l1 l2 = {                             // Lens composition
    Get = fun x -> l1.Get x |> Option.bind l2.Get
    GetAll = fun x -> l1.GetAll x |> List.collect l2.GetAll
}

// Usage - deep navigation through nested structures
let deepLens = workbookLens >-> sheetLens >-> rowLens >-> cellLens
let allCells = Lens.getAll deepLens workbook
```

### Document Builder Computation Expression
```fsharp
type DocumentBuilder() =
    member _.Yield(x) = [x]
    member _.Combine(a, b) = a @ b
    member _.For(xs, f) = List.collect f xs

let document = DocumentBuilder()

// Usage - declarative content building
let content = document {
    for para in paragraphs do
        yield Para para
    for table in tables do
        yield Tbl table
}
```

## 3. DOCX-Specific DSL

### Namespace-Aware Helpers (DRY)
```fsharp
module Docx =
    // Reusable namespace-aware element access
    let inline wElem name node = Ns.childNs name node
    let inline wElems name node = Ns.descendantsNs name node

    // Property extraction with defaults
    let hasProp propName node =
        wElem propName node |> Option.isSome

    let getPropVal propName attrName node =
        wElem propName node |> Option.bind (Xml.attr attrName)
```

### Functional Composition for Properties
```fsharp
// Compose property extractors
let runProps node =
    let rPr = wElem "rPr" node
    match rPr with
    | Some props ->
        {
            Bold = hasProp "b" props
            Italic = hasProp "i" props
            Underline = hasProp "u" props
            FontSize = getPropVal "sz" "val" props
        }
    | None -> defaultRunProps
```

## 4. XLSX-Specific DSL

### Cell Reference Parsing
```fsharp
// Mini-DSL for cell reference handling
let parseCellRef (ref: string) : CellReference =
    let letters = ref |> Seq.takeWhile System.Char.IsLetter |> ...
    let digits = ref |> Seq.skipWhile System.Char.IsLetter |> ...
    { Column = letters; Row = int digits }
```

### Value Discriminated Union
```fsharp
type CellValue =
    | Number of float
    | String of string
    | SharedString of int
    | Boolean of bool
    | Formula of string
    | Empty

// Pattern matching for type-safe value extraction
let cellValue (cNode: XmlNode) =
    let cellType = Xml.attr "t" cNode
    match cellType, valueNode with
    | Some "s", Some v -> SharedString (int v)
    | Some "b", Some v -> Boolean (v = "1")
    | Some "str", Some v -> String v
    | _, Some v -> Number (float v)
    | _ -> Empty
```

## 5. Promise Monad (for Async)

### Computation Expression for Promises
```fsharp
type PromiseBuilder() =
    member _.Bind(p, f) = Promise.bind f p
    member _.Return(x) = Promise.ret x

let promise = PromiseBuilder()

// Usage - async operations with do-notation style
let parseWorkbook archive = promise {
    let! wbXml = getFileText archive "xl/workbook.xml"
    let! sstXml = getFileText archive "xl/sharedStrings.xml"
    let! sheets = loadAllSheets wbXml
    return { Sheets = sheets; SharedStrings = sst }
}
```

## 6. Higher-Order Functions & Pipelines

### Text Extraction with Functional Pipeline
```fsharp
let extractText (doc: WordDocument) : string =
    doc.Content
    |> List.map contentText
    |> String.concat "\n\n"

// Recursive extraction with pattern matching
let rec contentText = function
    | Para p -> paraText p
    | Tbl t -> tableText t
```

## 7. Type-Safe Domain Models

### Discriminated Unions for Content Types
```fsharp
type DocumentContent =
    | Para of Paragraph
    | Tbl of Table

// Filtering with pattern matching
let paragraphs (doc: WordDocument) =
    doc.Content
    |> List.choose (function
        | Para p -> Some p
        | _ -> None)
```

## Key F# Features Leveraged

1. **Active Patterns**: Pattern matching on XML structures
2. **Computation Expressions**: Monadic navigation (xml, promise, document builders)
3. **Discriminated Unions**: Type-safe domain models
4. **Pattern Matching**: Exhaustive case analysis
5. **Partial Application**: Point-free style combinators
6. **Higher-Order Functions**: map, filter, fold abstractions
7. **Option Type**: Explicit handling of missing values
8. **Records**: Immutable data structures
9. **Modules**: Namespace organization
10. **Inline Functions**: Zero-cost abstractions

## DRY Principles Applied

1. **Shared XML Combinator**: Single implementation used across all readers
2. **Namespace Helpers**: Reusable namespace-stripping logic
3. **Property Extraction**: Generic prop/attr access patterns
4. **Lens Composition**: Reusable deep navigation
5. **Promise Monad**: Single async abstraction for all I/O
6. **Relationship Parsing**: Shared across DOCX and XLSX

## Result: Succinct, Type-Safe, Functional

The codebase achieves high abstraction with:
- Zero runtime reflection
- Compile-time type safety
- Fable-compatible JavaScript output
- Minimal boilerplate through DSLs
- Composable, reusable components
