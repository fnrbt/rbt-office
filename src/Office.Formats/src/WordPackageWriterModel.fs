module WordPackageWriterModel

open System
open System.Collections.Generic
open XmlGenerator
open WordEditorPackage

let roundToInt (value: float) : int =
    if Double.IsNaN value || Double.IsInfinity value then 0
    elif value >= 0.0 then int (Math.Floor (value + 0.5))
    else -(int (Math.Floor (0.5 - value)))

let pixelsToTwips value = string (roundToInt (value * 15.0))
let pixelsToPoints value = value * 72.0 / 96.0
let pixelsToEmus value = roundToInt (value * 9525.0)

module InverseConversion =
    let pixelsToHalfPoints pixels = string (max 1 (roundToInt (pixels * 1.5)))
    let pointsToHalfPoints points = string (roundToInt (points * 2.0))
    let pixelsToTwips pixels = string (roundToInt (pixels * 15.0))
    let pixelsToEighthPoints pixels = string (max 0 (roundToInt (pixels * 6.0)))
    let cssColorToWord color = Option.map ModelText.wordColor color
    let alignmentToWord (alignment: string) =
        match alignment.ToLowerInvariant() with
        | "center" -> "center"
        | "right" | "end" -> "right"
        | "justify" | "both" -> "both"
        | "distribute" -> "distribute"
        | _ -> "left"
    let pointsToTwips points = string (roundToInt (points * 20.0))
    let lineSpacingToWord spacing = string (roundToInt (spacing * 240.0))
    let fixedLineSpacingToTwips fontSizePixels spacing = string (roundToInt (spacing * max 1.0 fontSizePixels * 15.0))

type PageSetupInfo = {
    PageWidth: float
    PageHeight: float
    MarginLeft: float
    MarginRight: float
    MarginTop: float
    MarginBottom: float
    HeaderDistance: float
    FooterDistance: float
    DifferentFirstPage: bool
    EvenAndOddHeaders: bool
    ColumnCount: int
    ColumnSpace: float
    ColumnSeparator: bool
    ColumnWidths: (float * float) list
    SectionStart: string option
    HeaderReferences: (string * string) list
    FooterReferences: (string * string) list
    PageNumberStart: int option
    PageNumberFormat: string option
    VerticalAlign: string option
}

let defaultPageSetupInfo = {
    PageWidth = 816.0; PageHeight = 1056.0
    MarginLeft = 96.0; MarginRight = 96.0; MarginTop = 96.0; MarginBottom = 96.0
    HeaderDistance = 48.0; FooterDistance = 48.0
    DifferentFirstPage = false; EvenAndOddHeaders = false
    ColumnCount = 1; ColumnSpace = 0.0; ColumnSeparator = false; ColumnWidths = []
    SectionStart = None; HeaderReferences = []; FooterReferences = []
    PageNumberStart = None; PageNumberFormat = None; VerticalAlign = None
}

let pageSetupInfo (setup: WordEditorPackage.PageSetup) : PageSetupInfo =
    let columns = setup.Columns
    { PageWidth = setup.PageWidth; PageHeight = setup.PageHeight
      MarginLeft = setup.MarginLeft; MarginRight = setup.MarginRight
      MarginTop = setup.MarginTop; MarginBottom = setup.MarginBottom
      HeaderDistance = setup.HeaderDistance; FooterDistance = setup.FooterDistance
      DifferentFirstPage = setup.DifferentFirstPage; EvenAndOddHeaders = setup.EvenAndOddHeaders
      ColumnCount = columns |> Option.map (fun x -> x.Count) |> Option.defaultValue 1
      ColumnSpace = columns |> Option.map (fun x -> x.Gap) |> Option.defaultValue 0.0
      ColumnSeparator = columns |> Option.exists (fun x -> x.Separator)
      ColumnWidths = columns |> Option.map (fun x -> x.Widths) |> Option.defaultValue []
      SectionStart = setup.SectionStart
      HeaderReferences = setup.HeaderReferences |> List.map (fun r -> r.Kind, r.RelationshipId)
      FooterReferences = setup.FooterReferences |> List.map (fun r -> r.Kind, r.RelationshipId)
      PageNumberStart = setup.PageNumberStart; PageNumberFormat = setup.PageNumberFormat
      VerticalAlign = setup.VerticalAlign }

let optionalPageSetup (value: WordEditorPackage.PageSetup option) = Option.map pageSetupInfo value

let sectionColumnsElement (setup: PageSetupInfo) =
    if setup.ColumnCount < 2 then None
    else
        let attrs =
            [ yield "w:num", string setup.ColumnCount
              if setup.ColumnSeparator then yield "w:sep", "1"
              yield "w:space", pixelsToTwips setup.ColumnSpace
              if not setup.ColumnWidths.IsEmpty then yield "w:equalWidth", "0" ]
        let columns = setup.ColumnWidths |> List.map (fun (width, space) -> selfClosing "w:col" ["w:w", pixelsToTwips width; "w:space", pixelsToTwips space])
        Some (element "w:cols" attrs columns)

let sectionPropertiesChildren (references: string list) (sectionType: string option) (setup: PageSetupInfo option) (pageNumberStart: int option) (pageNumberFormat: string option) =
    [ yield! references
      match sectionType with
      | Some value when List.contains value ["continuous"; "nextPage"; "evenPage"; "oddPage"; "nextColumn"] -> yield selfClosing "w:type" ["w:val", value]
      | _ -> ()
      match setup with
      | Some value ->
          yield selfClosing "w:pgSz" [ yield "w:w", pixelsToTwips value.PageWidth; yield "w:h", pixelsToTwips value.PageHeight; if value.PageWidth > value.PageHeight then yield "w:orient", "landscape" ]
          yield selfClosing "w:pgMar" ["w:top", pixelsToTwips value.MarginTop; "w:right", pixelsToTwips value.MarginRight; "w:bottom", pixelsToTwips value.MarginBottom; "w:left", pixelsToTwips value.MarginLeft; "w:header", pixelsToTwips value.HeaderDistance; "w:footer", pixelsToTwips value.FooterDistance; "w:gutter", "0"]
      | None -> ()
      match pageNumberStart, pageNumberFormat with
      | None, None -> ()
      | start, format ->
          yield selfClosing "w:pgNumType" [
              match format with Some v when v <> "" -> yield "w:fmt", v | _ -> ()
              match start with Some v -> yield "w:start", string (max 0 v) | None -> () ]
      match setup |> Option.bind sectionColumnsElement with Some columns -> yield columns | None -> ()
      match setup |> Option.bind (fun x -> x.VerticalAlign) with Some v when List.contains v ["center"; "both"; "bottom"] -> yield selfClosing "w:vAlign" ["w:val", v] | _ -> ()
      match setup with Some x when x.DifferentFirstPage -> yield selfClosing "w:titlePg" [] | _ -> () ]

type PreservedRelationshipInfo = WordEditorPackage.Relationship
type PreservedPackage = WordEditorPackage.PreservedPackage

let emptyPreservedPackage : PreservedPackage =
    { Parts = []; RootRelationships = []; DocumentRelationshipId = "rId1"
      PartRelationships = Map.empty; Defaults = []; Overrides = []; MainContentType = None
      StylesXml = None; SettingsXml = None; SourceXml = Map.empty; SourceRelationships = Map.empty
      SourceParts = []; BodyContentChunks = [||]; BodyItemChunks = [||]; BodyChunkCount = 0
      PartItemChunks = Map.empty; NoteItemChunks = Map.empty; PartPaths = Map.empty
      OriginalText = Map.empty; OpenedBytes = None }

type CommentRef = { Id: string; Text: string; Author: string option; Date: string option }
type NoteRef = { Id: string; NoteType: string; Text: string }
type ImageAsset = { Src: string; RelId: string; FileName: string; Extension: string; ContentType: string; Data: byte[] option; Linked: bool }
type ImageScope = { Assets: ImageAsset array; AssetBySrc: Map<string, ImageAsset> }
let emptyImageScope = { Assets = [||]; AssetBySrc = Map.empty }

type ParagraphSource =
    abstract Of: EditorParagraph -> string option
    abstract SpanOf: EditorParagraph -> (int * int) option
type ParagraphOrigin =
    abstract Of: EditorParagraph -> string option
type ParagraphSplicer =
    abstract Splice: para: EditorParagraph -> pieces: (EditorRun * string * (string -> string)) list -> runXml: (EditorRun -> string) -> wrapOf: (EditorRun -> string -> string) -> (string list * string option * string * bool) option
    abstract MergeProperties: source: string -> was: string -> now: string -> string option
/// A table of the body as it was read, for a table that is written from the model (its SourceTable): the
/// text of its properties and grid, and of each of its rows and cells by where its element starts in the
/// part (the origin the model has of it: SourceRow, SourceCell); the layout of each (XmlPatch.layout: its
/// rows, cells, paragraphs among what stands around them - content controls, custom XML, marks, white
/// space). And how the properties that were read are changed (see mergedProperties).
/// CopyStartTag: the start tag for a copy - without the ids that are the element's alone (w14:paraId, w14:textId).
type SourceCellText = { CellXml: string; CellStartTag: string; CopyStartTag: string; CellProps: string; CellHead: string; CellChildren: XmlPatch.Layout }
/// RowHead / TableHead: the text from the start tag to the end of the properties (and grid), as it was read.
type SourceRowText = { RowXml: string; RowStartTag: string; CopyStartTag: string; RowExceptions: string; RowProps: string; RowHead: string; RowChildren: XmlPatch.Layout }
type SourceTableText = {
    /// What the table is in, as it was read (a content control, custom XML; the white space before it): to the table's start, and after the table.
    WrapStart: string
    WrapEnd: string
    TableHead: string
    TableStartTag: string
    TableProps: string
    TableGrid: string
    TableRows: System.Collections.Generic.Dictionary<int, SourceRowText>
    TableCells: System.Collections.Generic.Dictionary<int, SourceCellText>
    TableChildren: XmlPatch.Layout
    /// The table's element in the part's tree (w:tbl).
    TableNode: XmlDsl.XmlNode
}
type TableSource =
    abstract Of: WordEditorPackage.TableSource -> SourceTableText option
    /// The table of a w:tbl of the part being written (nothing around it: WrapStart and WrapEnd are "").
    abstract At: XmlDsl.XmlNode -> SourceTableText option
    /// container, the properties that were read, what the model wrote before and writes now -> the properties, or None (see mergedProperties).
    abstract Merge: string * string * string list * string list -> string option
type WriterContext = {
    mutable ReservedRelationshipIds: Set<string>
    mutable ReservedPartPaths: Set<string>
    mutable CurrentPartRelationshipIds: Set<string>
    mutable FieldResultParts: EditorRun list
    mutable MoveRunsLeft: Map<string, int>
    mutable MoveRangeIds: Map<string, string>
    mutable NextMoveRangeId: int
    mutable DefaultListNumIds: string * string
    mutable ParagraphSource: ParagraphSource option
    mutable ParagraphsInContext: bool
    mutable ReportGenerated: int
    mutable ReportDerived: int
    mutable ReportSpliced: int
    mutable ReportPatched: int
    mutable LiveCommentIds: Set<string>
    mutable DeletedCommentIds: Set<string>
    mutable ParagraphOrigin: ParagraphOrigin option
    mutable SpliceSourceText: string
    mutable ParagraphSplicer: ParagraphSplicer option
    mutable TableSource: TableSource option
    mutable CurrentSourceTable: SourceTableText option
    mutable NextShapeId: int
    mutable ShapesFromModel: bool
    CarriedSources: Dictionary<EditorParagraph, string>
    SpliceFieldParts: Dictionary<EditorRun, EditorRun list>
}

let createWriterContext () = {
    ReservedRelationshipIds = Set.empty; ReservedPartPaths = Set.empty; CurrentPartRelationshipIds = Set.empty
    FieldResultParts = []; MoveRunsLeft = Map.empty; MoveRangeIds = Map.empty; NextMoveRangeId = 900000
    DefaultListNumIds = "1", "2"; ParagraphSource = None; ParagraphsInContext = true
    ReportGenerated = 0; ReportDerived = 0; ReportSpliced = 0; ReportPatched = 0
    LiveCommentIds = Set.empty; DeletedCommentIds = Set.empty; ParagraphOrigin = None
    SpliceSourceText = ""; ParagraphSplicer = None; TableSource = None; CurrentSourceTable = None
    NextShapeId = 1000; ShapesFromModel = false
    CarriedSources = Dictionary<EditorParagraph, string>(HashIdentity.Reference)
    SpliceFieldParts = Dictionary<EditorRun, EditorRun list>(HashIdentity.Reference)
}

let freshRelationshipId (context: WriterContext) candidate =
    let mutable id = candidate
    while Set.contains id context.ReservedRelationshipIds do id <- id + "x"
    id
let preservedRelationshipElement (rel: PreservedRelationshipInfo) =
    selfClosing "Relationship" [ yield "Id", rel.Id; yield "Type", rel.Type; yield "Target", rel.Target; if rel.External then yield "TargetMode", "External" ]

let rec collectItemParagraphs (item: EditorBlock) : EditorParagraph list =
    match item with
    | Paragraph p -> [p]
    | TextBox shape -> shape.Paragraphs |> Array.toList |> List.collect collectItemParagraphs
    | Table table ->
        [ for row in table.Rows do
            for cell in row.Cells do
                for block in cell.Paragraphs do yield! collectItemParagraphs block
                for nested in cell.NestedTables do yield! collectItemParagraphs (Table nested) ]

let runHyperlinkTarget (run: WordEditorPackage.EditorRun) = Option.orElse run.InactiveHyperlink run.Hyperlink
let collectParagraphHyperlinks (para: EditorParagraph) = para.Runs |> Array.choose runHyperlinkTarget |> Array.filter (ModelText.isAnchor >> not) |> Array.toList
let collectItemHyperlinks item = collectItemParagraphs item |> List.collect collectParagraphHyperlinks
let hyperlinkMapWith (context: WriterContext) (existing: PreservedRelationshipInfo list) (items: EditorBlock array) =
    let known = existing |> List.filter (fun rel -> rel.External && OpcXml.relationshipKind rel.Type = "hyperlink")
    items |> Array.toList |> List.collect collectItemHyperlinks |> List.distinct
    |> List.mapi (fun index url ->
        match known |> List.tryFind (fun rel -> rel.Target = url) with
        | Some rel -> url, rel.Id
        | None -> url, freshRelationshipId context (sprintf "rIdLink%d" (index + 1))) |> Map.ofList
let hyperlinkMap context items = hyperlinkMapWith context [] items

let collectParagraphComments (para: EditorParagraph) =
    [ for run in para.Runs do
        match run.CommentId with
        | Some id -> yield { CommentRef.Id = id; Text = run.CommentText |> Option.defaultValue (sprintf "Comment %s" id); Author = run.CommentAuthor; Date = run.CommentDate }
        | None -> () ]
let collectParagraphNotes (para: EditorParagraph) =
    [ for run in para.Runs do
        match run.NoteId with
        | Some id ->
            let kind = run.NoteType |> Option.defaultValue "footnote"
            yield { NoteRef.Id = id; NoteType = kind; Text = run.NoteText |> Option.defaultValue (sprintf "%s %s" kind id) }
        | None -> () ]
let collectComments items = items |> Array.toList |> List.collect collectItemParagraphs |> List.collect collectParagraphComments |> List.distinctBy (fun x -> x.Id)
let collectNotes noteType items = items |> Array.toList |> List.collect collectItemParagraphs |> List.collect collectParagraphNotes |> List.filter (fun x -> x.NoteType = noteType) |> List.distinctBy (fun x -> x.Id)
let collectSectionReferenceIds items =
    items |> Array.toList |> List.collect collectItemParagraphs
    |> List.collect (fun p -> p.HeaderReferences @ p.FooterReferences)
    |> List.map (fun r -> r.RelationshipId) |> Set.ofList

let mimeToExtensionAndContentType (mime: string) =
    match mime.ToLowerInvariant() with
    | "image/jpeg" | "image/jpg" -> Some ("jpeg", "image/jpeg")
    | "image/png" -> Some ("png", "image/png")
    | "image/gif" -> Some ("gif", "image/gif")
    | "image/bmp" -> Some ("bmp", "image/bmp")
    | "image/svg+xml" -> Some ("svg", "image/svg+xml")
    | "image/webp" -> Some ("webp", "image/webp")
    | "image/x-emf" | "image/emf" -> Some ("emf", "image/x-emf")
    | "image/x-wmf" | "image/wmf" -> Some ("wmf", "image/x-wmf")
    | "image/tiff" -> Some ("tiff", "image/tiff")
    | _ -> None
let parseDataUrl src =
    match ModelText.dataUrl src with
    | Some { MediaType = mime; Base64 = Some data } ->
        mimeToExtensionAndContentType mime |> Option.map (fun (extension, contentType) -> extension, contentType, Convert.FromBase64String data)
    | _ -> None
let collectParagraphImageSources (para: EditorParagraph) =
    [ for run in para.Runs do
        match run.Content with
        | Picture image when image.Opaque.IsNone -> yield! Option.toList image.Source
        | _ -> () ]
let collectItemImageSources item = collectItemParagraphs item |> List.collect collectParagraphImageSources
let buildImageScope (context: WriterContext) mediaStartIndex (items: EditorBlock array) =
    let mutable nextMediaIndex = mediaStartIndex
    let assets =
        items |> Array.toList |> List.collect collectItemImageSources |> List.distinct
        |> List.choose (fun src ->
            parseDataUrl src
            |> Option.map (fun (extension, contentType, data) ->
                while Set.contains (sprintf "word/media/image%d.%s" nextMediaIndex extension) context.ReservedPartPaths do nextMediaIndex <- nextMediaIndex + 1
                let fileName = sprintf "image%d.%s" nextMediaIndex extension
                let relId = freshRelationshipId context (sprintf "rIdImage%d" nextMediaIndex)
                nextMediaIndex <- nextMediaIndex + 1
                { Src = src; RelId = relId; FileName = fileName; Extension = extension; ContentType = contentType; Data = Some data; Linked = false })
            |> Option.orElseWith (fun () ->
                if ModelText.isLinkedSource src then
                    let relId = freshRelationshipId context (sprintf "rIdImage%d" nextMediaIndex)
                    nextMediaIndex <- nextMediaIndex + 1
                    Some { Src = src; RelId = relId; FileName = ""; Extension = ""; ContentType = ""; Data = None; Linked = true }
                else None)) |> List.toArray
    { Assets = assets; AssetBySrc = assets |> Array.map (fun asset -> asset.Src, asset) |> Map.ofArray }, nextMediaIndex

// Run formatting and content serialization. Conversion helpers are shared with
// the paragraph and section writers.
let toggle name isOn =
    if isOn then selfClosing name [] else selfClosing name [("w:val", "0")]

let runToggle (formatting: RunFormatting) name explicitValue effective =
    if formatting.HasExplicit then
        match explicitValue with
        | Some value -> [toggle name value]
        | None -> []
    else [toggle name effective]

let orderedRunProps (run: EditorRun) : string list =
    let f = run.Formatting
    [ match f.Style with
      | Some style when style <> "" -> yield selfClosing "w:rStyle" [("w:val", style)]
      | _ -> ()
      if not (f.HasExplicit && f.FontFamilyInherited = Some true) then
          yield selfClosing "w:rFonts" [("w:ascii", f.FontFamily); ("w:hAnsi", f.FontFamily); ("w:cs", f.FontFamily)]
      yield! runToggle f "w:b" f.ExplicitBold f.Bold
      yield! runToggle f "w:i" f.ExplicitItalic f.Italic
      yield! runToggle f "w:caps" f.ExplicitAllCaps f.AllCaps
      yield! runToggle f "w:smallCaps" f.ExplicitSmallCaps f.SmallCaps
      yield! runToggle f "w:strike" f.ExplicitStrike f.Strike
      yield! runToggle f "w:dstrike" f.ExplicitDoubleStrike f.DoubleStrike
      yield! runToggle f "w:vanish" f.ExplicitHidden f.Hidden
      match InverseConversion.cssColorToWord f.Color with
      | Some color -> yield selfClosing "w:color" [("w:val", color)]
      | None -> if f.ExplicitColorAuto then yield selfClosing "w:color" [("w:val", "auto")]
      if abs f.CharacterSpacing >= 0.001 then
          yield selfClosing "w:spacing" [("w:val", InverseConversion.pixelsToTwips f.CharacterSpacing)]
      if abs f.Position >= 0.001 then
          yield selfClosing "w:position" [("w:val", InverseConversion.pointsToHalfPoints f.Position)]
      match (if f.HasExplicit then f.ExplicitFontSize else Some f.FontSize) with
      | Some size ->
          let sz = InverseConversion.pixelsToHalfPoints size
          yield selfClosing "w:sz" [("w:val", sz)]
          yield selfClosing "w:szCs" [("w:val", sz)]
      | None -> ()
      match f.Highlight with
      | Some color -> yield selfClosing "w:highlight" [("w:val", color)]
      | None -> ()
      let underline = if f.HasExplicit then f.ExplicitUnderline else Some f.Underline
      if underline = Some true then
          let style = match f.UnderlineStyle with Some value when value <> "" && value <> "none" -> value | _ -> "single"
          let color = match InverseConversion.cssColorToWord f.UnderlineColor with Some value when value <> "AUTO" -> [("w:color", value)] | _ -> []
          yield selfClosing "w:u" (("w:val", style) :: color)
      elif underline = Some false then
          yield selfClosing "w:u" [("w:val", "none")]
      match f.VertAlign with
      | Some value when value = "superscript" || value = "subscript" -> yield selfClosing "w:vertAlign" [("w:val", value)]
      | _ -> ()
      match f.EmphasisMark with
      | Some value when value <> "" && value <> "none" -> yield selfClosing "w:em" [("w:val", value)]
      | _ -> () ]

let generateRunProps (run: EditorRun) = orderedRunProps run

let modelRunProperties (run: EditorRun) =
    match run.Content with
    | Picture _ | Opaque _ -> []
    | _ when run.NoProps -> []
    | _ -> orderedRunProps run

let runTextElements deleted (text: string) =
    let name = if deleted then "w:delText" else "w:t"
    let textElement value = elementWithText name (if XmlPatch.needsPreserve value then [("xml:space", "preserve")] else []) value
    if isNull text || text = "" then [textElement ""]
    else
        ModelText.textPieces text
        |> List.map (function
            | ModelText.Text value -> textElement value
            | ModelText.Tab -> selfClosing "w:tab" []
            | ModelText.Break -> selfClosing "w:br" [])

let runElement rPr deleted text =
    element "w:r" [] ((if List.isEmpty rPr then [] else [element "w:rPr" [] rPr]) @ runTextElements deleted text)

let runText (run: EditorRun) =
    match run.Content with
    | Text text | Equation (text, _) -> text
    | _ -> ""

let isDeletedRun (run: EditorRun) =
    run.Revision |> Option.bind (fun revision -> revision.Kind) = Some "delete"

let generateRun (run: EditorRun) =
    runElement (generateRunProps run) false (runText run)

let generateRunWithObjectProps (run: EditorRun) =
    let rPr = orderedRunProps run
    let deleted = isDeletedRun run
    if run.RenderedPageBreakBefore && not deleted then
        element "w:r" [] ((if List.isEmpty rPr then [] else [element "w:rPr" [] rPr]) @ [selfClosing "w:lastRenderedPageBreak" []] @ runTextElements deleted (runText run))
    else runElement rPr deleted (runText run)

// Field result parts are supplied by the paragraph writer through the per-save context.
let generateFieldRun (context: WriterContext) (run: EditorRun) (fieldCode: string) =
    match context.FieldResultParts with
    | [] ->
        let text = runText run
        let placeholder = ModelText.isFieldPlaceholder text
        let result = run.FieldResult |> Option.defaultValue (if placeholder then "" else text)
        let result = if not placeholder && text <> "" && text.Length < result.Length && result.Contains text then text else result
        element "w:fldSimple" [("w:instr", fieldCode)] [generateRunWithObjectProps { run with Content = Text result }]
    | parts ->
        context.FieldResultParts <- []
        element "w:fldSimple" [("w:instr", fieldCode)] (generateRunWithObjectProps run :: (parts |> List.map generateRunWithObjectProps))

let runBookmarkStarts (run: EditorRun) = run.BookmarkStarts
let runBookmarkEnds (run: EditorRun) = run.BookmarkEnds
let hasBookmarks (run: EditorRun) = not run.BookmarkStarts.IsEmpty || not run.BookmarkEnds.IsEmpty

let bookmarkStartElements (start: ModelText.BookmarkStart) =
    let name = if start.Name <> "" then start.Name else "bookmark" + start.Id
    selfClosing "w:bookmarkStart" [("w:id", start.Id); ("w:name", name)] + (if start.EndsHere then selfClosing "w:bookmarkEnd" [("w:id", start.Id)] else "")

let bookmarkStartXml (run: EditorRun) =
    match run.BookmarkStarts with
    | [] -> None
    | starts -> Some (starts |> List.map bookmarkStartElements |> String.concat "")

let bookmarkEndXml (run: EditorRun) =
    let ends =
        match run.BookmarkEnds with
        | [] when run.FieldCode.IsSome ->
            match run.BookmarkStarts with
            | [start] when not start.EndsHere -> [ModelText.Ends start.Id]
            | _ -> []
        | ends -> ends
    match ends with
    | [] -> None
    | ends ->
        ends |> List.map (function ModelText.Opens start -> bookmarkStartElements start | ModelText.Ends id -> selfClosing "w:bookmarkEnd" [("w:id", id)]) |> String.concat "" |> Some

let private generateFieldBoundaryRun (run: EditorRun) (boundary: ModelText.FieldBoundary) =
    let child =
        match boundary with
        | ModelText.Begin -> selfClosing "w:fldChar" [ "w:fldCharType", "begin" ]
        | ModelText.Separate -> selfClosing "w:fldChar" [ "w:fldCharType", "separate" ]
        | ModelText.End -> selfClosing "w:fldChar" [ "w:fldCharType", "end" ]
        | ModelText.Instruction code -> elementWithText "w:instrText" [ "xml:space", "preserve" ] code
    let props = modelRunProperties run
    element "w:r" [] ((if List.isEmpty props then [] else [element "w:rPr" [] props]) @ [child])

let generateReferenceRun (context: WriterContext) (run: EditorRun) =
    let inner =
        match run.Content with
        | FieldBoundary boundary -> generateFieldBoundaryRun run boundary
        | Equation (_, xml) -> xml
        | _ ->
            match run.FieldCode with
            | Some fieldCode when fieldCode <> "" -> generateFieldRun context run fieldCode
            | _ ->
                match run.CommentId, run.NoteId with
                | Some commentId, _ ->
                    let f = run.Formatting
                    let ownColor = f.Color |> Option.filter (fun color -> color.ToLowerInvariant() <> DocxReader.commentMarkDisplayColor)
                    let rPr = orderedRunProps { run with Formatting = { f with Color = ownColor } }
                    element "w:r" [] ((if List.isEmpty rPr then [] else [element "w:rPr" [] rPr]) @ [selfClosing "w:commentReference" [("w:id", commentId)]])
                | None, Some noteId ->
                    let name = if run.NoteType = Some "endnote" then "w:endnoteReference" else "w:footnoteReference"
                    let rPr = orderedRunProps run
                    element "w:r" [] ((if List.isEmpty rPr then [] else [element "w:rPr" [] rPr]) @ [selfClosing name [("w:id", noteId)]])
                | None, None ->
                    if runText run = "" && (Option.isSome (bookmarkStartXml run) || Option.isSome (bookmarkEndXml run)) then ""
                    else generateRunWithObjectProps run
    [bookmarkStartXml run; (if inner <> "" then Some inner else None); bookmarkEndXml run] |> List.choose id |> String.concat ""

let revisionAttributes id author date =
    [ yield ("w:id", id |> Option.defaultValue "1")
      yield ("w:author", author |> Option.defaultValue "")
      match date with Some value -> yield ("w:date", value) | None -> () ]

// Reset counts and range IDs at each story boundary, and count all moves in the
// story before emitting it. The numeric counter persists across stories within the save.
let moveKey (run: EditorRun) =
    match run.Revision with
    | Some revision ->
        match revision.Kind, revision.MoveName with
        | Some kind, Some name when kind = "moveFrom" || kind = "moveTo" -> Some (kind + "\u0000" + name, kind, name)
        | _ -> None
    | None -> None

let resetMoveRanges (context: WriterContext) (runs: EditorRun seq) =
    context.MoveRangeIds <- Map.empty
    context.MoveRunsLeft <-
        runs |> Seq.fold (fun counts run ->
            match moveKey run with
            | Some (key, _, _) -> Map.add key (1 + (Map.tryFind key counts |> Option.defaultValue 0)) counts
            | None -> counts) Map.empty

let revisionWrappedRun (context: WriterContext) (run: EditorRun) (runXml: string) =
    match run.Revision with
    | None -> runXml
    | Some revision ->
        let attrs = revisionAttributes revision.Id revision.Author revision.Date
        match revision.Kind with
        | Some "insert" -> element "w:ins" attrs [runXml]
        | Some "delete" ->
            let deletion = element "w:del" attrs [runXml]
            match revision.Inserted with
            | Some inserted when inserted.Id.IsSome || inserted.Author.IsSome || inserted.Date.IsSome ->
                element "w:ins" (revisionAttributes inserted.Id inserted.Author inserted.Date) [deletion]
            | _ -> deletion
        | Some kind when kind = "moveFrom" || kind = "moveTo" ->
            let wrapped = element ("w:" + kind) attrs [runXml]
            match moveKey run with
            | Some (key, _, name) ->
                let rangeId, rangeStart =
                    match Map.tryFind key context.MoveRangeIds with
                    | Some id -> id, ""
                    | None ->
                        let id = string context.NextMoveRangeId
                        context.NextMoveRangeId <- context.NextMoveRangeId + 1
                        context.MoveRangeIds <- Map.add key id context.MoveRangeIds
                        id, selfClosing ("w:" + kind + "RangeStart") (revisionAttributes (Some id) revision.Author revision.Date @ [("w:name", name)])
                let left = (Map.tryFind key context.MoveRunsLeft |> Option.defaultValue 1) - 1
                context.MoveRunsLeft <- Map.add key left context.MoveRunsLeft
                let rangeEnd = if left <= 0 then selfClosing ("w:" + kind + "RangeEnd") [("w:id", rangeId)] else ""
                rangeStart + wrapped + rangeEnd
            | None -> wrapped
        | _ -> runXml

let private wrapHyperlink (hyperlinks: Map<string, string>) (run: EditorRun) xml =
    match run.Hyperlink |> Option.orElse run.InactiveHyperlink with
    | Some target ->
        match ModelText.anchorOf target with
        | Some anchor -> element "w:hyperlink" [("w:anchor", anchor); ("w:history", "1")] [xml]
        | None ->
            match Map.tryFind target hyperlinks with
            | Some id -> element "w:hyperlink" [("r:id", id); ("w:history", "1")] [xml]
            | None -> xml
    | None -> xml

let wrapRun (context: WriterContext) (hyperlinks: Map<string, string>) (run: EditorRun) inner =
    wrapHyperlink hyperlinks run (revisionWrappedRun context run inner)

/// A hyperlink enclosing a field encloses all its tokens, rather than a
/// separate hyperlink around every fldChar/instrText/result run.
let groupedHyperlinkRuns (context: WriterContext) (hyperlinks: Map<string, string>) (runs: (EditorRun * string * (string -> string)) list) =
    let output = ResizeArray<EditorRun * string>()
    let content = System.Text.StringBuilder()
    let mutable first: EditorRun option = None
    let flush () =
        match first with
        | Some run ->
            output.Add((run, wrapHyperlink hyperlinks run (content.ToString())))
            content.Clear() |> ignore
            first <- None
        | None -> ()
    for run, inner, _ in runs do
        let target = runHyperlinkTarget run
        match first with
        | Some previous when target.IsSome && target = runHyperlinkTarget previous && run.ContentControls = previous.ContentControls -> ()
        | _ -> flush ()
        let xml = revisionWrappedRun context run inner
        match target with
        | Some _ ->
            if first.IsNone then first <- Some run
            content.Append(xml) |> ignore
        | None -> output.Add((run, xml))
    flush ()
    List.ofSeq output

let flipAttributes flipH flipV =
    [ if flipH then yield "flipH", "1"
      if flipV then yield "flipV", "1" ]

let vmlFlip flipH flipV =
    match flipH, flipV with
    | true, true -> Some "x y"
    | true, false -> Some "x"
    | false, true -> Some "y"
    | false, false -> None

let anchorWrapElement (wrap: string option) (wrapText: string option) (points: float array) =
    let side = Option.defaultValue "bothSides" wrapText
    let number (value: float) = string (int (Math.Round value))
    let polygon =
        if points.Length >= 6 then
            element "wp:wrapPolygon" [ "edited", "0" ] [
                for i in 0 .. points.Length / 2 - 1 ->
                    selfClosing (if i = 0 then "wp:start" else "wp:lineTo") [ "x", number points.[i * 2]; "y", number points.[i * 2 + 1] ] ]
        else
            element "wp:wrapPolygon" [ "edited", "0" ] [
                selfClosing "wp:start" [ "x", "0"; "y", "0" ]
                selfClosing "wp:lineTo" [ "x", "0"; "y", "21600" ]
                selfClosing "wp:lineTo" [ "x", "21600"; "y", "21600" ]
                selfClosing "wp:lineTo" [ "x", "21600"; "y", "0" ]
                selfClosing "wp:lineTo" [ "x", "0"; "y", "0" ] ]
    match wrap with
    | Some "none" -> selfClosing "wp:wrapNone" []
    | Some "topAndBottom" -> selfClosing "wp:wrapTopAndBottom" []
    | Some "tight" -> element "wp:wrapTight" [ "wrapText", side ] [ polygon ]
    | Some "through" -> element "wp:wrapThrough" [ "wrapText", side ] [ polygon ]
    | _ -> selfClosing "wp:wrapSquare" [ "wrapText", side ]

let anchorPositionElement name relativeFrom (align: string option) offset =
    match align with
    | Some value when value <> "" -> element name [ "relativeFrom", relativeFrom ] [ elementWithText "wp:align" [] value ]
    | _ -> element name [ "relativeFrom", relativeFrom ] [ elementWithText "wp:posOffset" [] offset ]

let generateImageRun (asset: ImageAsset) (image: WordEditorPackage.Image) =
    let measure defaultValue value = value |> Option.defaultValue defaultValue |> pixelsToEmus |> string
    let flag defaultValue value = if Option.defaultValue defaultValue value then "1" else "0"
    let width, height = measure 96.0 image.Width, measure 48.0 image.Height
    let altAttrs = image.Alt |> Option.map (fun alt -> [ "descr", alt ]) |> Option.defaultValue []
    let titleAttrs = image.Title |> Option.map (fun title -> [ "title", title ]) |> Option.defaultValue []
    let imageName = Option.defaultValue asset.FileName image.Name
    let cropAttrs =
        [ "l", image.CropLeft; "t", image.CropTop; "r", image.CropRight; "b", image.CropBottom ]
        |> List.choose (fun (name, value) -> Option.map (fun v -> name, v) value)
    let drawingContainer children =
        if image.Floating then
            let points = image.WrapPolygon |> List.collect (fun (x, y) -> [x; y]) |> List.toArray
            element "wp:anchor"
                [ "xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                  "distT", measure 0.0 image.DistanceTop; "distB", measure 0.0 image.DistanceBottom
                  "distL", measure 12.0 image.DistanceLeft; "distR", measure 12.0 image.DistanceRight
                  "simplePos", flag false image.SimplePosition; "relativeHeight", Option.defaultValue "0" image.RelativeHeight
                  "behindDoc", flag false image.BehindDoc; "locked", flag false image.Locked
                  "layoutInCell", flag true image.LayoutInCell; "allowOverlap", flag true image.AllowOverlap ]
                ([ selfClosing "wp:simplePos" [ "x", measure 0.0 image.SimplePositionX; "y", measure 0.0 image.SimplePositionY ]
                   anchorPositionElement "wp:positionH" (Option.defaultValue "column" image.PositionHRelativeFrom) image.PositionHAlign (measure 0.0 image.PositionX)
                   anchorPositionElement "wp:positionV" (Option.defaultValue "paragraph" image.PositionVRelativeFrom) image.PositionVAlign (measure 0.0 image.PositionY)
                   selfClosing "wp:extent" [ "cx", width; "cy", height ]
                   anchorWrapElement image.Wrap image.WrapText points ] @ children)
        else
            element "wp:inline" [ "xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ]
                (selfClosing "wp:extent" [ "cx", width; "cy", height ] :: children)
    let id = asset.RelId.Replace("rIdImage", "").TrimEnd('x')
    element "w:r" [] [
        element "w:drawing" [] [
            drawingContainer [
                selfClosing "wp:docPr" ([ "id", id; "name", imageName ] @ altAttrs @ titleAttrs)
                element "a:graphic" [ "xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main" ] [
                    element "a:graphicData" [ "uri", "http://schemas.openxmlformats.org/drawingml/2006/picture" ] [
                        element "pic:pic" [ "xmlns:pic", "http://schemas.openxmlformats.org/drawingml/2006/picture" ] [
                            element "pic:nvPicPr" [] [
                                selfClosing "pic:cNvPr" ([ "id", id; "name", imageName ] @ altAttrs @ titleAttrs)
                                selfClosing "pic:cNvPicPr" [] ]
                            element "pic:blipFill" [] [
                                selfClosing "a:blip" [ (if asset.Linked then "r:link" else "r:embed"), asset.RelId ]
                                if not cropAttrs.IsEmpty then selfClosing "a:srcRect" cropAttrs
                                element "a:stretch" [] [ selfClosing "a:fillRect" [] ] ]
                            element "pic:spPr" [] [
                                element "a:xfrm" (flipAttributes image.FlipH image.FlipV) [
                                    selfClosing "a:off" [ "x", "0"; "y", "0" ]
                                    selfClosing "a:ext" [ "cx", width; "cy", height ] ]
                                selfClosing "a:prstGeom" [ "prst", "rect" ] ] ] ] ] ] ] ]

let includeHeaderFooterKind _differentFirstPage _evenAndOddHeaders kind =
    kind = "default" || kind = "first" || kind = "even"

let sectionReferences elementName differentFirstPage evenAndOddHeaders (references: HeaderFooterReference list) =
    references
    |> List.distinctBy (fun item -> item.Kind)
    |> List.choose (fun item ->
        if includeHeaderFooterKind differentFirstPage evenAndOddHeaders item.Kind then
            Some (selfClosing elementName [ "w:type", item.Kind; "r:id", item.RelationshipId ])
        else None)

let paragraphSectionProperties (packageSetup: PageSetupInfo) (para: EditorParagraph) =
    if para.SectionBreak.IsNone && para.SectionPageSetup.IsNone && para.HeaderReferences.IsEmpty && para.FooterReferences.IsEmpty then None
    else
        let setup =
            para.SectionPageSetup |> Option.map pageSetupInfo
            |> Option.defaultValue { packageSetup with DifferentFirstPage = false; ColumnCount = 1; ColumnWidths = []; SectionStart = None }
            |> fun setup ->
                { setup with HeaderReferences = []; FooterReferences = []
                             PageNumberStart = None; PageNumberFormat = None
                             VerticalAlign = para.SectionVerticalAlign }
            |> fun setup ->
                match para.SectionColumns with
                | None -> setup
                | Some columns ->
                    { setup with ColumnCount = columns.Count; ColumnSpace = columns.Gap
                                 ColumnSeparator = columns.Separator; ColumnWidths = columns.Widths }
        let refs =
            sectionReferences "w:headerReference" setup.DifferentFirstPage setup.EvenAndOddHeaders para.HeaderReferences
            @ sectionReferences "w:footerReference" setup.DifferentFirstPage setup.EvenAndOddHeaders para.FooterReferences
        Some (element "w:sectPr" [] (sectionPropertiesChildren refs para.SectionBreak (Some setup) para.SectionPageNumberStart para.SectionPageNumberFormat))

let defaultFontSizePixels = 11.0 * 96.0 / 72.0
let unsetFontSizePixels = 10.0 * 96.0 / 72.0
let documentDefaultsStyleId = "__docDefaults"

let paragraphFontSize (para: EditorParagraph) =
    match para.Formatting.ParagraphMarkFontSize with
    | Some size when size > 0.0 -> size
    | _ ->
        let sizes =
            para.Runs |> Array.choose (fun run ->
                match run.Content with Picture _ | Opaque _ -> None | _ -> Some run.Formatting.FontSize)
        if sizes.Length = 0 then unsetFontSizePixels else Array.max sizes

let paragraphBorderElement name (value: Border option) =
    value |> Option.map (fun border ->
        let style = Option.defaultValue "single" border.Style
        let hidden = style = "none" || style = "nil"
        let width = border.Width |> Option.map InverseConversion.pixelsToEighthPoints |> Option.defaultValue (if hidden then "0" else "4")
        let space = border.Space |> Option.map (fun amount -> string (max 0 (roundToInt amount))) |> Option.defaultValue "0"
        let color = InverseConversion.cssColorToWord border.Color |> Option.defaultValue "auto"
        selfClosing name [ "w:val", style; "w:sz", width; "w:space", space; "w:color", color ])

let paragraphBordersElement (value: Borders option) =
    value |> Option.bind (fun borders ->
        let sides =
            [ paragraphBorderElement "w:top" borders.Top; paragraphBorderElement "w:left" borders.Left
              paragraphBorderElement "w:bottom" borders.Bottom; paragraphBorderElement "w:right" borders.Right ]
            |> List.choose id
        if sides.IsEmpty then None else Some (element "w:pBdr" [] sides))

let tabStopsElement (para: EditorParagraph) =
    let tab alignment leader position =
        selfClosing "w:tab" [
            yield "w:val", (if alignment = "" then "left" else alignment)
            match leader with Some value when value <> "" && value <> "none" -> yield "w:leader", value | _ -> ()
            yield "w:pos", InverseConversion.pixelsToTwips position ]
    let stops = para.Formatting.TabStops |> List.map (fun stop -> tab stop.Alignment stop.Leader stop.Position)
    let stops =
        if not stops.IsEmpty then stops
        else para.Formatting.RightTabPosition |> Option.map (tab "right" para.Formatting.RightTabLeader >> List.singleton) |> Option.defaultValue []
    if stops.IsEmpty then None else Some (element "w:tabs" [] stops)

let statesExplicitFormat (para: EditorParagraph) = para.Formatting.ExplicitParagraphFormat
let writtenMeasure (para: EditorParagraph) explicitValue effective =
    if statesExplicitFormat para then explicitValue else effective

let paragraphMarkRunProps (para: EditorParagraph) =
    let f = para.Formatting
    let pinned =
        match f.LineSpacingRule, writtenMeasure para f.ExplicitLineSpacing f.LineSpacing with
        | Some rule, Some _ when (rule = "exact" || rule = "atLeast") && not (statesExplicitFormat para) -> Some (paragraphFontSize para)
        | _ -> None
    let stated = if f.HasMarkProperties then Some f.ExplicitMarkFontSize else None
    let tracked =
        match f.MarkRevision with
        | Some revision when revision.Kind = Some "insert" || revision.Kind = Some "delete" ->
            [ selfClosing (if revision.Kind = Some "delete" then "w:del" else "w:ins")
                  (revisionAttributes revision.Id revision.Author revision.Date) ]
        | _ -> []
    let size value =
        let sz = InverseConversion.pixelsToHalfPoints value
        [ selfClosing "w:sz" [ "w:val", sz ]; selfClosing "w:szCs" [ "w:val", sz ] ]
    match pinned, stated with
    | Some value, _ | None, Some (Some value) -> Some (element "w:rPr" [] (tracked @ size value))
    | None, Some None -> Some (element "w:rPr" [] tracked)
    | None, None -> if tracked.IsEmpty then None else Some (element "w:rPr" [] tracked)

let generateParaPropsWithPageSetup (context: WriterContext) (packageSetup: PageSetupInfo) (para: EditorParagraph) =
    let f = para.Formatting
    let alignment =
        if f.HasExplicit then Option.map InverseConversion.alignmentToWord f.ExplicitAlignment
        else Some (InverseConversion.alignmentToWord f.Alignment)
    let spacingAttributes =
        [ match writtenMeasure para f.ExplicitSpacingBefore f.SpacingBefore with
          | Some v -> yield "w:before", InverseConversion.pointsToTwips v; yield "w:beforeAutospacing", "0"
          | None -> ()
          match writtenMeasure para f.ExplicitSpacingAfter f.SpacingAfter with
          | Some v -> yield "w:after", InverseConversion.pointsToTwips v; yield "w:afterAutospacing", "0"
          | None -> ()
          match writtenMeasure para f.ExplicitLineSpacing f.LineSpacing with
          | Some value ->
              match f.LineSpacingRule with
              | Some rule when rule = "exact" || rule = "atLeast" ->
                  yield "w:line", InverseConversion.fixedLineSpacingToTwips (paragraphFontSize para) value
                  yield "w:lineRule", rule
              | _ -> yield "w:line", InverseConversion.lineSpacingToWord value; yield "w:lineRule", "auto"
          | None -> () ]
    let indentAttributes =
        [ match writtenMeasure para f.ExplicitIndentLeft f.IndentLeft with Some v -> yield "w:left", InverseConversion.pixelsToTwips v | None -> ()
          match writtenMeasure para f.ExplicitIndentRight f.IndentRight with Some v -> yield "w:right", InverseConversion.pixelsToTwips v | None -> ()
          match writtenMeasure para f.ExplicitFirstLineIndent f.FirstLineIndent with
          | Some v when v < 0.0 -> yield "w:hanging", InverseConversion.pixelsToTwips -v
          | Some v -> yield "w:firstLine", InverseConversion.pixelsToTwips v
          | None -> () ]
    let onOff name = Option.map (toggle name) >> Option.toList
    let children =
        [ match f.Style, para.HeadingLevel with
          | Some id, _ when id <> "" && id <> documentDefaultsStyleId -> yield selfClosing "w:pStyle" [ "w:val", id ]
          | Some id, _ when id = documentDefaultsStyleId -> ()
          | _, Some level when level >= 1 && level <= 3 -> yield selfClosing "w:pStyle" [ "w:val", sprintf "Heading%d" level ]
          | _ -> ()
          let inferred pick = f.StyleInfo |> Option.bind pick
          let keep (direct: bool option) (pick: InferredStyle -> bool option) =
              if f.HasExplicit || direct.IsSome then direct else inferred pick
          yield! onOff "w:keepNext" (keep f.KeepNext (fun s -> s.KeepNext))
          yield! onOff "w:keepLines" (keep f.KeepLines (fun s -> s.KeepLines))
          if f.PageBreakBefore && not f.PageBreakBeforeDerived && not f.PageBreakAtStart then yield selfClosing "w:pageBreakBefore" []
          match f.DropCapKind with
          | Some kind when kind = "drop" || kind = "margin" ->
              yield selfClosing "w:framePr" [
                  yield "w:dropCap", kind; yield "w:lines", string (max 1 (Option.defaultValue 1 f.DropCapLines))
                  match f.DropCapDistance with Some d when d > 0.0 -> yield "w:hSpace", pixelsToTwips d | _ -> ()
                  yield "w:wrap", "around"; yield "w:vAnchor", "text"; yield "w:hAnchor", "text" ]
          | _ -> ()
          yield! onOff "w:widowControl" f.WidowControl
          match para.List with
          | Some list when list.Style = "bullet" || list.Style = "numbered" ->
              let parsed = list.Id |> Option.bind (fun text -> match Int32.TryParse text with true, value -> Some (string value) | _ -> None)
              let numId = parsed |> Option.defaultValue (if list.Style = "bullet" then fst context.DefaultListNumIds else snd context.DefaultListNumIds)
              yield element "w:numPr" [] [ selfClosing "w:ilvl" [ "w:val", string (max 0 (list.Level - 1)) ]; selfClosing "w:numId" [ "w:val", numId ] ]
          | Some list when list.Id = Some "0" -> yield element "w:numPr" [] [ selfClosing "w:numId" [ "w:val", "0" ] ]
          | _ -> ()
          match paragraphBordersElement f.Borders with Some borders -> yield borders | None -> ()
          let pattern = if f.ExplicitParagraphFormat then f.ExplicitShadingPattern else f.ShadingPattern
          let foreground = if f.ExplicitParagraphFormat then f.ExplicitShadingForeground else f.ShadingForeground
          let background = if f.ExplicitParagraphFormat then f.ExplicitBackgroundColor else f.BackgroundColor
          if pattern.IsSome || foreground.IsSome || background.IsSome then
              let color value = InverseConversion.cssColorToWord value |> Option.defaultValue "auto"
              yield selfClosing "w:shd" [ "w:val", Option.defaultValue "clear" pattern; "w:color", color foreground; "w:fill", color background ]
          if not f.ExplicitParagraphFormat || f.ExplicitTabs then
              match tabStopsElement para with Some tabs -> yield tabs | None -> ()
          if f.Rtl then yield selfClosing "w:bidi" []
          if not spacingAttributes.IsEmpty then yield selfClosing "w:spacing" spacingAttributes
          if not indentAttributes.IsEmpty then yield selfClosing "w:ind" indentAttributes
          yield! onOff "w:contextualSpacing" (if f.ExplicitParagraphFormat then f.ExplicitContextualSpacing else f.ContextualSpacing)
          match alignment with Some value -> yield selfClosing "w:jc" [ "w:val", value ] | None -> ()
          match f.OutlineLevel with Some level when level >= 0 && level <= 9 -> yield selfClosing "w:outlineLvl" [ "w:val", string level ] | _ -> ()
          match paragraphMarkRunProps para with Some props -> yield props | None -> ()
          match paragraphSectionProperties packageSetup para with Some section -> yield section | None -> () ]
    element "w:pPr" [] children

let generateParaProps context para = generateParaPropsWithPageSetup context defaultPageSetupInfo para


type RangeMark = { Node: XmlDsl.XmlNode; Kind: string; Opens: bool; Id: string }
    with member mark.Key = mark.Kind + ":" + mark.Id

/// The marks of ranges the writer carries and balances: those it carries between blocks and paragraphs, and
/// the ranges of moved text too (which are the model's: balanced, not carried).
let carriedRangeKinds = [ "bookmark"; "commentRange"; "perm" ]
let allRangeKinds = carriedRangeKinds @ [ "moveFromRange"; "moveToRange" ]

/// The element as a mark of a range of one of these kinds (an empty w:...Start or w:...End with a w:id).
let rangeMarkOf (kinds: string list) (node: XmlDsl.XmlNode) : RangeMark option =
    if XmlPatch.prefix node <> "w" || not (XmlPatch.elements node).IsEmpty then None
    else
        let local = XmlPatch.localName node
        kinds
        |> List.tryPick (fun kind -> if local = kind + "Start" then Some (kind, true) elif local = kind + "End" then Some (kind, false) else None)
        |> Option.bind (fun (kind, opens) -> XmlPatch.attribute node "w:id" |> Option.map (fun id -> { Node = node; Kind = kind; Opens = opens; Id = id }))

/// The marks of ranges of these kinds in a tree, in document order.
let rangeMarksIn (kinds: string list) (node: XmlDsl.XmlNode) : RangeMark list =
    XmlPatch.descendantsWhere (rangeMarkOf kinds >> Option.isSome) node |> List.choose (rangeMarkOf kinds)

/// The end mark of a range.
let rangeEndXml (kind: string) (id: string) = selfClosing ("w:" + kind + "End") [ ("w:id", id) ]

/// The marks of a paragraph's XML that the model does not hold: those to write at the start of
/// the paragraph and those to write at its end. (What a shape, a picture or an object anchored in
/// the paragraph holds - the paragraphs of a text box - is not of the paragraph.)
let carriedMarksOf (context: WriterContext) (paragraphXml: string) : string * string =
    match XmlPatch.read paragraphXml with
    | None -> "", ""
    | Some paragraph ->
        // The marks and the runs (a run, an equation, a simple field) of the paragraph, in order.
        let tokens = ResizeArray<RangeMark option>()
        let rec visit (node: XmlDsl.XmlNode) =
            match XmlPatch.name node with
            | "mc:AlternateContent" | "w:drawing" | "w:pict" | "w:object" -> ()
            | name ->
                match rangeMarkOf carriedRangeKinds node with
                | Some mark -> tokens.Add (Some mark)
                | None ->
                    if name = "w:r" || name = "m:oMath" || name = "m:oMathPara" || name = "w:fldSimple" then tokens.Add None
                    List.iter visit (XmlPatch.elements node)
        List.iter visit (XmlPatch.elements paragraph.Node)
        let tokens = List.ofSeq tokens
        let firstRun = tokens |> List.tryFindIndex Option.isNone |> Option.defaultValue -1
        let lastRun = tokens |> List.tryFindIndexBack Option.isNone |> Option.defaultValue -1
        let atStart = System.Text.StringBuilder()
        let atEnd = System.Text.StringBuilder()
        let opened = System.Collections.Generic.HashSet<string>()
        tokens
        |> List.iteri (fun index token ->
            match token with
            | None -> ()
            | Some mark ->
                let bytes = XmlPatch.bytesOf paragraph.Text mark.Node
                let beforeRuns = firstRun < 0 || index < firstRun
                let afterRuns = lastRun >= 0 && index > lastRun
                if mark.Kind = "bookmark" then
                    // (Those of a run are the model's.)
                    if firstRun < 0 then atStart.Append(bytes) |> ignore
                    elif mark.Opens then opened.Add mark.Id |> ignore
                    elif beforeRuns && not (opened.Contains mark.Id) then atStart.Append(bytes) |> ignore
                elif mark.Kind <> "commentRange" || Set.contains mark.Id context.LiveCommentIds then
                    if beforeRuns then atStart.Append(bytes) |> ignore
                    elif afterRuns then atEnd.Append(bytes) |> ignore
                    elif mark.Opens then atStart.Append(bytes) |> ignore
                    else atEnd.Append(bytes) |> ignore)
        atStart.ToString(), atEnd.ToString()

/// A paragraph the model wrote with marks at its start (after its properties) and at its end.
let withCarriedMarks (atStart: string) (atEnd: string) (xml: string) : string =
    if atStart = "" && atEnd = "" then xml
    else
        match XmlPatch.read xml with
        | Some paragraph when XmlPatch.name paragraph.Node = "w:p" ->
            let edits =
                [ if atStart <> "" then yield XmlPatch.Prepend atStart
                  if atEnd <> "" then yield XmlPatch.Append atEnd ]
            XmlPatch.apply XmlPatch.Word.order paragraph edits
        | _ -> xml


let carriedMarksOfItem (context: WriterContext) (para: EditorParagraph) =
    let copied = para.Source.Copy || (match para.Source.Origin with Some (LikeAt _) -> true | _ -> false)
    let source =
        match Portable.Dict.tryFind para context.CarriedSources with
        | ValueSome xml -> Some xml
        | _ when para.SourceParagraph.IsSome || not copied ->
            context.ParagraphOrigin |> Option.bind (fun provider -> provider.Of para)
        | _ -> None
    source |> Option.map (carriedMarksOf context) |> Option.defaultValue ("", "")


let paragraphPropertiesEqual (a: EditorParagraph) (b: EditorParagraph) =
    a.HeadingLevel = b.HeadingLevel && a.Formatting = b.Formatting && a.List = b.List
    && a.SectionBreak = b.SectionBreak && a.SectionPageSetup = b.SectionPageSetup
    && a.SectionColumns = b.SectionColumns && a.SectionVerticalAlign = b.SectionVerticalAlign
    && a.SectionPageNumberStart = b.SectionPageNumberStart && a.SectionPageNumberFormat = b.SectionPageNumberFormat
    && a.HeaderReferences = b.HeaderReferences && a.FooterReferences = b.FooterReferences

let likePropertiesEqual = paragraphPropertiesEqual

let setSpliceCut (spec: ParagraphEdit) (para: EditorParagraph) =
    let para =
        match spec.Cut with
        | Some (start, finish) ->
            let length =
                spec.OldText |> Option.map String.length
                |> Option.defaultWith (fun () ->
                    para.SpliceOriginal |> Option.map (fun original -> original.Runs |> Array.sumBy (runText >> String.length)) |> Option.defaultValue 0)
            { para with SpliceLeading = (start = 0); SpliceTrailing = (finish >= length) }
        | None -> para
    if spec.Join.Length = 0 then para
    else { para with SpliceTrailing = false; SpliceJoined = spec.Join |> Array.map (fun join -> join.Original) }


let controlProperties (layer: ContentControlLayer) =
    let id, tag, alias, xml = layer.Id, layer.Tag, layer.Alias, layer.Xml
    let props =
        match xml with
        | Some sourceXml ->
            match XmlPatch.readFragmentRoot sourceXml with
            | None -> invalidOp "Cannot preserve content-control properties: invalid XML"
            | Some fragment ->
                match XmlPatch.elements fragment.Node |> List.tryFind (fun node -> XmlPatch.localName node = "sdtPr") with
                | None -> invalidOp "Cannot preserve content-control properties: missing sdtPr"
                | Some original ->
                    let child name =
                        XmlPatch.elements original |> List.tryFind (fun node -> XmlPatch.localName node = name)
                    let childValue name = child name |> Option.bind (DocxReader.Docx.attr "val")
                    let field name value =
                        let existing = child name
                        let qualified =
                            existing |> Option.map XmlPatch.name
                            |> Option.defaultWith (fun () -> XmlDsl.Ns.withPrefixOf (XmlPatch.name original) name)
                        let valueName =
                            existing
                            |> Option.bind (fun node ->
                                let own = XmlDsl.Ns.withPrefixOf (XmlPatch.name node) "val"
                                if (XmlPatch.attribute node own).IsSome then Some own
                                elif (XmlPatch.attribute node "val").IsSome then Some "val"
                                else XmlPatch.attributeNames node |> List.tryFind (fun name -> XmlDsl.Ns.stripNs name = "val"))
                            |> Option.defaultWith (fun () -> XmlDsl.Ns.withPrefixOf qualified "val")
                        match value with
                        | Some value when value <> "" ->
                            [ XmlPatch.EditChild (qualified, [ XmlPatch.SetAttribute (valueName, value) ],
                                                  Some (selfClosing qualified [ valueName, value ])) ]
                        | _ -> [ XmlPatch.RemoveChild qualified ]
                    if childValue "id" = id && childValue "tag" = tag && childValue "alias" = alias then sourceXml
                    else XmlPatch.applyFragment XmlPatch.Word.order fragment
                             (XmlPatch.editsAt fragment.Node [ original, field "id" id @ field "tag" tag @ field "alias" alias ])
        | None -> ""
    if props <> "" then props
    else element "w:sdtPr" [] [
        match id with Some value when value <> "" -> yield selfClosing "w:id" [ "w:val", value ] | _ -> ()
        match tag with Some value when value <> "" -> yield selfClosing "w:tag" [ "w:val", value ] | _ -> ()
        match alias with Some value when value <> "" -> yield selfClosing "w:alias" [ "w:val", value ] | _ -> () ]

let inlineControlRuns (runs: (EditorRun * string) list) =
    let output = System.Text.StringBuilder()
    let active = ResizeArray<ContentControlLayer>()
    for run, xml in runs do
        let mutable common = 0
        let mutable remaining = run.ContentControls
        while common < active.Count && not remaining.IsEmpty
              && active.[common].Group <> "" && active.[common] = remaining.Head do
            common <- common + 1
            remaining <- remaining.Tail
        while active.Count > common do
            output.Append("</w:sdtContent></w:sdt>") |> ignore
            Portable.Stack.dropLast active
        for layer in remaining do
            output.Append("<w:sdt>").Append(controlProperties layer).Append("<w:sdtContent>") |> ignore
            active.Add layer
        output.Append(xml) |> ignore
    while active.Count > 0 do
        output.Append("</w:sdtContent></w:sdt>") |> ignore
        Portable.Stack.dropLast active
    output.ToString()

let generateParagraphWithHyperlinksAndImagesWithPageSetup (context: WriterContext) (packageSetup: PageSetupInfo) (hyperlinks: Map<string, string>) (images: ImageScope) (para: EditorParagraph) =
    let inputPara = para
    match (if para.SourceParagraph.IsNone || not context.ParagraphsInContext then None else context.ParagraphSource |> Option.bind (fun provider -> provider.Of para)) with
    | Some xml ->
        context.ReportDerived <- context.ReportDerived + 1
        xml
    | None ->
        context.ReportGenerated <- context.ReportGenerated + 1
        let para =
            match para.SpliceFrom, para.SourceParagraph with
            | None, Some spec when context.ParagraphsInContext ->
                match context.ParagraphSource |> Option.bind (fun provider -> provider.SpanOf para), spec.After |> Option.bind (fun after -> after.SpliceOriginal) with
                | Some span, Some original -> setSpliceCut spec { para with SpliceFrom = Some span; SpliceOriginal = Some original }
                | _ -> para
            | _ -> para
        let pPr = generateParaPropsWithPageSetup context packageSetup para
        let all = para.Runs
        let sameMark (a: EditorRun) (b: EditorRun) =
            match a.CommentId, b.CommentId with
            | Some x, Some y -> x = y
            | None, None ->
                match a.NoteId, b.NoteId with
                | Some x, Some y -> x = y && a.NoteType = b.NoteType
                | _ -> false
            | _ -> false
        let runsArray = all |> Array.mapi (fun index run -> index, run) |> Array.filter (fun (index, run) -> index = 0 || not (sameMark all.[index - 1] run && all.[index - 1].ContentControls = run.ContentControls)) |> Array.map snd
        let fieldParts = Dictionary<int, EditorRun list>()
        let fieldPartRuns = HashSet<int>()
        let isPlainText (run: EditorRun) = match run.Content with Text _ -> true | _ -> false
        let wrapper (run: EditorRun) = runHyperlinkTarget run, (run.Revision |> Option.bind (fun r -> r.Kind)), (run.Revision |> Option.bind (fun r -> r.Id)), run.CommentId, run.NoteId
        let mutable index = 0
        while index < runsArray.Length do
            let head = runsArray.[index]
            let mutable next = index + 1
            match (if isPlainText head then head.FieldCode else None), head.FieldResult with
            | Some code, Some whole when code <> "" ->
                let mutable text = runText head
                while next < runsArray.Length && isPlainText runsArray.[next]
                      && runsArray.[next].FieldCode = Some code && runsArray.[next].FieldResult = Some whole
                      && wrapper runsArray.[next] = wrapper head && runsArray.[next].ContentControls = head.ContentControls
                      && not (hasBookmarks runsArray.[next]) && text <> whole do
                    text <- text + runText runsArray.[next]
                    next <- next + 1
                if next > index + 1 then
                    fieldParts.[index] <- [ for part in index + 1 .. next - 1 -> runsArray.[part] ]
                    for part in index + 1 .. next - 1 do fieldPartRuns.Add part |> ignore
                else next <- index + 1
            | _ -> ()
            index <- next
        let runs =
            runsArray |> Array.toList |> List.mapi (fun i run -> i, run)
            |> List.filter (fun (i, _) -> not (fieldPartRuns.Contains i))
            |> List.map (fun (i, run) ->
                context.FieldResultParts <- match fieldParts.TryGetValue i with true, parts -> parts | _ -> []
                if not context.FieldResultParts.IsEmpty then context.SpliceFieldParts.[run] <- context.FieldResultParts
                let opaque (value: DocxReader.OpaqueObject) fallback =
                    if OpcXml.relationshipIdsIn value.Xml |> List.forall (fun id -> Set.contains id context.CurrentPartRelationshipIds) then
                        if value.ParagraphLevel then value.Xml else "<w:r>" + value.Xml + "</w:r>"
                    else runElement [] (isDeletedRun run) fallback
                let xml =
                    match run.Content with
                    | Opaque value -> opaque value ""
                    | Picture image ->
                        match image.Opaque with
                        | Some value -> opaque value (Option.defaultValue "" image.Alt)
                        | None ->
                            match image.Source |> Option.bind (fun src -> Map.tryFind src images.AssetBySrc) with
                            | Some asset -> generateImageRun asset image
                            | None -> runElement [] (isDeletedRun run) (Option.defaultValue "Image" image.Alt)
                    | _ -> generateReferenceRun context run
                run, xml, wrapRun context hyperlinks run)
        let textRunXml run =
            context.FieldResultParts <- []
            generateReferenceRun context run
        let wrappingOf run inner =
            revisionWrappedRun context run inner
            + (runHyperlinkTarget run |> Option.map (fun target -> "\u0003" + target) |> Option.defaultValue "")
            + (run.FieldCode |> Option.map (fun code -> "\u0004" + code) |> Option.defaultValue "")
        let splice =
            if runs |> List.exists (fun (run, _, _) -> match run.Content with FieldBoundary _ -> true | _ -> false)
            then None
            else context.ParagraphSplicer |> Option.bind (fun splicer -> splicer.Splice para runs textRunXml wrappingOf)
        context.SpliceFieldParts.Clear()
        if splice.IsSome then
            context.ReportSpliced <- context.ReportSpliced + 1
            context.ReportPatched <- context.ReportPatched + 1
            context.ReportGenerated <- context.ReportGenerated - 1
        let runs =
            match splice with
            | Some (spliced, _, _, _) -> spliced
            | None when runs |> List.exists (fun (run, _, _) -> (runHyperlinkTarget run).IsSome) ->
                let grouped = groupedHyperlinkRuns context hyperlinks runs
                if grouped |> List.exists (fun (run, _) -> not run.ContentControls.IsEmpty) then
                    [inlineControlRuns grouped]
                else grouped |> List.map snd
            | None when runs |> List.exists (fun (run, _, _) -> not run.ContentControls.IsEmpty) ->
                runs |> List.map (fun (run, inner, wrap) -> run, wrap inner) |> inlineControlRuns |> List.singleton
            | None -> runs |> List.map (fun (_, inner, wrap) -> wrap inner)
        let pPr =
            match splice, para.SpliceOriginal with
            | Some (_, Some original, _, _), Some before when paragraphPropertiesEqual para before -> original
            | Some (_, None, _, _), Some before when paragraphPropertiesEqual para before -> ""
            | Some (_, Some original, _, _), Some before ->
                context.ParagraphSplicer |> Option.bind (fun splicer -> splicer.MergeProperties original (generateParaPropsWithPageSetup context packageSetup before) pPr) |> Option.defaultValue pPr
            | _ -> pPr
        let f = para.Formatting
        let breakKept = splice |> Option.exists (fun (_, _, _, kept) -> kept)
        let runs =
            if f.PageBreakAfter && not breakKept then
                runs @ [element "w:r" [] [selfClosing "w:br" [ "w:type", if f.ColumnBreakAfter then "column" else "page" ]]]
            else runs
        let runs =
            if f.RenderedPageBreakBefore.IsSome && not (para.Runs |> Array.exists (fun run -> run.RenderedPageBreakBefore)) && splice.IsNone then
                element "w:r" [] [selfClosing "w:lastRenderedPageBreak" []] :: runs
            else runs
        let runs =
            if f.PageBreakAtStart && not breakKept then element "w:r" [] [selfClosing "w:br" [ "w:type", "page" ]] :: runs
            else runs
        let written =
            match splice with
            | Some (_, _, head, _) -> head + (if XmlPatch.isEmpty pPr then "" else pPr) + String.concat "" runs + "</w:p>"
            | None ->
                match (if XmlPatch.isEmpty pPr then runs else pPr :: runs) with
                | [] -> "<w:p></w:p>"
                | children -> element "w:p" [] children
        match splice, carriedMarksOfItem context inputPara with
        | Some _, _ | None, ("", "") -> written
        | None, (start, finish) -> withCarriedMarks start finish written

let generateParagraphWithHyperlinksAndImages context hyperlinks images para =
    generateParagraphWithHyperlinksAndImagesWithPageSetup context defaultPageSetupInfo hyperlinks images para
let generateParagraphWithHyperlinks context hyperlinks para =
    generateParagraphWithHyperlinksAndImages context hyperlinks emptyImageScope para
let generateParagraph context para = generateParagraphWithHyperlinks context Map.empty para


let defaultTableBorders () =
    element "w:tblBorders" [] [
        selfClosing "w:top" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
        selfClosing "w:left" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
        selfClosing "w:bottom" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
        selfClosing "w:right" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
        selfClosing "w:insideH" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
        selfClosing "w:insideV" [("w:val", "single"); ("w:sz", "4"); ("w:color", "auto")]
    ]

let borderElement name (value: Border option) =
    value |> Option.map (fun value ->
        let style = value.Style |> Option.defaultValue "single"
        let hidden = style = "none" || style = "nil"
        let color = value.Color |> InverseConversion.cssColorToWord |> Option.defaultValue "auto"
        let attrs =
            [ yield ("w:val", style)
              match value.Width with
              | Some amount -> yield ("w:sz", InverseConversion.pixelsToEighthPoints amount)
              | None -> if not hidden then yield ("w:sz", "4")
              if not hidden || value.Color.IsSome then yield ("w:color", color) ]
        selfClosing name attrs)

let bordersElement elementName (value: Borders option) =
    value |> Option.bind (fun value ->
        let sides =
            [ borderElement "w:top" value.Top; borderElement "w:left" value.Left
              borderElement "w:bottom" value.Bottom; borderElement "w:right" value.Right
              borderElement "w:insideH" value.InsideH; borderElement "w:insideV" value.InsideV ] |> List.choose id
        if sides.IsEmpty then None else Some (element elementName [] sides))

let tableWidthValue widthType width =
    match widthType with
    | "pct" -> roundToInt (width * 50.0) |> string
    // Editor-made widths ("pt" kind) are px; DOCX has no such unit, so they go out as dxa.
    | "pt" -> InverseConversion.pixelsToTwips width
    | "auto" | "nil" -> "0"
    | _ -> roundToInt width |> string

let optionalTableWidth elementName value typeValue =
    match value, typeValue with
    | Some width, Some widthType -> Some (selfClosing elementName [("w:w", tableWidthValue widthType width); ("w:type", (if widthType = "pt" then "dxa" else widthType))])
    | Some width, None -> Some (selfClosing elementName [("w:w", string (roundToInt width)); ("w:type", "dxa")])
    | None, Some widthType when widthType = "auto" || widthType = "nil" -> Some (selfClosing elementName [("w:w", "0"); ("w:type", widthType)])
    | _ -> None

let wordColorOrAuto value =
    value
    |> Option.map (fun color ->
        if color = "auto" then "auto"
        else InverseConversion.cssColorToWord (Some color) |> Option.defaultValue color)
    |> Option.defaultValue "auto"

let shadingElement (value: Shading option) =
    value |> Option.map (fun value ->
        selfClosing "w:shd" [("w:val", Option.defaultValue "clear" value.Pattern); ("w:color", wordColorOrAuto value.ForegroundColor); ("w:fill", wordColorOrAuto value.BackgroundColor)])

let backgroundShadingElement value =
    value |> Option.map (fun color ->
        let wordColor = InverseConversion.cssColorToWord (Some color) |> Option.defaultValue color
        selfClosing "w:shd" [("w:val", "clear"); ("w:color", "auto"); ("w:fill", wordColor)])

let cellMarginsElement elementName top left bottom right =
    let margin name value =
        match value with
        | Some amount when amount >= 0.0 -> Some (selfClosing name [("w:w", string (roundToInt amount)); ("w:type", "dxa")])
        | _ -> None
    let children = [ margin "w:top" top; margin "w:left" left; margin "w:bottom" bottom; margin "w:right" right ] |> List.choose id
    if children.IsEmpty then None else Some (element elementName [] children)

let cellSpacingElement value =
    match value with
    | Some amount when amount > 0.0 -> Some (selfClosing "w:tblCellSpacing" [("w:w", string (roundToInt amount)); ("w:type", "dxa")])
    | _ -> None

let generateTableGrid (table: EditorTable) =
    let columns =
        table.GridColumns
        |> Array.toList
        |> List.map (fun width -> selfClosing "w:gridCol" [("w:w", string (roundToInt width))])
    if columns.IsEmpty then "" else element "w:tblGrid" [] columns

/// Whether a w:tblGrid that was read has the columns the table has now: as many, of the same
/// widths. (A table the model holds no grid of keeps the one it was read with.)
let sourceGridIs (sourceGrid: string) (table: EditorTable) : bool =
    let now = table.GridColumns |> Array.map roundToInt
    if now.Length = 0 then true
    else
        let columns = XmlPatch.read sourceGrid |> Option.map (fun grid -> XmlPatch.elements grid.Node |> List.filter (fun node -> XmlPatch.localName node = "gridCol")) |> Option.defaultValue []
        let widths = columns |> List.map (fun column -> XmlPatch.attribute column "w:w" |> Option.bind (fun value -> match System.Double.TryParse value with | true, width -> Some (roundToInt width) | _ -> None))
        let ok = widths |> List.forall Option.isSome
        let read = widths |> List.choose id |> ResizeArray
        ok && read.Count = now.Length && (Array.forall2 (fun (a: int) (b: int) -> abs (a - b) <= 1) (read.ToArray()) now)

/// The properties the model writes for a cell, a row and a table (the children of w:tcPr, w:trPr, w:tblPr).
let tableCellProperties (cell: EditorCell) : string list =
    // CT_TcPr order: tcW, gridSpan, vMerge, tcBorders, shd, tcMar, vAlign.
    [
        match optionalTableWidth "w:tcW" cell.PreferredWidth cell.PreferredWidthType with
        | Some width -> yield width
        | None ->
            match cell.Width with
            | Some width -> yield selfClosing "w:tcW" [("w:w", string (roundToInt width)); ("w:type", "dxa")]
            | None -> ()

        match cell.ColSpan with
        | span when span > 1 -> yield selfClosing "w:gridSpan" [("w:val", string (span))]
        | _ -> ()

        match cell.VerticalMerge with
        | Some "restart" -> yield selfClosing "w:vMerge" [("w:val", "restart")]
        | Some "continue" -> yield selfClosing "w:vMerge" []
        | _ -> ()

        match bordersElement "w:tcBorders" cell.Borders with
        | Some borders -> yield borders
        | None ->
            // Legacy single-border fields from editor-made cells.
            if cell.Borders.IsNone && ((cell.BorderColor).IsSome || (cell.BorderWidth).IsSome) then
                let color = cell.BorderColor |> InverseConversion.cssColorToWord |> Option.defaultValue "auto"
                let width = cell.BorderWidth |> Option.map InverseConversion.pixelsToEighthPoints |> Option.defaultValue "8"
                let border name = selfClosing name [("w:val", "single"); ("w:sz", width); ("w:color", color)]
                yield element "w:tcBorders" [] [ border "w:top"; border "w:left"; border "w:bottom"; border "w:right" ]

        match shadingElement cell.Shading |> Option.orElseWith (fun () -> backgroundShadingElement cell.BackgroundColor) with
        | Some shading -> yield shading
        | None -> ()

        match cellMarginsElement "w:tcMar" cell.PaddingTop cell.PaddingLeft cell.PaddingBottom cell.PaddingRight with
        | Some margins -> yield margins
        | None -> ()

        match cell.VerticalAlignment with
        | Some value when value = "top" || value = "center" || value = "bottom" || value = "both" ->
            yield selfClosing "w:vAlign" [("w:val", value)]
        | _ -> ()
    ]


/// (And the row's w:tblPrEx, "" when it has none.)
let tableRowProperties (row: EditorRow) : string * string list =
    // Row-level cell margins live in w:tblPrEx, which precedes w:trPr.
    let rowExceptions =
        match cellMarginsElement "w:tblCellMar" row.PaddingTop row.PaddingLeft row.PaddingBottom row.PaddingRight with
        | Some margins -> element "w:tblPrEx" [] [margins]
        | None -> ""
    let rowPropChildren = [
        if row.GridBefore > 0 then yield selfClosing "w:gridBefore" [("w:val", string row.GridBefore)]
        if row.GridAfter > 0 then yield selfClosing "w:gridAfter" [("w:val", string row.GridAfter)]
        match row.WidthBefore with
        | Some width -> yield selfClosing "w:wBefore" [("w:w", string (roundToInt width)); ("w:type", "dxa")]
        | None -> ()
        match row.WidthAfter with
        | Some width -> yield selfClosing "w:wAfter" [("w:w", string (roundToInt width)); ("w:type", "dxa")]
        | None -> ()
        match row.CanSplit with
        | Some canSplit -> yield selfClosing "w:cantSplit" (if canSplit then [("w:val", "0")] else [])
        | None -> ()
        match row.Height with
        | Some height ->
            let attrs =
                [("w:val", string (roundToInt height))]
                @ (row.HeightRule |> Option.map (fun rule -> [("w:hRule", rule)]) |> Option.defaultValue [])
            yield selfClosing "w:trHeight" attrs
        | None -> ()
        if row.IsHeader then yield selfClosing "w:tblHeader" []
        match cellSpacingElement row.CellSpacing with
        | Some spacing -> yield spacing
        | None -> ()
    ]
    rowExceptions, rowPropChildren

let tablePropertyElements (table: EditorTable) : string list =
    // CT_TblPr order: tblStyle, tblW, jc, tblCellSpacing, tblBorders, shd, tblLayout, tblCellMar, tblLook.
    [
        // The table keeps its style: the reader resolves the style's formatting
        // into the cells from the reference (and the style itself is kept in
        // styles.xml), so dropping it changed how the saved table read back.
        match table.StyleId with
        | Some styleId -> yield selfClosing "w:tblStyle" [("w:val", styleId)]
        | None -> ()
        // A floating table stays one: without its w:tblpPr the table came back in
        // the text flow and the paragraphs that ran beside it below it.
        match table.FloatingXml with
        | Some floatingXml when (XmlPatch.read floatingXml |> Option.map (fun floating -> XmlPatch.name floating.Node)) = Some "w:tblpPr" -> yield floatingXml
        | _ -> if table.Floating then yield selfClosing "w:tblpPr" [("w:leftFromText", "180"); ("w:rightFromText", "180"); ("w:vertAnchor", "text"); ("w:tblpY", "1")]
        match optionalTableWidth "w:tblW" table.PreferredWidth table.PreferredWidthType with
        | Some width -> yield width
        | None -> ()
        match table.Alignment with
        | Some value when value <> "" -> yield selfClosing "w:jc" [("w:val", value)]
        | _ -> ()
        match cellSpacingElement table.CellSpacing with
        | Some spacing -> yield spacing
        | None -> ()
        match table.Indent with
        | Some indent -> yield selfClosing "w:tblInd" [("w:w", string (roundToInt indent)); ("w:type", "dxa")]
        | None -> ()
        match bordersElement "w:tblBorders" table.Borders with
        | Some borders -> yield borders
        | None ->
            // Only a table made in the editor (no grid, no width: it never came
            // from a file) gets the default grid. An imported table without
            // borders is borderless and must stay that way.
            if table.GridColumns.Length = 0 && table.PreferredWidthType.IsNone then
                yield defaultTableBorders ()
        match shadingElement table.Shading |> Option.orElseWith (fun () -> backgroundShadingElement table.BackgroundColor) with
        | Some shading -> yield shading
        | None -> ()
        match table.LayoutMode with
        | Some mode when mode = "fixed" || mode = "autofit" || mode = "autoFit" ->
            yield selfClosing "w:tblLayout" [("w:type", (if mode = "fixed" then "fixed" else "autofit"))]
        | _ -> ()
        match cellMarginsElement "w:tblCellMar" table.PaddingTop table.PaddingLeft table.PaddingBottom table.PaddingRight with
        | Some margins -> yield margins
        | None -> ()
        match table.Look with
        | Some look ->
            let bits =
                (if look.FirstRow then 0x0020 else 0) ||| (if look.LastRow then 0x0040 else 0)
                ||| (if look.FirstColumn then 0x0080 else 0) ||| (if look.LastColumn then 0x0100 else 0)
                ||| (if look.HorizontalBanding then 0 else 0x0200) ||| (if look.VerticalBanding then 0 else 0x0400)
            let hexDigits = "0123456789ABCDEF"
            let hex = System.String([| for shift in [ 12; 8; 4; 0 ] -> hexDigits.[(bits >>> shift) &&& 0xF] |])
            let onOffAttr isOn = if isOn then "1" else "0"
            yield selfClosing "w:tblLook" [
                ("w:val", hex)
                ("w:firstRow", onOffAttr look.FirstRow); ("w:lastRow", onOffAttr look.LastRow)
                ("w:firstColumn", onOffAttr look.FirstColumn); ("w:lastColumn", onOffAttr look.LastColumn)
                ("w:noHBand", onOffAttr (not look.HorizontalBanding)); ("w:noVBand", onOffAttr (not look.VerticalBanding)) ]
        | None -> ()
    ]


let likePropertiesText (para: EditorParagraph) =
    para.HeadingLevel, para.Formatting, para.List, para.SectionBreak, para.SectionPageSetup,
    para.SectionColumns, para.SectionVerticalAlign, para.SectionPageNumberStart,
    para.SectionPageNumberFormat, para.HeaderReferences, para.FooterReferences

let paragraphIsEmpty (para: EditorParagraph) =
    para.Runs |> Array.forall (fun run ->
        (match run.Content with Text text -> text = "" | _ -> false)
        && run.FieldCode.IsNone && run.FieldResult.IsNone && run.NoteId.IsNone && run.CommentId.IsNone
        && List.isEmpty run.BookmarkStarts && List.isEmpty run.BookmarkEnds)

/// An empty paragraph made like one that was read (`paragraphXml`: the first of a cell next to which a row or a
/// column was put in): its element without its ids (w14:paraId, w14:textId) and without what it holds but its
/// properties - those of the paragraph and of its mark, as they were read.
let emptyParagraphLike (paragraphXml: string) : string option =
    match XmlPatch.read paragraphXml with
    | Some paragraph when XmlPatch.name paragraph.Node = "w:p" ->
        let span = XmlPatch.span paragraph.Node
        let startTag = XmlPatch.startTag paragraph.Text paragraph.Node [ XmlPatch.RemoveAttribute "w14:paraId"; XmlPatch.RemoveAttribute "w14:textId" ]
        // (<w:p/>: the start tag is the element.)
        if span.StartTagEnd >= span.End then Some startTag
        else
            let properties = XmlPatch.elements paragraph.Node |> List.tryFind (fun kid -> XmlPatch.localName kid = "pPr") |> Option.map (XmlPatch.bytesOf paragraph.Text) |> Option.defaultValue ""
            Some (startTag + properties + "</" + XmlPatch.name paragraph.Node + ">")
    | _ -> None

/// The table being written of a table that was read; and whether the paragraphs of what is being written
/// may be made of the XML they were read from: those of a cell that is written with the properties its
/// table, row and cell were read with - what a paragraph's XML leaves to its table's style it takes from there.

let sourceOfCell (part: CellSource option) =
    part |> Option.bind (fun part -> if part.At < 0 then None else Some (part.At, part.Copy, part.Same))
let sourceOfRow (part: RowSource option) =
    part |> Option.bind (fun part -> if part.At < 0 then None else Some (part.At, part.Copy, part.Same))
let sourceCopy (source: SourceFragment) =
    source.Copy || (match source.Origin with Some (LikeAt _) -> true | _ -> false)


let rec generateTableCell (context: WriterContext) (hyperlinks: Map<string, string>) (images: ImageScope) (cell: EditorCell) : string =
    let source =
        match context.CurrentSourceTable, sourceOfCell cell.SourceCell with
        | Some table, Some (at, copy, same) ->
            (match Portable.Dict.tryFind at table.TableCells with
             | ValueSome known -> Some (known, copy, same)
             | _ -> None)
        | _ -> None
    match source with
    | Some (known, false, true) -> known.CellXml
    | _ ->
        // Its properties: the ones it was read with (or the one it is a copy of) - changed where the model's changed - or the model's.
        let tcPrChildren = tableCellProperties cell
        let properties =
            match source with
            | Some (known, copy, _) ->
                let before = cell.SourceCell |> Option.bind (fun source -> source.Before)
                match before with
                | None -> Some (if copy then known.CopyStartTag + known.CellProps else known.CellHead)
                | Some before ->
                    context.TableSource |> Option.bind (fun provider -> provider.Merge ("w:tcPr", known.CellProps, tableCellProperties before, tcPrChildren))
                    |> Option.map (fun props -> (if copy then known.CopyStartTag else known.CellStartTag) + props)
            | None -> None
        let outerContext = context.ParagraphsInContext
        let outerTable = context.CurrentSourceTable
        context.ParagraphsInContext <- outerContext && properties.IsSome
        // (Each paragraph with the w:p it is of, when that is one of the cell's and it is the first of it.)
        let paragraphItems =
            cell.Paragraphs |> Array.map (function
                | Paragraph para -> para
                | _ -> invalidArg "cell" "Cell.Paragraphs must contain only paragraphs")
        // (A cell made like one that was read - a row or a column put in - that holds one empty paragraph made like
        // that cell's first, as the model has it: that paragraph's element, empty - see emptyParagraphLike.)
        let likeParagraph (para: EditorParagraph) : string option =
            match source with
            | Some (_, true, _) when paragraphItems.Length = 1 && Array.isEmpty cell.NestedTables && paragraphIsEmpty para ->
                let first =
                    cell.SourceCell |> Option.bind (fun source -> source.Was)
                    |> Option.bind (fun was ->
                        was.Paragraphs |> Array.tryHead |> Option.map (function
                            | Paragraph para -> para
                            | _ -> invalidArg "cell" "Source cell paragraphs must contain only paragraphs"))
                match first, (match para.Source.Origin with Some (ReadAt at | LikeAt at) -> Some at | None -> None) with
                | Some first, Some at when sourceCopy para.Source && (match first.Source.Origin with Some (ReadAt at | LikeAt at) -> Some at | None -> None) = Some at && likePropertiesText para = likePropertiesText first ->
                    context.ParagraphOrigin |> Option.bind (fun provider -> provider.Of para) |> Option.bind emptyParagraphLike
                | _ -> None
            | _ -> None
        let paragraphs =
            paragraphItems
            |> Array.map (fun para ->
                match (if context.ParagraphsInContext then likeParagraph para else None) with
                | Some xml ->
                    context.ReportDerived <- context.ReportDerived + 1
                    xml
                | None -> generateParagraphWithHyperlinksAndImagesWithPageSetup context defaultPageSetupInfo hyperlinks images para)
        let paragraphAt =
            let used = System.Collections.Generic.HashSet<int>()
            paragraphItems
            |> Array.map (fun item ->
                match source, (match item.Source.Origin with Some (ReadAt at | LikeAt at) -> Some at | None -> None) with
                | Some (known, false, _), Some at when not (sourceCopy item.Source) && XmlPatch.isItemOf known.CellChildren at && used.Add at -> XmlPatch.AtItem at
                | _ -> XmlPatch.WithPrevious)
        let nestedTables = cell.NestedTables
        // NestedTableAnchors.[i] = how many of the cell's paragraphs precede nested
        // table i. The writer used to read only Paragraphs, so a nested table and
        // everything in it was deleted on save.
        let anchors = cell.NestedTableAnchors
        let content = ResizeArray<XmlPatch.Placed * string>()
        let mutable lastIsTable = false
        let emitTablesAt index =
            nestedTables
            |> Array.iteri (fun tableIndex nested ->
                let anchor = if tableIndex < anchors.Length then anchors.[tableIndex] else paragraphs.Length
                if (min anchor paragraphs.Length) = index then
                    // (A table in the cell is written from the model: so are its paragraphs.)
                    let inContext = context.ParagraphsInContext
                    context.ParagraphsInContext <- false
                    context.CurrentSourceTable <- None
                    content.Add ((XmlPatch.WithPrevious, generateTable context hyperlinks images nested))
                    context.ParagraphsInContext <- inContext
                    context.CurrentSourceTable <- outerTable
                    lastIsTable <- true)
        for index in 0 .. paragraphs.Length - 1 do
            emitTablesAt index
            content.Add ((paragraphAt.[index], paragraphs.[index]))
            lastIsTable <- false
        emitTablesAt paragraphs.Length
        // A cell must contain, and end with, a paragraph.
        if content.Count = 0 || lastIsTable then content.Add ((XmlPatch.WithPrevious, element "w:p" [] []))
        context.ParagraphsInContext <- outerContext
        match properties, source with
        // (A cell that was read: its paragraphs and tables among what stood around them - content controls, marks.)
        | Some head, Some (known, false, _) -> head + XmlPatch.writeLayout known.CellChildren (content.ToArray()) + "</w:tc>"
        | Some head, _ -> head + (content |> Seq.map snd |> String.concat "") + "</w:tc>"
        | None, _ ->
            let propertiesXml = if tcPrChildren.IsEmpty then "" else element "w:tcPr" [] tcPrChildren
            element "w:tc" [] (propertiesXml :: (content |> Seq.map snd |> List.ofSeq))

and generateTableRow (context: WriterContext) (hyperlinks: Map<string, string>) (images: ImageScope) (row: EditorRow) : string =
    let source =
        match context.CurrentSourceTable, sourceOfRow row.SourceRow with
        | Some table, Some (at, copy, same) ->
            (match Portable.Dict.tryFind at table.TableRows with
             | ValueSome known -> Some (known, copy, same)
             | _ -> None)
        | _ -> None
    match source with
    | Some (known, false, true) -> known.RowXml
    | _ ->
        let rowExceptions, rowPropChildren = tableRowProperties row
        // Its start tag and properties: the ones it was read with (or the one it is a copy of) - changed where the model's changed - or the model's.
        let head =
            match source with
            | Some (known, copy, _) ->
                let startTag = if copy then known.CopyStartTag else known.RowStartTag
                let before = row.SourceRow |> Option.bind (fun source -> source.Before)
                match before with
                | None -> Some (if copy then startTag + known.RowExceptions + known.RowProps else known.RowHead)
                | Some before ->
                    let (exceptionsBefore, propsBefore) = tableRowProperties before
                    context.TableSource |> Option.bind (fun provider -> provider.Merge ("w:trPr", known.RowProps, propsBefore, rowPropChildren))
                    |> Option.map (fun props -> startTag + (if exceptionsBefore = rowExceptions then known.RowExceptions else rowExceptions) + props)
            | None -> None
        let outerContext = context.ParagraphsInContext
        context.ParagraphsInContext <- outerContext && head.IsSome
        let cells =
            row.Cells
            |> Array.map (fun cell ->
                let at = sourceOfCell cell.SourceCell |> Option.bind (fun (at, copy, _) -> if copy then None else Some at)
                (match at with Some at -> XmlPatch.AtItem at | None -> XmlPatch.WithPrevious), generateTableCell context hyperlinks images cell)
        context.ParagraphsInContext <- outerContext
        // (The cells of a row that was read among what stood among them; a copy of a row has its cells alone.)
        let cellsXml =
            match source with
            | Some (known, false, _) -> XmlPatch.writeLayout known.RowChildren cells
            | _ -> cells |> Array.map snd |> String.concat ""
        match head with
        | Some text -> text + cellsXml + "</w:tr>"
        | None ->
            let rowProps = if rowPropChildren.IsEmpty then "" else element "w:trPr" [] rowPropChildren
            element "w:tr" [] [rowExceptions; rowProps; cellsXml]

and generateTable (context: WriterContext) (hyperlinks: Map<string, string>) (images: ImageScope) (table: EditorTable) : string =
    let tblPrChildren = tablePropertyElements table
    // A table that was read: its properties are the ones it was read with - changed where the
    // model's changed - and its rows and cells may be (see generateTableRow).
    let source = table.SourceTable |> Option.bind (fun table -> context.TableSource |> Option.bind (fun provider -> provider.Of table))
    let properties =
        match source with
        | Some known ->
            let before = table.SourceTable |> Option.bind (fun source -> source.Before)
            // (The grid that was read, unless the table has another by now - a column put in or taken out.)
            let grid = if sourceGridIs known.TableGrid table then known.TableGrid else generateTableGrid table
            match before with
            | None -> Some (if grid = known.TableGrid then known.TableHead else known.TableStartTag + known.TableProps + grid)
            | Some before ->
                context.TableSource |> Option.bind (fun provider -> provider.Merge ("w:tblPr", known.TableProps, tablePropertyElements before, tblPrChildren))
                |> Option.map (fun props -> known.TableStartTag + props + generateTableGrid table)
        | None -> None
    let outerContext = context.ParagraphsInContext
    let outerTable = context.CurrentSourceTable
    context.ParagraphsInContext <- properties.IsSome
    context.CurrentSourceTable <- (if properties.IsSome then source else None)
    let rows =
        let written =
            table.Rows
            |> Array.map (fun row ->
                let at = sourceOfRow row.SourceRow |> Option.map (fun (at, _, _) -> at)
                (match at with Some at -> XmlPatch.AtItem at | None -> XmlPatch.WithPrevious), generateTableRow context hyperlinks images row)
        match context.CurrentSourceTable with
        | Some known -> XmlPatch.writeLayout known.TableChildren written
        | _ -> written |> Array.map snd |> String.concat ""
    context.ParagraphsInContext <- outerContext
    context.CurrentSourceTable <- outerTable
    match properties, source with
    | Some text, Some known -> known.WrapStart + text + rows + "</w:tbl>" + known.WrapEnd
    | _ ->
        element "w:tbl" [] [
            element "w:tblPr" [] tblPrChildren
            generateTableGrid table
            rows
        ]

/// Shape ids must be unique in the document; they are handed out per save.
/// Every shape is written from the model, none as the element it was read
/// from (a package that asks for everything from the model: RegenerateAll).

let drawingNamespaces =
    {| Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
       A = "http://schemas.openxmlformats.org/drawingml/2006/main"
       Wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
       Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006" |}

let srgbColor (color: string) =
    selfClosing "a:srgbClr" [("val", InverseConversion.cssColorToWord (Some color) |> Option.defaultValue "000000")]

/// a:prstGeom or a:custGeom of a shape item (Geometry / CustomPath / CustomW / CustomH).
let shapeGeometryElement (item: EditorTextBox) =
    let geometry = item.Geometry
    let coordinate (value: float) = string (roundToInt value)
    let point (points: float array) index = selfClosing "a:pt" [("x", coordinate ((if index < points.Length then points.[index] else 0.0))); ("y", coordinate ((if index + 1 < points.Length then points.[index + 1] else 0.0)))]
    let segments =
        item.CustomPath
        |> List.choose (fun (op, points) ->
            match Some op with
            | Some "M" -> Some (element "a:moveTo" [] [point points 0])
            | Some "L" -> Some (element "a:lnTo" [] [point points 0])
            | Some "C" -> Some (element "a:cubicBezTo" [] [point points 0; point points 2; point points 4])
            | Some "Q" -> Some (element "a:quadBezTo" [] [point points 0; point points 2])
            | Some "Z" -> Some (selfClosing "a:close" [])
            | _ -> None)
    let coordW = item.CustomWidth
    let coordH = item.CustomHeight
    if geometry = "custom" && not segments.IsEmpty && coordW > 0.0 && coordH > 0.0 then
        element "a:custGeom" [] [
            selfClosing "a:avLst" []
            selfClosing "a:gdLst" []
            selfClosing "a:ahLst" []
            selfClosing "a:cxnLst" []
            selfClosing "a:rect" [("l", "0"); ("t", "0"); ("r", "r"); ("b", "b")]
            element "a:pathLst" [] [element "a:path" [("w", coordinate coordW); ("h", coordinate coordH)] segments]
        ]
    else
        element "a:prstGeom" [("prst", (if geometry = "custom" || geometry = "" then "rect" else geometry))] [selfClosing "a:avLst" []]

/// The fill of a shape: gradient, solid colour or none.
let shapeFillElement (item: EditorTextBox) =
    let stops = item.Gradient
    if not stops.IsEmpty then
        let stopElements =
            stops |> List.map (fun (offset, color) ->
                element "a:gs" [("pos", string (roundToInt (offset * 100000.0)))] [srgbColor color])
        let direction =
            if item.GradientRadial then
                element "a:path" [("path", "circle")] [selfClosing "a:fillToRect" [("l", "50000"); ("t", "50000"); ("r", "50000"); ("b", "50000")]]
            else
                selfClosing "a:lin" [("ang", string (roundToInt (item.GradientAngle * 60000.0))); ("scaled", "0")]
        element "a:gradFill" [("rotWithShape", "1")] [element "a:gsLst" [] stopElements; direction]
    elif item.NoFill then selfClosing "a:noFill" []
    else element "a:solidFill" [] [srgbColor (item.FillColor |> Option.defaultValue "#FFFFFF")]

/// a:ln: width (px -> EMU), colour or none, preset dash.
let shapeLineElement (item: EditorTextBox) =
    let width = item.BorderWidth |> Option.defaultValue 1.0 |> pixelsToEmus
    let children =
        [ if item.NoLine then yield selfClosing "a:noFill" []
          else yield element "a:solidFill" [] [srgbColor (item.BorderColor |> Option.defaultValue "#808080")]
          match item.BorderDash with
          | Some dash when dash <> "" && dash <> "solid" -> yield selfClosing "a:prstDash" [("val", dash)]
          | _ -> () ]
    element "a:ln" [("w", string (max 0 width))] children

/// The VML form of a shape, for readers without DrawingML shapes (the
/// mc:Fallback Word writes too): a text box rectangle at the same place.
let generateVmlShape (shapeId: string) (item: EditorTextBox) (content: string) : string =
    let width = item.Width |> pixelsToPoints
    let height = item.Height |> pixelsToPoints
    let x = item.X |> Option.map pixelsToPoints |> Option.defaultValue 0.0
    let y = item.Y |> Option.map pixelsToPoints |> Option.defaultValue 0.0
    let strokeWeight = item.BorderWidth |> Option.map pixelsToPoints |> Option.defaultValue 0.75
    let rotation = item.Rotation
    let style =
        if item.Inline then sprintf "width:%.4fpt;height:%.4fpt" width height
        else sprintf "position:absolute;margin-left:%.4fpt;margin-top:%.4fpt;width:%.4fpt;height:%.4fpt;z-index:1" x y width height
    let style = if abs rotation >= 0.01 then style + sprintf ";rotation:%.2f" rotation else style
    let style = match vmlFlip item.FlipH item.FlipV with Some flip -> style + ";flip:" + flip | None -> style
    element "w:pict" [] [
        element "v:shape" [
            yield ("xmlns:v", "urn:schemas-microsoft-com:vml")
            yield ("id", shapeId)
            yield ("type", "#_x0000_t202")
            yield ("style", style)
            if item.NoFill then yield ("filled", "f")
            else yield ("fillcolor", item.FillColor |> Option.defaultValue "#FFFFFF")
            if item.NoLine then yield ("stroked", "f")
            else
                yield ("strokecolor", item.BorderColor |> Option.defaultValue "#808080")
                yield ("strokeweight", sprintf "%.2fpt" strokeWeight)
        ] [
            element "v:textbox" [
                // (The same insets and fit as the DrawingML choice says.)
                let inset (value: float option) (fallback: float) = value |> Option.map pixelsToPoints |> Option.defaultValue fallback
                if not (item.InsetLeft.IsNone && item.InsetTop.IsNone && item.InsetRight.IsNone && item.InsetBottom.IsNone) then
                    yield ("inset", sprintf "%.2fpt,%.2fpt,%.2fpt,%.2fpt" (inset item.InsetLeft 7.2) (inset item.InsetTop 3.6) (inset item.InsetRight 7.2) (inset item.InsetBottom 3.6))
                if item.AutoFit then yield ("style", "mso-fit-shape-to-text:t")
            ] [
                element "w:txbxContent" [] [content]
            ]
        ]
    ]

/// A text box / shape as Word writes it: DrawingML (wp:anchor or wp:inline
/// around a wps:wsp) in an mc:AlternateContent with a VML fallback. Geometry,
/// rotation, solid / gradient / no fill, line colour, width and dash, wrap
/// mode and position are written from the model. (It used to be a bare VML
/// text box rectangle, so an ellipse or a rotated arrow came back as a white box.)
/// The edits of a shape's VML fallback (the mc:Fallback of its alternate content) that give it the place
/// and the size the shape has now, said in VML's way: margin-left / margin-top / width / height of the
/// style of its shape element, in points, from the frame the anchor names
/// (mso-position-horizontal-relative / mso-position-vertical-relative); an alignment in a direction that
/// was moved in (mso-position-horizontal: center) and a size or place in percent go. Everything else of
/// the fallback - its text, fill, line, path - stays the text it was. None: a fallback this does not know
/// how to place (a line, a shape in a line of text, no style): it is left out then, as it says
/// the old place.
///
/// A VML group (v:group) is placed and sized by its own style the same way - the first, outermost one:
/// its coordsize stays, so its shapes are where they were in it and scale with it.
let vmlFallbackEdits (fallback: XmlDsl.XmlNode) (item: EditorTextBox) (moveH: bool) (moveV: bool) (resize: bool) : (XmlDsl.XmlNode * XmlPatch.ElementEdit list) option =
    let isShape (node: XmlDsl.XmlNode) =
        XmlPatch.prefix node = "v" && List.contains (XmlPatch.localName node) [ "shape"; "rect"; "oval"; "roundrect" ] && (XmlPatch.attribute node "style").IsSome
    let isGroup (node: XmlDsl.XmlNode) = XmlPatch.name node = "v:group" && (XmlPatch.attribute node "style").IsSome
    let target =
        match XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "v:group") fallback with
        | group :: _ -> if isGroup group then Some group else None
        | [] -> XmlPatch.descendantsWhere isShape fallback |> List.tryHead
    match target with
    | Some shape ->
        let mutable declarations = XmlPatch.Style.parse (XmlPatch.attribute shape "style" |> Option.defaultValue "")
        let set property value = declarations <- XmlPatch.Style.set property value declarations
        let remove property = declarations <- XmlPatch.Style.remove property declarations
        let points (pixels: float) = string (System.Math.Round(pixels * 75.0) / 100.0) + "pt"
        if (moveH || moveV) && not (declarations |> List.contains ("position", "absolute")) then None
        else
            let mutable ok = true
            if moveH then
                match item.X with
                | Some x ->
                    set "margin-left" (points x)
                    remove "mso-position-horizontal"
                    remove "mso-left-percent"
                    match item.PositionHRelativeFrom with
                    | Some "page" -> set "mso-position-horizontal-relative" "page"
                    | Some "margin" -> set "mso-position-horizontal-relative" "margin"
                    | Some "column" -> set "mso-position-horizontal-relative" "text"
                    | Some "character" -> set "mso-position-horizontal-relative" "char"
                    | Some "leftMargin" -> set "mso-position-horizontal-relative" "left-margin-area"
                    | Some "rightMargin" -> set "mso-position-horizontal-relative" "right-margin-area"
                    | Some "insideMargin" -> set "mso-position-horizontal-relative" "inner-margin-area"
                    | Some "outsideMargin" -> set "mso-position-horizontal-relative" "outer-margin-area"
                    | Some _ -> ok <- false
                    | None -> ()
                | None -> ok <- false
            if moveV then
                match item.Y with
                | Some y ->
                    set "margin-top" (points y)
                    remove "mso-position-vertical"
                    remove "mso-top-percent"
                    match item.PositionVRelativeFrom with
                    | Some "page" -> set "mso-position-vertical-relative" "page"
                    | Some "margin" -> set "mso-position-vertical-relative" "margin"
                    | Some "paragraph" -> set "mso-position-vertical-relative" "text"
                    | Some "line" -> set "mso-position-vertical-relative" "line"
                    | Some "topMargin" -> set "mso-position-vertical-relative" "top-margin-area"
                    | Some "bottomMargin" -> set "mso-position-vertical-relative" "bottom-margin-area"
                    | Some "insideMargin" -> set "mso-position-vertical-relative" "inner-margin-area"
                    | Some "outsideMargin" -> set "mso-position-vertical-relative" "outer-margin-area"
                    | Some _ -> ok <- false
                    | None -> ()
                | None -> ok <- false
            if resize then
                set "width" (points item.Width)
                set "height" (points item.Height)
                remove "mso-width-percent"
                remove "mso-height-percent"
            if ok then Some (shape, [ XmlPatch.SetAttribute ("style", XmlPatch.Style.write declarations) ]) else None
    | _ -> None

/// The element a shape was read from with the shape's place and size as they are now (a shape
/// that was moved or resized and is otherwise what was read: Docx.textBoxMovedFrom): the
/// offsets of wp:positionH / wp:positionV, wp:extent and the extent of the shape's own
/// a:xfrm. Everything else of the element stays the text it was. Its VML fallback says the
/// place and the size in its own way: they are changed in it too (vmlFallbackEdits), and a
/// fallback that cannot be is left out (the DrawingML is what a reader that knows it takes;
/// Word writes the fallback anew). None: not in this element - a shape in a line
/// of text that is moved, an anchor without the offsets.
///
/// A group or a drawing canvas (an item with Children: wpg:wgp, wpc:wpc, v:group) is placed the same
/// way; resized, its own extent changes - wp:extent and the a:ext of the group's a:xfrm (the first
/// wpg:grpSpPr: its a:chExt stays, so its shapes scale with it; a canvas has wp:extent alone) -, and a
/// VML group (a w:pict of its own, or the fallback) its style's place and size. Its shapes' XML is never
/// touched. (A shape that is not a group: a wpg / wpc / w:pict in its DrawingML is not placed here.)
let patchShapeAnchor (xml: string) (item: EditorTextBox) : string option =
    let moveH = item.SourceMoved |> Option.exists (fun moved -> moved.H)
    let moveV = item.SourceMoved |> Option.exists (fun moved -> moved.V)
    let resize = item.SourceMoved |> Option.exists (fun moved -> moved.Size)
    let isGroup = not (Array.isEmpty item.Children)
    match XmlPatch.readFragmentRoot xml with
    | None -> None
    | Some fragment ->
        let named (qualified: string) (node: XmlDsl.XmlNode) = XmlPatch.descendantsWhere (fun inner -> XmlPatch.name inner = qualified) node
        // (The DrawingML: the choice of the alternate content, or the drawing itself.)
        let drawing = named "mc:Choice" fragment.Node |> List.tryHead |> Option.defaultValue fragment.Node
        let fallback = named "mc:Fallback" fragment.Node |> List.tryHead
        // (A VML group alone: a w:pict and no DrawingML. The fragment's root is a wrapper of the run's content.)
        let vmlOnly = (named "w:drawing" drawing).IsEmpty && not (named "w:pict" drawing).IsEmpty
        let unplaceable =
            if isGroup then []
            else XmlPatch.descendantsWhere (fun node -> XmlPatch.prefix node = "wpg" || XmlPatch.prefix node = "wpc" || XmlPatch.name node = "w:pict") drawing
        let anchor = named "wp:anchor" drawing |> List.tryHead
        if isGroup && vmlOnly then
            // (A VML group alone: its style says where it stands and how large it is.)
            vmlFallbackEdits drawing item moveH moveV resize
            |> Option.map (fun placed -> XmlPatch.applyFragment (fun _ -> []) fragment (XmlPatch.editsAt fragment.Node [ placed ]))
        elif not unplaceable.IsEmpty then None
        elif (moveH || moveV) && anchor.IsNone then None
        else
            let emu (pixels: float) = string (int64 (System.Math.Round(pixels * 9525.0)))
            let mutable ok = true
            let edits = ResizeArray<XmlDsl.XmlNode * XmlPatch.ElementEdit list>()
            // (The first element of that name in the drawing, with these edits; there being none, the shape is not placed here.)
            let edit (qualified: string) (within: XmlDsl.XmlNode) (made: XmlDsl.XmlNode -> XmlPatch.ElementEdit list) =
                match named qualified within |> List.tryHead with
                | Some node -> edits.Add ((node, made node))
                | None -> ok <- false
            // (A position: its frame, and an offset in the place of an alignment.)
            let position (offset: float) (relativeFrom: string option) (node: XmlDsl.XmlNode) =
                [ match relativeFrom with Some frame -> XmlPatch.SetAttribute ("relativeFrom", frame) | None -> ()
                  XmlPatch.RemoveChild "wp:align"
                  XmlPatch.RemoveChild "wp:posOffset"
                  XmlPatch.AddChild ("wp:posOffset", "<wp:posOffset>" + emu offset + "</wp:posOffset>") ]
            if moveH then
                match item.X with
                | Some x -> edit "wp:positionH" drawing (position x (item.PositionHRelativeFrom))
                | None -> ok <- false
            if moveV then
                match item.Y with
                | Some y -> edit "wp:positionV" drawing (position y (item.PositionVRelativeFrom))
                | None -> ok <- false
            if resize then
                let width, height = item.Width, item.Height
                let size = [ XmlPatch.SetAttribute ("cx", emu width); XmlPatch.SetAttribute ("cy", emu height) ]
                edit "wp:extent" drawing (fun _ -> size)
                if isGroup then
                    // (A group's own extent: the a:ext of the a:xfrm of its properties - the outermost group's.
                    // A canvas has none: its shapes are placed in it, not scaled.)
                    match named "wpg:grpSpPr" drawing with
                    | properties :: _ ->
                        match XmlPatch.child "xfrm" properties |> Option.bind (XmlPatch.child "ext") with
                        | Some extent when XmlPatch.name extent = "a:ext" -> edits.Add ((extent, size))
                        | _ -> ok <- false
                    | [] -> if (named "wpc:wpc" drawing).IsEmpty then ok <- false
                else
                    // (The shape's own extent: in its properties.)
                    match XmlPatch.descendantsWhere (fun node -> XmlPatch.localName node = "spPr" && (XmlPatch.prefix node = "wps" || XmlPatch.prefix node = "pic")) drawing with
                    | properties :: _ ->
                        match named "a:ext" properties with
                        | extent :: _ -> edits.Add ((extent, size))
                        | [] -> ()
                    | [] -> ()
            if not ok then None
            else
                // (The fallback at the shape's place too, or left out.)
                match fallback with
                | Some fallback ->
                    match vmlFallbackEdits fallback item moveH moveV resize with
                    | Some placed -> edits.Add placed
                    | None -> edits.Add ((fallback, [ XmlPatch.Removed ]))
                | None -> ()
                Some (XmlPatch.applyFragment (fun _ -> []) fragment (XmlPatch.editsAt fragment.Node (List.ofSeq edits)))

let rec generateTextBoxRun (context: WriterContext) (packageSetup: PageSetupInfo) (hyperlinks: Map<string, string>) (images: ImageScope) (item: EditorTextBox) : string =
    // A shape that is what was read goes back as the element it was read from
    // (its DrawingML and its VML fallback, with what the model does not hold
    // of them) - when everything that element refers to is in this part.
    let source =
        (if context.ShapesFromModel && Array.isEmpty item.Children then None else item.SourceXml)
        |> Option.filter (fun xml ->
            OpcXml.relationshipIdsIn xml
            |> List.forall (fun id -> Set.contains id context.CurrentPartRelationshipIds))
    // (Moved or resized, and otherwise what was read: its element with the anchor changed in it.)
    // (A group or a canvas is never written from the model, which has no XML of its shapes: not placed
    // anew in its element, it is the element as it was read.)
    let isGroup = not (Array.isEmpty item.Children)
    let source =
        match source with
        | Some xml when item.SourceMoved.IsSome ->
            match patchShapeAnchor xml item with
            | Some patched -> Some patched
            | None when isGroup -> Some xml
            | None -> None
        | other -> other
    match source with
    | Some xml -> "<w:r>" + xml + "</w:r>"
    | None when item.FitText -> generateWordArtRun context item
    | None -> generateTextBoxRunFromModel context packageSetup hyperlinks images item

/// Text drawn as large as its shape (WordArt, a watermark) from the model: a
/// VML shape with a v:textpath, as Word writes a watermark - where it stands
/// (aligned or by its margins, relative to the margins, the page or the
/// text), turned, behind the text, the colour and opacity of its letters,
/// their face. (Written as a text box it came back as small text in a frame.)
and generateWordArtRun (context: WriterContext) (item: EditorTextBox) : string =
    let shapeNumber = context.NextShapeId
    context.NextShapeId <- context.NextShapeId + 1
    let paragraphs = item.Paragraphs
    let runs = paragraphs |> Array.collect (function Paragraph paragraph -> paragraph.Runs | _ -> [||])
    let text = runs |> Array.map (fun run -> (match run.Content with Text text -> text | _ -> "")) |> String.concat ""
    let first = runs |> Array.tryHead
    let face = first |> Option.bind (fun run -> Some run.Formatting.FontFamily) |> Option.defaultValue "Calibri"
    let color = first |> Option.bind (fun run -> run.Formatting.Color) |> Option.defaultValue "#C0C0C0"
    let bold = first |> Option.exists (fun run -> run.Formatting.Bold)
    let italic = first |> Option.exists (fun run -> run.Formatting.Italic)
    let points (value: float option) (fallback: float) = value |> Option.map pixelsToPoints |> Option.defaultValue fallback
    let rotation = item.Rotation
    let relative (value: string option) (ofText: string) =
        match value with
        | Some "margin" -> Some "margin"
        | Some "page" -> Some "page"
        | Some "column" | Some "paragraph" -> Some ofText
        | Some "character" -> Some "char"
        | Some "line" -> Some "line"
        | _ -> None
    let style =
        [ if not (item.Inline) then yield "position:absolute"
          yield sprintf "margin-left:%.2fpt" (points item.X 0.0)
          yield sprintf "margin-top:%.2fpt" (points item.Y 0.0)
          yield sprintf "width:%.2fpt" (points (Some item.Width) 400.0)
          yield sprintf "height:%.2fpt" (points (Some item.Height) 100.0)
          if abs rotation >= 0.01 then yield sprintf "rotation:%.2f" rotation
          match vmlFlip item.FlipH item.FlipV with
          | Some flip -> yield "flip:" + flip
          | None -> ()
          yield (if Option.defaultValue false item.BehindDoc then "z-index:-251658752" else "z-index:251658752")
          match item.PositionHAlign with
          | Some align -> yield "mso-position-horizontal:" + align
          | None -> ()
          match relative item.PositionHRelativeFrom "text" with
          | Some name -> yield "mso-position-horizontal-relative:" + name
          | None -> ()
          match item.PositionVAlign with
          | Some align -> yield "mso-position-vertical:" + align
          | None -> ()
          match relative item.PositionVRelativeFrom "text" with
          | Some name -> yield "mso-position-vertical-relative:" + name
          | None -> () ]
        |> String.concat ";"
    let pathStyle =
        sprintf "font-family:\"%s\";font-size:1pt" face
        + (if bold then ";font-weight:bold" else "")
        + (if italic then ";font-style:italic" else "")
    let shapeType =
        "<v:shapetype id=\"_x0000_t136\" coordsize=\"21600,21600\" o:spt=\"136\" adj=\"10800\" path=\"m@7,l@8,m@5,21600l@6,21600e\">"
        + "<v:formulas><v:f eqn=\"sum #0 0 10800\"/><v:f eqn=\"prod #0 2 1\"/><v:f eqn=\"sum 21600 0 @1\"/><v:f eqn=\"sum 0 0 @2\"/><v:f eqn=\"sum 21600 0 @3\"/><v:f eqn=\"if @0 @3 0\"/><v:f eqn=\"if @0 21600 @1\"/>"
        + "<v:f eqn=\"if @0 0 @2\"/><v:f eqn=\"if @0 @4 21600\"/><v:f eqn=\"mid @5 @6\"/><v:f eqn=\"mid @8 @5\"/><v:f eqn=\"mid @7 @8\"/><v:f eqn=\"mid @6 @7\"/><v:f eqn=\"sum @6 0 @5\"/></v:formulas>"
        + "<v:path textpathok=\"t\" o:connecttype=\"custom\" o:connectlocs=\"@9,0;@10,10800;@11,21600;@12,10800\" o:connectangles=\"270,180,90,0\"/><v:textpath on=\"t\" fitshape=\"t\"/>"
        + "<v:handles><v:h position=\"#0,bottomRight\" xrange=\"6629,14971\"/></v:handles><o:lock v:ext=\"edit\" text=\"t\" shapetype=\"t\"/></v:shapetype>"
    let opacity = Some item.Opacity |> Option.filter (fun value -> value >= 0.0 && value < 0.999)
    element "w:r" [] [
        element "w:pict" [("xmlns:v", "urn:schemas-microsoft-com:vml"); ("xmlns:o", "urn:schemas-microsoft-com:office:office")] [
            shapeType
            element "v:shape" [
                ("id", item.Name |> Option.defaultValue (sprintf "WordArt%d" shapeNumber)); ("o:spid", sprintf "_x0000_s%d" shapeNumber); ("type", "#_x0000_t136")
                ("style", style); ("o:allowincell", "f"); ("fillcolor", color); ("stroked", "f")
            ] [
                match opacity with
                | Some value -> yield selfClosing "v:fill" [("opacity", sprintf "%.3f" value)]
                | None -> ()
                yield selfClosing "v:textpath" [("style", pathStyle); ("string", text)]
            ]
        ]
    ]

and generateTextBoxRunFromModel (context: WriterContext) (packageSetup: PageSetupInfo) (hyperlinks: Map<string, string>) (images: ImageScope) (item: EditorTextBox) : string =
    let widthEmu = item.Width |> pixelsToEmus |> max 0 |> string
    let heightEmu = item.Height |> pixelsToEmus |> max 0 |> string
    let shapeNumber = context.NextShapeId
    context.NextShapeId <- context.NextShapeId + 1
    let shapeName = item.Name |> Option.defaultValue (sprintf "Shape %d" shapeNumber)
    let paragraphs = item.Paragraphs
    let content =
        if paragraphs.Length = 0 then ""
        else paragraphs |> Array.map (function Paragraph paragraph -> generateParagraphWithHyperlinksAndImagesWithPageSetup context packageSetup hyperlinks images paragraph | Table table -> generateTable context hyperlinks images table | TextBox box -> generateTextBox context packageSetup hyperlinks images box) |> String.concat ""
    let emptyParagraph = element "w:p" [] []
    let rotation = item.Rotation
    let shape =
        element "wps:wsp" [] [
            yield selfClosing "wps:cNvSpPr" (if paragraphs.Length > 0 then [("txBox", "1")] else [])
            yield element "wps:spPr" [] [
                element "a:xfrm" ((if abs rotation >= 0.0001 then [("rot", string (roundToInt (rotation * 60000.0)))] else []) @ flipAttributes item.FlipH item.FlipV) [
                    selfClosing "a:off" [("x", "0"); ("y", "0")]
                    selfClosing "a:ext" [("cx", widthEmu); ("cy", heightEmu)]
                ]
                shapeGeometryElement item
                shapeFillElement item
                shapeLineElement item
            ]
            if paragraphs.Length > 0 then
                yield element "wps:txbx" [] [element "w:txbxContent" [] [content]]
            // The room around the text, where it stands in the shape's height, whether the shape follows it.
            let insetEmu (value: float option) (fallback: int) = value |> Option.map (fun px -> pixelsToEmus px |> max 0) |> Option.defaultValue fallback |> string
            let bodyAttributes =
                [ ("rot", "0"); ("vert", "horz"); ("wrap", (if item.NoTextWrap then "none" else "square"))
                  ("lIns", insetEmu item.InsetLeft 91440); ("tIns", insetEmu item.InsetTop 45720); ("rIns", insetEmu item.InsetRight 91440); ("bIns", insetEmu item.InsetBottom 45720)
                  ("anchor", item.VerticalAnchor |> Option.defaultValue "t"); ("anchorCtr", "0") ]
            if item.AutoFit then yield element "wps:bodyPr" bodyAttributes [ selfClosing "a:spAutoFit" [] ]
            else yield selfClosing "wps:bodyPr" bodyAttributes
        ]
    let graphic =
        element "a:graphic" [("xmlns:a", drawingNamespaces.A)] [
            element "a:graphicData" [("uri", drawingNamespaces.Wps)] [shape]
        ]
    let docPr = selfClosing "wp:docPr" [("id", string shapeNumber); ("name", shapeName)]
    let distance (value: float option) = value |> Option.defaultValue 0.0 |> pixelsToEmus |> max 0 |> string
    let drawing =
        if item.Inline then
            element "wp:inline" [("xmlns:wp", drawingNamespaces.Wp); ("distT", distance item.DistanceTop); ("distB", distance item.DistanceBottom); ("distL", distance item.DistanceLeft); ("distR", distance item.DistanceRight)] [
                selfClosing "wp:extent" [("cx", widthEmu); ("cy", heightEmu)]
                docPr
                selfClosing "wp:cNvGraphicFramePr" []
                graphic
            ]
        else
            let offset (value: float option) = value |> Option.defaultValue 0.0 |> pixelsToEmus |> string
            let relativeFrom (value: string option) fallback =
                value |> Option.defaultValue fallback
            element "wp:anchor" [
                ("xmlns:wp", drawingNamespaces.Wp)
                ("distT", distance item.DistanceTop); ("distB", distance item.DistanceBottom); ("distL", distance item.DistanceLeft); ("distR", distance item.DistanceRight)
                ("simplePos", "0")
                ("relativeHeight", item.RelativeHeight |> Option.defaultValue "251658240")
                ("behindDoc", (if Option.defaultValue false item.BehindDoc then "1" else "0")); ("locked", "0")
                ("layoutInCell", (if Option.defaultValue true item.LayoutInCell then "1" else "0")); ("allowOverlap", (if Option.defaultValue true item.AllowOverlap then "1" else "0"))
            ] [
                selfClosing "wp:simplePos" [("x", "0"); ("y", "0")]
                anchorPositionElement "wp:positionH" (relativeFrom item.PositionHRelativeFrom "column") (item.PositionHAlign) (offset item.X)
                anchorPositionElement "wp:positionV" (relativeFrom item.PositionVRelativeFrom "paragraph") (item.PositionVAlign) (offset item.Y)
                selfClosing "wp:extent" [("cx", widthEmu); ("cy", heightEmu)]
                selfClosing "wp:effectExtent" [("l", "0"); ("t", "0"); ("r", "0"); ("b", "0")]
                // A text box made in the editor has no wrap mode; it floats over the text.
                anchorWrapElement (item.Wrap |> Option.orElse (Some "none")) None [||]
                docPr
                selfClosing "wp:cNvGraphicFramePr" []
                graphic
            ]
    element "w:r" [] [
        element "mc:AlternateContent" [("xmlns:mc", drawingNamespaces.Mc); ("xmlns:wps", drawingNamespaces.Wps)] [
            element "mc:Choice" [("Requires", "wps")] [element "w:drawing" [] [drawing]]
            element "mc:Fallback" [] [generateVmlShape (sprintf "_x0000_s%d" shapeNumber) item (if content = "" then emptyParagraph else content)]
        ]
    ]

and generateTextBox (context: WriterContext) (packageSetup: PageSetupInfo) (hyperlinks: Map<string, string>) (images: ImageScope) (item: EditorTextBox) : string =
    element "w:p" [] [generateTextBoxRun context packageSetup hyperlinks images item]

let contentControlLayers (item: EditorBlock) : ContentControlLayer list =
    let data =
        match item with
        | Paragraph para ->
            let f = para.Formatting
            Some (f.ContentControlAncestors, f.ContentControlGroup, f.ContentControlId, f.ContentControlTag, f.ContentControlAlias, f.ContentControlXml)
        | Table table ->
            Some (table.ContentControlAncestors, table.ContentControlGroup, table.ContentControlId, table.ContentControlTag, table.ContentControlAlias, table.ContentControlXml)
        | TextBox _ -> None
    match data with
    | None -> []
    | Some (ancestors, group, id, tag, alias, xml) ->
        if group.IsSome || id.IsSome || tag.IsSome || alias.IsSome || xml.IsSome then
            ancestors @ [ { Group = defaultArg group ""; Id = id; Tag = tag; Alias = alias; Xml = xml } ]
        else ancestors

let generateBlocks (generateItem: EditorBlock -> string) (generateTextBoxRun: EditorTextBox -> string) (items: EditorBlock array) =
    let output = System.Text.StringBuilder()
    let active = ResizeArray<ContentControlLayer>()
    // Keep one plain block pending: a following text box is inserted in its
    // paragraph (or the table's last paragraph), as before.
    let mutable pendingPlain: string option = None
    let flushPlain () =
        match pendingPlain with
        | Some xml ->
            output.Append xml |> ignore
            pendingPlain <- None
        | None -> ()
    let closeTo depth =
        while active.Count > depth do
            output.Append("</w:sdtContent></w:sdt>") |> ignore
            Portable.Stack.dropLast active
    let sameLayer (a: ContentControlLayer) (b: ContentControlLayer) =
        a.Id = b.Id && a.Tag = b.Tag && a.Alias = b.Alias && a.Xml = b.Xml &&
        (if a.Group <> "" || b.Group <> "" then a.Group <> "" && a.Group = b.Group
         else a.Id.IsSome && b.Id.IsSome)
    let anchorBox box =
        match pendingPlain |> Option.bind XmlPatch.read with
        | None -> false
        | Some block ->
            let paragraph =
                match XmlPatch.name block.Node with
                | "w:p" -> Some block.Node
                | "w:tbl" ->
                    XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "w:p") block.Node
                    |> List.sortBy (fun para -> (XmlPatch.span para).End)
                    |> List.tryLast
                | _ -> None
            match paragraph with
            | Some para ->
                pendingPlain <- Some (XmlPatch.apply XmlPatch.Word.order block
                    (XmlPatch.editsAt block.Node [ para, [ XmlPatch.Append (generateTextBoxRun box) ] ]))
                true
            | None -> false
    for item in items do
        let anchored =
            match item with
            | TextBox box -> anchorBox box
            | _ -> false
        if not anchored then
            let layers = contentControlLayers item
            let levelCount = layers.Length
            flushPlain ()
            if layers.IsEmpty then
                closeTo 0
                pendingPlain <- Some (generateItem item)
            else
                let mutable common = 0
                let mutable remaining = layers
                while common < active.Count && not remaining.IsEmpty && sameLayer active.[common] remaining.Head do
                    common <- common + 1
                    remaining <- remaining.Tail
                closeTo common
                // The model generates each item once. Read its wrappers once,
                // then stream only newly opened wrappers and its own block text.
                // Never reparse or copy an accumulated control's content.
                let xml = generateItem item
                match XmlPatch.read xml with
                | None -> invalidOp "Cannot assemble content control: invalid generated XML"
                | Some source ->
                    let mutable current = source.Node
                    let mutable depth = 0
                    for layer in layers do
                        if XmlPatch.name current <> "w:sdt" then
                            invalidOp "Cannot assemble content control: missing nested w:sdt"
                        let content =
                            XmlPatch.child "sdtContent" current
                            |> Option.defaultWith (fun () -> invalidOp "Cannot assemble content control: missing w:sdtContent")
                        let controlSpan = XmlPatch.span current
                        let contentSpan = XmlPatch.span content
                        if depth >= common then
                            output.Append(xml, controlSpan.Start, contentSpan.StartTagEnd - controlSpan.Start) |> ignore
                            active.Add layer
                        depth <- depth + 1
                        if depth = levelCount then
                            output.Append(xml, contentSpan.StartTagEnd, contentSpan.EndTagStart - contentSpan.StartTagEnd) |> ignore
                        else
                            current <-
                                XmlPatch.elements content
                                |> List.tryHead
                                |> Option.defaultWith (fun () -> invalidOp "Cannot assemble content control: missing inner block")
    flushPlain ()
    closeTo 0
    output.ToString()


let wrapContentControl (item: EditorBlock) content =
    (contentControlLayers item, content)
    ||> List.foldBack (fun layer inner ->
        "<w:sdt>" + controlProperties layer + "<w:sdtContent>" + inner + "</w:sdtContent></w:sdt>")

let generateDocumentItemWithImagesAndPageSetup context packageSetup hyperlinks images item =
    let content =
        match item with
        | Table table -> generateTable context hyperlinks images table
        | TextBox box -> generateTextBox context packageSetup hyperlinks images box
        | Paragraph para -> generateParagraphWithHyperlinksAndImagesWithPageSetup context packageSetup hyperlinks images para
    wrapContentControl item content
let generateDocumentItemWithImages context hyperlinks images item =
    generateDocumentItemWithImagesAndPageSetup context defaultPageSetupInfo hyperlinks images item
let generateDocumentItem context hyperlinks item = generateDocumentItemWithImages context hyperlinks emptyImageScope item

type HeaderFooterPartInfo = {
    Kind: string
    RelId: string
    Target: string
    PartName: string
    Blocks: EditorBlock array
    IsGlobal: bool
    Unchanged: bool
}
type HeaderFooterPartPackage = { Info: HeaderFooterPartInfo; Images: ImageScope; Hyperlinks: Map<string, string> }

let partBlocks (part: HeaderFooterPart) = part.Blocks
let headerFooterInfos context (originalPaths: Map<string, string>) scopedRelationshipIds prefix startIndex (parts: HeaderFooterPart array) =
    let originalName (part: HeaderFooterPart) =
        part.RelationshipId |> Option.bind (fun relId -> Map.tryFind relId originalPaths)
        |> Option.bind (fun path -> if PartPath.directory path = "word" && PartPath.extension path = "xml" then Some (PartPath.fileName path) else None)
    let taken = HashSet<string>()
    for KeyValue (_, path) in originalPaths do
        if PartPath.directory path = "word" then taken.Add ((PartPath.fileName path).ToLowerInvariant()) |> ignore
    let given = HashSet<string>()
    let names =
        parts |> Array.map (fun part ->
            match originalName part with Some name when given.Add (name.ToLowerInvariant()) -> Some name | _ -> None)
    let mutable nextNumber = startIndex
    let freshName () =
        let mutable name = sprintf "%s%d.xml" prefix nextNumber
        while taken.Contains name || given.Contains name do
            nextNumber <- nextNumber + 1
            name <- sprintf "%s%d.xml" prefix nextNumber
        nextNumber <- nextNumber + 1
        given.Add name |> ignore
        name
    parts |> Array.mapi (fun index part ->
        let explicitRelId = part.RelationshipId
        let relId = explicitRelId |> Option.defaultValue (freshRelationshipId context (sprintf "rId%s%d" prefix (startIndex + index)))
        let fileName = match names.[index] with Some name -> name | None -> freshName ()
        { Kind = part.Kind; RelId = relId; Target = fileName; PartName = sprintf "/word/%s" fileName
          Blocks = partBlocks part; IsGlobal = not (Set.contains relId scopedRelationshipIds)
          Unchanged = explicitRelId.IsSome && part.SourceUnchanged })

let generateHeaderFooterXml context rootName hyperlinks images blocks =
    let content = blocks |> generateBlocks (generateDocumentItemWithImages context hyperlinks images) (generateTextBoxRun context defaultPageSetupInfo hyperlinks images)
    let content = if content = "" then element "w:p" [] [] else content
    document (element rootName [ "xmlns:w", Namespaces.word; "xmlns:r", Namespaces.relationships; "xmlns:m", Namespaces.math ] [ content ])

let generateSectionProperties (pageSetup: PageSetupInfo option) (headers: HeaderFooterPartInfo array) (footers: HeaderFooterPartInfo array) =
    let setup = Option.defaultValue defaultPageSetupInfo pageSetup
    let refs name (explicit: (string * string) list) (parts: HeaderFooterPartInfo array) =
        let fromParts = parts |> Array.toList |> List.map (fun part -> part.Kind, part.RelId)
        explicit @ fromParts
        |> List.filter (fun (kind, _) -> includeHeaderFooterKind setup.DifferentFirstPage setup.EvenAndOddHeaders kind)
        |> List.distinctBy fst |> List.map (fun (kind, relId) -> selfClosing name [ "w:type", kind; "r:id", relId ])
    let children =
        sectionPropertiesChildren (refs "w:headerReference" setup.HeaderReferences headers @ refs "w:footerReference" setup.FooterReferences footers)
            (pageSetup |> Option.bind (fun value -> value.SectionStart)) pageSetup setup.PageNumberStart setup.PageNumberFormat
    if children.IsEmpty then "" else element "w:sectPr" [] children
