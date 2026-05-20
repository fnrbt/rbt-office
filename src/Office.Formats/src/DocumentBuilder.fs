module DocumentBuilder

open Fable.Core.JsInterop
open DocumentProcessors.EditorProcessor
open DocumentWriter

/// Options for paragraph formatting
type ParagraphOptions = {
    Text: string option
    Alignment: string option
    IndentLeft: float option
    IndentRight: float option
    FirstLineIndent: float option
    SpacingBefore: float option
    SpacingAfter: float option
    LineSpacing: float option
    Bold: bool option
    Italic: bool option
    Underline: bool option
    FontSize: float option
    FontFamily: string option
    Color: string option
    VertAlign: string option
}

/// Document builder state
type DocumentBuilderState = {
    mutable Paragraphs: EditorParagraph list
}

/// Document builder
type DocumentBuilder(state: DocumentBuilderState) =
    /// Add a paragraph with options
    member this.AddParagraph(options: ParagraphOptions) =
        let text = defaultArg options.Text ""
        let alignment = defaultArg options.Alignment "left"
        let bold = defaultArg options.Bold false
        let italic = defaultArg options.Italic false
        let underline = defaultArg options.Underline false
        let fontSize = defaultArg options.FontSize 11.0
        let fontFamily = defaultArg options.FontFamily "Calibri"

        let run = {
            Text = text
            Bold = bold
            Italic = italic
            Underline = underline
            Strike = false
            FontSize = fontSize
            CharacterSpacing = 0.0
            FontFamily = fontFamily
            Color = options.Color
            Highlight = None
            VertAlign = options.VertAlign
            CommentId = None
            CommentText = None
            CommentAuthor = None
            CommentDate = None
            NoteId = None
            NoteType = None
            NoteText = None
            Hyperlink = None
            BookmarkStartId = None
            BookmarkName = None
            BookmarkEndId = None
            FieldCode = None
            FieldResult = None
            RenderedPageBreakBefore = false
            IsImage = false
            ImageSrc = None
            ImageLinked = false
            ImageWidth = None
            ImageHeight = None
            ImageAlt = None
            ImageName = None
            ImageTitle = None
            ImageFloating = false
            ImageWrap = None
            ImagePositionX = None
            ImagePositionY = None
            ImagePositionHRelativeFrom = None
            ImagePositionVRelativeFrom = None
            ImageBehindDoc = None
            ImageAllowOverlap = None
            ImageLayoutInCell = None
            ImageDistanceTop = None
            ImageDistanceBottom = None
            ImageDistanceLeft = None
            ImageDistanceRight = None
            ImageRelativeHeight = None
            ImageLocked = None
            ImageSimplePosition = None
            ImageSimplePositionX = None
            ImageSimplePositionY = None
            ImageCropLeft = None
            ImageCropTop = None
            ImageCropRight = None
            ImageCropBottom = None
        }

        let para = {
            Runs = [run]
            Alignment = alignment
            IndentLeft = options.IndentLeft
            IndentRight = options.IndentRight
            FirstLineIndent = options.FirstLineIndent
            RightTabPosition = None
            RightTabLeader = None
            SpacingBefore = options.SpacingBefore
            SpacingAfter = options.SpacingAfter
            LineSpacing = options.LineSpacing
        }

        state.Paragraphs <- state.Paragraphs @ [para]
        this

    /// Add a paragraph with multiple runs
    member this.AddParagraphWithRuns(alignment: string, runs: EditorRun list) =
        let para = {
            Runs = runs
            Alignment = alignment
            IndentLeft = None
            IndentRight = None
            FirstLineIndent = None
            RightTabPosition = None
            RightTabLeader = None
            SpacingBefore = None
            SpacingAfter = None
            LineSpacing = None
        }

        state.Paragraphs <- state.Paragraphs @ [para]
        this

    /// Save the document to a blob
    member this.Save() =
        let paraObjects =
            state.Paragraphs
            |> List.map (fun para ->
                createObj [
                    "Runs" ==> (para.Runs |> List.toArray)
                    "Alignment" ==> para.Alignment
                    "IndentLeft" ==> para.IndentLeft
                    "IndentRight" ==> para.IndentRight
                    "FirstLineIndent" ==> para.FirstLineIndent
                    "SpacingBefore" ==> para.SpacingBefore
                    "SpacingAfter" ==> para.SpacingAfter
                    "LineSpacing" ==> para.LineSpacing
                ]
            )
            |> List.toArray
        writeFromEditor paraObjects

/// Create a new document builder
let create () =
    let state = { Paragraphs = [] }
    DocumentBuilder(state)

/// Add paragraph (for JavaScript interop)
let addParagraph (builder: DocumentBuilder) (options: ParagraphOptions) =
    builder.AddParagraph(options)

/// Add paragraph with runs (for JavaScript interop)
let addParagraphWithRuns (builder: DocumentBuilder) (alignment: string) (runs: EditorRun list) =
    builder.AddParagraphWithRuns(alignment, runs)

/// Save document (for JavaScript interop)
let save (builder: DocumentBuilder) =
    builder.Save()

/// Helper to create a run
let createRun (text: string) (bold: bool) (italic: bool) (fontSize: float) (fontFamily: string) (color: string option) : EditorRun =
    {
        Text = text
        Bold = bold
        Italic = italic
        Underline = false
        Strike = false
        FontSize = fontSize
        CharacterSpacing = 0.0
        FontFamily = fontFamily
        Color = color
        Highlight = None
        VertAlign = None
        CommentId = None
        CommentText = None
        CommentAuthor = None
        CommentDate = None
        NoteId = None
        NoteType = None
        NoteText = None
        Hyperlink = None
        BookmarkStartId = None
        BookmarkName = None
        BookmarkEndId = None
        FieldCode = None
        FieldResult = None
        RenderedPageBreakBefore = false
        IsImage = false
        ImageSrc = None
        ImageLinked = false
        ImageWidth = None
        ImageHeight = None
        ImageAlt = None
        ImageName = None
        ImageTitle = None
        ImageFloating = false
        ImageWrap = None
        ImagePositionX = None
        ImagePositionY = None
        ImagePositionHRelativeFrom = None
        ImagePositionVRelativeFrom = None
        ImageBehindDoc = None
        ImageAllowOverlap = None
        ImageLayoutInCell = None
        ImageDistanceTop = None
        ImageDistanceBottom = None
        ImageDistanceLeft = None
        ImageDistanceRight = None
        ImageRelativeHeight = None
        ImageLocked = None
        ImageSimplePosition = None
        ImageSimplePositionX = None
        ImageSimplePositionY = None
        ImageCropLeft = None
        ImageCropTop = None
        ImageCropRight = None
        ImageCropBottom = None
    }
