module Rbt.Word.Docx

open System.Collections.Generic
open Fable.Core
open Fable.Core.JsInterop
open Rbt.Word.Document
open Rbt.Word.TextBuffer

#if FABLE_COMPILER
let private preservedOfEditorObject = WordEditorPackage.preservedOfEditorObject
#else
let private preservedOfEditorObject (_: obj) : WordEditorPackage.PreservedPackage option =
    invalidOp "JavaScript package adapters are only available in the browser"
#endif

[<Emit("$0 === null || $0 === undefined")>]
let private isNullOrUndefined (_value: obj) : bool = jsNative

let private intOrDefault defaultValue (value: obj) =
    if isNullOrUndefined value then defaultValue
    else unbox<int> value

let private addStyleIfNew (styles: ResizeArray<NamedStyle>) (seen: Dictionary<string, Dictionary<string, unit>>) style =
    let mutable ids = Unchecked.defaultof<Dictionary<string, unit>>
    if not (seen.TryGetValue(style.StyleType, &ids)) then
        ids <- Dictionary<string, unit>()
        seen.[style.StyleType] <- ids
    if not (ids.ContainsKey style.Id) then
        ids.[style.Id] <- ()
        styles.Add style

let private exportEditorObject (doc: Document) : obj =
#if FABLE_COMPILER
    TypedDocxExport.exportPackage doc |> WordEditorPackage.packageToEditorObject
#else
    invalidOp "JavaScript package adapters are only available in the browser"
#endif

let exportStylesToEditor (styles: NamedStyle list) : obj array =
    let package = exportEditorObject { Document.empty with Styles = styles }
    unbox package?Styles

let private paragraphsFromBlocks blocks =
    blocks
    |> List.choose (function
        | Paragraph para
        | Heading (_, para) -> Some para
        | Table _
        | TextBox _ -> None)

/// Convert the JavaScript package shape at the browser boundary.
let private importEditorObjectWithCache cache (package: obj) : Document =
#if FABLE_COMPILER
    package |> WordEditorPackage.packageOfEditorObject |> TypedDocxImport.importPackageWithCache cache
#else
    invalidOp "JavaScript package adapters are only available in the browser"
#endif

let private importEditorObject package =
    importEditorObjectWithCache (TypedDocxImport.newLevelCache ()) package

let importFromEditor (editorParagraphs: obj array) : Document =
    importEditorObject (createObj [ "Paragraphs" ==> editorParagraphs ])

let importFromEditorWithStyles (editorParagraphs: obj array) (styleCatalog: obj array) : Document =
    importEditorObject (createObj [ "Paragraphs" ==> editorParagraphs; "Styles" ==> styleCatalog ])

let importFromEditorWithHeaderFooters (editorParagraphs: obj array) (headerParagraphs: obj array) (footerParagraphs: obj array) : Document =
    let doc = importFromEditor editorParagraphs
    let header = importFromEditor headerParagraphs
    let footer = importFromEditor footerParagraphs
    { doc with Headers = paragraphsFromBlocks header.Blocks; Footers = paragraphsFromBlocks footer.Blocks }

let importFromEditorWithHeaderFooterParts (editorParagraphs: obj array) (headerParts: obj array) (footerParts: obj array) : Document =
    importEditorObject (createObj [ "Paragraphs" ==> editorParagraphs; "HeaderParts" ==> headerParts; "FooterParts" ==> footerParts ])

let importFromEditorWithPageSetupAndHeaderFooterParts (editorParagraphs: obj array) (pageSetup: obj) (headerParts: obj array) (footerParts: obj array) : Document =
    importEditorObject (createObj [ "Paragraphs" ==> editorParagraphs; "PageSetup" ==> pageSetup; "HeaderParts" ==> headerParts; "FooterParts" ==> footerParts ])

let private optionalArray (value: obj) =
    if isNullOrUndefined value then [||] else unbox<obj array> value

/// The reader's comment reference run only carries the comment id; give it the
/// text, author and date of the comment so the comments pane can show them.
let private withCommentDetails (comments: CommentBody list) (doc: Document) : Document =
    if comments.IsEmpty then doc
    else
        let byId = comments |> List.map (fun comment -> comment.Id, comment) |> Map.ofList
        let fill (para: ParagraphBlock) =
            let needsFill =
                para.Buffer.Runs
                |> List.exists (function
                    | Text run -> run.Style.CommentId.IsSome && run.Style.CommentText.IsNone
                    | Image _ -> false)
            if not needsFill then para
            else
                let runs =
                    para.Buffer.Runs
                    |> List.map (function
                        | Text run when run.Style.CommentId.IsSome && run.Style.CommentText.IsNone ->
                            match run.Style.CommentId |> Option.bind (fun id -> Map.tryFind id byId) with
                            | Some comment ->
                                Text { run with
                                        Style = { run.Style with
                                                    CommentText = Some comment.Text
                                                    CommentAuthor = run.Style.CommentAuthor |> Option.orElse comment.Author
                                                    CommentDate = run.Style.CommentDate |> Option.orElse comment.Date } }
                            | None -> Text run
                        | other -> other)
                { para with Buffer = { para.Buffer with Runs = runs } }
        let mapPart (part: HeaderFooterPart) = { part with Blocks = mapParagraphsOfBlocks fill part.Blocks }
        { doc with
            Blocks = mapParagraphsOfBlocks fill doc.Blocks
            HeaderParts = doc.HeaderParts |> List.map mapPart
            FooterParts = doc.FooterParts |> List.map mapPart }


[<Emit("new WeakMap()")>]
let private newPackageMap () : obj = jsNative
[<Emit("((typeof $1 === 'object' && $1 !== null) ? $0.get($1) : undefined)")>]
let private packageMapGet (_map: obj) (_package: obj) : Document = jsNative
[<Emit("((typeof $1 === 'object' && $1 !== null) ? $0.set($1, $2) : undefined)")>]
let private packageMapSet (_map: obj) (_package: obj) (_doc: Document) : unit = jsNative
#if FABLE_COMPILER
let private convertedPackages = newPackageMap ()
#else
let private convertedPackages : obj = null
#endif

// ---- a package that is still being read ---------------------------------------
// Office.Formats' openProgressive hands a package over with the first blocks
// of the body in Paragraphs and reads the rest in slices (its Progressive
// member; the text boxes come last, when the body is read). Its document is
// made as the items arrive: the styles they bring, the comments' details,
// the anchors of the text boxes - what importEditorObject does at once -
// and, the package read to its end, it is the document importEditorObject
// makes of the complete package.


/// More of a package was read.
type ImportStep = {
    /// The blocks that were read, in their order (at the end: the text boxes anchored in the body).
    Appended: Block list
    /// What was read, given to a document made of what was read before - the
    /// imported one, or one edited since: the blocks after its blocks, the
    /// styles they brought after its styles. The same blocks (and styles)
    /// always get the same extended list, so documents that shared their
    /// content before share it after.
    Extend: Document -> Document
    /// The package is read to its end.
    Complete: bool
    /// The package cannot be read to its end (Office.Formats' Progressive.Failed):
    /// why. This is its last step - the document is what was read.
    Failed: string option
    /// The font faces of the blocks that were read, once something asked for them (null until
    /// then; see Viewer.facesOfStep): they are found once for all who follow the package.
    mutable Faces: obj
}

/// The document of a package that is being read (see above).
type ProgressiveImport = {
    Package: obj
    /// What is read so far, as a document.
    mutable Document: Document
    mutable Complete: bool
    /// Why the package could not be read to its end, when it could not: the document stays
    /// what was read (and cannot be saved, see PackageData.Imported).
    mutable Failed: string option
    /// Told after each step: those who look at what the step brought before anything is done
    /// with it (the font faces it needs, see Viewer.onFontsDiscovered), then those who show it.
    FirstListeners: ResizeArray<ImportStep -> unit>
    Listeners: ResizeArray<ImportStep -> unit>
    // (Kept between the steps.)
    StyleList: ResizeArray<NamedStyle>
    StylesSeen: Dictionary<string, Dictionary<string, unit>>
    LevelCache: TypedDocxImport.LevelCache
    /// The id of the block of each item of the body ("": an item that has none).
    ItemBlockIds: ResizeArray<string>
    CommentBodies: CommentBody list
    Imported: ImportedDocument
    mutable Consumed: int
    /// The package's count of the changes of its page setup, as last looked at.
    mutable PageSetupVersion: int
    /// How many editors and viewers show the document while it is read (see holdProgressiveImport).
    mutable Holders: int
}

[<Emit("($0 != null && typeof $0 === 'object' && $0.Progressive != null)")>]
let private hasProgressive (_package: obj) : bool = jsNative
[<Emit("$0.get($1)")>]
let private importMapGet (_map: obj) (_package: obj) : ProgressiveImport = jsNative
[<Emit("$0.set($1, $2)")>]
let private importMapSet (_map: obj) (_package: obj) (_state: ProgressiveImport) : unit = jsNative
#if FABLE_COMPILER
let private progressiveImports = newPackageMap ()
#else
let private progressiveImports : obj = null
#endif

/// The details of the comments for the comment references of blocks (see withCommentDetails).
let private withCommentDetailsOfBlocks (comments: CommentBody list) (blocks: Block list) : Block list =
    if comments.IsEmpty then blocks
    else ({ Document.empty with Blocks = blocks; HeaderParts = []; FooterParts = [] } |> withCommentDetails comments).Blocks

/// What a step gives a document (ImportStep.Extend).
let private extension (appended: Block list) (stylesBefore: NamedStyle list) (stylesAfter: NamedStyle list)
                      (pageSetupBefore: PageSetup) (pageSetupAfter: PageSetup) (selection: Selection option)
                      (completedPreservation: WordEditorPackage.PreservedPackage option option) : Document -> Document =
    let addedStyles =
        if obj.ReferenceEquals(stylesBefore, stylesAfter) then []
        else stylesAfter |> List.skip (List.length stylesBefore)
    let knownBlocks = ResizeArray<Block list * Block list>()
    let knownStyles = ResizeArray<NamedStyle list * NamedStyle list>()
    let knownPackages = ResizeArray<PackageData * PackageData>()
    // (Used here as well as in the function below: a list that is named once, inside the
    // function, is made anew by the JavaScript build at every call of it - and nothing
    // that was extended would be found again.)
    knownBlocks.Clear()
    knownStyles.Clear()
    knownPackages.Clear()
    let extended (known: ResizeArray<'T * 'T>) (value: 'T) (make: unit -> 'T) : 'T =
        let mutable found = Unchecked.defaultof<'T>
        let mutable has = false
        let mutable i = 0
        while not has && i < known.Count do
            let before, after = known.[i]
            // (What was extended, or the extended one itself: a document is given a step once.)
            if obj.ReferenceEquals(before, value) || obj.ReferenceEquals(after, value) then
                found <- after
                has <- true
            i <- i + 1
        if has then found
        else
            let made = make ()
            known.Add ((value, made))
            made
    fun (doc: Document) ->
        let blocks = if appended.IsEmpty then doc.Blocks else extended knownBlocks doc.Blocks (fun () -> doc.Blocks @ appended)
        let styles =
            if addedStyles.IsEmpty then doc.Styles
            elif obj.ReferenceEquals(doc.Styles, stylesBefore) then stylesAfter
            else extended knownStyles doc.Styles (fun () -> doc.Styles @ addedStyles)
        // (The style catalog as imported: with the styles the blocks brought.)
        let package =
            if addedStyles.IsEmpty && completedPreservation.IsNone then doc.Package
            else
                extended knownPackages doc.Package (fun () ->
                    { doc.Package with
                        ImportedStyles = if addedStyles.IsEmpty then doc.Package.ImportedStyles else stylesAfter
                        Preserved = completedPreservation |> Option.defaultValue doc.Package.Preserved })
        let pageSetup = if obj.ReferenceEquals(doc.PageSetup, pageSetupBefore) then pageSetupAfter else doc.PageSetup
        let selection = match doc.Selection, selection with | None, Some _ -> selection | current, _ -> current
        if obj.ReferenceEquals(blocks, doc.Blocks) && obj.ReferenceEquals(styles, doc.Styles) && obj.ReferenceEquals(package, doc.Package)
           && obj.ReferenceEquals(pageSetup, doc.PageSetup) && obj.ReferenceEquals(selection, doc.Selection) then doc
        else { doc with Blocks = blocks; Styles = styles; Package = package; PageSetup = pageSetup; Selection = selection }

/// Take the items of the package that the document does not have yet.
let private consumeProgressive (state: ProgressiveImport) : unit =
    let items = optionalArray (state.Package?Paragraphs)
    let complete = unbox<bool> (state.Package?Progressive?IsComplete)
    let failed =
        let flag : obj = state.Package?Progressive?Failed
        if state.Failed.IsNone && not complete && not (isNullOrUndefined flag) && unbox<bool> flag then
            let error : obj = state.Package?Progressive?Error
            Some (if isNullOrUndefined error then "The document could not be read to its end" else string error)
        else None
    if items.Length > state.Consumed || (complete && not state.Complete) || failed.IsSome || (state.Package?Progressive?PageSetupVersion |> intOrDefault 0) <> state.PageSetupVersion then
        let stylesBefore = state.Document.Styles
        let styleCount = state.StyleList.Count
        let appended = ResizeArray<Block>()
        let freshItems = items.[state.Consumed ..]
        let typedItems =
#if FABLE_COMPILER
            (WordEditorPackage.packageOfEditorObject (createObj [ "Paragraphs" ==> freshItems ])).Blocks
#else
            invalidOp "JavaScript package adapters are only available in the browser"
#endif
        for item in typedItems do
            for style in TypedDocxImport.inferredStylesOfBlock item do
                addStyleIfNew state.StyleList state.StylesSeen style
            match TypedDocxImport.importBlock state.LevelCache item with
            | TextBox textBox ->
                // (The reader gives the anchor as an index into the items.)
                let anchor =
                    (match item with WordEditorPackage.TextBox shape -> shape.AnchorBlock | _ -> None)
                    |> Option.bind (fun at -> if at >= 0 && at < state.ItemBlockIds.Count && state.ItemBlockIds.[at] <> "" then Some state.ItemBlockIds.[at] else None)
                state.ItemBlockIds.Add ""
                appended.Add (match anchor with Some anchorId -> TextBox { textBox with AnchorBlockId = Some anchorId } | None -> TextBox textBox)
            | Paragraph para
            | Heading (_, para) as block ->
                state.ItemBlockIds.Add para.Id
                appended.Add block
            | Table table as block ->
                state.ItemBlockIds.Add table.Id
                appended.Add block
        state.Consumed <- items.Length
        let blocks = withCommentDetailsOfBlocks state.CommentBodies (List.ofSeq appended)
        let stylesAfter = if state.StyleList.Count = styleCount then stylesBefore else List.ofSeq state.StyleList
        // The caret a document is imported with: at the start of its first paragraph.
        let selection =
            if state.Document.Selection.IsSome then None
            else
                blocks
                |> List.tryPick (function Paragraph para -> Some para | _ -> None)
                |> Option.map (fun para -> { Anchor = { BlockId = para.Id; Offset = 0 }; Focus = { BlockId = para.Id; Offset = 0 } })
        let pageSetupBefore = state.Document.PageSetup
        // (Until the body is read the package states the page setup of the section the blocks
        // read so far end in; it says when that is another one.)
        let pageSetupVersion = state.Package?Progressive?PageSetupVersion |> intOrDefault 0
        let pageSetupAfter =
            if pageSetupVersion <> state.PageSetupVersion && not (isNullOrUndefined state.Package?PageSetup) then
                state.PageSetupVersion <- pageSetupVersion
                let read = (importEditorObject (createObj [ "PageSetup" ==> state.Package?PageSetup ])).PageSetup
                if read = pageSetupBefore then pageSetupBefore else read
            else pageSetupBefore
        let completedPreservation =
            if complete && not state.Complete then Some (preservedOfEditorObject state.Package?Preserved)
            else None
        let extend = extension blocks stylesBefore stylesAfter pageSetupBefore pageSetupAfter selection completedPreservation
        let document = extend state.Document
        state.Document <- document
        state.Complete <- complete
        if failed.IsSome then state.Failed <- failed
        if complete || failed.IsSome then state.LevelCache.Clear()
        // What the document was when it was read (see PackageData.Imported).
        state.Imported.Document <- { document with Package = { document.Package with Imported = None } }
        let step = { Appended = blocks; Extend = extend; Complete = complete; Failed = failed; Faces = null }
        for listener in state.FirstListeners.ToArray() do
            if state.FirstListeners.Contains listener then listener step
        for listener in state.Listeners.ToArray() do
            if state.Listeners.Contains listener then listener step

let private beginProgressive (package: obj) : ProgressiveImport =
    let levelCache = TypedDocxImport.newLevelCache ()
    let front = importEditorObjectWithCache levelCache (createObj [
        "Paragraphs" ==> [||]
        "Styles" ==> package?Styles
        "Theme" ==> package?Theme
        "HeaderParts" ==> package?HeaderParts
        "FooterParts" ==> package?FooterParts
        "Footnotes" ==> package?Footnotes
        "Endnotes" ==> package?Endnotes
        "Comments" ==> package?Comments
        "PageSetup" ==> package?PageSetup
        "Preserved" ==> package?Preserved ])
    let comments = front.Package.Comments
    let styleList = ResizeArray<NamedStyle>()
    let stylesSeen = Dictionary<string, Dictionary<string, unit>>()
    for style in front.Styles do addStyleIfNew styleList stylesSeen style
    let imported = { ImportedDocument.Document = Document.empty }
    let doc =
        { front with
            Blocks = []
            Selection = None
            Package = { front.Package with Imported = Some imported } }
    let state =
        { Package = package
          Document = doc
          Complete = false
          Failed = None
          FirstListeners = ResizeArray()
          Listeners = ResizeArray()
          StyleList = styleList
          StylesSeen = stylesSeen
          LevelCache = levelCache
          ItemBlockIds = ResizeArray()
          CommentBodies = comments
          Imported = imported
          Consumed = 0
          PageSetupVersion = package?Progressive?PageSetupVersion |> intOrDefault 0
          Holders = 0 }
    consumeProgressive state
    package?Progressive?subscribe (System.Func<int, int, bool, unit>(fun _ _ _ -> consumeProgressive state)) |> ignore
    state

/// The document of a package that is still being read: None for a package
/// that is read (forEditorPackage's, or one opened front first that was read
/// to its end before anything asked for its document).
let progressiveImportOf (package: obj) : ProgressiveImport option =
    if not (hasProgressive package) then None
    else
        let known = importMapGet progressiveImports package
        if not (isNull (box known)) then Some known
        elif unbox<bool> (package?Progressive?IsComplete) then None
        else
            let state = beginProgressive package
            importMapSet progressiveImports package state
            Some state

/// Read the rest of the package now. (One that cannot be read to its end: as far as it can
/// be - `state.Failed` says so afterwards.)
let completeProgressiveImport (state: ProgressiveImport) : unit =
    if not state.Complete && state.Failed.IsNone then
        state.Package?Progressive?complete () |> ignore
        consumeProgressive state

/// Stop reading the package: its document stays what was read, as that of a
/// package that failed (Office.Formats' Progressive.cancel; the listeners are
/// told with ImportStep.Failed). False when there was nothing to cancel.
let cancelProgressiveImport (state: ProgressiveImport) : bool =
    if state.Complete || state.Failed.IsSome then false
    else
        let cancel : obj = state.Package?Progressive?cancel
        if isNullOrUndefined cancel then false
        else
            let cancelled = unbox<bool> (state.Package?Progressive?cancel ())
            consumeProgressive state
            cancelled

/// What an editor or a viewer that was disposed of while its document was
/// still being read says of it (getLoadError): the package's own words for a
/// reading that was cancelled.
let readingCancelledMessage = "The reading of the document was cancelled"

/// An editor or a viewer shows the document while it is read ...
let holdProgressiveImport (state: ProgressiveImport) : unit =
    state.Holders <- state.Holders + 1

/// ... and is done with it (another document was loaded, it was disposed of):
/// when nothing shows a package that is still being read, the reading is
/// cancelled - no slices for a document nobody looks at, and the package lets
/// go of the file it reads.
let releaseProgressiveImport (state: ProgressiveImport) : unit =
    if state.Holders > 0 then
        state.Holders <- state.Holders - 1
        if state.Holders = 0 then cancelProgressiveImport state |> ignore

/// The document of an editor package. A package object is converted once:
/// a host that asks for a package's fonts (getRequiredFonts) and then loads
/// it into an editor or a viewer pays for one conversion, not two. (A
/// Document is immutable; the package object is not expected to change.)
/// For a package that is still being read (Office.Formats' openProgressive)
/// it is the document of what is read so far - see progressiveImportOf.
let importFromEditorPackage (package: obj) : Document =
    match progressiveImportOf package with
    | Some state -> state.Document
    | None ->
        let known = packageMapGet convertedPackages package
        if not (isNull (box known)) then known
        else
            let doc = importEditorObject package
            packageMapSet convertedPackages package doc
            doc

let exportToEditor (doc: Document) : obj array =
    let package = exportEditorObject doc
    unbox package?Paragraphs

/// Header / footer parts without their document: nothing is known to be
/// inherited from a style, so all formatting is written as direct formatting.
let exportHeaderFooterPartsToEditor (parts: HeaderFooterPart list) : obj array =
    let package = exportEditorObject { Document.empty with HeaderParts = parts }
    unbox package?HeaderParts

/// Page setup for the JavaScript editor package, in layout pixels.
let exportPageSetupToEditor (setup: PageSetup) : obj =
    let package = exportEditorObject { Document.empty with PageSetup = setup }
    package?PageSetup

/// Compatibility adapter for the browser's JavaScript package API.
let exportDocumentPackage (doc: Document) : obj =
    exportEditorObject doc

/// Export the complete editor package, including source-preservation metadata.
let exportTypedDocumentPackage (doc: Document) : WordEditorPackage.EditorPackage =
    TypedDocxExport.exportPackage doc

#if !FABLE_COMPILER
/// Open a DOCX in the same editor model used by the browser.
let loadBytesNative (bytes: byte[]) : Document option =
    WordEditorPackageReader.loadBytesNative bytes
    |> Option.map TypedDocxImport.importPackage
#endif
