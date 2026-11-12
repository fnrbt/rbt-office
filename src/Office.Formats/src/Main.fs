module Main

open Fable.Core
open Fable.Core.JsInterop
open ZipInterop
open System.Collections.Generic

// Public API for DOCX reading
module Word =
    /// Load a DOCX file from array buffer or blob
    let load = DocxReader.load

    /// Extract all text from a Word document
    let extractText = DocxReader.extractText

    /// Get document content (paragraphs and tables)
    let getContent (doc: DocxReader.WordDocument) = doc.Content |> List.toArray

    /// Get all paragraphs from document
    let paragraphs (doc: DocxReader.WordDocument) =
        doc.Content
        |> List.choose (function
            | DocxReader.Para p -> Some p
            | _ -> None)

    /// Get all tables from document
    let tables (doc: DocxReader.WordDocument) =
        doc.Content
        |> List.choose (function
            | DocxReader.Tbl t -> Some t
            | _ -> None)

    /// Get document styles
    let styles (doc: DocxReader.WordDocument) = doc.Styles

    /// Get document headers
    let headers (doc: DocxReader.WordDocument) = doc.Headers

    /// Get document footers
    let footers (doc: DocxReader.WordDocument) = doc.Footers

    /// Get document properties (title, author, etc.)
    let properties (doc: DocxReader.WordDocument) = doc.Properties

    /// Get all hyperlinks from document
    let hyperlinks (doc: DocxReader.WordDocument) =
        let extractFromPara (para: DocxReader.Paragraph) =
            para.Runs
            |> List.choose (fun run ->
                match run.Content with
                | DocxReader.Link link -> Some link
                | _ -> None)

        doc.Content
        |> List.collect (function
            | DocxReader.Para p -> extractFromPara p
            | DocxReader.Tbl t ->
                t.Rows
                |> List.collect (fun row ->
                    row.Cells
                    |> List.collect (fun cell ->
                        cell.Paragraphs
                        |> List.collect extractFromPara))
        )

    /// Get all images from document
    let images (doc: DocxReader.WordDocument) =
        let extractFromPara (para: DocxReader.Paragraph) =
            para.Runs
            |> List.choose (fun run ->
                match run.Content with
                | DocxReader.Image img -> Some img
                | _ -> None)

        doc.Content
        |> List.collect (function
            | DocxReader.Para p -> extractFromPara p
            | DocxReader.Tbl t ->
                t.Rows
                |> List.collect (fun row ->
                    row.Cells
                    |> List.collect (fun cell ->
                        cell.Paragraphs
                        |> List.collect extractFromPara))
        )

    /// Get all comments from document
    let comments (doc: DocxReader.WordDocument) = doc.Comments

    /// Get all revisions/track changes from document
    let revisions (doc: DocxReader.WordDocument) = doc.Revisions

    /// Get footnotes from document
    let footnotes (doc: DocxReader.WordDocument) = doc.Footnotes

    /// Get endnotes from document
    let endnotes (doc: DocxReader.WordDocument) = doc.Endnotes

    /// Get shapes and text boxes from document
    let shapes (doc: DocxReader.WordDocument) = doc.Shapes

// Public API for Document Processors
module Processors =
    /// Portable document model; native callers retain typed source/package
    /// information without passing through the browser object adapter.
    let forTypedEditorPackage = WordEditorPackageReader.ofDocument
#if !FABLE_COMPILER
    let forTypedEditorPackageBytes = WordEditorPackageReader.loadBytesNative
#endif
#if FABLE_COMPILER
    /// A copy of the file a package was opened from, when it was given as bytes: kept with what
    /// the writer needs (Preserved.OpenedBytes), it is what an unchanged document is saved as
    /// (see DocumentWriter.writeFromEditorPackage).
    [<Emit("($0 instanceof Uint8Array ? $0.slice() : ($0 instanceof ArrayBuffer ? new Uint8Array($0.slice(0)) : null))")>]
    let private openedBytesOf (_fileData: obj) : obj = jsNative

    let forEditorPackage (fileData: obj) : Fable.Core.JS.Promise<obj> =
        promise {
            let! archive = loadArchive fileData
            let! docOpt = DocxReader.parseDocument archive

            match docOpt with
            | Some doc ->
                let typed = WordEditorPackageReader.ofDocument doc
                let opened = openedBytesOf fileData
                let typed =
                    if isNull opened then typed
                    else { typed with
                             Preserved = typed.Preserved |> Option.map (fun preserved ->
                                 { preserved with OpenedBytes = Some (unbox<byte[]> opened) }) }
                return WordEditorPackage.packageToEditorObject typed
            | None ->
                return createObj [
                    "Paragraphs" ==> [||]
                    "Styles" ==> [||]
                ]
        }

    // ---- a package opened front first ---------------------------------------------
    // forEditorPackage reads the whole main part, makes blocks of all of it and
    // editor items of all the blocks before it answers: for a long document that
    // is most of the time until anything is shown. openProgressive answers once
    // the parts around the body are read and the first blocks are there; the
    // rest is read in slices of a few milliseconds between which input and
    // frames come first. The object it gives is the editor package - the same
    // fields; Paragraphs holds what is read so far - with a Progressive member:
    //
    //   Progressive.IsComplete      everything is read: the package is then, field for
    //                               field, what forEditorPackage gives for the document
    //   Progressive.AtOnce          the main part could not be read front first and was
    //                               read whole when the package was opened
    //   Progressive.Progress()      how far the body is read, 0 .. 1
    //   Progressive.readMore(ms)    read on for about that long now; true when complete
    //   Progressive.complete()      read the rest now
    //   Progressive.start()         read the rest in slices from now on (done when the
    //                               package is opened, unless { autoStart: false })
    //   Progressive.subscribe(f)    f(first, count, complete) after items were added to
    //                               Paragraphs (count of them from index first; the shapes
    //                               come last, with complete); returns the unsubscribe function
    //   Progressive.whenComplete()  a promise of the package, complete
    //   Progressive.Slices / LongestSliceMs / LongestSlice   how many slices read it, what the longest
    //                               took (with what its listeners did: the page could not answer for
    //                               that long) and which it was (1: the first); of that slice,
    //                               LongestSliceItems (the items it added), LongestSlicePassMs (its
    //                               longest pass of reading) and LongestSliceNotifyMs (the listeners)
    //   Progressive.Failed / Error  the body could not be read to its end (its text is not well
    //                               formed from some element on, or an element could not be made
    //                               blocks of): Failed is true and Error says why. Paragraphs holds
    //                               what was read before that element (the shapes anchored in it
    //                               after the blocks), IsComplete stays false and nothing more is
    //                               read: readMore and complete return at once, the listeners were
    //                               called a last time (complete false, Failed set by then), and
    //                               whenComplete - a promise of the complete package - rejects with
    //                               the error. Preserved is not filled in: such a package cannot be
    //                               written back. (A main part that is read at once fails as
    //                               forEditorPackage does: openProgressive rejects.)
    //   Progressive.cancel()        stop reading: no further slice is scheduled (one that is
    //                               waiting does nothing), and the package lets go of what it
    //                               read from - the archive's parts, the text of the main part.
    //                               The package is then as one that failed (Failed, and
    //                               Cancelled, are true; Error says that it was cancelled;
    //                               Paragraphs holds what was read; the listeners are called a
    //                               last time; whenComplete rejects). True when it cancelled
    //                               anything: false for a package that is read, failed or
    //                               cancelled already. An editor or a viewer that shows the
    //                               package cancels it when it is done with it (see Word's
    //                               importFromDocxPackage).
    //
    // Preserved (what the writer needs of the package) is an object from the
    // start and is filled in when the body is read. PageSetup is, until then,
    // the page setup of the section the blocks read so far end in - the pages
    // of a section are on its paper before the paragraph that ends it is read -
    // and Progressive.PageSetupVersion counts its changes.

    [<Emit("performance.now()")>]
    let private nowMs () : float = jsNative
    [<Emit("Object.assign($0, $1)")>]
    let private assignFields (_target: obj) (_source: obj) : unit = jsNative
    [<Emit("$0.push($1)")>]
    let private pushItem (_target: obj array) (_item: obj) : unit = jsNative
    [<Emit("$0($1, $2, $3)")>]
    let private callListener (_listener: obj) (_first: int) (_count: int) (_complete: bool) : unit = jsNative
    [<Emit("(typeof document !== 'undefined' && document.hidden === true)")>]
    let private pageIsHidden () : bool = jsNative
    // (As the layout's slices are scheduled, see Renderer.afterThisTask: a task of its own after
    // the current one, so that input events and frames that are waiting come first. In two hops:
    // a timer that came due during the slice is made ready when the slice's task ends - after a
    // next slice posted from inside it, which then ran first, and the page's timers saw two
    // slices as one. The next slice is posted from a task of nothing, behind the timers.)
    [<Emit("""(function (run) {
        var hop = function (next) {
            if (typeof scheduler !== 'undefined' && scheduler && typeof scheduler.postTask === 'function') { scheduler.postTask(next, { priority: 'user-visible' }); return; }
            if (typeof MessageChannel === 'function') { var channel = new MessageChannel(); channel.port1.onmessage = function () { channel.port1.close(); next(); }; channel.port2.postMessage(0); return; }
            setTimeout(next, 0);
        };
        hop(function () { hop(run); });
    })($0)""")>]
    let private afterThisTask (_callback: unit -> unit) : unit = jsNative
    [<Emit("(($0 != null && typeof $0[$1] === 'number') ? $0[$1] : $2)")>]
    let private numberOption (_options: obj) (_name: string) (_fallback: float) : float = jsNative
    [<Emit("(($0 != null && typeof $0[$1] === 'boolean') ? $0[$1] : $2)")>]
    let private booleanOption (_options: obj) (_name: string) (_fallback: bool) : bool = jsNative

    /// Open a DOCX front first (see above). `options` (optional):
    ///   firstBlocks  how many blocks of the body are read before the package is handed over
    ///                (default 64: two screens)
    ///   autoStart    false: the rest is read only when asked for (start / readMore / complete)
    ///   sliceMs      how long a slice aims to take, with what its listeners do (default 10: a
    ///                slice that meets a block that costs many times the others, or a collection
    ///                of garbage, still ends within a frame)
    let openProgressive (fileData: obj) (options: obj) : Fable.Core.JS.Promise<obj> =
        promise {
            let! archive = loadArchive fileData
            let! source = DocxReader.sourceOfArchive archive
            match DocxReader.openProgressiveFromSource false source with
            | None ->
                return createObj [
                    "Paragraphs" ==> [||]
                    "Styles" ==> [||]
                ]
            | Some opened ->
                let front = opened.Front
                // (Let go of when the reading is cancelled: the reader holds the archive's parts
                // and the text of the main part.)
                let mutable current : DocxReader.ProgressiveReader option = Some opened
                let cancelledReader : DocxReader.ProgressiveReader =
                    { opened with
                        IsDone = (fun () -> true)
                        Failure = (fun () -> Some "The reading of the document was cancelled")
                        Read = (fun _ -> { First = 0; Blocks = []; Shapes = [] })
                        OpenSection = (fun () -> None)
                        BlockCount = (fun () -> 0)
                        Progress = (fun () -> 0.0)
                        ImagesOf = (fun _ -> Map.empty)
                        ImageOf = (fun _ -> None)
                        Finish = (fun () -> failwith "The reading of the document was cancelled") }
                let reader () : DocxReader.ProgressiveReader = match current with | Some reader -> reader | None -> cancelledReader
                let firstBlocks = max 1 (int (numberOption options "firstBlocks" 64.0))
                let sliceMs = max 1.0 (numberOption options "sliceMs" 10.0)
                let convert = WordEditorPackageReader.converter front
                let paragraphs : obj array = [||]
                let shapeItems = ResizeArray<obj>()
                let listeners = ResizeArray<obj>()
                let waiters = ResizeArray<(obj -> unit) * (exn -> unit)>()
                let preserved = createObj []
                let openedBytes = openedBytesOf fileData
                let state = createObj [ "IsComplete" ==> false; "AtOnce" ==> opened.AtOnce; "Failed" ==> false; "Cancelled" ==> false; "Error" ==> null; "PageSetupVersion" ==> 0; "Slices" ==> 0; "LongestSliceMs" ==> 0.0; "LongestSlice" ==> 0 ]
                let blocksOf (parts: DocxReader.HeaderFooter list) = parts |> List.collect (fun part -> part.Blocks)
                let noteBlocks (notes: DocxReader.Note list) = notes |> List.collect (fun note -> note.Blocks)
                let headerImages = (reader ()).ImagesOf (blocksOf front.Headers @ blocksOf front.Footers)
                let noteImages = (reader ()).ImagesOf (noteBlocks front.Footnotes @ noteBlocks front.Endnotes)
                let commentImages = (reader ()).ImagesOf (front.Comments |> List.collect (fun comment -> comment.Blocks))
                let mergeImages images additional =
                    additional |> Map.fold (fun kept id source -> Map.add id source kept) images
                let frontWithImages =
                    { front with Images = mergeImages (mergeImages (mergeImages front.Images headerImages) noteImages) commentImages }
                let typedFront = WordEditorPackageReader.ofDocument frontWithImages
                let package =
                    WordEditorPackage.packageToEditorObject
                        { typedFront with Blocks = [||]; Preserved = None }
                package?Paragraphs <- paragraphs
                package?Preserved <- preserved
                package?Progressive <- state
                let mutable complete = false
                let mutable failed : exn option = None
                // The page setup the package states: that of the section the blocks read so far end
                // in, until the body is read (see ProgressiveReader.OpenSection).
                let mutable pageSetup = front.PageSetup
                let statePageSetup (setup: DocxReader.PageSetup option) =
                    if setup <> pageSetup then
                        pageSetup <- setup
                        package?PageSetup <- (setup |> Option.map WordEditorPackageReader.pageSetup |> WordEditorPackage.pageSetupToEditorObject)
                        state?PageSetupVersion <- unbox<int> (state?PageSetupVersion) + 1
                // Blocks a pass of about a millisecond reads: what the last ones say.
                let mutable step = 8
                let add (batch: DocxReader.ProgressiveBatch) =
                    if not batch.Blocks.IsEmpty then
                        for item in convert batch.Blocks ((reader ()).ImagesOf batch.Blocks) |> WordEditorPackage.blocksToEditorObjects do
                            pushItem paragraphs item
                    if not batch.Shapes.IsEmpty then
                        let shaped =
                            WordEditorPackageReader.shapes convert front.Images (reader ()).ImageOf batch.Shapes
                            |> WordEditorPackage.blocksToEditorObjects
                        shapeItems.AddRange shaped
                    if not ((reader ()).IsDone ()) then statePageSetup ((reader ()).OpenSection ())
                    // (What the reader could not read: the blocks before it were in the batch.)
                    match (reader ()).Failure () with
                    | Some message -> failwith message
                    | None -> ()
                /// The body is read: the shapes after the blocks, what is kept for the writer, the page setup.
                let finish () =
                    let final = (reader ()).Finish ()
                    let shapes = shapeItems.ToArray()
                    assignFields preserved (final |> WordEditorPackage.preservedOfDocument |> Some |> WordEditorPackage.preservedToEditorObject)
                    // (With the rest of what the writer needs, as forEditorPackage has it.)
                    preserved?OpenedBytes <- openedBytes
                    for item in shapes do pushItem paragraphs item
                    statePageSetup final.PageSetup
                    complete <- true
                    state?IsComplete <- true
                    // (Read: the reader, and what it read from, is not needed any more.)
                    let blockCount = (reader ()).BlockCount ()
                    state?BlockCount <- (fun () -> blockCount)
                    current <- None
                let notify (first: int) =
                    let count = paragraphs.Length - first
                    if count > 0 || complete || failed.IsSome then
                        for listener in listeners.ToArray() do
                            if listeners.Contains listener then callListener listener first count complete
                    if complete && waiters.Count > 0 then
                        let waiting = waiters.ToArray()
                        waiters.Clear()
                        for resolve, _ in waiting do resolve package
                /// The body cannot be read on: what was read is what there is (the shapes anchored in
                /// it after the blocks, as at the end).
                let fail (error: exn) =
                    if failed.IsNone then
                        failed <- Some error
                        for item in shapeItems do pushItem paragraphs item
                        state?Failed <- true
                        state?Error <- error.Message
                let rejectWaiters () =
                    match failed with
                    | Some error when waiters.Count > 0 ->
                        let waiting = waiters.ToArray()
                        waiters.Clear()
                        for _, reject in waiting do reject error
                    | _ -> ()
                // What the listeners took after the last slice (they make their blocks of the items):
                // a slice reads for its time less that.
                let mutable notifyMs = 0.0
                let mutable longestPass = 0.0
                /// Read on for about `budget` milliseconds (at least one step).
                let readMore (budget: float) : bool =
                    if not complete && failed.IsNone then
                        let first = paragraphs.Length
                        let started = nowMs ()
                        let mutable more = true
                        try
                            while more do
                                let passStart = nowMs ()
                                let batch = (reader ()).Read step
                                add batch
                                if (reader ()).IsDone () then
                                    finish ()
                                    more <- false
                                else
                                    let took = max 0.05 (nowMs () - passStart)
                                    if took > longestPass then longestPass <- took
                                    // (Steps of about half a millisecond, no more than twice the last one: the
                                    // slice ends near its time, and blocks that cost more than the last ones
                                    // said - a table after headings - are met a few at a time.)
                                    // (And never many: a run of empty paragraphs says that hundreds fit, the
                                    // paragraphs of text after them do not.)
                                    step <- max 1 (min (min 32 (step * 2)) (int (float (max 1 batch.Blocks.Length) * 0.5 / took)))
                                    if nowMs () - started >= budget then more <- false
                        with error -> fail error
                        let told = nowMs ()
                        notify first
                        notifyMs <- nowMs () - told
                        rejectWaiters ()
                    complete
                let mutable scheduled = false
                let rec pump () =
                    if not complete && failed.IsNone && not scheduled then
                        scheduled <- true
                        afterThisTask (fun () ->
                            scheduled <- false
                            if not complete && failed.IsNone then
                                let sliceStart = nowMs ()
                                let itemsBefore = paragraphs.Length
                                longestPass <- 0.0
                                // A page that is not shown gets its tasks rarely: longer slices then.
                                (try readMore (if pageIsHidden () then 200.0 else max 2.0 (sliceMs - notifyMs)) |> ignore with _ -> ())
                                // (What a slice took with everything that followed it: the listeners make
                                // their blocks of the items and show them.)
                                let took = nowMs () - sliceStart
                                state?Slices <- unbox<int> (state?Slices) + 1
                                if took > unbox<float> (state?LongestSliceMs) then
                                    state?LongestSliceMs <- took
                                    state?LongestSlice <- unbox<int> (state?Slices)
                                    // (What it was made of: the listeners' part, the items it read, its longest pass.)
                                    state?LongestSliceNotifyMs <- notifyMs
                                    state?LongestSliceItems <- paragraphs.Length - itemsBefore
                                    state?LongestSlicePassMs <- longestPass
                                pump ())
                state?Progress <- (fun () -> if complete then 1.0 else (reader ()).Progress ())
                state?BlockCount <- (fun () -> (reader ()).BlockCount ())
                state?readMore <- (fun (budget: float) -> readMore (if isNull (box budget) then sliceMs else budget))
                state?complete <- (fun () -> readMore System.Double.PositiveInfinity |> ignore)
                state?start <- (fun () -> pump ())
                state?cancel <- (fun () ->
                    if complete || failed.IsSome then false
                    else
                        let first = paragraphs.Length
                        let progress = (reader ()).Progress ()
                        let blockCount = (reader ()).BlockCount ()
                        fail (exn "The reading of the document was cancelled")
                        state?Cancelled <- true
                        // (What says how far it got stays what it was.)
                        state?Progress <- (fun () -> progress)
                        state?BlockCount <- (fun () -> blockCount)
                        current <- None
                        notify first
                        rejectWaiters ()
                        true)
                state?subscribe <- (fun (listener: obj) ->
                    listeners.Add listener
                    fun () -> listeners.Remove listener |> ignore)
                state?whenComplete <- (fun () ->
                    Fable.Core.JS.Constructors.Promise.Create(fun resolve reject ->
                        match failed with
                        | Some error -> reject error
                        | None ->
                            if complete then resolve package
                            else
                                waiters.Add ((resolve, reject))
                                pump ()))
                // The start of the document.
                let begin' () =
                    try
                        add ((reader ()).Read firstBlocks)
                        if (reader ()).IsDone () then finish ()
                        elif booleanOption options "autoStart" true then pump ()
                    with error -> fail error
                begin' ()
                return package
        }

    /// Whether a package is read to its end (one that forEditorPackage made always is).
    let isLoadComplete (package: obj) : bool =
        isNull package || isNull (package?Progressive) || unbox<bool> (package?Progressive?IsComplete)

    /// A promise of the package once it is read to its end (reading on in slices until then).
    /// It rejects, with the error, for a package whose body cannot be read to its end
    /// (Progressive.Failed; what was read stays in the package).
    let whenLoadComplete (package: obj) : Fable.Core.JS.Promise<obj> =
        if isLoadComplete package then Fable.Core.JS.Constructors.Promise.Create(fun resolve _ -> resolve package)
        else unbox<Fable.Core.JS.Promise<obj>> (package?Progressive?whenComplete ())

    /// Read the rest of a package now. (Of a package that cannot be read to its end it reads what
    /// there is: isLoadComplete stays false, see isLoadFailed.)
    let completeLoad (package: obj) : unit =
        if not (isLoadComplete package) then package?Progressive?complete () |> ignore

    /// Whether the body of a package could not be read to its end (Progressive.Failed).
    let isLoadFailed (package: obj) : bool =
        not (isNull package) && not (isNull (package?Progressive)) && unbox<bool> (package?Progressive?Failed)

    /// Why not (Progressive.Error); null for a package that did not fail.
    let getLoadError (package: obj) : string =
        if isLoadFailed package then unbox<string> (package?Progressive?Error) else null
#endif

// Public API for Word Writing
module WordWriter =
    /// Write paragraphs to DOCX (Phase 1: plain text only)
    let write (paragraphs: string list) = DocxWriter.writeMinimalDocx paragraphs


#if FABLE_COMPILER
    let writeFromEditorPackage (package: obj) =
        DocumentWriter.writeFromEditorPackage package

    /// What the last writeFromEditorPackage made of the body of a document
    /// that was opened: { PatchedParagraphs, ModelParagraphsInTables,
    /// ModelBlocks, ModelBlockParagraphs, SourceBlocks, DerivedParagraphs,
    /// GeneratedParagraphs, Reasons } - the paragraphs that were changed in
    /// the XML they were read from, the paragraphs of tables that were
    /// written from the model into the table's own XML, the blocks (and the
    /// paragraphs in them) that were written from the model whole, the
    /// blocks that went back as the XML they were read from, the paragraphs
    /// of blocks written from the model that were made of the XML of a
    /// paragraph that was read (a paragraph of a table whose rows changed,
    /// the second half of a split paragraph), the paragraphs the model
    /// wrote (of all parts), and why the first few blocks did not go back
    /// as they were read.
    let getLastSaveReport () : obj =
        DocumentWriter.lastSaveReport ()
#endif

    /// Create document builder (Phase 3)
    let createDocument () = DocumentBuilder.create ()

    /// Add paragraph to builder
    let addParagraph = DocumentBuilder.addParagraph

    /// Add paragraph with runs to builder
    let addParagraphWithRuns = DocumentBuilder.addParagraphWithRuns

#if FABLE_COMPILER
    /// Save document
    let saveDocument = DocumentBuilder.save
#endif

    /// Helper to create a run
    let createRun = DocumentBuilder.createRun

    /// Create a simple table from 2D array (Phase 4)
    let createSimpleTable = TableWriter.createSimpleTable

    /// Create a table with borders (Phase 4)
    let createTableWithBorders = TableWriter.createTableWithBorders

    /// Generate table XML (Phase 4)
    let generateTableXml = TableWriter.generateTable

// Public API for Excel Writing
module ExcelWriter =
    /// Write 2D array to XLSX (Phase 1)
    let write (sheets: XlsxWriter.SheetData list) = XlsxWriter.writeMinimalXlsx sheets

    /// Write formatted XLSX (Phase 2)
    let writeFormatted (sheets: XlsxWriter.FormattedSheet list) = XlsxWriter.writeFormattedXlsx sheets

    /// Create workbook builder (Phase 3)
    let createWorkbook () = WorkbookBuilder.create ()

    /// Add sheet to workbook
    let addSheet = WorkbookBuilder.addSheet

    /// Add simple sheet to workbook
    let addSimpleSheet = WorkbookBuilder.addSimpleSheet

    /// Save workbook
    let saveWorkbook = WorkbookBuilder.saveWorkbook

    /// Add row to sheet builder
    let addRow = WorkbookBuilder.addRow

    /// Add simple row to sheet builder
    let addSimpleRow = WorkbookBuilder.addSimpleRow

    /// Helper to create a formatted cell
    let createCell = WorkbookBuilder.createCell

    /// Helper to create a header cell
    let createHeaderCell = WorkbookBuilder.createHeaderCell

    /// Helper to create a numeric cell with format
    let createNumericCell = WorkbookBuilder.createNumericCell

    /// Helper to create a formula cell (Phase 4)
    let createFormulaCell = WorkbookBuilder.createFormulaCell

    /// Add merged range to sheet builder (Phase 4)
    let addMergedRange = WorkbookBuilder.addMergedRange

    /// Create merged cell range (Phase 4)
    let createMergedRange = XlsxWriter.createMergedRange

// Public API for XLSX reading
module Excel =
    type private CellRangeBounds = {
        RowMin: int
        RowMax: int
        ColMin: int
        ColMax: int
    }

    let private columnIndex (letters: string) =
        let mutable acc = 0
        for ch in letters do
            if System.Char.IsLetter ch then
                acc <- acc * 26 + (int (System.Char.ToUpperInvariant ch) - int 'A' + 1)
        acc - 1

    let private refToCoord (ref: string) =
        let mutable index = 0
        let mutable col = 0
        let mutable row = 0
        let mutable hasCol = false
        let mutable hasRow = false
        let mutable inRow = false
        let mutable valid = not (System.String.IsNullOrWhiteSpace ref)
        while valid && index < ref.Length do
            let ch = ref.[index]
            if ch = '$' then
                ()
            elif System.Char.IsLetter ch && not inRow then
                hasCol <- true
                let digit = int (System.Char.ToUpperInvariant ch) - int 'A' + 1
                if digit < 1 || digit > 26 || col > (System.Int32.MaxValue - digit) / 26 then
                    valid <- false
                else
                    col <- col * 26 + digit
            elif System.Char.IsDigit ch then
                inRow <- true
                hasRow <- true
                let digit = int ch - int '0'
                if row > (System.Int32.MaxValue - digit) / 10 then
                    valid <- false
                else
                    row <- row * 10 + digit
            else
                valid <- false
            index <- index + 1
        if valid && hasCol && hasRow && row > 0 && col > 0 then Some (row - 1, col - 1)
        else None

    let private tryParseRangeBounds (range: string) =
        let parts = range.Split(':')
        match parts.Length with
        | 1 ->
            refToCoord parts.[0]
            |> Option.map (fun (r, c) -> { RowMin = r; RowMax = r; ColMin = c; ColMax = c })
        | 2 ->
            match refToCoord parts.[0], refToCoord parts.[1] with
            | Some (r1, c1), Some (r2, c2) ->
                Some {
                    RowMin = min r1 r2
                    RowMax = max r1 r2
                    ColMin = min c1 c2
                    ColMax = max c1 c2
                }
            | _ -> None
        | _ -> None

    let private rangeBoundsContains row col range =
        row >= range.RowMin
        && row <= range.RowMax
        && col >= range.ColMin
        && col <= range.ColMax

    let private rangeContains (row: int) (col: int) (range: string) =
        tryParseRangeBounds range
        |> Option.exists (rangeBoundsContains row col)

    let private rangeListContains row col (ranges: string) =
        ranges.Split([| ' ' |], System.StringSplitOptions.RemoveEmptyEntries)
        |> Array.exists (rangeContains row col)

    let private listDropdownRangeBounds (sheet: XlsxReader.Worksheet) =
        let ranges = ResizeArray<CellRangeBounds>()
        for validation in sheet.DataValidations do
            if validation.ValidationType = "list" then
                for range in validation.Range.Split([| ' ' |], System.StringSplitOptions.RemoveEmptyEntries) do
                    match tryParseRangeBounds range with
                    | Some bounds -> ranges.Add bounds
                    | None -> ()
        ranges.ToArray()

    let private rangeBoundsArrayContains row col (ranges: CellRangeBounds array) =
        let mutable found = false
        let mutable index = 0
        while not found && index < ranges.Length do
            if rangeBoundsContains row col ranges.[index] then
                found <- true
            index <- index + 1
        found

    let private optString (value: string option) =
        value |> Option.defaultValue null

    let private objField name (value: obj) =
#if FABLE_COMPILER
        name ==> value
#else
        name, value
#endif

    let private objFromFields fields : obj =
#if FABLE_COMPILER
        createObj fields
#else
        let result = Dictionary<string, obj>()
        for name, value in fields do
            result.[name] <- value
        box result
#endif

    let private addField (fields: ResizeArray<_>) name value =
        fields.Add(objField name value)

    let private addFieldWhen (fields: ResizeArray<_>) condition name value =
        if condition then addField fields name value

    let private addFieldOption (fields: ResizeArray<_>) name value =
        match value with
        | Some v -> addField fields name (box v)
        | None -> ()

    let private addFieldObjOption (fields: ResizeArray<_>) name value =
        match value with
        | Some v -> addField fields name v
        | None -> ()

    let private cellRefText (cellRef: XlsxReader.CellReference) =
        sprintf "%s%d" cellRef.Column cellRef.Row

    let private optFloat (value: float option) : obj =
        match value with
        | Some v -> box v
        | None -> null

    let private borderObj (side: XlsxReader.CellBorderSide option) : obj =
        match side with
        | Some s ->
            objFromFields [
                objField "style" s.Style
                objField "color" (optString s.Color)
            ]
        | None -> null

    let private cssColor (value: string option) : obj =
        match value with
        | Some raw ->
            let trimmed = raw.TrimStart('#')
            let cleaned = if trimmed.Length = 8 then trimmed.Substring(2) else trimmed
            if cleaned.Length = 6 then box ("#" + cleaned) else null
        | None -> null

    let private richRunObj (text: string, props: XlsxReader.RunProps) =
        let fields = ResizeArray()
        addField fields "text" text
        addFieldWhen fields props.Bold "bold" true
        addFieldWhen fields props.Italic "italic" true
        addFieldWhen fields props.Underline "underline" true
        addFieldWhen fields props.UnderlineDouble "underlineDouble" true
        addFieldWhen fields props.Strikethrough "strikethrough" true
        let color = cssColor props.Color
        if not (isNull color) then
            addField fields "color" color
        addFieldOption fields "fontName" props.FontName
        addFieldObjOption fields "fontSize" (props.FontSize |> Option.map box)
        match props.VertAlign with
        | XlsxReader.VertAlignSuperscript -> addField fields "vertAlign" "superscript"
        | XlsxReader.VertAlignSubscript -> addField fields "vertAlign" "subscript"
        | XlsxReader.VertAlignBaseline -> ()
        objFromFields fields

    let private richRunsForCell (workbook: XlsxReader.Workbook) (cell: XlsxReader.Cell) =
        let lift runs =
            if List.length runs <= 1 then [||]
            else runs |> List.map richRunObj |> List.toArray
        if not (List.isEmpty cell.InlineRuns) then
            lift cell.InlineRuns
        else
            match cell.Value with
            | XlsxReader.SharedString idx
                when idx >= 0
                     && idx < workbook.SharedStringRuns.Length
                     && not (List.isEmpty workbook.SharedStringRuns.[idx]) ->
                lift workbook.SharedStringRuns.[idx]
            | _ -> [||]

    let private richRunsObjForCell (workbook: XlsxReader.Workbook) (cell: XlsxReader.Cell) =
        let lift runs =
            match runs with
            | _ :: _ :: _ -> box (runs |> List.map richRunObj |> List.toArray)
            | _ -> null
        if not (List.isEmpty cell.InlineRuns) then
            lift cell.InlineRuns
        else
            match cell.Value with
            | XlsxReader.SharedString idx
                when idx >= 0
                     && idx < workbook.SharedStringRuns.Length
                     && not (List.isEmpty workbook.SharedStringRuns.[idx]) ->
                lift workbook.SharedStringRuns.[idx]
            | _ -> null

    /// Formula text (no leading "="), or null for a plain value cell.
    let private cellFormula (cell: XlsxReader.Cell) =
        match cell.Formula with
        | Some f -> f
        | None -> null

    /// The cell's value — for a formula cell, the result cached in the file.
    /// null for empty cells, error cells (see `cellErrorObj`) and formulas
    /// the file carries no cached result for.
    let private cellValueObj (workbook: XlsxReader.Workbook) (cell: XlsxReader.Cell) : obj =
        match XlsxReader.resolveValue workbook cell.Value with
        | XlsxReader.String s -> box s
        | XlsxReader.Number n -> box n
        | XlsxReader.Boolean b -> box b
        | XlsxReader.Date d -> box d
        | _ -> null

    /// Error code without the leading "#" ("DIV/0!", "N/A") for an error
    /// cell or a formula whose cached result is an error; null otherwise.
    let private cellErrorObj (cell: XlsxReader.Cell) : obj =
        match cell.Value with
        | XlsxReader.Error e -> box e
        | _ -> null

    let private listDropdownAt row col (sheet: XlsxReader.Worksheet) =
        sheet.DataValidations
        |> List.exists (fun dv ->
            dv.ValidationType = "list"
            && rangeListContains row col dv.Range)

    let private commentAtCoord row col (sheet: XlsxReader.Worksheet) =
        sheet.Comments
                |> List.tryFind (fun c ->
            c.CellRef.Row - 1 = row
            && columnIndex c.CellRef.Column = col)

    let private addStyleFields fields (style: XlsxReader.CellStyle option) =
        match style with
        | Some s ->
            addFieldWhen fields s.Bold "bold" true
            addFieldWhen fields s.Italic "italic" true
            addFieldWhen fields s.Underline "underline" true
            addFieldWhen fields s.UnderlineDouble "underlineDouble" true
            addFieldWhen fields s.Strikethrough "strikethrough" true
            addFieldOption fields "fontName" s.FontName
            addFieldOption fields "fontColor" s.FontColor
            addFieldOption fields "fillColor" s.FillColor
            addFieldObjOption fields "fontSize" (s.FontSize |> Option.map box)
            addFieldOption fields "numberFormat" s.NumberFormat
            addFieldOption fields "horizontalAlign" s.HorizontalAlign
            addFieldOption fields "verticalAlign" s.VerticalAlign
            addFieldWhen fields s.WrapText "wrapText" true
            addFieldObjOption fields "borderTop" (s.BorderTop |> Option.map (fun side -> borderObj (Some side)))
            addFieldObjOption fields "borderBottom" (s.BorderBottom |> Option.map (fun side -> borderObj (Some side)))
            addFieldObjOption fields "borderLeft" (s.BorderLeft |> Option.map (fun side -> borderObj (Some side)))
            addFieldObjOption fields "borderRight" (s.BorderRight |> Option.map (fun side -> borderObj (Some side)))
            addFieldObjOption fields "borderDiagonal" (s.BorderDiagonal |> Option.map (fun side -> borderObj (Some side)))
            addFieldWhen fields s.DiagonalUp "diagonalUp" true
            addFieldWhen fields s.DiagonalDown "diagonalDown" true
        | None -> ()

    let private styleObj (style: XlsxReader.CellStyle option) =
        let fields = ResizeArray()
        addStyleFields fields style
        objFromFields fields

    let private effectiveStyleIndex (workbook: XlsxReader.Workbook) (cell: XlsxReader.Cell) =
        match cell.StyleIndex with
        | Some idx when idx >= 0 && idx < workbook.Styles.Length -> idx
        | None when workbook.Styles.Length > 0 && cell.Style.IsSome -> 0
        | _ -> -1

    let private commentCoordSet (sheet: XlsxReader.Worksheet) =
        let comments = HashSet<struct (int * int)>()
        for comment in sheet.Comments do
            let row = comment.CellRef.Row - 1
            let col = columnIndex comment.CellRef.Column
            if row >= 0 && col >= 0 then
                comments.Add(struct (row, col)) |> ignore
        comments

    let private cellWithStyleObjWithFlags (workbook: XlsxReader.Workbook option) (c: XlsxReader.Cell) hasComment hasListDropdown : obj =
        let style = c.Style
        let fields = ResizeArray()
        addField fields "ref" (cellRefText c.Reference)
        match workbook with
        | Some wb ->
            let value = cellValueObj wb c
            if not (isNull value) then
                addField fields "value" value
        | None -> ()
        let formula = cellFormula c
        if not (isNull formula) then
            addField fields "formula" formula
        let error = cellErrorObj c
        if not (isNull error) then
            addField fields "error" error
        match style with
        | Some _ -> addStyleFields fields style
        | None -> ()
        addFieldWhen fields hasComment "hasComment" true
        addFieldWhen fields hasListDropdown "hasListDropdown" true
        match workbook with
        | Some wb ->
            let richRuns = richRunsForCell wb c
            if richRuns.Length > 0 then
                addField fields "richRuns" richRuns
        | None -> ()
        objFromFields fields

    let private cellWithStyleObj (workbook: XlsxReader.Workbook option) (sheet: XlsxReader.Worksheet) (c: XlsxReader.Cell) : obj =
        let row = c.Reference.Row - 1
        let col = columnIndex c.Reference.Column
        cellWithStyleObjWithFlags workbook c (commentAtCoord row col sheet |> Option.isSome) (listDropdownAt row col sheet)

    /// Load an XLSX file from array buffer or blob
    let load = XlsxReader.load

    /// Load only one worksheet from an XLSX file. None selects the first sheet.
    let loadSelected = XlsxReader.loadSelected

    /// Get all worksheet names, hidden ones included (formulas may refer to
    /// them). Pair with `sheetStates`, or use `visibleSheetNames` for tabs.
    let sheetNames (wb: XlsxReader.Workbook) =
        wb.Sheets |> List.map (fun s -> s.Name) |> List.toArray

    let private sheetStateText (state: XlsxReader.SheetState) =
        match state with
        | XlsxReader.SheetVisible -> "visible"
        | XlsxReader.SheetHidden -> "hidden"
        | XlsxReader.SheetVeryHidden -> "veryHidden"

    /// "visible" | "hidden" | "veryHidden" for each sheet, parallel to `sheetNames`.
    let sheetStates (wb: XlsxReader.Workbook) =
        wb.Sheets |> List.map (fun s -> sheetStateText s.State) |> List.toArray

    /// Names of the sheets Excel shows a tab for.
    let visibleSheetNames (wb: XlsxReader.Workbook) =
        wb.Sheets
        |> List.filter (fun s -> s.State = XlsxReader.SheetVisible)
        |> List.map (fun s -> s.Name)
        |> List.toArray

    /// Get a worksheet by name
    let sheet name (wb: XlsxReader.Workbook) =
        wb.Sheets |> List.tryFind (fun s -> s.Name = name)

    /// Convert worksheet to 2D array
    let toArray = XlsxReader.sheetToArray

    /// Used range dimensions as one-based row/column counts.
    let dimensions (sheet: XlsxReader.Worksheet) =
        let mutable rowCount = 0
        let mutable colCount = 0
        for row in sheet.Rows do
            if row.Index > rowCount then
                rowCount <- row.Index
            for cell in row.Cells do
                let col = columnIndex cell.Reference.Column + 1
                if col > colCount then
                    colCount <- col
        objFromFields [
            objField "rowCount" (box rowCount)
            objField "colCount" (box colCount)
        ]

    let private findCellByRef (cellRef: XlsxReader.CellReference) (sheet: XlsxReader.Worksheet) =
        sheet.Rows
        |> List.tryFind (fun row -> row.Index = cellRef.Row)
        |> Option.bind (fun row ->
            row.Cells
            |> List.tryFind (fun cell -> cell.Reference = cellRef))

    /// Get cell value at reference (e.g., "A1")
    let cellAt (ref: string) (sheet: XlsxReader.Worksheet) =
        let cellRef = XlsxReader.Xlsx.parseCellRef ref
        sheet
        |> findCellByRef cellRef
        |> Option.map (fun c -> c.Value)

    /// Get cell with style at reference
    let cellWithStyleAt (ref: string) (sheet: XlsxReader.Worksheet) : obj =
        let cellRef = XlsxReader.Xlsx.parseCellRef ref
        sheet
        |> findCellByRef cellRef
        |> Option.map (cellWithStyleObj None sheet)
        |> Option.defaultValue null

    /// Stream actual worksheet cells with values/styles in document order.
    /// This avoids the browser importer's old coordinate-by-coordinate scan,
    /// which is pathological for wide style-only sheets.
    let cellsWithStyle (workbook: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) =
        let comments = commentCoordSet sheet
        let dropdownRanges = listDropdownRangeBounds sheet
        let cells = ResizeArray<obj>()
        for row in sheet.Rows do
            for cell in row.Cells do
                let rowIndex = cell.Reference.Row - 1
                let colIndex = columnIndex cell.Reference.Column
                let hasComment =
                    rowIndex >= 0
                    && colIndex >= 0
                    && comments.Contains(struct (rowIndex, colIndex))
                let hasListDropdown =
                    rowIndex >= 0
                    && colIndex >= 0
                    && rangeBoundsArrayContains rowIndex colIndex dropdownRanges
                cells.Add(cellWithStyleObjWithFlags (Some workbook) cell hasComment hasListDropdown)
        cells.ToArray()

    /// Compact cell stream for importers. Slots are:
    /// row0, col0, value, formula, styleIndex, hasComment, hasListDropdown, richRuns.
    /// (Error results are only reported by `cellsColumnar` / `cellsWithStyle`.)
    let cellsCompact (workbook: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) =
        let comments = commentCoordSet sheet
        let dropdownRanges = listDropdownRangeBounds sheet
        let cells = ResizeArray<obj>()
        for row in sheet.Rows do
            for cell in row.Cells do
                let rowIndex = cell.Reference.Row - 1
                let colIndex = columnIndex cell.Reference.Column
                let hasComment =
                    rowIndex >= 0
                    && colIndex >= 0
                    && comments.Contains(struct (rowIndex, colIndex))
                let hasListDropdown =
                    rowIndex >= 0
                    && colIndex >= 0
                    && rangeBoundsArrayContains rowIndex colIndex dropdownRanges
                cells.Add(box [|
                    box rowIndex
                    box colIndex
                    cellValueObj workbook cell
                    box (cellFormula cell)
                    box (effectiveStyleIndex workbook cell)
                    box hasComment
                    box hasListDropdown
                    richRunsObjForCell workbook cell
                |])
        cells.ToArray()

    /// Columnar compact cell stream for importers. This avoids one tiny
    /// array/object per cell when importing very large sheets.
    let cellsColumnar (workbook: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) =
        let comments = commentCoordSet sheet
        let dropdownRanges = listDropdownRangeBounds sheet
        let rows = ResizeArray<int>()
        let cols = ResizeArray<int>()
        let values = ResizeArray<obj>()
        let formulas = ResizeArray<obj>()
        let styleIndexes = ResizeArray<int>()
        let hasComments = ResizeArray<bool>()
        let hasListDropdowns = ResizeArray<bool>()
        let richRuns = ResizeArray<obj>()
        let errors = ResizeArray<obj>()
        for row in sheet.Rows do
            for cell in row.Cells do
                let rowIndex = cell.Reference.Row - 1
                let colIndex = columnIndex cell.Reference.Column
                let hasComment =
                    rowIndex >= 0
                    && colIndex >= 0
                    && comments.Contains(struct (rowIndex, colIndex))
                let hasListDropdown =
                    rowIndex >= 0
                    && colIndex >= 0
                    && rangeBoundsArrayContains rowIndex colIndex dropdownRanges
                rows.Add rowIndex
                cols.Add colIndex
                values.Add(cellValueObj workbook cell)
                formulas.Add(box (cellFormula cell))
                styleIndexes.Add(effectiveStyleIndex workbook cell)
                hasComments.Add hasComment
                hasListDropdowns.Add hasListDropdown
                richRuns.Add(richRunsObjForCell workbook cell)
                errors.Add(cellErrorObj cell)
        objFromFields [
            objField "rows" (box (rows.ToArray()))
            objField "cols" (box (cols.ToArray()))
            objField "values" (box (values.ToArray()))
            objField "formulas" (box (formulas.ToArray()))
            objField "styleIndexes" (box (styleIndexes.ToArray()))
            objField "hasComments" (box (hasComments.ToArray()))
            objField "hasListDropdowns" (box (hasListDropdowns.ToArray()))
            objField "richRuns" (box (richRuns.ToArray()))
            objField "errors" (box (errors.ToArray()))
        ]

    let styleAt (index: int) (workbook: XlsxReader.Workbook) =
        if index >= 0 && index < workbook.Styles.Length then
            styleObj (Some workbook.Styles.[index])
        else
            null

    /// Get per-run rich text for a cell. Returns an empty JS array for plain cells.
    let richRunsAt (ref: string) (workbook: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) =
        let cellRef = XlsxReader.Xlsx.parseCellRef ref
        sheet
        |> findCellByRef cellRef
        |> Option.map (richRunsForCell workbook)
        |> Option.defaultValue [||]

    /// Get a cell comment/note at reference.
    let commentAt (ref: string) (sheet: XlsxReader.Worksheet) : obj =
        match refToCoord ref with
        | Some (row, col) ->
            match commentAtCoord row col sheet with
            | Some c ->
                createObj [
                    "text" ==> c.Text
                    "author" ==> optString c.Author
                ]
            | None -> null
        | None -> null

    /// Effective column width in Excel character units.
    let private defaultColumnWidth (sheet: XlsxReader.Worksheet) =
        match sheet.Format.DefaultColWidth with
        | Some width -> width
        | None ->
            match sheet.Format.BaseColWidth with
            | Some width -> width + 1.0
            | None -> 8.43

    let columnWidth (col: int) (sheet: XlsxReader.Worksheet) =
        let col1 = col + 1
        match sheet.Columns |> List.tryFind (fun c -> col1 >= c.Index && col1 <= c.MaxIndex) with
        | Some c when c.Hidden -> 0.0
        | Some c -> c.Width |> Option.defaultValue 0.0
        | None -> defaultColumnWidth sheet

    /// Effective column widths in Excel character units for zero-based columns [0, count).
    let columnWidths (count: int) (sheet: XlsxReader.Worksheet) =
        let count = max 0 count
        let widths = Array.create count (defaultColumnWidth sheet)
        for column in sheet.Columns do
            let first = max 0 (column.Index - 1)
            let last = min (count - 1) (column.MaxIndex - 1)
            if first <= last then
                let width = if column.Hidden then 0.0 else column.Width |> Option.defaultValue 0.0
                for col = first to last do
                    widths.[col] <- width
        widths

    /// Effective row height in points.
    let private defaultRowHeight (sheet: XlsxReader.Worksheet) =
        sheet.Format.DefaultRowHeight |> Option.defaultValue 15.0

    let private effectiveRowHeight (sheet: XlsxReader.Worksheet) (row: XlsxReader.Row) =
        row.Height |> Option.orElse sheet.Format.DefaultRowHeight |> Option.defaultValue 15.0

    let rowHeight (row: int) (sheet: XlsxReader.Worksheet) =
        let row1 = row + 1
        match sheet.Rows |> List.tryFind (fun r -> r.Index = row1) with
        | Some r -> effectiveRowHeight sheet r
        | None -> defaultRowHeight sheet

    /// Effective row heights in points for zero-based rows [0, count).
    let rowHeights (count: int) (sheet: XlsxReader.Worksheet) =
        let count = max 0 count
        let heights = Array.create count (defaultRowHeight sheet)
        for row in sheet.Rows do
            let index = row.Index - 1
            if index >= 0 && index < count then
                heights.[index] <- effectiveRowHeight sheet row
        heights

    /// Get merged cells from worksheet
    let mergedCells (sheet: XlsxReader.Worksheet) =
        sheet.MergedCells
        |> List.map (fun mc -> sprintf "%s:%s" (cellRefText mc.StartRef) (cellRefText mc.EndRef))
        |> List.toArray

    /// Get column information
    let columns (sheet: XlsxReader.Worksheet) = sheet.Columns |> List.toArray

    /// Get workbook styles
    let styles (wb: XlsxReader.Workbook) = wb.Styles

    /// Resolve shared strings and cell values
    let resolveCell (wb: XlsxReader.Workbook) (cell: XlsxReader.Cell) =
        let value = XlsxReader.resolveValue wb cell.Value
        { cell with Value = value }

    /// Get cell comments from a worksheet
    let comments (sheet: XlsxReader.Worksheet) = sheet.Comments

    /// Get conditional formatting rules from a worksheet
    let conditionalFormats (sheet: XlsxReader.Worksheet) = sheet.ConditionalFormats

    /// Get data validations from a worksheet
    let dataValidations (sheet: XlsxReader.Worksheet) = sheet.DataValidations

    /// Get charts from a worksheet
    let charts (sheet: XlsxReader.Worksheet) = sheet.Charts

    let private anchorObj (anchor: XlsxReader.SheetAnchor) : obj =
        createObj [
            "kind" ==> anchor.Kind
            "fromCol" ==> anchor.FromCol
            "fromColOffset" ==> anchor.FromColOffset
            "fromRow" ==> anchor.FromRow
            "fromRowOffset" ==> anchor.FromRowOffset
            "toCol" ==> anchor.ToCol
            "toColOffset" ==> anchor.ToColOffset
            "toRow" ==> anchor.ToRow
            "toRowOffset" ==> anchor.ToRowOffset
            "x" ==> anchor.X
            "y" ==> anchor.Y
            // EMU, as in the file (9525 per px).
            "width" ==> anchor.Width
            "height" ==> anchor.Height
            "editAs" ==> anchor.EditAs
        ]

    /// Pictures anchored to cells: { anchor, name, description, contentType, data (Uint8Array) }.
    let images (sheet: XlsxReader.Worksheet) : obj array =
        sheet.Extras.Images
        |> List.map (fun image ->
            createObj [
                "anchor" ==> anchorObj image.Anchor
                "name" ==> image.Name
                "description" ==> image.Description
                "contentType" ==> image.ContentType
                "data" ==> image.Data
            ])
        |> List.toArray

    /// Drawing objects that are kept, not modelled (charts, shapes): { anchor, kind, name }.
    let drawingObjects (sheet: XlsxReader.Worksheet) : obj array =
        sheet.Extras.Objects
        |> List.map (fun item -> createObj [ "anchor" ==> anchorObj item.Anchor; "kind" ==> item.Kind; "name" ==> item.Name ])
        |> List.toArray

    /// Threaded comments: { ref, author, date, text, replies: [{ author, date, text }] }.
    let threadedComments (sheet: XlsxReader.Worksheet) : obj array =
        sheet.Extras.ThreadedComments
        |> List.map (fun thread ->
            createObj [
                "ref" ==> sprintf "%s%d" thread.CellRef.Column thread.CellRef.Row
                "author" ==> thread.Author
                "date" ==> thread.Date
                "text" ==> thread.Text
                "replies" ==> (thread.Replies
                               |> List.map (fun reply -> createObj [ "author" ==> reply.Author; "date" ==> reply.Date; "text" ==> reply.Text ])
                               |> List.toArray)
            ])
        |> List.toArray

    /// Write a workbook as read (pictures, kept chart parts, threaded comments,
    /// named styles and theme included) back to an XLSX package.
    let writeWorkbook (workbook: XlsxReader.Workbook) =
        XlsxWriter.packageFromReaderWorkbook workbook |> XlsxWriter.writeWorkbookPackage

    /// Get pivot tables from a worksheet
    let pivotTables (sheet: XlsxReader.Worksheet) = sheet.PivotTables

    /// Get named ranges from workbook
    let namedRanges (wb: XlsxReader.Workbook) = wb.NamedRanges

    let private optObj (value: string option) : obj =
        match value with
        | Some v -> box v
        | None -> null

    let private optIntObj (value: int option) : obj =
        match value with
        | Some v -> box v
        | None -> null

    /// Header / band fills for a built-in table style, from the workbook
    /// theme: Excel's `TableStyle{Light,Medium,Dark}N` cycle through black
    /// and the six accents (N mod 7).
    let private tableFills (workbook: XlsxReader.Workbook) (styleName: string) : string * string =
        let digits = styleName |> Seq.filter System.Char.IsDigit |> Seq.toArray |> System.String
        let n = match System.Int32.TryParse(digits, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture) with | true, v -> v | _ -> 2
        let slot = (max 1 n - 1) % 7
        let accent =
            if slot = 0 || slot - 1 >= workbook.ThemeAccents.Length then "000000"
            else workbook.ThemeAccents.[slot - 1]
        "#" + accent, "#" + XlsxReader.Xlsx.applyColorTint accent 0.8

    /// Workbook-level facts the cell streams do not carry: the date system,
    /// every sheet with its visibility, and the defined names. `localSheetId`
    /// is the index of the owning sheet in the workbook (null = global).
    let workbookInfo (wb: XlsxReader.Workbook) : obj =
        objFromFields [
            objField "date1904" (box wb.Date1904)
            objField "sheets"
                (box [| for s in wb.Sheets ->
                          objFromFields [ objField "name" (box s.Name); objField "state" (box (sheetStateText s.State)) ] |])
            objField "definedNames"
                (box [| for n in wb.NamedRanges ->
                          objFromFields [
                              objField "name" (box n.Name)
                              objField "refersTo" (box n.RefersTo)
                              objField "localSheetId"
                                  (n.Scope
                                   |> Option.bind (fun s ->
                                       match System.Int32.TryParse(s, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture) with
                                       | true, v -> Some v
                                       | _ -> None)
                                   |> optIntObj)
                          ] |])
            objField "themeAccents" (box wb.ThemeAccents)
        ]

    /// Everything about a sheet besides its cells, as plain objects/arrays an
    /// importer can read without knowing the reader's record types: tab
    /// state, freeze panes, hidden rows/columns (0-based), tables,
    /// data validations, conditional formats (with the dxf each rule
    /// applies already resolved), hyperlinks and print setup.
    let sheetMeta (workbook: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) : obj =
        let tables =
            [| for t in sheet.Tables ->
                 objFromFields [
                     objField "name" (box t.Name)
                     objField "startRow" (box t.StartRow)
                     objField "startCol" (box t.StartCol)
                     objField "endRow" (box t.EndRow)
                     objField "endCol" (box t.EndCol)
                     objField "headerRowCount" (box t.HeaderRowCount)
                     objField "totalsRowCount" (box t.TotalsRowCount)
                     objField "columns" (box (List.toArray t.Columns))
                     objField "autoFilter"
                         (match t.AutoFilter with
                          | Some (c1, c2) -> box [| c1; c2 |]
                          | None -> null)
                     objField "style"
                         (match t.Style with
                          | Some s ->
                              let headerFill, stripeFill = tableFills workbook s.Name
                              objFromFields [
                                  objField "name" (box s.Name)
                                  objField "showFirstColumn" (box s.ShowFirstColumn)
                                  objField "showLastColumn" (box s.ShowLastColumn)
                                  objField "showRowStripes" (box s.ShowRowStripes)
                                  objField "showColumnStripes" (box s.ShowColumnStripes)
                                  objField "headerFill" (box headerFill)
                                  objField "stripeFill" (box stripeFill)
                              ]
                          | None -> null)
                 ] |]
        let validations =
            [| for dv in sheet.DataValidations ->
                 objFromFields [
                     objField "range" (box dv.Range)
                     objField "type" (box dv.ValidationType)
                     objField "operator" (optObj dv.Operator)
                     objField "formula1" (optObj dv.Formula1)
                     objField "formula2" (optObj dv.Formula2)
                     objField "allowBlank" (box dv.AllowBlank)
                     objField "showDropDown" (box dv.ShowDropDown)
                     objField "error" (optObj dv.ErrorMessage)
                     objField "errorTitle" (optObj dv.ErrorTitle)
                     objField "errorStyle" (optObj dv.ErrorStyle)
                     objField "showErrorMessage" (box dv.ShowErrorMessage)
                     objField "prompt" (optObj dv.Prompt)
                     objField "promptTitle" (optObj dv.PromptTitle)
                     objField "showInputMessage" (box dv.ShowInputMessage)
                 ] |]
        let conditionalFormats =
            [| for cf in sheet.ConditionalFormats ->
                 objFromFields [
                     objField "range" (box cf.Range)
                     objField "rules"
                         (box [| for rule in cf.Rules ->
                                   let dxf =
                                       rule.DxfId
                                       |> Option.bind (fun id -> if id >= 0 && id < workbook.Dxfs.Length then Some workbook.Dxfs.[id] else None)
                                   objFromFields [
                                       objField "type" (box rule.Type)
                                       objField "operator" (optObj rule.Operator)
                                       objField "formulas" (box (List.toArray rule.Formulas))
                                       objField "text" (optObj rule.Text)
                                       objField "priority" (optIntObj rule.Priority)
                                       objField "stopIfTrue" (box rule.StopIfTrue)
                                       objField "rank" (optIntObj rule.Rank)
                                       objField "percent" (box rule.Percent)
                                       objField "bottom" (box rule.Bottom)
                                       objField "aboveAverage" (box rule.AboveAverage)
                                       objField "equalAverage" (box rule.EqualAverage)
                                       objField "stdDev" (optIntObj rule.StdDev)
                                       objField "colorScale"
                                           (box [| for stop in rule.ColorScale ->
                                                     objFromFields [
                                                         objField "type" (box stop.CfvoType)
                                                         objField "value" (optObj stop.CfvoValue)
                                                         objField "color" (box stop.Color)
                                                     ] |])
                                       objField "dataBarColor" (optObj rule.DataBarColor)
                                       objField "dataBarCfvo"
                                           (box [| for (kind, value) in rule.DataBarCfvo ->
                                                     objFromFields [ objField "type" (box kind); objField "value" (optObj value) ] |])
                                       objField "iconSet" (optObj rule.IconSet)
                                       objField "iconSetCfvo"
                                           (box [| for (kind, value) in rule.IconSetCfvo ->
                                                     objFromFields [ objField "type" (box kind); objField "value" (optObj value) ] |])
                                       objField "iconSetReverse" (box rule.IconSetReverse)
                                       objField "style"
                                           (match dxf with
                                            | Some d ->
                                                objFromFields [
                                                    objField "fontColor" (cssColor d.FontColor)
                                                    objField "fillColor" (cssColor d.FillColor)
                                                    objField "bold" (box (d.Bold = Some true))
                                                    objField "italic" (box (d.Italic = Some true))
                                                ]
                                            | None -> null)
                                   ] |])
                 ] |]
        let hyperlinks =
            [| for h in sheet.Hyperlinks ->
                 objFromFields [
                     objField "range" (box h.Range)
                     objField "target" (box h.Target)
                     objField "isInternal" (box h.IsInternal)
                 ] |]
        let hiddenRows = [| for r in sheet.Rows do if r.Hidden then yield r.Index - 1 |]
        let hiddenCols =
            [| for c in sheet.Columns do
                 // A hidden range can run to the last column (16384); the
                 // importer only needs the ones near the data.
                 if c.Hidden then
                     for idx in c.Index .. min c.MaxIndex (c.Index + 1023) do
                         yield idx - 1 |]
        let pageSetup =
            match sheet.PageSetup with
            | None -> null
            | Some p ->
                objFromFields [
                    objField "paperSize" (optIntObj p.PaperSize)
                    objField "orientation"
                        (match p.Orientation with
                         | XlsxReader.OrientationLandscape -> box "landscape"
                         | XlsxReader.OrientationPortrait -> box "portrait"
                         | XlsxReader.OrientationDefault -> null)
                    objField "fitToPage" (box p.FitToPage)
                    objField "fitToWidth" (box p.FitToWidth)
                    objField "fitToHeight" (box p.FitToHeight)
                    objField "scale" (box p.Scale)
                    objField "firstPageNumber" (box p.FirstPageNumber)
                    objField "useFirstPageNumber" (box p.UseFirstPageNumber)
                    objField "pageOrder" (box (match p.PageOrder with XlsxReader.OverThenDown -> "overThenDown" | XlsxReader.DownThenOver -> "downThenOver"))
                    objField "blackAndWhite" (box p.BlackAndWhite)
                    objField "cellComments"
                        (box (match p.CellComments with
                              | XlsxReader.CellCommentsAsDisplayed -> "asDisplayed"
                              | XlsxReader.CellCommentsAtEnd -> "atEnd"
                              | XlsxReader.CellCommentsNone -> "none"))
                ]
        let pageMargins =
            match sheet.PageMargins with
            | None -> null
            | Some m ->
                objFromFields [
                    objField "top" (box m.Top)
                    objField "bottom" (box m.Bottom)
                    objField "left" (box m.Left)
                    objField "right" (box m.Right)
                    objField "header" (box m.Header)
                    objField "footer" (box m.Footer)
                ]
        let hf = sheet.HeaderFooter
        objFromFields [
            objField "name" (box sheet.Name)
            objField "state" (box (sheetStateText sheet.State))
            objField "frozenRows" (box (sheet.FrozenPane |> Option.map (fun p -> p.YSplit) |> Option.defaultValue 0))
            objField "frozenCols" (box (sheet.FrozenPane |> Option.map (fun p -> p.XSplit) |> Option.defaultValue 0))
            objField "showGridLines" (box sheet.ShowGridLines)
            objField "showZeros" (box sheet.ShowZeros)
            objField "hiddenRows" (box hiddenRows)
            objField "hiddenCols" (box hiddenCols)
            objField "tables" (box tables)
            objField "validations" (box validations)
            objField "conditionalFormats" (box conditionalFormats)
            objField "hyperlinks" (box hyperlinks)
            objField "pageSetup" pageSetup
            objField "pageMargins" pageMargins
            objField "printOptions"
                (objFromFields [
                    objField "gridLines" (box sheet.PrintOptions.GridLines)
                    objField "headings" (box sheet.PrintOptions.Headings)
                    objField "horizontalCentered" (box sheet.PrintOptions.HorizontalCentered)
                    objField "verticalCentered" (box sheet.PrintOptions.VerticalCentered)
                ])
            objField "headerFooter"
                (objFromFields [
                    objField "oddHeader" (optObj hf.OddHeader)
                    objField "oddFooter" (optObj hf.OddFooter)
                    objField "evenHeader" (optObj hf.EvenHeader)
                    objField "evenFooter" (optObj hf.EvenFooter)
                    objField "firstHeader" (optObj hf.FirstHeader)
                    objField "firstFooter" (optObj hf.FirstFooter)
                    objField "differentOddEven" (box hf.DifferentOddEven)
                    objField "differentFirst" (box hf.DifferentFirst)
                    objField "scaleWithDoc" (box hf.ScaleWithDoc)
                ])
            objField "rowBreaks" (box (List.toArray sheet.RowBreaks))
            objField "colBreaks" (box (List.toArray sheet.ColBreaks))
            objField "printArea" (optObj sheet.PrintArea)
            objField "printTitleRows" (optObj sheet.PrintTitleRows)
            objField "printTitleCols" (optObj sheet.PrintTitleCols)
        ]

    /// Evaluate a cell's formula (stub)
    let evaluateFormula (wb: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) (cell: XlsxReader.Cell) =
        XlsxReader.FormulaEngine.evaluateCell wb sheet cell

#if FABLE_COMPILER
// Example usage demonstrating the DSL
[<ExportDefault>]
let api () = {|
    word = {|
        load = Word.load
        extractText = Word.extractText
        getContent = Word.getContent
        paragraphs = Word.paragraphs
        tables = Word.tables
        styles = Word.styles
        headers = Word.headers
        footers = Word.footers
        properties = Word.properties
        hyperlinks = Word.hyperlinks
        images = Word.images
        comments = Word.comments
        revisions = Word.revisions
        footnotes = Word.footnotes
        endnotes = Word.endnotes
        shapes = Word.shapes
        write = WordWriter.write
        writeFromEditorPackage = WordWriter.writeFromEditorPackage
        getLastSaveReport = WordWriter.getLastSaveReport
        openProgressive = Processors.openProgressive
        isLoadComplete = Processors.isLoadComplete
        whenLoadComplete = Processors.whenLoadComplete
        completeLoad = Processors.completeLoad
        isLoadFailed = Processors.isLoadFailed
        getLoadError = Processors.getLoadError
        createDocument = WordWriter.createDocument
        addParagraph = WordWriter.addParagraph
        addParagraphWithRuns = WordWriter.addParagraphWithRuns
        saveDocument = WordWriter.saveDocument
        createRun = WordWriter.createRun
        createSimpleTable = WordWriter.createSimpleTable
        createTableWithBorders = WordWriter.createTableWithBorders
        generateTableXml = WordWriter.generateTableXml
    |}
    processors = {|
        forEditorPackage = Processors.forEditorPackage
        openProgressive = Processors.openProgressive
    |}
    excel = {|
        load = Excel.load
        loadSelected = Excel.loadSelected
        sheetNames = Excel.sheetNames
        sheetStates = Excel.sheetStates
        visibleSheetNames = Excel.visibleSheetNames
        workbookInfo = Excel.workbookInfo
        sheetMeta = Excel.sheetMeta
        sheet = Excel.sheet
        toArray = Excel.toArray
        dimensions = Excel.dimensions
        cellAt = Excel.cellAt
        cellWithStyleAt = Excel.cellWithStyleAt
        cellsWithStyle = Excel.cellsWithStyle
        cellsCompact = Excel.cellsCompact
        cellsColumnar = Excel.cellsColumnar
        styleAt = Excel.styleAt
        richRunsAt = Excel.richRunsAt
        commentAt = Excel.commentAt
        columnWidth = Excel.columnWidth
        columnWidths = Excel.columnWidths
        rowHeight = Excel.rowHeight
        rowHeights = Excel.rowHeights
        mergedCells = Excel.mergedCells
        columns = Excel.columns
        styles = Excel.styles
        resolveCell = Excel.resolveCell
        comments = Excel.comments
        conditionalFormats = Excel.conditionalFormats
        dataValidations = Excel.dataValidations
        charts = Excel.charts
        images = Excel.images
        drawingObjects = Excel.drawingObjects
        threadedComments = Excel.threadedComments
        writeWorkbook = Excel.writeWorkbook
        pivotTables = Excel.pivotTables
        namedRanges = Excel.namedRanges
        evaluateFormula = Excel.evaluateFormula
        write = ExcelWriter.write
        writeFormatted = ExcelWriter.writeFormatted
        createWorkbook = ExcelWriter.createWorkbook
        addSheet = ExcelWriter.addSheet
        addSimpleSheet = ExcelWriter.addSimpleSheet
        saveWorkbook = ExcelWriter.saveWorkbook
        addRow = ExcelWriter.addRow
        addSimpleRow = ExcelWriter.addSimpleRow
        createCell = ExcelWriter.createCell
        createHeaderCell = ExcelWriter.createHeaderCell
        createNumericCell = ExcelWriter.createNumericCell
        createFormulaCell = ExcelWriter.createFormulaCell
        addMergedRange = ExcelWriter.addMergedRange
        createMergedRange = ExcelWriter.createMergedRange
    |}
|}
#endif

// Browser-friendly usage example
[<Emit("console.log($0)")>]
let log (msg: obj) : unit = jsNative

// Example: Read DOCX from file input
let readWordDoc (file: obj) : Fable.Core.JS.Promise<unit> =
    promise {
        let! doc = Word.load file
        match doc with
        | Some d ->
            let text = Word.extractText d
            log {| docText = text; paragraphCount = (Word.paragraphs d).Length |}
        | None ->
            log "Failed to load document"
    }

// Example: Read XLSX from file input
let readExcelDoc (file: obj) : Fable.Core.JS.Promise<unit> =
    promise {
        let! wb = Excel.load file
        match wb with
        | Some w ->
            let sheets = Excel.sheetNames w
            log {| workbook = w; sheetNames = sheets |}

            // Get first sheet as array
            match w.Sheets with
            | sheet :: _ ->
                let arr = Excel.toArray w sheet
                log {| sheetData = arr |}
            | [] -> ()
        | None ->
            log "Failed to load workbook"
    }
