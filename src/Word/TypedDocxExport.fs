module Rbt.Word.TypedDocxExport

open Rbt.Word.Document
open Rbt.Word.TextBuffer

/// Undo the layout-time distribution of vertically merged cell content.
let private gatherVerticalMergeContent (rows: TableRow list) : TableRow list =
    let cells = rows |> List.map (fun row -> List.toArray row.Cells) |> List.toArray
    let starts =
        rows |> List.map (fun row ->
            let mutable column = row.GridBefore
            row.Cells |> List.map (fun cell ->
                let start = column
                column <- column + max 1 cell.ColSpan
                start) |> List.toArray) |> List.toArray
    if not (cells |> Array.exists (Array.exists (fun cell -> cell.VerticalMerge = Some ContinueMerge))) then rows
    else
        for rowIndex in 0 .. cells.Length - 1 do
            for cellIndex in 0 .. cells.[rowIndex].Length - 1 do
                let cell = cells.[rowIndex].[cellIndex]
                if cell.VerticalMerge = Some RestartMerge then
                    let column = starts.[rowIndex].[cellIndex]
                    let gathered = ResizeArray<ParagraphBlock>()
                    let mutable nextRow = rowIndex + 1
                    let mutable more = true
                    while more && nextRow < cells.Length do
                        match starts.[nextRow] |> Array.tryFindIndex ((=) column) with
                        | Some index when cells.[nextRow].[index].VerticalMerge = Some ContinueMerge ->
                            let continuation = cells.[nextRow].[index]
                            gathered.AddRange continuation.Content
                            let placeholder =
                                match continuation.Content with
                                | first :: _ -> [ { first with Buffer = TextBuffer.empty } ]
                                | [] -> []
                            cells.[nextRow].[index] <- { continuation with Content = placeholder }
                            nextRow <- nextRow + 1
                        | _ -> more <- false
                    while gathered.Count > 0 && TextBuffer.isWhiteSpace gathered.[gathered.Count - 1].Buffer do
                        Portable.Stack.dropLast gathered
                    if gathered.Count > 0 then
                        cells.[rowIndex].[cellIndex] <- { cell with Content = cell.Content @ List.ofSeq gathered }
        rows |> List.mapi (fun index row -> { row with Cells = List.ofArray cells.[index] })

let private blocksInAnchorOrder (blocks: Block list) =
    let idOf = function
        | Paragraph paragraph | Heading (_, paragraph) -> Some paragraph.Id
        | Table table -> Some table.Id
        | TextBox _ -> None
    let ids = blocks |> List.choose idOf |> Set.ofList
    let isAnchored (shape: TextBoxBlock) = shape.AnchorBlockId |> Option.exists (fun id -> Set.contains id ids)
    let shapes = blocks |> List.choose (function TextBox shape when isAnchored shape -> Some shape | _ -> None)
    if shapes.IsEmpty then blocks
    else
        blocks |> List.collect (fun block ->
            match block with
            | TextBox shape when isAnchored shape -> []
            | _ ->
                match idOf block with
                | Some id -> block :: (shapes |> List.filter (fun shape -> shape.AnchorBlockId = Some id) |> List.map TextBox)
                | None -> [ block ])

let private storyParagraphs (block: Block) : ParagraphBlock list option =
    let rec ofTable (table: TableBlock) =
        gatherVerticalMergeContent table.Rows
        |> List.collect (fun row ->
            row.Cells |> List.collect (fun cell ->
                if cell.NestedTables.IsEmpty then cell.Content
                else
                    let count = cell.Content.Length
                    let anchors = List.toArray cell.NestedTableAnchors
                    let nested = cell.NestedTables |> List.mapi (fun i inner -> inner, (if i < anchors.Length then anchors.[i] else count) |> max 0 |> min count) |> List.sortBy snd
                    let before index = nested |> List.filter (fun (_, anchor) -> anchor = index) |> List.collect (fun (inner, _) -> ofTable inner)
                    (cell.Content |> List.mapi (fun index para -> before index @ [ para ]) |> List.concat) @ before count))
    match block with
    | Paragraph para | Heading (_, para) -> Some [ para ]
    | Table table -> Some (ofTable table)
    | TextBox _ -> None

let private runsByCharacter (buffer: TextBuffer.TextBuffer) =
    let result = ResizeArray<TextBuffer.Run>()
    for run in buffer.Runs do
        match run with
        | TextBuffer.Text text -> for _ in 1 .. text.Length do result.Add run
        | TextBuffer.Image _ -> result.Add run
    result.ToArray()

let private withFormatOf (format: TextBuffer.TextStyle) (style: TextBuffer.TextStyle) =
    { style with
        Style = format.Style; Bold = format.Bold; Italic = format.Italic; Underline = format.Underline
        UnderlineStyle = format.UnderlineStyle; UnderlineColor = format.UnderlineColor; AllCaps = format.AllCaps
        SmallCaps = format.SmallCaps; Strike = format.Strike; DoubleStrike = format.DoubleStrike; Hidden = format.Hidden
        Position = format.Position; EmphasisMark = format.EmphasisMark; FontSize = format.FontSize
        CharacterSpacing = format.CharacterSpacing; FontFamily = format.FontFamily; Color = format.Color
        Highlight = format.Highlight; VertAlign = format.VertAlign; Explicit = format.Explicit
        RenderedPageBreakBefore = format.RenderedPageBreakBefore }

let private bufferOfCharacters (text: string) (runs: TextBuffer.Run array) =
    let result = ResizeArray<TextBuffer.Run>()
    let mutable i = 0
    while i < runs.Length do
        match runs.[i] with
        | TextBuffer.Image image ->
            result.Add (TextBuffer.Image { image with Start = i })
            i <- i + 1
        | TextBuffer.Text run ->
            let mutable j = i + 1
            while j < runs.Length && (match runs.[j] with
                                      | TextBuffer.Text other -> run.Style.Checkbox.IsNone && (obj.ReferenceEquals(other, run) || obj.ReferenceEquals(other.Style, run.Style) || other.Style = run.Style)
                                      | TextBuffer.Image _ -> false) do
                j <- j + 1
            result.Add (TextBuffer.Text { Start = i; Length = j - i; Style = run.Style })
            i <- j
    { Rope = Rope.ofString text; Runs = List.ofSeq result }

type private Derivation = {
    Block: int
    Para: int
    Original: ParagraphBlock
    Level: int
    From: int
    Until: int
    Joined: (int * int * ParagraphBlock * int) list
}

let private originalStory (originalBlocks: Block list) =
    let story = ResizeArray<int * int * ParagraphBlock * int>()
    let counts = ResizeArray<int>()
    originalBlocks |> List.iteri (fun index block ->
        let level = match block with Heading (level, _) -> level | _ -> 0
        match storyParagraphs block with
        | Some paragraphs ->
            paragraphs |> List.iteri (fun at para -> story.Add (index, at, para, level))
            counts.Add paragraphs.Length
        | None -> counts.Add 0)
    story.ToArray(), counts.ToArray()

let private derivationsOf (story: (int * int * ParagraphBlock * int) array) (blocks: Block list) =
    let mutable result = Map.empty<string, Derivation>
    let grouped = story |> Array.mapi (fun at (_, _, para, _) -> para.Id, at) |> Array.groupBy fst
    let byId = grouped |> Array.choose (fun (id, entries) -> if entries.Length = 1 then Some (id, snd entries.[0]) else None) |> Map.ofArray
    let present = paragraphsOfBlocks blocks |> List.map (fun para -> para.Id) |> Set.ofList
    let textOf (para: ParagraphBlock) = TextBuffer.toString para.Buffer
    let consider (siblings: ParagraphBlock array) =
        let mutable i = 0
        while i < siblings.Length do
            let para = siblings.[i]
            let mutable next = i + 1
            match Map.tryFind para.Id byId with
            | Some at when not (Map.containsKey para.Id result) ->
                let block, index, original, level = story.[at]
                let whole = { Block = block; Para = index; Original = original; Level = level; From = 0; Until = TextBuffer.length original.Buffer; Joined = [] }
                let was = if obj.ReferenceEquals(para, original) then "" else textOf original
                let now = if obj.ReferenceEquals(para, original) then "" else textOf para
                if now = was then result <- Map.add para.Id whole result
                elif now.Length < was.Length && was.StartsWith(now, System.StringComparison.Ordinal) then
                    let parts = ResizeArray<ParagraphBlock * int * int>()
                    parts.Add (para, 0, now.Length)
                    let mutable covered = now.Length
                    let mutable more = true
                    while more && covered < was.Length && next < siblings.Length do
                        let candidate = siblings.[next]
                        let text = textOf candidate
                        if not (Map.containsKey candidate.Id byId) && not (Map.containsKey candidate.Id result)
                           && covered + text.Length <= was.Length && was.Substring(covered, text.Length) = text then
                            parts.Add (candidate, covered, covered + text.Length)
                            covered <- covered + text.Length
                            next <- next + 1
                        else more <- false
                    if covered = was.Length && parts.Count > 1 then
                        for part, from, until in parts do result <- Map.add part.Id { whole with From = from; Until = until } result
                    else
                        next <- i + 1
                        result <- Map.add para.Id whole result
                elif now.Length > was.Length && now.StartsWith(was, System.StringComparison.Ordinal) then
                    let joined = ResizeArray<int * int * ParagraphBlock * int>()
                    let mutable covered = was.Length
                    let mutable follower = at + 1
                    let mutable more = true
                    while more && covered < now.Length && follower < story.Length do
                        let otherBlock, otherIndex, other, otherLevel = story.[follower]
                        let text = textOf other
                        let lastBlock, _, _, _ = story.[follower - 1]
                        let adjacent = otherBlock = lastBlock || (otherBlock = lastBlock + 1 && otherIndex = 0)
                        if adjacent && other.Metadata.Style = original.Metadata.Style && otherLevel = level
                           && other.Metadata.SectionBreak.IsNone && not other.Metadata.PageBreakAfter && not original.Metadata.PageBreakAfter
                           && not (Set.contains other.Id present) && covered + text.Length <= now.Length && now.Substring(covered, text.Length) = text then
                            joined.Add (otherBlock, otherIndex, other, otherLevel)
                            covered <- covered + text.Length
                            follower <- follower + 1
                        else more <- false
                    result <- Map.add para.Id (if covered = now.Length && joined.Count > 0 then { whole with Joined = List.ofSeq joined } else whole) result
                else result <- Map.add para.Id whole result
            | _ -> ()
            i <- next
    let rec ofTable (table: TableBlock) =
        for row in gatherVerticalMergeContent table.Rows do
            for cell in row.Cells do
                consider (List.toArray cell.Content)
                for inner in cell.NestedTables do ofTable inner
    consider (blocks |> List.choose (function Paragraph para | Heading (_, para) -> Some para | _ -> None) |> List.toArray)
    for block in blocks do match block with Table table -> ofTable table | _ -> ()
    result

let private derivedOriginal (derivation: Derivation) =
    let original = derivation.Original
    let length = TextBuffer.length original.Buffer
    if derivation.From = 0 && derivation.Until = length && derivation.Joined.IsEmpty then original
    else
        let runs = ResizeArray<TextBuffer.Run>()
        let text = System.Text.StringBuilder()
        let own = runsByCharacter original.Buffer
        let ownText = TextBuffer.toString original.Buffer
        if own.Length = ownText.Length then
            for at in derivation.From .. derivation.Until - 1 do runs.Add own.[at]
            text.Append(ownText.Substring(derivation.From, derivation.Until - derivation.From)) |> ignore
        for _, _, other, _ in derivation.Joined do
            runs.AddRange (runsByCharacter other.Buffer)
            text.Append(TextBuffer.toString other.Buffer) |> ignore
        { original with Buffer = bufferOfCharacters (text.ToString()) (runs.ToArray()) }

let private paragraphTextChange index (before: ParagraphBlock) (after: ParagraphBlock) : WordEditorPackage.ParagraphTextEdit option =
    if obj.ReferenceEquals(before, after) || before.Metadata <> after.Metadata then None
    else
        let oldText = TextBuffer.toString before.Buffer
        let newText = TextBuffer.toString after.Buffer
        let oldRuns = runsByCharacter before.Buffer
        let newRuns = runsByCharacter after.Buffer
        if oldRuns.Length <> oldText.Length || newRuns.Length <> newText.Length then None
        else
            let mutable lastA = Unchecked.defaultof<TextBuffer.Run>
            let mutable lastB = Unchecked.defaultof<TextBuffer.Run>
            let mutable lastSame = false
            let sameRun a b =
                if obj.ReferenceEquals(a, b) then true
                elif obj.ReferenceEquals(a, lastA) && obj.ReferenceEquals(b, lastB) then lastSame
                else
                    let same =
                        match a, b with
                        | TextBuffer.Text x, TextBuffer.Text y ->
                            obj.ReferenceEquals(x.Style, y.Style) || x.Style = y.Style
                            || (x.Style.RenderedPageBreakBefore <> y.Style.RenderedPageBreakBefore && { x.Style with RenderedPageBreakBefore = y.Style.RenderedPageBreakBefore } = y.Style)
                        | TextBuffer.Image x, TextBuffer.Image y -> { x with Start = 0 } = { y with Start = 0 }
                        | _ -> false
                    lastA <- a
                    lastB <- b
                    lastSame <- same
                    same
            let same i j = oldText.[i] = newText.[j] && sameRun oldRuns.[i] newRuns.[j]
            let limit = min oldText.Length newText.Length
            let mutable prefix = 0
            while prefix < limit && same prefix prefix do prefix <- prefix + 1
            let mutable suffix = 0
            while suffix < limit - prefix && same (oldText.Length - 1 - suffix) (newText.Length - 1 - suffix) do suffix <- suffix + 1
            let deleted = oldText.Length - prefix - suffix
            let insertedLength = newText.Length - prefix - suffix
            let isText = function TextBuffer.Text _ -> true | TextBuffer.Image _ -> false
            let allText (runs: TextBuffer.Run array) start count =
                let mutable ok = true
                for i in start .. start + count - 1 do
                    if not (isText runs.[i]) then ok <- false
                ok
            let marks (buffer: TextBuffer.TextBuffer) =
                buffer.Runs |> List.choose (function TextBuffer.Text run when run.Style.RenderedPageBreakBefore && run.Length > 0 -> Some run.Start | _ -> None)
            let marksKept =
                let previous = marks before.Buffer
                let current = marks after.Buffer
                if previous.IsEmpty && current.IsEmpty then true
                else
                    let moved = previous |> List.filter (fun at -> at < prefix || at >= prefix + deleted) |> List.map (fun at -> if at < prefix then at else at - deleted + insertedLength)
                    let inside = previous |> List.exists (fun at -> at >= prefix && at < prefix + deleted)
                    not inside && (moved = current ||
                        (insertedLength > 0 && List.contains (prefix + deleted) previous &&
                         (moved |> List.map (fun at -> if at = prefix + insertedLength then prefix else at)) = current))
            if deleted = 0 && insertedLength = 0 || not marksKept || not (allText oldRuns prefix deleted) || not (allText newRuns prefix insertedLength) then None
            else
                let oneRun =
                    insertedLength > 0 &&
                        (let mutable one = true
                         for i in prefix + 1 .. prefix + insertedLength - 1 do
                             if not (sameRun newRuns.[prefix] newRuns.[i]) then one <- false
                         one)
                let joinsBefore = oneRun && prefix > 0 && sameRun oldRuns.[prefix - 1] newRuns.[prefix]
                let joinsAfter = oneRun && suffix > 0 && sameRun oldRuns.[prefix + deleted] newRuns.[prefix]
                let plain =
                    oneRun && (match newRuns.[prefix] with
                               | TextBuffer.Text run -> not (TextBuffer.hasAnnotationIdentity run.Style)
                               | TextBuffer.Image _ -> false)
                let witnesses =
                    if not plain then [||]
                    else
                        let formatted run =
                            match run, newRuns.[prefix] with
                            | TextBuffer.Text candidate, TextBuffer.Text typed ->
                                obj.ReferenceEquals(candidate.Style, typed.Style) ||
                                (candidate.Style.Hyperlink.IsNone && candidate.Style.InactiveHyperlink.IsNone && TextBuffer.withoutAnnotations candidate.Style = typed.Style)
                            | _ -> false
                        let ordinary c = c >= ' ' && c <> '\uFFFC' && c <> '\u00AD' && c <> '\u2011' && c <> '\u2028'
                        let found = ResizeArray<int>()
                        let mutable lastRun = Unchecked.defaultof<TextBuffer.Run>
                        let mutable lastFormatted = false
                        let mutable taken = false
                        let order = seq { yield! seq { prefix + deleted .. oldRuns.Length - 1 }; yield! seq { prefix - 1 .. -1 .. 0 } }
                        for i in order do
                            if found.Count < 6 then
                                if not (obj.ReferenceEquals(oldRuns.[i], lastRun)) then
                                    lastRun <- oldRuns.[i]
                                    lastFormatted <- formatted oldRuns.[i]
                                    taken <- false
                                if lastFormatted && not taken && ordinary oldText.[i] then
                                    found.Add i
                                    taken <- true
                        found.ToArray()
                Some { Paragraph = index; Offset = prefix; Deleted = deleted
                       Inserted = newText.Substring(prefix, insertedLength); JoinBefore = joinsBefore
                       JoinAfter = joinsAfter; RunsPlain = plain; Witnesses = witnesses
                       MarkStyle = plain && witnesses.Length = 0 && not (oldRuns |> Array.exists isText)
                       NewText = newText }

let private paragraphEdit
    (exportBefore: ParagraphBlock -> WordEditorPackage.EditorParagraph)
    (export: ParagraphBlock -> WordEditorPackage.EditorParagraph)
    index (a: ParagraphBlock) (b: ParagraphBlock) : WordEditorPackage.ParagraphEdit =
    let afterItem = { export b with SpliceOriginal = Some (exportBefore a) }
    let baseEdit : WordEditorPackage.ParagraphEdit =
        { Paragraph = index; Block = None; Para = None; Text = None; OldText = None
          Before = None; After = Some afterItem; NewText = None; Regenerate = false
          Verbatim = false; Why = None; Cut = None; Join = [||] }
    let regenerate why = { baseEdit with Regenerate = true; Why = Some why }
    let ma, mb = a.Metadata, b.Metadata
    if ma.ContentControlId <> mb.ContentControlId || ma.ContentControlTag <> mb.ContentControlTag
       || ma.ContentControlAlias <> mb.ContentControlAlias || ma.ContentControlXml <> mb.ContentControlXml
       || ma.ContentControlGroup <> mb.ContentControlGroup || ma.ContentControlAncestors <> mb.ContentControlAncestors
       || ma.PageBreakAfter <> mb.PageBreakAfter || ma.RenderedPageBreakBefore <> mb.RenderedPageBreakBefore then
        regenerate "its content control or the page break after it was changed"
    else
        let oldText = TextBuffer.toString a.Buffer
        let newText = TextBuffer.toString b.Buffer
        let oldRuns = runsByCharacter a.Buffer
        let newRuns = runsByCharacter b.Buffer
        if oldRuns.Length <> oldText.Length || newRuns.Length <> newText.Length then regenerate "its runs do not cover its text"
        else
            let mutable lastA = Unchecked.defaultof<TextBuffer.Run>
            let mutable lastB = Unchecked.defaultof<TextBuffer.Run>
            let mutable lastSame = false
            let sameKind x y =
                if obj.ReferenceEquals(x, y) then true
                elif obj.ReferenceEquals(x, lastA) && obj.ReferenceEquals(y, lastB) then lastSame
                else
                    let same =
                        match x, y with
                        | TextBuffer.Text p, TextBuffer.Text q -> obj.ReferenceEquals(p.Style, q.Style) || withFormatOf p.Style q.Style = p.Style
                        | TextBuffer.Image p, TextBuffer.Image q -> { p with Start = 0 } = { q with Start = 0 }
                        | _ -> false
                    lastA <- x
                    lastB <- y
                    lastSame <- same
                    same
            let limit = min oldText.Length newText.Length
            let mutable prefix = 0
            while prefix < limit && oldText.[prefix] = newText.[prefix] do prefix <- prefix + 1
            let mutable suffix = 0
            while suffix < limit - prefix && oldText.[oldText.Length - 1 - suffix] = newText.[newText.Length - 1 - suffix] do suffix <- suffix + 1
            let deleted = oldText.Length - prefix - suffix
            let inserted = newText.Length - prefix - suffix
            let middle =
                if deleted = 0 && inserted = 0 then a
                else
                    let neighbour typed =
                        let candidate at = if at >= 0 && at < oldRuns.Length && sameKind oldRuns.[at] typed then Some oldRuns.[at] else None
                        match typed with
                        | TextBuffer.Image _ -> typed
                        | TextBuffer.Text _ -> candidate (prefix - 1) |> Option.orElse (candidate (prefix + deleted)) |> Option.defaultValue typed
                    let runs = Array.init newText.Length (fun j ->
                        if j < prefix then oldRuns.[j]
                        elif j < prefix + inserted then neighbour newRuns.[j]
                        else oldRuns.[j - inserted + deleted])
                    { a with Buffer = bufferOfCharacters newText runs }
            let middleRuns = if obj.ReferenceEquals(middle, a) then oldRuns else runsByCharacter middle.Buffer
            let mutable kinds = middleRuns.Length = newRuns.Length
            if kinds then
                for j in 0 .. newRuns.Length - 1 do
                    if kinds && not (sameKind middleRuns.[j] newRuns.[j]) then kinds <- false
            if not kinds then regenerate "a link, a field, a note, a comment, a bookmark, a tracked change or a picture of it was changed"
            else
                let textEdit = if obj.ReferenceEquals(middle, a) then Some None else paragraphTextChange index a middle |> Option.map Some
                match textEdit with
                | None -> regenerate "its text was changed in a way that is not one stretch of plain text (a picture, the mark of a rendered page break)"
                | Some text ->
                    let middleItem = exportBefore middle
                    { baseEdit with Text = text; Before = (if middleItem = afterItem then None else Some middleItem); NewText = Some newText }

let private derivedEdit exportBefore export index (derivation: Derivation) (para: ParagraphBlock) : WordEditorPackage.ParagraphEdit =
    let before = derivedOriginal derivation
    let isPart = derivation.From > 0 || derivation.Until < TextBuffer.length derivation.Original.Buffer
    let basic () : WordEditorPackage.ParagraphEdit =
        { Paragraph = index; Block = None; Para = None; Text = None; OldText = None
          Before = None; After = Some (export para); NewText = None; Regenerate = false
          Verbatim = false; Why = None; Cut = None; Join = [||] }
    let edit =
        if isPart && derivation.Original.Metadata.PageBreakAfter then
            { basic () with Regenerate = true; Why = Some "a paragraph with a page break after it was split" }
        elif obj.ReferenceEquals(before, para) || exportBefore before = export para then
            { basic () with NewText = Some (TextBuffer.toString para.Buffer) }
        else paragraphEdit exportBefore export index before para
    { edit with
        After = edit.After |> Option.map (fun after -> { after with SpliceOriginal = Some (exportBefore derivation.Original) })
        Cut = if isPart then Some (derivation.From, derivation.Until) else None
        Join =
            derivation.Joined
            |> List.map (fun (block, at, other, _) ->
                let joined: WordEditorPackage.JoinedParagraph =
                    { Block = block; Para = at; Text = TextBuffer.toString other.Buffer; Original = exportBefore other }
                joined)
            |> List.toArray }

let private alignment = function Left -> "left" | Center -> "center" | Right -> "right" | Justify -> "justify"
let private listStyle = function NoList -> "none" | Bullet -> "bullet" | Numbered -> "numbered"
let private source origin : WordEditorPackage.SourceFragment =
    { Origin = origin |> Option.map (function ReadAt at -> WordEditorPackage.ReadAt at | LikeAt at -> WordEditorPackage.LikeAt at)
      Xml = None; Copy = (match origin with Some (LikeAt _) -> true | _ -> false); Same = false }
let private columns (value: SectionColumns) : WordEditorPackage.SectionColumns =
    { Count = value.Count; Gap = value.Gap; Separator = value.Separator; Widths = value.Widths }
let private references values : WordEditorPackage.HeaderFooterReference list =
    values |> List.map (fun (kind, id) -> { Kind = kind; RelationshipId = id })
let private pageSetup (value: PageSetup) : WordEditorPackage.PageSetup =
    { PageWidth = value.PageWidth; PageHeight = value.PageHeight; MarginLeft = value.MarginLeft; MarginRight = value.MarginRight
      MarginTop = value.MarginTop; MarginBottom = value.MarginBottom; HeaderDistance = value.HeaderDistance; FooterDistance = value.FooterDistance
      DifferentFirstPage = value.DifferentFirstPage; EvenAndOddHeaders = value.EvenAndOddHeaders
      DefaultTabStop = Some value.DefaultTabStop; CompatibilityMode = Some value.CompatibilityMode; Columns = Option.map columns value.Columns; SectionStart = value.SectionStart
      HeaderReferences = references value.HeaderReferences; FooterReferences = references value.FooterReferences
      PageNumberStart = value.PageNumberStart; PageNumberFormat = value.PageNumberFormat; VerticalAlign = value.VerticalAlign }
let private sectionSetup (value: SectionPageSetup) : WordEditorPackage.PageSetup =
    { PageWidth = value.PageWidth; PageHeight = value.PageHeight; MarginLeft = value.MarginLeft; MarginRight = value.MarginRight
      MarginTop = value.MarginTop; MarginBottom = value.MarginBottom; HeaderDistance = value.HeaderDistance; FooterDistance = value.FooterDistance
      DifferentFirstPage = value.DifferentFirstPage; EvenAndOddHeaders = value.EvenAndOddHeaders
      DefaultTabStop = None; CompatibilityMode = None; Columns = None; SectionStart = None; HeaderReferences = []; FooterReferences = []
      PageNumberStart = None; PageNumberFormat = None; VerticalAlign = None }
let private paragraphBorders (value: ParagraphBorders option) : WordEditorPackage.Borders option =
    let border (value: ParagraphBorder) : WordEditorPackage.Border =
        { Style = value.Style; Color = value.Color; Width = value.Width; Space = value.Space }
    value |> Option.map (fun value ->
        { Top = Option.map border value.Top; Right = Option.map border value.Right; Bottom = Option.map border value.Bottom
          Left = Option.map border value.Left; InsideH = Option.map border value.Between; InsideV = None })
let private tableBorders (value: TableBorders option) : WordEditorPackage.Borders option =
    let border (value: TableBorder) : WordEditorPackage.Border =
        { Style = value.Style; Color = value.Color; Width = value.Width; Space = None }
    value |> Option.map (fun value ->
        { Top = Option.map border value.Top; Right = Option.map border value.Right; Bottom = Option.map border value.Bottom
          Left = Option.map border value.Left; InsideH = Option.map border value.InsideH; InsideV = Option.map border value.InsideV })
let private shading (value: Shading option) : WordEditorPackage.Shading option =
    value |> Option.map (fun value ->
        { Pattern = value.Pattern; ForegroundColor = value.ForegroundColor; BackgroundColor = value.BackgroundColor })

let private paragraphToEditor (context: WordStyles.DocumentStyleContext option) (headingLevel: int option) (para: ParagraphBlock) : WordEditorPackage.EditorParagraph =
    let metadata = para.Metadata
    let buffer =
        let rec merge runs =
            match runs with
            | Text first :: Text rest :: tail
                when first.Style.RenderedPageBreakBefore && first.Style.FieldCode.IsSome &&
                     rest.Start = first.Start + first.Length && rest.Style = { first.Style with RenderedPageBreakBefore = false } ->
                Text { first with Length = first.Length + rest.Length } :: merge tail
            | run :: tail -> run :: merge tail
            | [] -> []
        if para.Buffer.Runs |> List.exists (function Text run -> run.Style.RenderedPageBreakBefore && run.Style.FieldCode.IsSome | _ -> false) then
            { para.Buffer with Runs = merge para.Buffer.Runs }
        else para.Buffer
    let namedStyle =
        context |> Option.bind (fun ctx ->
            metadata.Style |> Option.orElse ctx.DefaultParagraphStyleId |> Option.bind (fun id -> Map.tryFind id ctx.ResolvedParagraphStyles))
    let fromStyles (pick: NamedStyle -> float option) =
        namedStyle |> Option.bind pick |> Option.orElseWith (fun () -> context |> Option.bind (fun ctx -> ctx.Defaults) |> Option.bind pick)
    let explicitMeasure flag value inherited =
        match value with
        | None -> None
        | Some v when ParagraphExplicit.has flag metadata -> Some v
        | Some v -> match inherited with Some i when abs (i - v) < 0.01 -> None | None when abs v < 0.01 -> None | _ -> Some v
    let levelIndents = if metadata.ListStyle <> NoList then metadata.ListLevelIndents else None
    let explicitIndent flag value styled level =
        match value, level with
        | Some v, Some (Some l) when not (ParagraphExplicit.has flag metadata) && abs (v - l) < 0.01 -> None
        | _ -> explicitMeasure flag value styled
    let paragraphStyle =
        match context with
        | Some ctx when not (headingLevel.IsSome && metadata.Style.IsNone) ->
            let block = match headingLevel with Some level -> Heading (level, para) | None -> Paragraph para
            Some (ctx, WordStyles.forBlockInContext ctx block)
        | _ -> None
    let inheritedStyle (style: TextBuffer.TextStyle) =
        paragraphStyle |> Option.map (fun (ctx, blockStyle) -> WordStyles.effectiveTextStyleInContext ctx blockStyle { TextBuffer.defaultStyle with Style = style.Style })
    let explicitToggle flag (pick: TextBuffer.TextStyle -> bool) inherited style =
        if ExplicitFormat.has flag style then Some (pick style)
        elif pick style && not (inherited |> Option.exists pick) then Some true else None
    let explicitSize inherited (style: TextBuffer.TextStyle) =
        if ExplicitFormat.has ExplicitFormat.FontSize style then Some style.FontSize
        elif style.FontSize = TextBuffer.defaultStyle.FontSize then None
        else
            match inherited with
            | Some (inherited: TextBuffer.TextStyle) ->
                let size = if inherited.FontSize = TextBuffer.defaultStyle.FontSize then 10.0 * 96.0 / 72.0 else inherited.FontSize
                if abs (size - style.FontSize) < 0.01 then None else Some style.FontSize
            | None -> Some style.FontSize
    let explicitAlignment =
        if metadata.ExplicitAlignment then Some metadata.Alignment
        elif metadata.Alignment = Left then None
        else match paragraphStyle with Some (_, style) when style.Alignment = Some metadata.Alignment -> None | _ -> Some metadata.Alignment
    let startKey (start: TextBuffer.BookmarkStart) = (if start.EndsHere then "point " else "start ") + start.Id
    let endKey = function TextBuffer.Ends id -> "end " + id | TextBuffer.Opens start -> (if start.EndsHere then "opens point " else "opens ") + start.Id
    let textRuns = buffer.Runs |> List.choose (function Text run -> Some run | _ -> None)
    let firstStarts =
        textRuns |> List.collect (fun run -> run.Style.BookmarkStarts |> List.map (fun start -> startKey start, run))
        |> List.fold (fun owners (key, run) -> if Map.containsKey key owners then owners else Map.add key run owners) Map.empty
    let lastEnds =
        textRuns |> List.collect (fun run -> run.Style.BookmarkEnds |> List.map (fun ending -> endKey ending, run))
        |> List.fold (fun owners (key, run) -> Map.add key run owners) Map.empty
    let bookmarkStart (value: TextBuffer.BookmarkStart) : ModelText.BookmarkStart =
        { Id = value.Id; Name = value.Name; EndsHere = value.EndsHere }
    let bookmarkEnd = function
        | TextBuffer.Ends id -> ModelText.Ends id
        | TextBuffer.Opens value -> ModelText.Opens (bookmarkStart value)
    let formatting inherited (style: TextBuffer.TextStyle) : WordEditorPackage.RunFormatting =
        { Style = style.Style; StyleInfo = None; Bold = style.Bold; Italic = style.Italic; Underline = style.Underline
          UnderlineStyle = style.UnderlineStyle; UnderlineColor = style.UnderlineColor; AllCaps = style.AllCaps
          SmallCaps = style.SmallCaps; Strike = style.Strike; DoubleStrike = style.DoubleStrike; Hidden = style.Hidden
          Position = style.Position; EmphasisMark = style.EmphasisMark; FontSize = style.FontSize; CharacterSpacing = style.CharacterSpacing
          FontFamily = Option.defaultValue "Calibri" style.FontFamily
          FontFamilyInherited = Some (not (ExplicitFormat.has ExplicitFormat.FontFamily style) && (style.FontFamily.IsNone || inherited |> Option.exists (fun (s: TextBuffer.TextStyle) -> s.FontFamily = style.FontFamily)))
          Color = style.Color; Highlight = style.Highlight; VertAlign = style.VertAlign; HasExplicit = true
          ExplicitBold = explicitToggle ExplicitFormat.Bold (fun s -> s.Bold) inherited style
          ExplicitItalic = explicitToggle ExplicitFormat.Italic (fun s -> s.Italic) inherited style
          ExplicitUnderline = explicitToggle ExplicitFormat.Underline (fun s -> s.Underline) inherited style
          ExplicitAllCaps = explicitToggle ExplicitFormat.AllCaps (fun s -> s.AllCaps) inherited style
          ExplicitSmallCaps = explicitToggle ExplicitFormat.SmallCaps (fun s -> s.SmallCaps) inherited style
          ExplicitStrike = explicitToggle ExplicitFormat.Strike (fun s -> s.Strike) inherited style
          ExplicitDoubleStrike = explicitToggle ExplicitFormat.DoubleStrike (fun s -> s.DoubleStrike) inherited style
          ExplicitHidden = explicitToggle ExplicitFormat.Hidden (fun s -> s.Hidden) inherited style
          ExplicitFontSize = explicitSize inherited style
          ExplicitColorAuto = ExplicitFormat.has ExplicitFormat.ColorAuto style && style.Color.IsNone }
    let emptyRun content format : WordEditorPackage.EditorRun =
        { Content = content; Formatting = format; Hyperlink = None; InactiveHyperlink = None; CommentId = None
          CommentText = None; CommentAuthor = None; CommentDate = None; NoteId = None; NoteType = None; NoteText = None
          BookmarkStarts = []; BookmarkEnds = []; FieldCode = None; FieldResult = None; Revision = None
          RenderedPageBreakBefore = false; Checkbox = None; ContentControls = []; SourceRun = None; NoProps = false }
    let runs =
        buffer.Runs |> List.map (function
            | Text run ->
                let s = run.Style
                let revision =
                    if s.RevisionId.IsNone && s.RevisionType.IsNone && s.RevisionAuthor.IsNone && s.RevisionDate.IsNone && s.DeletedInsertion.IsNone && s.MoveName.IsNone then None
                    else
                        let value: WordEditorPackage.Revision =
                            { Id = s.RevisionId; Kind = s.RevisionType; Author = s.RevisionAuthor; Date = s.RevisionDate
                              Inserted = s.DeletedInsertion |> Option.map (fun mark -> ({ Id = mark.Id; Author = mark.Author; Date = mark.Date } : DocxReader.RevisionMark))
                              MoveName = s.MoveName }
                        Some value
                let text = Rope.slice run.Start run.Length buffer.Rope
                let content =
                    match s.FieldBoundary with
                    | Some boundary ->
                        WordEditorPackage.FieldBoundary (
                            match boundary with
                            | TextBuffer.FieldBoundary.Begin -> ModelText.Begin
                            | TextBuffer.FieldBoundary.Instruction text -> ModelText.Instruction text
                            | TextBuffer.FieldBoundary.Separate -> ModelText.Separate
                            | TextBuffer.FieldBoundary.End -> ModelText.End)
                    | None ->
                        match s.EquationXml with
                        | Some xml -> WordEditorPackage.Equation (text, xml)
                        | None -> WordEditorPackage.Text text
                { emptyRun content (formatting (inheritedStyle s) s) with
                    Hyperlink = s.Hyperlink; InactiveHyperlink = s.InactiveHyperlink; CommentId = s.CommentId; CommentText = s.CommentText
                    CommentAuthor = s.CommentAuthor; CommentDate = s.CommentDate; NoteId = s.NoteId; NoteType = s.NoteType; NoteText = s.NoteText
                    BookmarkStarts = s.BookmarkStarts |> List.filter (fun value -> obj.ReferenceEquals(Map.find (startKey value) firstStarts, run)) |> List.map bookmarkStart
                    BookmarkEnds = s.BookmarkEnds |> List.filter (fun value -> obj.ReferenceEquals(Map.find (endKey value) lastEnds, run)) |> List.map bookmarkEnd
                    FieldCode = s.FieldCode; FieldResult = s.FieldResult; Revision = revision
                    RenderedPageBreakBefore = s.RenderedPageBreakBefore; Checkbox = s.Checkbox; SourceRun = s.SourceRun
                    ContentControls = s.ContentControls |> List.map (fun layer ->
                        ({ Group = layer.Group; Id = layer.Id; Tag = layer.Tag; Alias = layer.Alias; Xml = layer.Xml } : WordEditorPackage.ContentControlLayer)) }
            | Image img ->
                let picture : WordEditorPackage.Image =
                    { Source = Some img.Src; Linked = img.Linked; Width = img.Width; Height = img.Height; Alt = img.Alt; Name = img.Name; Title = img.Title
                      Floating = img.Floating; Wrap = img.Wrap; PositionX = img.PositionX; PositionY = img.PositionY
                      PositionHRelativeFrom = img.PositionHRelativeFrom; PositionVRelativeFrom = img.PositionVRelativeFrom
                      PositionHAlign = img.PositionHAlign; PositionVAlign = img.PositionVAlign; BehindDoc = img.BehindDoc
                      AllowOverlap = img.AllowOverlap; LayoutInCell = img.LayoutInCell; DistanceTop = img.DistanceTop; DistanceBottom = img.DistanceBottom
                      DistanceLeft = img.DistanceLeft; DistanceRight = img.DistanceRight; RelativeHeight = img.RelativeHeight; Locked = img.Locked
                      SimplePosition = img.SimplePosition; SimplePositionX = img.SimplePositionX; SimplePositionY = img.SimplePositionY
                      CropLeft = img.CropLeft; CropTop = img.CropTop; CropRight = img.CropRight; CropBottom = img.CropBottom
                      WrapText = img.WrapText; WrapPolygon = img.WrapPolygon; FlipH = img.FlipH; FlipV = img.FlipV
                      Opaque = img.Opaque |> Option.map (fun value -> ({ Xml = value.Xml; Kind = value.Kind; ParagraphLevel = value.ParagraphLevel; ShapeCandidate = false } : DocxReader.OpaqueObject)) }
                { emptyRun (WordEditorPackage.Picture picture) (formatting None TextBuffer.defaultStyle) with SourceRun = img.SourceRun })
        |> List.toArray
    let paragraphFormat : WordEditorPackage.ParagraphFormatting =
        { Style = metadata.Style; StyleInfo = None; Alignment = alignment metadata.Alignment; HasExplicit = true
          ExplicitAlignment = Option.map alignment explicitAlignment
          IndentLeft = metadata.IndentLeft; IndentRight = metadata.IndentRight; FirstLineIndent = metadata.FirstLineIndent
          SpacingBefore = metadata.SpacingBefore; SpacingAfter = metadata.SpacingAfter; LineSpacing = metadata.LineSpacing; LineSpacingRule = metadata.LineSpacingRule
          RightTabPosition = metadata.RightTabPosition; RightTabLeader = metadata.RightTabLeader
          TabStops = metadata.TabStops |> List.map (fun stop -> ({ Position = stop.Position; Alignment = stop.Alignment; Leader = stop.Leader } : WordEditorPackage.TabStop))
          Rtl = metadata.Rtl; ContextualSpacing = metadata.ContextualSpacing; KeepNext = metadata.KeepNext; KeepLines = metadata.KeepLines; WidowControl = metadata.WidowControl
          PageBreakBefore = metadata.PageBreakBefore; PageBreakBeforeDerived = metadata.PageBreakBefore && ParagraphExplicit.has ParagraphExplicit.DerivedPageBreakBefore metadata
          PageBreakAtStart = metadata.PageBreakBefore && ParagraphExplicit.has ParagraphExplicit.PageBreakRunBefore metadata
          PageBreakAfter = metadata.PageBreakAfter; ColumnBreakAfter = metadata.PageBreakAfter && metadata.ColumnBreakAfter
          RenderedPageBreakBefore = metadata.RenderedPageBreakBefore; BackgroundColor = metadata.BackgroundColor
          ShadingPattern = metadata.ParagraphShadingPattern; ShadingForeground = metadata.ParagraphShadingForeground
          Borders = paragraphBorders metadata.ParagraphBorders; EffectiveBorders = paragraphBorders metadata.EffectiveParagraphBorders
          OutlineLevel = metadata.OutlineLevel; ExplicitParagraphFormat = true
          ExplicitIndentLeft = explicitIndent ParagraphExplicit.IndentLeft metadata.IndentLeft (fromStyles (fun s -> s.IndentLeft)) (levelIndents |> Option.map fst)
          ExplicitIndentRight = explicitMeasure ParagraphExplicit.IndentRight metadata.IndentRight (fromStyles (fun s -> s.IndentRight))
          ExplicitFirstLineIndent = explicitIndent ParagraphExplicit.FirstLineIndent metadata.FirstLineIndent (fromStyles (fun s -> s.FirstLineIndent)) (levelIndents |> Option.map snd)
          ExplicitSpacingBefore = explicitMeasure ParagraphExplicit.SpacingBefore metadata.SpacingBefore (fromStyles (fun s -> s.SpacingBefore))
          ExplicitSpacingAfter = explicitMeasure ParagraphExplicit.SpacingAfter metadata.SpacingAfter (fromStyles (fun s -> s.SpacingAfter))
          ExplicitLineSpacing =
            match metadata.LineSpacing with
            | Some value when ParagraphExplicit.has ParagraphExplicit.LineSpacing metadata -> Some value
            | Some value ->
                let rule = namedStyle |> Option.bind (fun s -> s.LineSpacingRule)
                let fixedRule r = r = Some "exact" || r = Some "atLeast"
                match fromStyles (fun s -> s.LineSpacing) with
                | Some inherited when not (fixedRule metadata.LineSpacingRule) && not (fixedRule rule) && abs (inherited - value) < 0.001 -> None
                | Some _ when fixedRule metadata.LineSpacingRule && metadata.LineSpacingRule = rule -> None
                | _ -> Some value
            | None -> None
          ExplicitContextualSpacing = if ParagraphExplicit.has ParagraphExplicit.ContextualSpacing metadata then metadata.ContextualSpacing else None
          ExplicitBackgroundColor =
            match metadata.BackgroundColor with
            | Some color when ParagraphExplicit.has ParagraphExplicit.BackgroundColor metadata || (namedStyle |> Option.bind (fun s -> s.BackgroundColor)) <> Some color -> Some color
            | _ -> None
          ExplicitShadingPattern = metadata.ExplicitParagraphShadingPattern; ExplicitShadingForeground = metadata.ExplicitParagraphShadingForeground
          ExplicitTabs = ParagraphExplicit.has ParagraphExplicit.Tabs metadata; HasMarkProperties = metadata.HasMarkProperties
          ExplicitMarkFontSize = metadata.ExplicitMarkFontSize; ParagraphMarkFontSize = metadata.ParagraphMarkFontSize
          MarkRevision = metadata.MarkRevision |> Option.map (fun revision ->
            ({ Id = revision.Mark.Id; Kind = Some revision.RevisionType; Author = revision.Mark.Author; Date = revision.Mark.Date; Inserted = None; MoveName = None } : WordEditorPackage.Revision))
          DropCapKind = metadata.DropCap |> Option.map (fun cap -> if cap.InMargin then "margin" else "drop")
          DropCapLines = metadata.DropCap |> Option.map (fun cap -> cap.Lines); DropCapDistance = metadata.DropCap |> Option.map (fun cap -> cap.Distance)
          ContentControlId = metadata.ContentControlId; ContentControlTag = metadata.ContentControlTag
          ContentControlAlias = metadata.ContentControlAlias; ContentControlXml = metadata.ContentControlXml
          ContentControlGroup = metadata.ContentControlGroup
          ContentControlAncestors = metadata.ContentControlAncestors |> List.map (fun layer ->
              ({ Group = layer.Group; Id = layer.Id; Tag = layer.Tag; Alias = layer.Alias; Xml = layer.Xml } : WordEditorPackage.ContentControlLayer))
          EnclosingCellControls = metadata.EnclosingCellControls |> List.map (fun layer ->
              ({ Group = layer.Group; Id = layer.Id; Tag = layer.Tag; Alias = layer.Alias; Xml = layer.Xml } : WordEditorPackage.ContentControlLayer)) }
    let list =
        if metadata.ListStyle = NoList && metadata.ListId <> Some "0" then None
        else Some ({
            Style = listStyle metadata.ListStyle; Level = metadata.ListLevel; Id = metadata.ListId
            SourceId = metadata.ListSourceId; Restart = metadata.ListRestart; FormatEdited = metadata.ListFormatEdited
            NumberFormat = metadata.ListNumberFormat; LevelText = metadata.ListLevelText; Start = metadata.ListStart
            StartOverrides = metadata.ListStartOverrides; ListLevelIndents = metadata.ListLevelIndents
            Levels = metadata.ListLevels |> List.map (fun level ->
                let item: WordEditorPackage.ListLevel =
                    { Level = level.Level; NumberFormat = Some level.NumberFormat; LevelText = level.LevelText
                      Start = level.Start; IndentLeft = level.IndentLeft; FirstLineIndent = level.FirstLineIndent; Suffix = level.Suffix
                      Justification = level.Justification; Legal = level.Legal; Restart = level.Restart; MarkFont = level.MarkFont
                      MarkBold = level.MarkBold; MarkItalic = level.MarkItalic; MarkColor = level.MarkColor; MarkSize = level.MarkSize }
                item) } : WordEditorPackage.ParagraphList)
    { Runs = runs; HeadingLevel = headingLevel; Formatting = paragraphFormat; List = list
      SectionBreak = metadata.SectionBreak; SectionPageSetup = Option.map sectionSetup metadata.SectionPageSetup
      SectionColumns = Option.map columns metadata.SectionColumns; SectionVerticalAlign = metadata.SectionVerticalAlign
      SectionPageNumberStart = metadata.SectionPageNumberStart; SectionPageNumberFormat = metadata.SectionPageNumberFormat
      HeaderReferences = references metadata.HeaderReferences; FooterReferences = references metadata.FooterReferences
      Source = source metadata.Origin; SourceParagraph = None; SpliceOriginal = None; SpliceFrom = None
      SpliceLeading = true; SpliceTrailing = true; SpliceJoined = [||]
      SourceIndex = None; SourceUnchanged = false; SourceWhy = None; SourceTextChanges = None }

type private ExportContext = {
    Styles: WordStyles.DocumentStyleContext option
    ImportedShapes: Map<string, TextBoxBlock>
}

let private movedShape (context: ExportContext) (shape: TextBoxBlock) =
    match Map.tryFind shape.Id context.ImportedShapes with
    | Some original when shape.SourceXml.IsSome && not (obj.ReferenceEquals(original, shape)) &&
                         obj.ReferenceEquals(original.Content, shape.Content) && obj.ReferenceEquals(original.Children, shape.Children) ->
        let placed = { original with Content = []; Children = []; X = shape.X; Y = shape.Y; Width = shape.Width; Height = shape.Height
                                     PositionHAlign = shape.PositionHAlign; PositionVAlign = shape.PositionVAlign
                                     PositionHRelativeFrom = shape.PositionHRelativeFrom; PositionVRelativeFrom = shape.PositionVRelativeFrom }
        if (placed = { shape with Content = []; Children = [] } || not original.Children.IsEmpty) &&
           placed <> { original with Content = []; Children = [] } then Some original else None
    | _ -> None
let private unchangedShape (context: ExportContext) (shape: TextBoxBlock) =
    shape.SourceXml.IsSome &&
    (match Map.tryFind shape.Id context.ImportedShapes with
     | Some original ->
         obj.ReferenceEquals(original, shape) ||
         (obj.ReferenceEquals(original.Content, shape.Content) && obj.ReferenceEquals(original.Children, shape.Children) &&
          { original with Content = []; Children = [] } = { shape with Content = []; Children = [] })
     | None -> false)
let private width (value: TableWidth option) = value |> Option.map (fun width -> Option.defaultValue 0.0 width.Value)
let private widthType (value: TableWidth option) =
    value |> Option.map (fun width -> match width.Kind with DxaWidth -> "dxa" | PercentWidth -> "pct" | AutoWidth -> "auto" | PointWidth -> "pt")
let private dxa value = value |> Option.map (fun pixels -> System.Math.Floor (pixels * WordLayout.canvasScale.DxaPerUnit + 0.5))

let rec private blockToEditor (context: ExportContext) (block: Block) : WordEditorPackage.EditorBlock =
    match block with
    | Paragraph para -> WordEditorPackage.Paragraph (paragraphToEditor context.Styles None para)
    | Heading (level, para) -> WordEditorPackage.Paragraph (paragraphToEditor context.Styles (Some level) para)
    | Table table -> WordEditorPackage.Table (tableToEditor context table)
    | TextBox shape -> WordEditorPackage.TextBox (shapeToEditor context shape)
and private tableToEditor context (table: TableBlock) : WordEditorPackage.EditorTable =
    let makeCell (cell: TableCell) : WordEditorPackage.EditorCell =
        { Paragraphs = cell.Content |> List.map (paragraphToEditor context.Styles None >> WordEditorPackage.Paragraph) |> List.toArray
          NestedTables = cell.NestedTables |> List.map (tableToEditor context) |> List.toArray
          NestedTableAnchors = List.toArray cell.NestedTableAnchors; ColSpan = cell.ColSpan; BackgroundColor = cell.BackgroundColor
          EnclosingControls = cell.EnclosingControls |> List.map (fun layer ->
              ({ Group = layer.Group; Id = layer.Id; Tag = layer.Tag; Alias = layer.Alias; Xml = layer.Xml } : WordEditorPackage.ContentControlLayer))
          Shading = shading cell.Shading; Width = cell.Width; PreferredWidth = width cell.PreferredWidth; PreferredWidthType = widthType cell.PreferredWidth
          BorderColor = cell.BorderColor; BorderWidth = cell.BorderWidth; Borders = tableBorders cell.Borders
          VerticalAlignment = cell.VerticalAlignment
          VerticalMerge = cell.VerticalMerge |> Option.map (function RestartMerge -> "restart" | ContinueMerge -> "continue")
          PaddingTop = dxa cell.PaddingTop; PaddingRight = dxa cell.PaddingRight; PaddingBottom = dxa cell.PaddingBottom; PaddingLeft = dxa cell.PaddingLeft
          Source = source cell.Origin; SourceCell = None }
    let makeRow (row: TableRow) : WordEditorPackage.EditorRow =
        { Cells = row.Cells |> List.map makeCell |> List.toArray
          IsHeader = row.IsHeader; Height = row.Height; HeightRule = row.HeightRule; CanSplit = row.CanSplit
          CellSpacing = dxa row.CellSpacing; PaddingTop = dxa row.PaddingTop; PaddingRight = dxa row.PaddingRight
          PaddingBottom = dxa row.PaddingBottom; PaddingLeft = dxa row.PaddingLeft
          GridBefore = row.GridBefore; GridAfter = row.GridAfter; WidthBefore = dxa row.WidthBefore; WidthAfter = dxa row.WidthAfter
          BordersOverride = tableBorders row.BordersOverride; Source = source row.Origin; SourceRow = None }
    let rows = gatherVerticalMergeContent table.Rows |> List.map makeRow |> List.toArray
    { Rows = rows; PreferredWidth = width table.PreferredWidth; PreferredWidthType = widthType table.PreferredWidth
      Alignment = table.Alignment; Indent = dxa table.Indent
      LayoutMode = table.LayoutMode |> Option.map (function FixedTableLayout -> "fixed" | AutoFitTableLayout -> "autofit")
      GridColumns = table.GridColumns |> List.map (fun w -> Option.defaultValue 0.0 w.Value) |> List.toArray
      Borders = tableBorders table.Borders; BackgroundColor = table.BackgroundColor; Shading = shading table.Shading
      CellSpacing = dxa table.CellSpacing; PaddingTop = dxa table.PaddingTop; PaddingRight = dxa table.PaddingRight
      PaddingBottom = dxa table.PaddingBottom; PaddingLeft = dxa table.PaddingLeft; Floating = table.Floating; FloatingXml = table.FloatingXml
      StyleId = table.StyleId
      ContentControlId = table.ContentControlId; ContentControlTag = table.ContentControlTag
      ContentControlAlias = table.ContentControlAlias; ContentControlXml = table.ContentControlXml
      ContentControlGroup = table.ContentControlGroup
      ContentControlAncestors = table.ContentControlAncestors |> List.map (fun layer ->
          ({ Group = layer.Group; Id = layer.Id; Tag = layer.Tag; Alias = layer.Alias; Xml = layer.Xml } : WordEditorPackage.ContentControlLayer))
      Look = table.Look |> Option.map (fun look ->
        ({ FirstRow = look.FirstRow; LastRow = look.LastRow; FirstColumn = look.FirstColumn; LastColumn = look.LastColumn
           HorizontalBanding = look.HorizontalBanding; VerticalBanding = look.VerticalBanding } : WordEditorPackage.EditorTableLook))
      Source = source None; SourceIndex = None; SourceUnchanged = false; SourceWhy = None; SourceTextChanges = None; SourceTable = None }
and private shapeToEditor context (shape: TextBoxBlock) : WordEditorPackage.EditorTextBox =
    let geometry, path, customWidth, customHeight =
        match shape.Geometry with
        | GeoRect -> "rect", [], 0.0, 0.0
        | GeoPreset preset -> preset, [], 0.0, 0.0
        | GeoCustom (segments, w, h) ->
            let path = segments |> List.map (function
                | PMove (x, y) -> "M", [| x; y |]
                | PLine (x, y) -> "L", [| x; y |]
                | PCubic (a, b, c, d, x, y) -> "C", [| a; b; c; d; x; y |]
                | PQuad (a, b, x, y) -> "Q", [| a; b; x; y |]
                | PClose -> "Z", [||])
            "custom", path, w, h
    let moved = movedShape context shape
    { Paragraphs = shape.Content |> List.map (paragraphToEditor context.Styles None >> WordEditorPackage.Paragraph) |> List.toArray
      Children = shape.Children |> List.map (shapeToEditor context) |> List.toArray
      ShapeId = shape.SourceShapeId; ShapeKind = shape.SourceShapeKind; Name = shape.Name; Geometry = geometry; CustomPath = path
      CustomWidth = customWidth; CustomHeight = customHeight; X = shape.X; Y = shape.Y
      PositionHRelativeFrom = shape.PositionHRelativeFrom; PositionVRelativeFrom = shape.PositionVRelativeFrom
      PositionHAlign = shape.PositionHAlign; PositionVAlign = shape.PositionVAlign; Width = shape.Width; Height = shape.Height
      FillColor = shape.FillColor; BorderColor = shape.BorderColor
      NoFill = shape.FillColor.IsNone; NoLine = shape.BorderColor.IsNone
      BorderWidth = shape.BorderWidth; BorderDash = shape.BorderDash
      Gradient = shape.Gradient |> Option.map (fun g -> g.Stops |> List.map (fun stop -> stop.Offset, stop.Color)) |> Option.defaultValue []
      GradientAngle = shape.Gradient |> Option.map (fun g -> g.Angle) |> Option.defaultValue 0.0
      GradientRadial = shape.Gradient |> Option.map (fun g -> g.Radial) |> Option.defaultValue false
      Rotation = shape.Rotation; FlipH = shape.FlipH; FlipV = shape.FlipV; Inline = shape.Inline; Wrap = shape.Wrap
      BehindDoc = shape.BehindDoc; AllowOverlap = shape.AllowOverlap; LayoutInCell = shape.LayoutInCell
      DistanceTop = shape.DistanceTop; DistanceBottom = shape.DistanceBottom; DistanceLeft = shape.DistanceLeft; DistanceRight = shape.DistanceRight
      RelativeHeight = shape.RelativeHeight; AnchorBlock = None; InsetLeft = shape.InsetLeft; InsetTop = shape.InsetTop
      InsetRight = shape.InsetRight; InsetBottom = shape.InsetBottom; VerticalAnchor = shape.VerticalAnchor; AutoFit = shape.AutoFit
      NoTextWrap = shape.NoTextWrap; FitText = shape.FitText; Opacity = shape.Opacity; ChildSpace = shape.ChildSpace; ImageSource = shape.ImageSource
      SourceXml = if unchangedShape context shape || (not shape.Children.IsEmpty && shape.SourceXml.IsSome) || moved.IsSome then shape.SourceXml else None
      SourceMoved = moved |> Option.map (fun original ->
          ({ H = original.X <> shape.X || original.PositionHAlign <> shape.PositionHAlign || original.PositionHRelativeFrom <> shape.PositionHRelativeFrom
             V = original.Y <> shape.Y || original.PositionVAlign <> shape.PositionVAlign || original.PositionVRelativeFrom <> shape.PositionVRelativeFrom
             Size = original.Width <> shape.Width || original.Height <> shape.Height } : WordEditorPackage.ShapeMovement))
      SourceIndex = None; SourceUnchanged = false; SourceWhy = None; SourceTextChanges = None }

let private blockId = function
    | Paragraph para | Heading (_, para) -> para.Id
    | Table table -> table.Id
    | TextBox shape -> shape.Id
let private sameBlockObject a b =
    match a, b with
    | Paragraph x, Paragraph y -> obj.ReferenceEquals(x, y)
    | Heading (lx, x), Heading (ly, y) -> lx = ly && obj.ReferenceEquals(x, y)
    | Table x, Table y -> obj.ReferenceEquals(x, y)
    | TextBox x, TextBox y -> obj.ReferenceEquals(x, y)
    | _ -> false

let private blockTextChanges context derivations original current : Result<WordEditorPackage.SourceTextChanges, string> =
    match storyParagraphs original, storyParagraphs current with
    | Some before, Some after when before.Length = after.Length && List.forall2 (fun (a: ParagraphBlock) (b: ParagraphBlock) -> a.Id = b.Id) before after ->
        let changed = List.zip before after |> List.mapi (fun index pair -> index, pair) |> List.filter (fun (_, (a, b)) -> not (obj.ReferenceEquals(a, b)))
        let restored =
            [ current ] |> mapParagraphsOfBlocks (fun para ->
                changed |> List.tryPick (fun (_, (a, b)) -> if obj.ReferenceEquals(b, para) then Some a else None) |> Option.defaultValue para) |> List.head
        let sameKind =
            match original, current with
            | Heading (x, _), Heading (y, _) -> x = y
            | Paragraph _, Paragraph _ | Table _, Table _ -> true
            | _ -> false
        if changed.IsEmpty then Error "something other than its paragraphs was changed"
        elif not sameKind then Error "it is another kind of block"
        elif blockToEditor context restored <> blockToEditor context original then Error "more than its paragraphs was changed"
        else
            let export para =
                let level = match current with Heading (level, _) -> Some level | _ -> None
                paragraphToEditor context.Styles level para
            let edits = changed |> List.filter (fun (_, (a, b)) -> export a <> export b) |> List.map (fun (index, (a, b)) ->
                match Map.tryFind b.Id derivations with
                | Some derivation when obj.ReferenceEquals(derivation.Original, a) &&
                                       (derivation.From > 0 || derivation.Until < TextBuffer.length a.Buffer || not derivation.Joined.IsEmpty) ->
                    derivedEdit export export index derivation b
                | _ -> paragraphEdit export export index a b) |> List.toArray
            Ok ({ Texts = before |> List.map (fun para -> TextBuffer.toString para.Buffer) |> List.toArray; Edits = edits } : WordEditorPackage.SourceTextChanges)
    | Some before, Some after -> Error (sprintf "its paragraphs are not the paragraphs it had (%d, it had %d)" after.Length before.Length)
    | _ -> Error "a shape"

let private markDerivedTable index (original: WordEditorPackage.EditorTable) (table: WordEditorPackage.EditorTable) =
    let at origin = match origin with Some (WordEditorPackage.ReadAt at | WordEditorPackage.LikeAt at) when at >= 0 -> Some at | _ -> None
    let originalRows = original.Rows |> Array.choose (fun row -> at row.Source.Origin |> Option.map (fun at -> at, row)) |> Map.ofArray
    let originalCells = original.Rows |> Array.collect (fun row -> row.Cells) |> Array.choose (fun cell -> at cell.Source.Origin |> Option.map (fun at -> at, cell)) |> Map.ofArray
    let mutable usedRows = Set.empty
    let mutable usedCells = Set.empty
    let cellOwn (cell: WordEditorPackage.EditorCell) = { cell with Paragraphs = [||]; NestedTables = [||]; NestedTableAnchors = [||]; SourceCell = None }
    let rowOwn (row: WordEditorPackage.EditorRow) = { row with Cells = [||]; SourceRow = None }
    let tableOwn (table: WordEditorPackage.EditorTable) =
        { table with Rows = [||]; SourceTable = None; SourceIndex = None; SourceUnchanged = false; SourceWhy = None; SourceTextChanges = None }
    let rows = table.Rows |> Array.map (fun row ->
        let sourceRow =
            match at row.Source.Origin |> Option.bind (fun at -> Map.tryFind at originalRows |> Option.map (fun original -> at, original)) with
            | Some (at, was) ->
                let copy = row.Source.Copy || Set.contains at usedRows
                if not copy then usedRows <- Set.add at usedRows
                let value: WordEditorPackage.RowSource =
                    { At = at; Copy = copy; Same = not copy && row = was
                      Before = (if rowOwn row = rowOwn was then None else Some was); Was = None }
                Some value
            | None -> None
        let cells = row.Cells |> Array.map (fun cell ->
            let sourceCell =
                match at cell.Source.Origin |> Option.bind (fun at -> Map.tryFind at originalCells |> Option.map (fun original -> at, original)) with
                | Some (at, was) ->
                    let copy = cell.Source.Copy || Set.contains at usedCells
                    if not copy then usedCells <- Set.add at usedCells
                    let value: WordEditorPackage.CellSource =
                        { At = at; Copy = copy; Same = not copy && cell = was
                          Before = (if cellOwn cell = cellOwn was then None else Some was); Was = Some was }
                    Some value
                | None -> None
            { cell with SourceCell = sourceCell })
        { row with Cells = cells; SourceRow = sourceRow })
    { table with
        Rows = rows
        SourceTable =
            Some { At = at table.Source.Origin |> Option.defaultValue -1
                   Block = Some index; Copy = false; Same = false; Was = None
                   Before = if tableOwn table = tableOwn original then None else Some original } }

let private markDerivedParagraphs (context: ExportContext) derivations block item =
    let level = match block with Heading (level, _) -> Some level | _ -> None
    let export = paragraphToEditor context.Styles level
    let mutable remaining = storyParagraphs block |> Option.defaultValue []
    let rec visitBlock item =
        match item with
        | WordEditorPackage.TextBox _ -> item
        | WordEditorPackage.Paragraph paragraph ->
            match remaining with
            | [] -> item
            | para :: rest ->
                remaining <- rest
                match Map.tryFind para.Id derivations with
                | None -> item
                | Some derivation ->
                    let oldLevel = if derivation.Level > 0 then Some derivation.Level else None
                    let exportBefore = paragraphToEditor context.Styles oldLevel
                    let edit = derivedEdit exportBefore export 0 derivation para
                    let sourceEdit =
                        { edit with
                            Block = Some derivation.Block; Para = Some derivation.Para
                            OldText = Some (TextBuffer.toString derivation.Original.Buffer)
                            Verbatim = obj.ReferenceEquals(derivation.Original, para) && derivation.From = 0 && derivation.Joined.IsEmpty }
                    WordEditorPackage.Paragraph { paragraph with SourceParagraph = Some sourceEdit }
        | WordEditorPackage.Table table -> WordEditorPackage.Table (visitTable table)
    and visitTable table =
        let rows = table.Rows |> Array.map (fun row ->
            let cells = row.Cells |> Array.map (fun cell ->
                let paras = Array.copy cell.Paragraphs
                let nested = Array.copy cell.NestedTables
                let count = paras.Length
                let order = nested |> Array.mapi (fun i _ -> i, (if i < cell.NestedTableAnchors.Length then cell.NestedTableAnchors.[i] else count) |> max 0 |> min count) |> Array.sortBy snd
                for index in 0 .. count do
                    for i, anchor in order do
                        if anchor = index then nested.[i] <- visitTable nested.[i]
                    if index < count then paras.[index] <- visitBlock paras.[index]
                { cell with Paragraphs = paras; NestedTables = nested })
            { row with Cells = cells })
        { table with Rows = rows }
    visitBlock item

let private blocksWithSource context originalBlocks blocks =
    let known = originalBlocks |> List.mapi (fun index block -> blockId block, (index, block)) |> Map.ofList
    let mutable counts = [||]
    let mutable derived = None
    let derivations () =
        match derived with
        | Some value -> value
        | None ->
            let story, foundCounts = originalStory originalBlocks
            counts <- foundCounts
            let value = derivationsOf story blocks
            derived <- Some value
            value
    let items = blocksInAnchorOrder blocks |> List.map (fun block ->
        let item = blockToEditor context block
        match Map.tryFind (blockId block) known with
        | None -> markDerivedParagraphs context (derivations ()) block item
        | Some (index, original) ->
            let originalItem = blockToEditor context original
            let unchanged = sameBlockObject block original || item = originalItem
            let changes = if unchanged then None else Some (blockTextChanges context (derivations ()) original block)
            let sourceWhy = match changes with Some (Error why) -> Some why | _ -> None
            let textChanges = match changes with Some (Ok edits) -> Some edits | _ -> None
            let marked =
                match item, originalItem, sourceWhy with
                | WordEditorPackage.Table table, WordEditorPackage.Table was, Some _ -> WordEditorPackage.Table (markDerivedTable index was table)
                | _ -> item
            let marked = if sourceWhy.IsSome then markDerivedParagraphs context (derivations ()) block marked else marked
            match marked with
            | WordEditorPackage.Paragraph p -> WordEditorPackage.Paragraph { p with SourceIndex = Some index; SourceUnchanged = unchanged; SourceWhy = sourceWhy; SourceTextChanges = textChanges }
            | WordEditorPackage.Table t -> WordEditorPackage.Table { t with SourceIndex = Some index; SourceUnchanged = unchanged; SourceWhy = sourceWhy; SourceTextChanges = textChanges }
            | WordEditorPackage.TextBox s -> WordEditorPackage.TextBox { s with SourceIndex = Some index; SourceUnchanged = unchanged; SourceWhy = sourceWhy; SourceTextChanges = textChanges }) |> List.toArray
    items, counts

let private styleToEditor (style: NamedStyle) : WordEditorPackage.Style =
    let on (pick: TextBuffer.TextStyle -> bool) = style.Text |> Option.bind (fun text -> if pick text then Some true else None)
    let text = style.Text
    let run : WordEditorPackage.StyleRun =
        { Bold = style.TextBold |> Option.orElse (on (fun t -> t.Bold))
          Italic = style.TextItalic |> Option.orElse (on (fun t -> t.Italic))
          Underline = style.TextUnderline |> Option.orElse (on (fun t -> t.Underline))
          UnderlineStyle = None; UnderlineColor = None
          AllCaps = style.TextAllCaps |> Option.orElse (on (fun t -> t.AllCaps))
          SmallCaps = style.TextSmallCaps |> Option.orElse (on (fun t -> t.SmallCaps))
          Strike = style.TextStrike |> Option.orElse (on (fun t -> t.Strike))
          DoubleStrike = None; Hidden = None
          Position = None; EmphasisMark = None
          VertAlign = text |> Option.bind (fun t -> t.VertAlign)
          FontSize = text |> Option.bind (fun t -> if ExplicitFormat.has ExplicitFormat.FontSize t || t.FontSize <> TextBuffer.defaultStyle.FontSize then Some t.FontSize else None)
          CharacterSpacing = text |> Option.map (fun t -> t.CharacterSpacing)
          FontFamily = text |> Option.bind (fun t -> t.FontFamily); Color = text |> Option.bind (fun t -> t.Color); Highlight = text |> Option.bind (fun t -> t.Highlight) }
    let para : WordEditorPackage.StyleParagraph =
        { SpacingBefore = style.SpacingBefore; SpacingAfter = style.SpacingAfter; ContextualSpacing = style.ContextualSpacing
          LineSpacing = style.LineSpacing; LineSpacingRule = style.LineSpacingRule; Alignment = Option.map alignment style.Alignment
          IndentLeft = style.IndentLeft; IndentRight = style.IndentRight; FirstLineIndent = style.FirstLineIndent; BackgroundColor = style.BackgroundColor
          OutlineLevel = style.OutlineLevel; KeepNext = style.KeepNext; KeepLines = style.KeepLines
          PageBreakBefore = style.PageBreakBefore; WidowControl = style.WidowControl }
    { Id = style.Id; Name = style.Name; StyleType = style.StyleType; IsDefault = style.IsDefault; BasedOn = style.BasedOn
      Next = style.Next; Linked = style.Linked; IsQuickStyle = Some style.IsQuickStyle; Aliases = style.Aliases; UiPriority = style.UiPriority
      SemiHidden = Some style.SemiHidden; UnhideWhenUsed = Some style.UnhideWhenUsed; Run = Some run; Paragraph = Some para }

let private plainBodyBlocks (context: ExportContext) style text =
    let para = { Id = ""; Buffer = TextBuffer.ofString text; Metadata = { defaultMetadata with Style = style } }
    [| WordEditorPackage.Paragraph (paragraphToEditor context.Styles None para) |]

let private notesWithBodies context noteType (doc: Document) : (WordEditorPackage.Story * Block list option) array =
    let bodies = doc.Package.Notes |> List.filter (fun note -> note.NoteType = noteType)
    let referenced =
        allParagraphs doc |> List.collect (fun para ->
            para.Buffer.Runs |> List.choose (function
                | Text run -> match run.Style.NoteId with Some id when Option.defaultValue "footnote" run.Style.NoteType = noteType -> Some (id, run.Style.NoteText) | _ -> None
                | Image _ -> None)) |> List.distinctBy fst
    let styleId = if noteType = "endnote" then "EndnoteText" else "FootnoteText"
    let style = if doc.Styles |> List.exists (fun s -> s.Id = styleId && s.StyleType = "paragraph") then Some styleId else None
    let make id kind hasMark markStyle text blocks : WordEditorPackage.Story =
        { Id = id; Kind = kind; HasMark = Some hasMark; MarkStyle = markStyle; Text = Some text; Blocks = blocks; SourceStoryCounts = [||]; Unchanged = false }
    let rich (body: NoteBody) = make body.Id body.Kind body.HasMark body.MarkStyle body.Text (body.Blocks |> List.map (blockToEditor context) |> List.toArray), Some body.Blocks
    [| for body in bodies do
           if body.Kind.IsSome then yield rich body
       for id, text in referenced do
           match bodies |> List.tryFind (fun body -> body.Id = id && body.Kind.IsNone) with
           | Some body when text.IsNone || text = Some body.Text -> yield rich body
           | Some body ->
               let text = Option.defaultValue "" text
               yield make id None body.HasMark body.MarkStyle text (plainBodyBlocks context style text), None
           | None ->
               let text = Option.defaultValue "" text
               yield make id None true None text (plainBodyBlocks context style text), None |]

let private commentsWithBodies context (doc: Document) : (WordEditorPackage.CommentPart * Block list option) array =
    let referenced =
        allParagraphs doc |> List.collect (fun para ->
            para.Buffer.Runs |> List.choose (function Text run -> run.Style.CommentId |> Option.map (fun id -> id, run.Style) | Image _ -> None)) |> List.distinctBy fst
    [| for id, style in referenced do
           match doc.Package.Comments |> List.tryFind (fun body -> body.Id = id) with
           | Some body when style.CommentText.IsNone || style.CommentText = Some body.Text ->
               let item: WordEditorPackage.CommentPart =
                   { Id = id; Author = style.CommentAuthor |> Option.orElse body.Author
                     Date = style.CommentDate |> Option.orElse body.Date; Initials = body.Initials; Text = Some body.Text
                     Blocks = body.Blocks |> List.map (blockToEditor context) |> List.toArray; SourceStoryCounts = [||]; Unchanged = false }
               yield item, Some body.Blocks
           | body ->
               let text = Option.defaultValue "" style.CommentText
               let item: WordEditorPackage.CommentPart =
                   { Id = id; Author = style.CommentAuthor; Date = style.CommentDate
                     Initials = body |> Option.bind (fun body -> body.Initials); Text = Some text
                     Blocks = plainBodyBlocks context None text; SourceStoryCounts = [||]; Unchanged = false }
               yield item, None |]

let private listSignatures (doc: Document) =
    let blocks = doc.Blocks @ (doc.HeaderParts @ doc.FooterParts |> List.collect (fun part -> part.Blocks)) @
                 (doc.Package.Notes |> List.collect (fun note -> note.Blocks)) @ (doc.Package.Comments |> List.collect (fun comment -> comment.Blocks))
    paragraphsOfBlocks blocks |> List.fold (fun signatures para ->
        let m = para.Metadata
        match m.ListId with
        | Some id when m.ListStyle <> NoList ->
            let key = id + "/" + string m.ListLevel
            if Map.containsKey key signatures then signatures
            else Map.add key (String.concat "\u0000" [ listStyle m.ListStyle; Option.defaultValue "" m.ListNumberFormat;
                                                        Option.defaultValue "" m.ListLevelText; m.ListStart |> Option.map string |> Option.defaultValue "" ]) signatures
        | _ -> signatures) Map.empty

let exportPackage (doc: Document) : WordEditorPackage.EditorPackage =
    let imported = doc.Package.Imported |> Option.map (fun value -> value.Document)
    let shapes =
        imported |> Option.map (fun original ->
            let blocks = original.Blocks @ (original.HeaderParts @ original.FooterParts |> List.collect (fun part -> part.Blocks))
            blocks |> List.choose (function TextBox shape when shape.SourceXml.IsSome -> Some (shape.Id, shape) | _ -> None) |> Map.ofList)
        |> Option.defaultValue Map.empty
    let context = { Styles = Some (WordStyles.styleContext doc); ImportedShapes = shapes }
    let plainBlocks blocks = blocksInAnchorOrder blocks |> List.map (blockToEditor context) |> List.toArray
    let parts (current: HeaderFooterPart list) (original: HeaderFooterPart list) =
        current |> List.map (fun part ->
            let blocks = plainBlocks part.Blocks
            match part.RelationshipId |> Option.bind (fun id -> original |> List.tryFind (fun candidate -> candidate.RelationshipId = Some id)) with
            | None ->
                ({ RelationshipId = part.RelationshipId; Kind = part.Kind; Blocks = blocks; SourceStoryCounts = [||]; SourceUnchanged = false } : WordEditorPackage.HeaderFooterPart)
            | Some before ->
                let unchanged = obj.ReferenceEquals(part, before) || (obj.ReferenceEquals(part.Blocks, before.Blocks) && part.Kind = before.Kind) ||
                                (part.Kind = before.Kind && blocks = plainBlocks before.Blocks)
                let tagged, counts = if unchanged then blocks, [||] else blocksWithSource context before.Blocks part.Blocks
                let result: WordEditorPackage.HeaderFooterPart =
                    { RelationshipId = part.RelationshipId; Kind = part.Kind; Blocks = tagged
                      SourceStoryCounts = counts; SourceUnchanged = unchanged }
                result) |> List.toArray
    let withNoteSource noteType =
        let current = notesWithBodies context noteType doc
        let before = imported |> Option.map (notesWithBodies context noteType) |> Option.defaultValue [||]
        let oldItems = before |> Array.map fst
        let unchanged = imported.IsSome && (current |> Array.map fst) = oldItems
        let items = current |> Array.map (fun (item, blocks) ->
            let same = oldItems |> Array.exists (fun old -> old.Id = item.Id && old = item)
            let sourceBlocks =
                imported |> Option.bind (fun original ->
                    original.Package.Notes |> List.tryFind (fun note -> note.NoteType = noteType && note.Id = item.Id) |> Option.map (fun note -> note.Blocks))
            match blocks, sourceBlocks with
            | Some now, Some was when not same ->
                let tagged, counts = blocksWithSource context was now
                { item with Blocks = tagged; SourceStoryCounts = counts; Unchanged = same }
            | _ -> { item with Unchanged = same })
        items, unchanged
    let withCommentSource () =
        let current = commentsWithBodies context doc
        let before = imported |> Option.map (commentsWithBodies context) |> Option.defaultValue [||]
        let oldItems = before |> Array.map fst
        let unchanged = imported.IsSome && (current |> Array.map fst) = oldItems
        let items = current |> Array.map (fun (item, blocks) ->
            let same = oldItems |> Array.exists (fun old -> old.Id = item.Id && old = item)
            let sourceBlocks =
                imported |> Option.bind (fun original ->
                    original.Package.Comments |> List.tryFind (fun comment -> comment.Id = item.Id) |> Option.map (fun comment -> comment.Blocks))
            match blocks, sourceBlocks with
            | Some now, Some was when not same ->
                let tagged, counts = blocksWithSource context was now
                { item with Blocks = tagged; SourceStoryCounts = counts; Unchanged = same }
            | _ -> { item with Unchanged = same })
        items, unchanged
    let footnotes, footnotesUnchanged = withNoteSource "footnote"
    let endnotes, endnotesUnchanged = withNoteSource "endnote"
    let comments, commentsUnchanged = withCommentSource ()
    let changedListLevels =
        match imported with
        | None -> [||]
        | Some original ->
            let before = listSignatures original
            listSignatures doc |> Map.toArray |> Array.choose (fun (key, value) ->
                match Map.tryFind key before with Some previous when previous <> value -> Some key | _ -> None)
    let blocks, counts = match imported with Some original -> blocksWithSource context original.Blocks doc.Blocks | None -> plainBlocks doc.Blocks, [||]
    let partIds parts = parts |> List.map (fun (part: HeaderFooterPart) -> part.Kind, part.RelationshipId)
    let pageSetupUnchanged =
        imported |> Option.exists (fun original -> doc.PageSetup = original.PageSetup && partIds doc.HeaderParts = partIds original.HeaderParts && partIds doc.FooterParts = partIds original.FooterParts)
    { Blocks = blocks; Styles = doc.Styles |> List.map styleToEditor |> List.toArray
      HeaderParts = parts doc.HeaderParts (imported |> Option.map (fun original -> original.HeaderParts) |> Option.defaultValue [])
      FooterParts = parts doc.FooterParts (imported |> Option.map (fun original -> original.FooterParts) |> Option.defaultValue [])
      Footnotes = footnotes; Endnotes = endnotes; Comments = comments; PageSetup = Some (pageSetup doc.PageSetup)
      OriginalPageSetup = imported |> Option.map (fun original -> pageSetup original.PageSetup)
      Theme = doc.Package.Theme; Preserved = doc.Package.Preserved; SourceStoryCounts = counts
      ChangedStyles = Some (doc.Styles |> List.filter (fun style -> not (doc.Package.ImportedStyles |> List.contains style)) |> List.map (fun style -> style.StyleType + ":" + style.Id) |> List.toArray)
      ChangedListLevels = changedListLevels; PageSetupUnchanged = pageSetupUnchanged; ListsUnchanged = imported.IsSome && changedListLevels.Length = 0
      FootnotesUnchanged = footnotesUnchanged; EndnotesUnchanged = endnotesUnchanged; CommentsUnchanged = commentsUnchanged
      Unchanged = imported |> Option.exists (fun original -> sameContent doc original); RegenerateAll = false }
