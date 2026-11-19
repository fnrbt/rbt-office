/// Word's outline of a document: the level of each block of the body (Heading 1-9 and paragraphs
/// with an outline level are headings, 1-9; every other paragraph, a table, a text box is body
/// text, Word's level 10), the headings, the subtext of a heading (every block after it up to the
/// next heading of the same or a higher level), and the commands that change the outline: promote,
/// demote, promote to Heading 1, demote to body text, a level; a heading moved with its subtext past
/// its sibling, blocks moved, a heading deleted with its subtext. What the navigation pane, outline
/// view and VBA's OutlinePromote / OutlineDemote / OutlineDemoteToBody / OutlineLevel /
/// Range.Relocate use. Shared by the browser and the native hosts.
///
/// Levels are Word's: a paragraph's own w:outlineLvl, else its style's (inherited along BasedOn);
/// a built-in heading style ("heading 1" .. "heading 9") is its number whatever the paragraph says.
/// The body is flat: paragraphs in table cells, text boxes, headers, footers and notes are not
/// headings of the outline (a table is a block of body text under the heading before it).
module Rbt.Word.Outline

open System
open System.Collections.Generic
open Rbt.Word.Document
module TB = Rbt.Word.TextBuffer

/// Word's level of body text (wdOutlineLevelBodyText).
[<Literal>]
let BodyLevel = 10

/// The number of one of Word's built-in heading styles - "heading 1" .. "heading 9" by name, or,
/// for a style without a name, "Heading1" .. "Heading9" by id - else 0.
let builtInHeadingNumber (styleId: string) (name: string option) =
    let numbered (text: string) (prefix: int) =
        if text.Length = prefix + 1 && Char.IsDigit text.[prefix] && text.[prefix] <> '0' then int text.[prefix] - int '0' else 0
    match name with
    | Some name when name.StartsWith("heading ", StringComparison.OrdinalIgnoreCase) -> numbered name 8
    | Some _ -> 0
    | None -> if styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) then numbered styleId 7 else 0

let private pxOfPt (pt: float) = pt / 0.75

/// Word's built-in Heading 1 .. Heading 9 (2016 and later), for a document that does not define the
/// one a command applies: it is added to the document's styles and saved with it.
let builtInHeadingStyle (level: int) : NamedStyle =
    let size, color, italic =
        match level with
        | 1 -> 16.0, "#2F5496", false
        | 2 -> 13.0, "#2F5496", false
        | 3 -> 12.0, "#1F3763", false
        | 4 -> 11.0, "#2F5496", true
        | 5 -> 11.0, "#2F5496", false
        | 6 -> 11.0, "#1F3763", false
        | 7 -> 11.0, "#1F3763", true
        | 8 -> 10.5, "#272727", false
        | _ -> 10.5, "#272727", true
    let text = TB.withExplicitFontSize (pxOfPt size) { TB.defaultStyle with FontFamily = Some "Calibri Light"; Color = Some color }
    { Id = sprintf "Heading%d" level; Name = Some (sprintf "heading %d" level); StyleType = "paragraph"; IsDefault = false
      BasedOn = Some "Normal"; Next = Some "Normal"; Linked = Some (sprintf "Heading%dChar" level); IsQuickStyle = true
      Aliases = None; UiPriority = Some 9; SemiHidden = level > 1; UnhideWhenUsed = level > 1
      Text = Some text; TextBold = None; TextItalic = (if italic then Some true else None); TextUnderline = None
      TextAllCaps = None; TextSmallCaps = None; TextStrike = None
      SpacingBefore = Some (if level = 1 then 12.0 else 2.0); SpacingAfter = Some 0.0; ContextualSpacing = None
      LineSpacing = None; LineSpacingRule = None; Alignment = None; IndentLeft = None; IndentRight = None
      FirstLineIndent = None; BackgroundColor = None; OutlineLevel = Some (level - 1)
      KeepNext = Some true; KeepLines = Some true; PageBreakBefore = None; WidowControl = None }

/// Where a paragraph's level comes from: what a promote or a demote changes.
type LevelSource =
    /// A built-in heading style (or a heading made here without one): another Heading N.
    | HeadingStyle = 0
    /// The paragraph's own w:outlineLvl: changed.
    | Direct = 1
    /// The outline level of a style of the document's own: the paragraph is given its own.
    | StyleLevel = 2
    /// Body text.
    | Body = 3

/// The levels of the paragraphs of one state of the styles: each style's looked up once.
type LevelResolver(doc: Document) =
    let context = WordStyles.styleContext doc
    let styleLevels = Dictionary<string, int>()
    let styleBuiltIn = Dictionary<string, bool>()
    let resolve (styleId: string) =
        match Map.tryFind styleId context.ResolvedParagraphStyles with
        | Some style ->
            // (The name a style has is its own, not the style it is based on: a resolved style keeps it.)
            let own = Map.tryFind ("paragraph", styleId) context.Styles |> Option.bind (fun s -> s.Name) |> Option.orElse style.Name
            let number = builtInHeadingNumber styleId own
            if number > 0 then number, true
            else
                match style.OutlineLevel with
                | Some level when level >= 0 && level <= 8 -> level + 1, false
                | _ -> BodyLevel, false
        | None ->
            // A style the document does not define: Word's own heading, by its id.
            let number = builtInHeadingNumber styleId None
            if number > 0 then number, true else BodyLevel, false
    member _.Context = context
    /// The level of a paragraph style (BodyLevel for one without a level) and whether it is a built-in heading.
    member _.Style(styleId: string) : int * bool =
        match Portable.Dict.tryFind styleId styleLevels with
        | ValueSome level -> level, (match Portable.Dict.tryFind styleId styleBuiltIn with ValueSome b -> b | ValueNone -> false)
        | ValueNone ->
            let level, builtIn = resolve styleId
            styleLevels.[styleId] <- level
            styleBuiltIn.[styleId] <- builtIn
            level, builtIn
    /// The level of a paragraph and where it comes from. `headingBlock`: the level of the Heading
    /// block that holds it (0: none).
    member this.Paragraph(para: ParagraphBlock, headingBlock: int) : struct (int * LevelSource) =
        let styleId =
            match para.Metadata.Style with
            | Some id -> Some id
            | None -> context.DefaultParagraphStyleId
        let styleLevel, builtIn =
            match styleId with
            | Some id -> this.Style id
            | None -> BodyLevel, false
        if builtIn then struct (styleLevel, LevelSource.HeadingStyle)
        elif headingBlock > 0 && para.Metadata.Style.IsNone then struct (headingBlock, LevelSource.HeadingStyle)
        else
            match para.Metadata.OutlineLevel with
            | Some level when level >= 0 && level <= 8 -> struct (level + 1, LevelSource.Direct)
            | Some _ -> struct (BodyLevel, LevelSource.Direct)
            | None when styleLevel < BodyLevel -> struct (styleLevel, LevelSource.StyleLevel)
            | None -> struct (BodyLevel, LevelSource.Body)
    member this.Block(block: Block) : int =
        match block with
        | Paragraph p -> let struct (level, _) = this.Paragraph(p, 0) in level
        | Heading (level, p) -> let struct (level, _) = this.Paragraph(p, level) in level
        | Table _ | TextBox _ -> BodyLevel

/// The outline of one state of a document's body.
type DocumentOutline = {
    /// The body's blocks and styles it is of.
    Blocks: Block list
    Styles: NamedStyle list
    /// Each block's level, by its place in the body (BodyLevel for body text, tables, text boxes).
    Levels: int array
    BlockIds: BlockId array
    /// The places of the headings, in order, and their text (see lineText).
    Headings: int array
    HeadingTexts: string array
}

let private blockId (block: Block) =
    match block with
    | Paragraph p | Heading (_, p) -> p.Id
    | Table t -> t.Id
    | TextBox t -> t.Id

/// The text of a paragraph as a line of the outline: tabs and breaks as spaces.
let lineText (para: ParagraphBlock) =
    let text = TB.toString para.Buffer
    let chars = text.ToCharArray()
    let mutable changed = false
    for i = 0 to chars.Length - 1 do
        let c = chars.[i]
        if c < ' ' then
            chars.[i] <- ' '
            changed <- true
    (if changed then String(chars) else text).Trim()

let private compute (doc: Document) =
    let resolver = LevelResolver doc
    let levels = ResizeArray<int>()
    let ids = ResizeArray<BlockId>()
    let headings = ResizeArray<int>()
    let texts = ResizeArray<string>()
    let mutable rest = doc.Blocks
    while not rest.IsEmpty do
        let block = rest.Head
        let level = resolver.Block block
        if level < BodyLevel then
            headings.Add levels.Count
            texts.Add (match block with Paragraph p | Heading (_, p) -> lineText p | _ -> "")
        levels.Add level
        ids.Add (blockId block)
        rest <- rest.Tail
    { Blocks = doc.Blocks; Styles = doc.Styles; Levels = levels.ToArray(); BlockIds = ids.ToArray(); Headings = headings.ToArray(); HeadingTexts = texts.ToArray() }

/// The outline last computed (by the blocks and styles it is of: the same lists from one query to
/// the next while nothing changes).
let mutable private lastOutline : DocumentOutline option = None

/// The outline of the document's body.
let ofDocument (doc: Document) : DocumentOutline =
    match lastOutline with
    | Some known when obj.ReferenceEquals(known.Blocks, doc.Blocks) && obj.ReferenceEquals(known.Styles, doc.Styles) -> known
    | _ ->
        let outline = compute doc
        lastOutline <- Some outline
        outline

/// The end (exclusive) of the subtext of the block at `index`: the place of the next heading of
/// its level or a higher one (the end of the body when there is none). A block of body text has
/// no subtext (index + 1).
let subtextEnd (outline: DocumentOutline) (index: int) =
    let levels = outline.Levels
    if index < 0 || index >= levels.Length then index + 1
    else
        let level = levels.[index]
        if level >= BodyLevel then index + 1
        else
            let mutable j = index + 1
            while j < levels.Length && levels.[j] > level do j <- j + 1
            j

/// The place of the heading a block at `index` is under (the heading itself for a heading); -1
/// before the first heading.
let headingAt (outline: DocumentOutline) (index: int) =
    let headings = outline.Headings
    let mutable lo = 0
    let mutable hi = headings.Length - 1
    let mutable found = -1
    while lo <= hi do
        let mid = lo + (hi - lo) / 2
        if headings.[mid] <= index then
            found <- headings.[mid]
            lo <- mid + 1
        else hi <- mid - 1
    found

/// The level of the heading before the block at `index` (1 when there is none: what body text
/// promoted there becomes, as in Word).
let private previousHeadingLevel (outline: DocumentOutline) (index: int) =
    let at = headingAt outline (index - 1)
    if at < 0 then 1 else outline.Levels.[at]

/// The place of a block of the body, by its id or the id of what it holds (a paragraph of a
/// table cell is in its table).
let indexOf (doc: Document) (id: BlockId) =
    match locateBlock doc id with
    | ValueSome (struct (index, _)) -> index
    | ValueNone -> -1

/// The body block at `index` as a paragraph (None for a table or a text box).
let paragraphAt (doc: Document) (index: int) : ParagraphBlock option =
    match List.tryItem index doc.Blocks with
    | Some (Paragraph p) | Some (Heading (_, p)) -> Some p
    | _ -> None

// ---- levels changed ------------------------------------------------------------------------------

/// A change of level: Word's Promote, Demote, Promote to Heading 1, Demote to Body Text; a level of
/// the outlining toolbar's box (1-9, or 10 for body text); the Paragraph.OutlineLevel of VBA (the
/// paragraph's own level: a built-in heading keeps its style's).
type LevelChange =
    | Promote
    | Demote
    | PromoteToHeading1
    | DemoteToBody
    | ToLevel of int
    | OwnLevel of int

/// The paragraph style the document has for Heading `level`, and the styles with it (Word's built-in
/// one added when the document has none).
let headingStyleFor (level: int) (doc: Document) : string * NamedStyle list =
    let existing =
        doc.Styles |> List.tryFind (fun s ->
            s.StyleType = "paragraph" && builtInHeadingNumber s.Id s.Name = level)
    match existing with
    | Some style -> style.Id, doc.Styles
    | None ->
        let style = builtInHeadingStyle level
        // (An id the document gives another style is not taken.)
        let taken = doc.Styles |> List.exists (fun s -> s.Id = style.Id)
        let style = if taken then { style with Id = style.Id + "Builtin" } else style
        style.Id, doc.Styles @ [ style ]

/// Direct paragraph formatting a paragraph style takes the place of (Demote to Body Text, as Word).
let private withoutParagraphFormatting (metadata: BlockMetadata) =
    let flags =
        ParagraphExplicit.SpacingBefore ||| ParagraphExplicit.SpacingAfter ||| ParagraphExplicit.LineSpacing
        ||| ParagraphExplicit.IndentLeft ||| ParagraphExplicit.IndentRight ||| ParagraphExplicit.FirstLineIndent
        ||| ParagraphExplicit.ContextualSpacing
    { metadata with
        SpacingBefore = None; SpacingAfter = None; LineSpacing = None; LineSpacingRule = None
        IndentLeft = None; IndentRight = None; FirstLineIndent = None; ContextualSpacing = None
        KeepNext = None; KeepLines = None
        Explicit = metadata.Explicit &&& ~~~flags }

/// What becomes of one paragraph.
type private Target =
    | Unchanged
    | HeadingLevel of int
    | BodyText
    | OwnOutline of int

/// Change the level of the paragraphs with these ids (of the body, or of a table cell: there the
/// style changes, as in Word, though such a paragraph is no heading of the outline). One new
/// document; the same document when nothing changes.
let changeLevels (change: LevelChange) (ids: BlockId list) (doc: Document) : Document =
    if ids.IsEmpty then doc
    else
    let outline = ofDocument doc
    let resolver = LevelResolver doc
    // (Each paragraph with its level, where it comes from, and its place in the body.)
    let targets =
        ids |> List.choose (fun id ->
            match locateBlock doc id with
            | ValueNone -> None
            | ValueSome (struct (index, block)) ->
                match block with
                | Paragraph p when p.Id = id -> let struct (l, s) = resolver.Paragraph(p, 0) in Some (p, index, l, s, true)
                | Heading (h, p) when p.Id = id -> let struct (l, s) = resolver.Paragraph(p, h) in Some (p, index, l, s, true)
                | _ ->
                    match findParagraph doc id with
                    | Some p -> let struct (l, s) = resolver.Paragraph(p, 0) in Some (p, index, l, s, false)
                    | None -> None)
    // (Body text among headings is left as it is: Word promotes body text only when no heading is with it.)
    let anyHeading = targets |> List.exists (fun (_, _, level, _, _) -> level < BodyLevel)
    let decide (_: ParagraphBlock, index: int, level: int, source: LevelSource, _: bool) =
        let byLevel (wanted: int) =
            match source with
            | LevelSource.Direct | LevelSource.StyleLevel -> OwnOutline wanted
            | _ -> if wanted >= BodyLevel then BodyText else HeadingLevel wanted
        match change with
        | Promote ->
            if level < BodyLevel then (if level > 1 then byLevel (level - 1) else Unchanged)
            elif anyHeading then Unchanged
            else HeadingLevel (previousHeadingLevel outline index)
        | Demote ->
            if level < BodyLevel then (if level < 9 then byLevel (level + 1) else Unchanged)
            elif anyHeading then Unchanged
            else HeadingLevel (min 9 (previousHeadingLevel outline index + 1))
        | PromoteToHeading1 -> HeadingLevel 1
        | DemoteToBody -> BodyText
        | ToLevel wanted ->
            let wanted = max 1 (min BodyLevel wanted)
            if wanted = level then Unchanged
            else
                match source with
                | LevelSource.Direct | LevelSource.StyleLevel -> OwnOutline wanted
                | _ -> if wanted >= BodyLevel then BodyText else HeadingLevel wanted
        | OwnLevel wanted ->
            let wanted = max 1 (min BodyLevel wanted)
            if source = LevelSource.HeadingStyle || wanted = level then Unchanged else OwnOutline wanted
    let decisions = targets |> List.map (fun target -> target, decide target)
    if decisions |> List.forall (fun (_, d) -> d = Unchanged) then doc
    else
    // (Styles the change needs that the document has not got are added first.)
    let mutable styles = doc.Styles
    let headingIds = Dictionary<int, string>()
    for _, decision in decisions do
        match decision with
        | HeadingLevel level when not (headingIds.ContainsKey level) ->
            let id, withStyle = headingStyleFor level { doc with Styles = styles }
            styles <- withStyle
            headingIds.[level] <- id
        | _ -> ()
    let context = resolver.Context
    let defaultStyle = context.DefaultParagraphStyleId
    // The style's own level without the paragraph's: what a paragraph's own level of body text is against.
    let styleLevelOf (para: ParagraphBlock) =
        match (match para.Metadata.Style with Some id -> Some id | None -> defaultStyle) with
        | Some id -> fst (resolver.Style id)
        | None -> BodyLevel
    let replaced = Dictionary<BlockId, ParagraphBlock * int>()
    for (para, _, _, _, _), decision in decisions do
        match decision with
        | Unchanged -> ()
        | HeadingLevel level ->
            let styleId = headingIds.[level]
            let restyled = WordStyles.restyleParagraph context doc (Some styleId) para
            let restyled = { restyled with Metadata = { restyled.Metadata with OutlineLevel = None } }
            replaced.[para.Id] <- (restyled, level)
        | BodyText ->
            let restyled = WordStyles.restyleParagraph context doc defaultStyle para
            let restyled = { restyled with Metadata = { withoutParagraphFormatting restyled.Metadata with OutlineLevel = None } }
            replaced.[para.Id] <- (restyled, 0)
        | OwnOutline level ->
            let own =
                if level < BodyLevel then Some (level - 1)
                elif styleLevelOf para < BodyLevel then Some 9
                else None
            replaced.[para.Id] <- ({ para with Metadata = { para.Metadata with OutlineLevel = own } }, -1)
    // The body's blocks in one pass (a Heading block for a heading style, as the reader makes them);
    // a paragraph of a table cell in its table.
    let mutable remaining = replaced.Count
    let blocks =
        doc.Blocks |> List.map (fun block ->
            if remaining = 0 then block
            else
                match block with
                | Paragraph p | Heading (_, p) ->
                    match Portable.Dict.tryFind p.Id replaced with
                    | ValueSome (para, level) ->
                        remaining <- remaining - 1
                        replaced.Remove p.Id |> ignore
                        if level > 0 then Heading (min 3 level, para)
                        elif level = 0 then Paragraph para
                        else (match block with Heading (h, _) -> Heading (h, para) | _ -> Paragraph para)
                    | ValueNone -> block
                | _ -> block)
    let mutable result = { doc with Blocks = blocks; Styles = styles }
    for KeyValue (id, (para, _)) in replaced |> Seq.toArray do
        result <- replaceParagraph result id para
    result

/// The paragraphs of the body the selection is in, from the block its start is in to the block its
/// end is in (a table among them is not a paragraph; a selection in a table cell has none).
let selectedBodyParagraphs (doc: Document) : BlockId list =
    match doc.Selection with
    | None -> []
    | Some selection ->
        let a = indexOf doc selection.Anchor.BlockId
        let f = indexOf doc selection.Focus.BlockId
        if a < 0 || f < 0 then []
        else
            let first, last = min a f, max a f
            doc.Blocks
            |> List.skip first
            |> List.truncate (last - first + 1)
            |> List.choose (function
                | Paragraph p | Heading (_, p) -> Some p.Id
                | _ -> None)

/// The levels of the selection's paragraphs of the body changed (see changeLevels).
let changeSelectedLevels (change: LevelChange) (doc: Document) : Document =
    changeLevels change (selectedBodyParagraphs doc) doc

/// A heading and its subheadings (its subtext's headings) promoted or demoted together, as the
/// navigation pane does: each by one level, none above Heading 1 or below Heading 9 (the branch
/// is not changed when one of them cannot go).
let changeBranch (promote: bool) (id: BlockId) (doc: Document) : Document =
    let outline = ofDocument doc
    let index = indexOf doc id
    if index < 0 || outline.Levels.[index] >= BodyLevel then doc
    else
        let last = subtextEnd outline index
        let members = [ for i in index .. last - 1 do if outline.Levels.[i] < BodyLevel then yield i ]
        let canMove = members |> List.forall (fun i -> if promote then outline.Levels.[i] > 1 else outline.Levels.[i] < 9)
        if not canMove then doc
        else changeLevels (if promote then Promote else Demote) [ for i in members -> outline.BlockIds.[i] ] doc

// ---- blocks moved ----------------------------------------------------------------------------------

/// The blocks of the body at [first, last) moved to stand before the block at `before` (the
/// number of blocks: at the end). The same document when that is where they are.
let moveBlocks (first: int) (last: int) (before: int) (doc: Document) : Document =
    let blocks = List.toArray doc.Blocks
    let n = blocks.Length
    if first < 0 || last > n || first >= last || before < 0 || before > n || (before >= first && before <= last) then doc
    else
        let moved = blocks.[first .. last - 1]
        let rest = Array.append blocks.[0 .. first - 1] blocks.[last ..]
        let at = if before > last then before - (last - first) else before
        let result = Array.concat [ rest.[0 .. at - 1]; moved; rest.[at ..] ]
        { doc with Blocks = List.ofArray result }

/// Where a heading (with its subtext) moves to past its previous or next sibling - the heading of
/// the same level before it or after its subtext, under the same heading: (first, last, before),
/// or None where it has no such sibling. A block of body text moves past the block before or
/// after it.
let siblingMove (up: bool) (index: int) (doc: Document) : (int * int * int) option =
    let outline = ofDocument doc
    let levels = outline.Levels
    let n = levels.Length
    if index < 0 || index >= n then None
    else
        let level = levels.[index]
        let last = subtextEnd outline index
        if level >= BodyLevel then
            if up then (if index > 0 then Some (index, index + 1, index - 1) else None)
            else (if index + 1 < n then Some (index, index + 1, index + 2) else None)
        elif up then
            let mutable j = index - 1
            while j >= 0 && levels.[j] > level do j <- j - 1
            if j >= 0 && levels.[j] = level then Some (index, last, j) else None
        else
            if last < n && levels.[last] = level then Some (index, last, subtextEnd outline last) else None

/// The heading (with its subtext) or the block of body text with this id moved past its previous
/// or next sibling (see siblingMove).
let moveHeading (up: bool) (id: BlockId) (doc: Document) : Document =
    match siblingMove up (indexOf doc id) doc with
    | Some (first, last, before) -> moveBlocks first last before doc
    | None -> doc

/// The selection's blocks of the body moved past the block before or after them (Word's
/// Alt+Shift+Up / Down in print layout, Range.Relocate).
let moveSelectedBlocks (up: bool) (doc: Document) : Document =
    match doc.Selection with
    | None -> doc
    | Some selection ->
        let a = indexOf doc selection.Anchor.BlockId
        let f = indexOf doc selection.Focus.BlockId
        if a < 0 || f < 0 then doc
        else
            let first, last = min a f, max a f + 1
            if up then (if first > 0 then moveBlocks first last (first - 1) doc else doc)
            else (if last < List.length doc.Blocks then moveBlocks first last (last + 1) doc else doc)

/// The caret at the start of the block at `index` (a table's first paragraph), else at the start of
/// the block before it.
let private caretNear (index: int) (doc: Document) =
    let blocks = List.toArray doc.Blocks
    let paragraphOf (block: Block) = paragraphsOfBlocks [ block ] |> List.tryHead
    let at (i: int) = if i >= 0 && i < blocks.Length then (match blocks.[i] with TextBox _ -> None | b -> paragraphOf b) else None
    let mutable target = None
    let mutable i = index
    while target.IsNone && i < blocks.Length do
        target <- at i
        i <- i + 1
    i <- index - 1
    while target.IsNone && i >= 0 do
        target <- at i
        i <- i - 1
    match target with
    | Some para -> setCursor para.Id 0 doc
    | None -> doc

/// A heading and its subtext deleted (the navigation pane's Delete); the caret where they stood.
/// A document is not left without a block: an empty paragraph stays.
let deleteHeading (id: BlockId) (doc: Document) : Document =
    let outline = ofDocument doc
    let index = indexOf doc id
    if index < 0 || outline.Levels.[index] >= BodyLevel then doc
    else
        let last = subtextEnd outline index
        let blocks = List.toArray doc.Blocks
        let kept = Array.append blocks.[0 .. index - 1] blocks.[last ..]
        // (The section properties of the body's last paragraph stay with the body.)
        let lastBlock = blocks.[blocks.Length - 1]
        let keptBlocks, at =
            if kept.Length = 0 then
                let empty = newEmptyParagraph ()
                let metadata = match lastBlock with Paragraph p | Heading (_, p) -> { defaultMetadata with SectionPageSetup = p.Metadata.SectionPageSetup; SectionBreak = p.Metadata.SectionBreak; HeaderReferences = p.Metadata.HeaderReferences; FooterReferences = p.Metadata.FooterReferences } | _ -> defaultMetadata
                [| Paragraph { empty with Metadata = metadata } |], 0
            else kept, index
        caretNear at { doc with Blocks = List.ofArray keptBlocks }

/// The caret at the start of the heading (or block) with this id.
let goTo (id: BlockId) (doc: Document) : Document =
    let index = indexOf doc id
    if index < 0 then doc else caretNear index doc
