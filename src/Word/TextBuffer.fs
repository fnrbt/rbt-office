module Rbt.Word.TextBuffer

open Rbt.Word.Rope

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

/// Id, author and date of a tracked change.
type RevisionMark = {
    Id: string option
    Author: string option
    Date: string option
}

/// Text style for a run
type TextStyle = {
    Style: string option
    Bold: bool
    Italic: bool
    Underline: bool
    /// w:u@val — single|double|thick|dotted|dash|dotDash|dotDotDash|wave|wavyDouble|… (None = single).
    UnderlineStyle: string option
    /// w:u@color — distinct underline colour (None = text colour).
    UnderlineColor: string option
    AllCaps: bool
    SmallCaps: bool
    Strike: bool
    DoubleStrike: bool          // w:dstrike
    Hidden: bool                // w:vanish — not rendered, takes no space
    Position: float             // w:position — baseline raise(+)/lower(-) in points
    EmphasisMark: string option // w:em@val — dot|comma|circle|underDot
    FontSize: float  // in points
    CharacterSpacing: float
    FontFamily: string option
    Color: string option
    Highlight: string option
    /// The target of a link that may be followed: http, https, mailto, tel or an
    /// anchor inside the document ("#bookmark"). See isActiveHyperlinkTarget.
    Hyperlink: string option
    /// The target of a hyperlink of the file that is not followed (javascript:,
    /// file:, data:, a relative path, ...): kept so that the document saves as it
    /// was, never navigated.
    InactiveHyperlink: string option
    CommentId: string option
    CommentText: string option
    CommentAuthor: string option
    CommentDate: string option
    NoteId: string option
    NoteType: string option
    NoteText: string option
    BookmarkStartId: string option
    BookmarkName: string option
    BookmarkEndId: string option
    FieldCode: string option
    FieldResult: string option
    RenderedPageBreakBefore: bool
    EquationXml: string option
    RevisionId: string option
    RevisionType: string option
    RevisionAuthor: string option
    RevisionDate: string option
    /// Deleted text that was itself a tracked insertion (w:del inside w:ins): the
    /// insertion it sits in. RevisionType is "delete" and the Revision* fields
    /// describe the deletion. Accepting everything deletes the text, and so does
    /// rejecting everything (the insertion never happened).
    DeletedInsertion: RevisionMark option
    /// Name of the tracked move (w:moveFromRangeStart / w:moveToRangeStart) a run
    /// of RevisionType "moveFrom" / "moveTo" belongs to; it pairs the two ends.
    MoveName: string option
    VertAlign: string option
    /// Bit set of ExplicitFormat flags: the on/off properties and the font size
    /// that were set directly on this run. A flagged property overrides the
    /// paragraph / character style even when its value equals the default
    /// (non-bold under a bold style, an 11-unit run under a larger style). An
    /// unflagged property at its default value is "unset" and inherits.
    Explicit: int
}

/// Flags for TextStyle.Explicit.
module ExplicitFormat =
    [<Literal>]
    let Bold = 1
    [<Literal>]
    let Italic = 2
    [<Literal>]
    let Underline = 4
    [<Literal>]
    let AllCaps = 8
    [<Literal>]
    let SmallCaps = 16
    [<Literal>]
    let Strike = 32
    [<Literal>]
    let DoubleStrike = 64
    [<Literal>]
    let Hidden = 128
    [<Literal>]
    let FontSize = 256
    /// The run states "automatic" colour (w:color w:val="auto"): Color is None,
    /// and a colour from a style does not apply.
    [<Literal>]
    let ColorAuto = 512

    let has (flag: int) (style: TextStyle) = style.Explicit &&& flag <> 0

/// The colour the reader gives the mark of a comment whose run has none of its
/// own (DocxReader.commentMarkDisplayColor): of the display, not of the file.
let commentMarkDisplayColor = "#7c3aed"

/// Default text style
let defaultStyle = {
    Style = None
    Bold = false
    Italic = false
    Underline = false
    UnderlineStyle = None
    UnderlineColor = None
    AllCaps = false
    SmallCaps = false
    Strike = false
    DoubleStrike = false
    Hidden = false
    Position = 0.0
    EmphasisMark = None
    FontSize = 11.0  // 11pt default (like Word)
    CharacterSpacing = 0.0
    FontFamily = None
    Color = None
    Highlight = None
    Hyperlink = None
    InactiveHyperlink = None
    CommentId = None
    CommentText = None
    CommentAuthor = None
    CommentDate = None
    NoteId = None
    NoteType = None
    NoteText = None
    BookmarkStartId = None
    BookmarkName = None
    BookmarkEndId = None
    FieldCode = None
    FieldResult = None
    RenderedPageBreakBefore = false
    EquationXml = None
    RevisionId = None
    RevisionType = None
    RevisionAuthor = None
    RevisionDate = None
    DeletedInsertion = None
    MoveName = None
    VertAlign = None
    Explicit = 0
}

/// Whether a hyperlink target may become an active link: http, https, mailto,
/// tel, or an anchor inside the document. Everything else (javascript:,
/// vbscript:, file:, data:, relative paths, anything with a scheme not listed)
/// is kept as text and never followed.
let isActiveHyperlinkTarget (target: string) =
    if System.String.IsNullOrWhiteSpace target then false
    else
        let lower = target.ToLowerInvariant()
        lower.StartsWith "http://" || lower.StartsWith "https://" || lower.StartsWith "mailto:" || lower.StartsWith "tel:"
        || (lower.StartsWith "#" && lower.Length > 1)

/// Pictures that are linked rather than embedded have a URL the document
/// chose as their source. They are not loaded (a placeholder is shown) unless
/// the host opts in: opening a document must not make requests to an address
/// the document names.
let mutable allowLinkedImages = false

/// Tracked text that is shown as removed: a deletion, or the source of a move.
let isRemovalRevision (style: TextStyle) =
    style.RevisionType = Some "delete" || style.RevisionType = Some "moveFrom"

/// Tracked text that is shown as added: an insertion, or the destination of a move.
let isAdditionRevision (style: TextStyle) =
    style.RevisionType = Some "insert" || style.RevisionType = Some "moveTo"

/// Set on/off character properties directly on a run: the value is recorded as
/// explicit, so "off" survives under a style that turns the property on.
let withExplicitBold on (style: TextStyle) = { style with Bold = on; Explicit = style.Explicit ||| ExplicitFormat.Bold }
let withExplicitItalic on (style: TextStyle) = { style with Italic = on; Explicit = style.Explicit ||| ExplicitFormat.Italic }
let withExplicitUnderline on (style: TextStyle) = { style with Underline = on; Explicit = style.Explicit ||| ExplicitFormat.Underline }
/// Whether text is shown underlined: when it is underlined, and when it is a
/// link that may be followed - unless its text says that it is not underlined
/// (an underline switched off in the editor, <w:u w:val="none"/> in a file).
let underlineShown (style: TextStyle) =
    style.Underline || (style.Hyperlink.IsSome && (style.Explicit &&& ExplicitFormat.Underline) = 0)
let withExplicitAllCaps on (style: TextStyle) =
    { style with AllCaps = on; SmallCaps = false; Explicit = style.Explicit ||| ExplicitFormat.AllCaps ||| ExplicitFormat.SmallCaps }
let withExplicitSmallCaps on (style: TextStyle) =
    { style with SmallCaps = on; AllCaps = false; Explicit = style.Explicit ||| ExplicitFormat.AllCaps ||| ExplicitFormat.SmallCaps }
let withExplicitStrike on (style: TextStyle) = { style with Strike = on; Explicit = style.Explicit ||| ExplicitFormat.Strike }
let withExplicitDoubleStrike on (style: TextStyle) = { style with DoubleStrike = on; Explicit = style.Explicit ||| ExplicitFormat.DoubleStrike }
let withExplicitHidden on (style: TextStyle) = { style with Hidden = on; Explicit = style.Explicit ||| ExplicitFormat.Hidden }
let withExplicitFontSize size (style: TextStyle) = { style with FontSize = size; Explicit = style.Explicit ||| ExplicitFormat.FontSize }

let characterSpacingSlots (text: string) =
    if System.String.IsNullOrEmpty text then 0.0
    else
        let mutable count = 0
        let mutable last = '\000'
        for i = 0 to text.Length - 1 do
            let ch = text.[i]
            if ch <> '\t' && ch <> '\r' && ch <> '\n' then
                count <- count + 1
                last <- ch
        if count = 0 then 0.0
        elif System.Char.IsWhiteSpace last then float count
        else float (count - 1)

let characterSpacingAdvance (text: string) (style: TextStyle) =
    style.CharacterSpacing * characterSpacingSlots text

/// An object of the file the editor does not model (chart, SmartArt, OLE
/// object, grouped drawing, ...): the element as it was read, written back
/// unchanged on save. It sits in the text as an image run that shows the
/// object's fallback picture or a placeholder frame.
type OpaqueObject = {
    /// The element, with the namespace declarations it needs on its root.
    Xml: string
    /// chart | diagram | ole | shape | drawing
    Kind: string
    /// A child of the paragraph (it holds its own runs) rather than of a run.
    ParagraphLevel: bool
}

/// Src of an opaque object without a picture: "opaque:<kind>".
let opaqueSourcePrefix = "opaque:"

/// Image data for inline images
type ImageRun = {
    Start: int           // Character offset (position in text)
    Src: string          // Data URL or image source
    Linked: bool
    Width: float option  // Width in pixels (None = intrinsic)
    Height: float option // Height in pixels (None = intrinsic)
    Alt: string option   // Alt text/description
    Name: string option
    Title: string option
    Floating: bool       // Imported floating/anchored image metadata
    Wrap: string option  // DOCX wrap mode when Floating is true
    PositionX: float option
    PositionY: float option
    PositionHRelativeFrom: string option
    PositionVRelativeFrom: string option
    BehindDoc: bool option
    AllowOverlap: bool option
    LayoutInCell: bool option
    DistanceTop: float option
    DistanceBottom: float option
    DistanceLeft: float option
    DistanceRight: float option
    RelativeHeight: string option
    Locked: bool option
    SimplePosition: bool option
    SimplePositionX: float option
    SimplePositionY: float option
    CropLeft: string option
    CropTop: string option
    CropRight: string option
    CropBottom: string option
    /// wp:align of wp:positionH / wp:positionV (left, center, right, top, ...): the
    /// image is placed by alignment instead of PositionX / PositionY.
    PositionHAlign: string option
    PositionVAlign: string option
    /// Which sides text wraps on (bothSides, left, right, largest).
    WrapText: string option
    /// The outline text wraps along with "tight" and "through" wrapping
    /// (wp:wrapPolygon): its points in the file's units, 21600 across the
    /// picture's width and down its height. Empty: the picture's box.
    WrapPolygon: (float * float) list
    /// Set when this is not a picture but a preserved object of the file.
    Opaque: OpaqueObject option
}

/// Text run with consistent styling
type TextRun = {
    Start: int       // Character offset start
    Length: int      // Number of characters
    Style: TextStyle
}

/// Run can be either text or image
type Run =
    | Text of TextRun
    | Image of ImageRun

/// Text buffer with formatting information
type TextBuffer = {
    Rope: Rope
    Runs: Run list  // Non-overlapping runs covering the text (images use U+FFFC placeholder)
}

and FormatChange =
    | Bold
    | NoBold
    | Italic
    | NoItalic
    | Underline
    | NoUnderline
    | SetColor of string
    | ClearColor
    | SetFontSize of float

/// Src of a linked picture that is not loaded (see allowLinkedImages).
let linkedImagePlaceholderSource = opaqueSourcePrefix + "linked"

/// What an image run shows: its picture, or for a linked picture the host has
/// not allowed, a placeholder.
let displayedImageSource (image: ImageRun) =
    if image.Linked && image.Opaque.IsNone && not allowLinkedImages then linkedImagePlaceholderSource else image.Src

// Helper functions for Run type
module Run =
    let start = function
        | Text tr -> tr.Start
        | Image ir -> ir.Start

    let length = function
        | Text tr -> tr.Length
        | Image _ -> 1  // Images occupy 1 character (U+FFFC placeholder)

    let endPos run = start run + length run

    let shiftStart offset = function
        | Text tr -> Text { tr with Start = tr.Start + offset }
        | Image ir -> Image { ir with Start = ir.Start + offset }

/// Whether two styles are equal: what `left = right` says, field by field.
/// (The record's own equality gathers the forty fields of both sides before
/// it compares the first - most of the time of a format given to a whole long
/// document; here the fields that differ most often come first. A test goes
/// through the record's fields and fails for one that is not compared here.)
let sameTextStyle (left: TextStyle) (right: TextStyle) : bool =
    obj.ReferenceEquals(left, right)
    || (left.Bold = right.Bold
        && left.Italic = right.Italic
        && left.Underline = right.Underline
        && left.AllCaps = right.AllCaps
        && left.SmallCaps = right.SmallCaps
        && left.Strike = right.Strike
        && left.DoubleStrike = right.DoubleStrike
        && left.Hidden = right.Hidden
        && left.Position = right.Position
        && left.FontSize = right.FontSize
        && left.CharacterSpacing = right.CharacterSpacing
        && left.RenderedPageBreakBefore = right.RenderedPageBreakBefore
        && left.Explicit = right.Explicit
        && left.Style = right.Style
        && left.UnderlineStyle = right.UnderlineStyle
        && left.UnderlineColor = right.UnderlineColor
        && left.EmphasisMark = right.EmphasisMark
        && left.FontFamily = right.FontFamily
        && left.Color = right.Color
        && left.Highlight = right.Highlight
        && left.Hyperlink = right.Hyperlink
        && left.InactiveHyperlink = right.InactiveHyperlink
        && left.CommentId = right.CommentId
        && left.CommentText = right.CommentText
        && left.CommentAuthor = right.CommentAuthor
        && left.CommentDate = right.CommentDate
        && left.NoteId = right.NoteId
        && left.NoteType = right.NoteType
        && left.NoteText = right.NoteText
        && left.BookmarkStartId = right.BookmarkStartId
        && left.BookmarkName = right.BookmarkName
        && left.BookmarkEndId = right.BookmarkEndId
        && left.FieldCode = right.FieldCode
        && left.FieldResult = right.FieldResult
        && left.EquationXml = right.EquationXml
        && left.RevisionId = right.RevisionId
        && left.RevisionType = right.RevisionType
        && left.RevisionAuthor = right.RevisionAuthor
        && left.RevisionDate = right.RevisionDate
        && left.DeletedInsertion = right.DeletedInsertion
        && left.MoveName = right.MoveName
        && left.VertAlign = right.VertAlign)

let private mergedRunList (runs: ResizeArray<Run>) =
    let mutable result = []
    for i = runs.Count - 1 downto 0 do
        match runs.[i], result with
        | Text current, Text next :: rest when current.Start + current.Length = next.Start && sameTextStyle current.Style next.Style ->
            result <- Text { current with Length = current.Length + next.Length } :: rest
        | run, _ ->
            result <- run :: result
    result

let private tryExtendLastTextRun (canExtendStyle: TextStyle -> bool) bufferLen textLen (runs: Run list) =
    let values = ResizeArray<Run>()
    let mutable remaining = runs
    let mutable extendedLast = false
    let mutable canExtend = true
    while canExtend && not remaining.IsEmpty do
        match remaining with
        | [ Text tr ] when tr.Start + tr.Length = bufferLen && canExtendStyle tr.Style ->
            values.Add(Text { tr with Length = tr.Length + textLen })
            extendedLast <- true
            remaining <- []
        | [ Text _ ] ->
            canExtend <- false
        | [ Image _ ] ->
            canExtend <- false
        | run :: tail ->
            values.Add run
            remaining <- tail
        | [] -> ()
    if extendedLast then Some (resizeArrayToList values) else None

/// Create empty buffer
let empty = {
    Rope = Empty
    Runs = []
}

/// Create from string
let ofString text =
    let len = String.length text
    {
        Rope = Rope.ofString text
        Runs = if len > 0 then [Text { Start = 0; Length = len; Style = defaultStyle }] else []
    }

/// Get length
let length buffer = Rope.length buffer.Rope

/// Get text content
let toString buffer = Rope.toString buffer.Rope

let appendToStringBuilder (sb: System.Text.StringBuilder) buffer =
    Rope.appendToStringBuilder sb buffer.Rope

let appendSliceToStringBuilder (sb: System.Text.StringBuilder) start len buffer =
    Rope.appendSliceToStringBuilder sb start len buffer.Rope

let sliceEqualsString start len (text: string) buffer =
    Rope.sliceEqualsString start len text buffer.Rope

/// Find non-overlapping literal matches without materializing the whole rope.
let iterMatchRanges (matchCase: bool) (query: string) buffer (emit: int -> int -> unit) =
    if System.String.IsNullOrEmpty query || query.Length > length buffer then
        ()
    else
        let normalize ch =
            if matchCase then ch else System.Char.ToUpperInvariant ch
        let firstQueryChar = normalize query.[0]
        let indexOfInString (text: string) start =
#if FABLE_COMPILER
            if start > text.Length - query.Length then
                -1
            elif matchCase then
                text.IndexOf(query, start)
            else
                let charEqualsIgnoreCase a b =
                    a = b || System.Char.ToUpperInvariant a = System.Char.ToUpperInvariant b
                let mutable found = -1
                let mutable index = start
                while found < 0 && index <= text.Length - query.Length do
                    let mutable offset = 0
                    let mutable matches = true
                    while matches && offset < query.Length do
                        if not (charEqualsIgnoreCase text.[index + offset] query.[offset]) then
                            matches <- false
                        offset <- offset + 1
                    if matches then
                        found <- index
                    index <- index + 1
                found
#else
            let comparison = if matchCase then System.StringComparison.Ordinal else System.StringComparison.OrdinalIgnoreCase
            text.IndexOf(query, start, comparison)
#endif

        let mutable nextAllowedStart = 0
        let emitIfAllowed start finish =
            if start >= nextAllowedStart then
                emit start finish
                nextAllowedStart <- finish

        let searchText absoluteStart (text: string) startLocal =
            let mutable index = indexOfInString text startLocal
            while index >= 0 do
                let start = absoluteStart + index
                emitIfAllowed start (start + query.Length)
                let nextLocal = max (index + 1) (nextAllowedStart - absoluteStart)
                index <- indexOfInString text nextLocal

        match buffer.Rope with
        | Leaf text ->
            searchText 0 text 0
        | _ ->
            let carryLimit = query.Length - 1
            let carryBuffer = if carryLimit > 0 then Array.zeroCreate<char> carryLimit else [||]
            let mutable carryLength = 0
            let mutable carryStart = 0
            let equalsQuery ch queryIndex =
                normalize ch = normalize query.[queryIndex]
            let updateCarry absoluteStart (text: string) =
                if carryLimit > 0 then
                    let takeFromText = min carryLimit text.Length
                    if takeFromText = carryLimit then
                        let sourceStart = text.Length - takeFromText
                        for i = 0 to takeFromText - 1 do
                            carryBuffer.[i] <- text.[sourceStart + i]
                        carryLength <- takeFromText
                        carryStart <- absoluteStart + sourceStart
                    else
                        let keepFromCarry = min carryLength (carryLimit - takeFromText)
                        let previousCarryLength = carryLength
                        let previousCarryStart = carryStart
                        if keepFromCarry > 0 then
                            let sourceStart = previousCarryLength - keepFromCarry
                            for i = 0 to keepFromCarry - 1 do
                                carryBuffer.[i] <- carryBuffer.[sourceStart + i]
                        for i = 0 to takeFromText - 1 do
                            carryBuffer.[keepFromCarry + i] <- text.[i]
                        carryLength <- keepFromCarry + takeFromText
                        carryStart <-
                            if keepFromCarry > 0 then previousCarryStart + previousCarryLength - keepFromCarry
                            else absoluteStart

            let mutable absolute = 0
            let rec walk rope =
                match rope with
                | Empty -> ()
                | Leaf text ->
                    let leafStart = absolute
                    if carryLimit > 0 && carryLength > 0 && text.Length > 0 then
                        let mutable startInCarry = 0
                        while startInCarry < carryLength do
                            let carryAvailable = carryLength - startInCarry
                            if carryAvailable < query.Length
                               && carryAvailable + text.Length >= query.Length
                               && normalize carryBuffer.[startInCarry] = firstQueryChar then
                                let mutable queryIndex = 0
                                let mutable matches = true
                                while matches && queryIndex < query.Length do
                                    let ch =
                                        if queryIndex < carryAvailable then
                                            carryBuffer.[startInCarry + queryIndex]
                                        else
                                            text.[queryIndex - carryAvailable]
                                    if not (equalsQuery ch queryIndex) then
                                        matches <- false
                                    queryIndex <- queryIndex + 1
                                if matches then
                                    let start = carryStart + startInCarry
                                    emitIfAllowed start (start + query.Length)
                            startInCarry <- startInCarry + 1

                    let startLocal = max 0 (nextAllowedStart - leafStart)
                    searchText leafStart text startLocal
                    updateCarry leafStart text
                    absolute <- absolute + text.Length
                | Node (left, right, _) ->
                    walk left
                    walk right
            walk buffer.Rope

/// True when the buffer contains no non-whitespace characters, without
/// materializing the rope as one string.
let isWhiteSpace buffer =
    let rec ropeIsWhiteSpace rope =
        match rope with
        | Empty -> true
        | Leaf text ->
            let mutable index = 0
            let mutable onlyWhitespace = true
            while onlyWhitespace && index < text.Length do
                if not (System.Char.IsWhiteSpace text.[index]) then
                    onlyWhitespace <- false
                index <- index + 1
            onlyWhitespace
        | Node (left, right, _) ->
            ropeIsWhiteSpace left && ropeIsWhiteSpace right
    ropeIsWhiteSpace buffer.Rope

/// True when the buffer is non-empty and contains only whitespace or NBSP,
/// without materializing the rope as one string.
let isNonbreakingBlank buffer =
    let isWhitespaceOrNbsp ch =
        System.Char.IsWhiteSpace ch || ch = '\u00A0'
    let rec ropeIsBlank rope =
        match rope with
        | Empty -> true
        | Leaf text ->
            let mutable index = 0
            let mutable onlyBlank = true
            while onlyBlank && index < text.Length do
                if not (isWhitespaceOrNbsp text.[index]) then
                    onlyBlank <- false
                index <- index + 1
            onlyBlank
        | Node (left, right, _) ->
            ropeIsBlank left && ropeIsBlank right
    length buffer > 0 && ropeIsBlank buffer.Rope

let private clearRevisionFields (style: TextStyle) =
    if style.RevisionType.IsNone && style.RevisionId.IsNone && style.RevisionAuthor.IsNone && style.RevisionDate.IsNone && style.DeletedInsertion.IsNone && style.MoveName.IsNone then
        style
    else
        { style with RevisionType = None; RevisionId = None; RevisionAuthor = None; RevisionDate = None; DeletedInsertion = None; MoveName = None }

/// True when a run carries annotation identity (a field, hyperlink, comment,
/// note reference, bookmark, equation or revision) rather than only
/// character formatting.
let hasAnnotationIdentity (style: TextStyle) =
    style.Hyperlink.IsSome
    || style.InactiveHyperlink.IsSome
    || style.CommentId.IsSome
    || style.CommentText.IsSome
    || style.CommentAuthor.IsSome
    || style.CommentDate.IsSome
    || style.NoteId.IsSome
    || style.NoteType.IsSome
    || style.NoteText.IsSome
    || style.BookmarkStartId.IsSome
    || style.BookmarkName.IsSome
    || style.BookmarkEndId.IsSome
    || style.FieldCode.IsSome
    || style.FieldResult.IsSome
    || style.EquationXml.IsSome
    || style.RevisionType.IsSome
    || style.RevisionId.IsSome
    || style.RevisionAuthor.IsSome
    || style.RevisionDate.IsSome
    || style.DeletedInsertion.IsSome
    || style.MoveName.IsSome
    || style.RenderedPageBreakBefore

/// Style for text typed at the edge of a run. The new text takes the run's
/// character formatting but not its annotation identity: text typed after a
/// field must not become part of the field (it would be thrown away when the
/// field refreshes), and the same goes for note references, bookmarks,
/// equations and revisions. A hyperlink or comment is kept only when the run
/// on the other side of the caret belongs to the same one, i.e. the caret is
/// really inside it.
let private boundaryInsertStyle (style: TextStyle) (otherSide: TextStyle option) =
    if not (hasAnnotationIdentity style) then
        style
    else
        let keepHyperlink =
            match otherSide with
            | Some other -> (style.Hyperlink.IsSome || style.InactiveHyperlink.IsSome) && other.Hyperlink = style.Hyperlink && other.InactiveHyperlink = style.InactiveHyperlink
            | None -> false
        let keepComment =
            match otherSide with
            | Some other -> style.CommentId.IsSome && other.CommentId = style.CommentId
            | None -> false
        { style with
            Hyperlink = if keepHyperlink then style.Hyperlink else None
            InactiveHyperlink = if keepHyperlink then style.InactiveHyperlink else None
            CommentId = if keepComment then style.CommentId else None
            CommentText = if keepComment then style.CommentText else None
            CommentAuthor = if keepComment then style.CommentAuthor else None
            CommentDate = if keepComment then style.CommentDate else None
            // (The colour a comment's mark is shown in is the mark's: text typed beside it is not of that colour.)
            Color = (if style.CommentId.IsSome && not keepComment && style.Color = Some commentMarkDisplayColor then None else style.Color)
            NoteId = None
            NoteType = None
            NoteText = None
            BookmarkStartId = None
            BookmarkName = None
            BookmarkEndId = None
            FieldCode = None
            FieldResult = None
            RenderedPageBreakBefore = false
            EquationXml = None
            RevisionId = None
            RevisionType = None
            RevisionAuthor = None
            RevisionDate = None
            DeletedInsertion = None
            MoveName = None }

/// A run's character formatting alone: what text typed next to it has when
/// it is part of none of the run's field, link, comment, note, bookmark or
/// tracked change.
let withoutAnnotations (style: TextStyle) : TextStyle = boundaryInsertStyle style None

/// Insert text at position
let insert pos (text: string) buffer =
    if String.length text = 0 then buffer
    else
        let textLen = String.length text
        let bufferLen = length buffer

        let insertByScan () =
            let newRuns = ResizeArray<Run>()
            let mutable inserted = false
            let mutable remainingRuns = buffer.Runs
            for run in buffer.Runs do
                remainingRuns <- remainingRuns.Tail
                let runStart = Run.start run
                let runEnd = Run.endPos run
                if not inserted && pos >= runStart && pos <= runEnd then
                    // Insert within this run
                    match run with
                    | Image _ ->
                        // Can't split an image: the text goes on whichever
                        // side of it the position falls, matching the rope.
                        if pos = runStart then
                            newRuns.Add(Text { Start = pos; Length = textLen; Style = defaultStyle })
                            newRuns.Add(Run.shiftStart textLen run)
                        else
                            newRuns.Add(run)
                            newRuns.Add(Text { Start = pos; Length = textLen; Style = defaultStyle })
                    | Text tr ->
                        if pos = runStart then
                            // Insert at start of run: nothing precedes the
                            // caret, so the text is outside this run's
                            // annotations.
                            // (The mark of a page break as it was rendered stays
                            // where it is: before the text typed there.)
                            let typed = boundaryInsertStyle tr.Style None
                            if tr.Style.RenderedPageBreakBefore then
                                newRuns.Add(Text { Start = tr.Start; Length = textLen; Style = { typed with RenderedPageBreakBefore = true } })
                                newRuns.Add(Text { tr with Start = tr.Start + textLen; Style = { tr.Style with RenderedPageBreakBefore = false } })
                            else
                                newRuns.Add(Text { Start = tr.Start; Length = textLen; Style = typed })
                                newRuns.Add(Text { tr with Start = tr.Start + textLen })
                        elif pos = runEnd then
                            // Insert at end of run
                            let following =
                                match remainingRuns with
                                | Text next :: _ when next.Start = runEnd -> Some next.Style
                                | _ -> None
                            newRuns.Add(run)
                            newRuns.Add(Text { Start = runEnd; Length = textLen; Style = boundaryInsertStyle tr.Style following })
                        else
                            // Insert in middle - split into three runs. The
                            // new text is not itself a tracked change.
                            let beforeLen = pos - tr.Start
                            let afterLen = tr.Length - beforeLen
                            newRuns.Add(Text { tr with Length = beforeLen })
                            newRuns.Add(Text { Start = pos; Length = textLen; Style = clearRevisionFields tr.Style })
                            newRuns.Add(Text { Start = pos + textLen; Length = afterLen; Style = tr.Style })
                    inserted <- true
                elif inserted then
                    // Shift runs after insertion
                    newRuns.Add(Run.shiftStart textLen run)
                else
                    // Before insertion point
                    newRuns.Add(run)

            // If no runs or insertion at end, add new run
            if not inserted then
                newRuns.Add(Text { Start = pos; Length = textLen; Style = defaultStyle })

            { buffer with
                Rope = Rope.insert pos text buffer.Rope
                Runs = mergedRunList newRuns }

        if bufferLen = 0 && pos = 0 then
            ofString text
        elif pos = bufferLen then
            match buffer.Runs with
            | [ Text tr ] when tr.Start = 0 && tr.Length = bufferLen && not (hasAnnotationIdentity tr.Style) ->
                { buffer with
                    Rope = Rope.insert pos text buffer.Rope
                    Runs = [ Text { tr with Length = tr.Length + textLen } ] }
            | _ ->
                match tryExtendLastTextRun (fun style -> not (hasAnnotationIdentity style)) bufferLen textLen buffer.Runs with
                | Some runs ->
                    { buffer with
                        Rope = Rope.insert pos text buffer.Rope
                        Runs = runs }
                | None -> insertByScan ()
        else
            insertByScan ()

/// Delete range [start, start + len)
let delete start len buffer =
    if len <= 0 then buffer
    else
        let bufferLen = length buffer
        let deleteEnd = start + len
        let actualStart = max 0 (min bufferLen start)
        let actualEnd = max actualStart (min bufferLen deleteEnd)
        let actualLen = actualEnd - actualStart

        if actualLen <= 0 then
            buffer
        else
            let start = actualStart
            let len = actualLen
            let deleteEnd = actualEnd
            match buffer.Runs with
            | [ Text tr ] when tr.Start = 0 && tr.Length = bufferLen ->
                let remaining = bufferLen - actualLen
                { buffer with
                    Rope = Rope.delete actualStart actualLen buffer.Rope
                    Runs = if remaining > 0 then [ Text { tr with Length = remaining } ] else [] }
            | _ ->
                let newRuns = ResizeArray<Run>()

                for run in buffer.Runs do
                    let runStart = Run.start run
                    let runEnd = Run.endPos run
                    if runEnd <= start then
                        // Run is entirely before deletion
                        newRuns.Add(run)
                    elif runStart >= deleteEnd then
                        // Run is entirely after deletion - shift back
                        newRuns.Add(Run.shiftStart -len run)
                    elif runStart >= start && runEnd <= deleteEnd then
                        // Run is entirely within deletion - remove it
                        ()
                    elif runStart < start && runEnd > deleteEnd then
                        // Deletion is within run - shrink it
                        match run with
                        | Image _ -> () // Image deleted
                        | Text tr -> newRuns.Add(Text { tr with Length = tr.Length - len })
                    elif runStart < start then
                        // Run starts before, ends within deletion
                        match run with
                        | Image _ -> () // Image deleted
                        | Text tr ->
                            let newLen = start - tr.Start
                            newRuns.Add(Text { tr with Length = newLen })
                    else
                        // Run starts within deletion, ends after
                        match run with
                        | Image _ -> () // Image deleted
                        | Text tr ->
                            let overlap = deleteEnd - tr.Start
                            let newLen = tr.Length - overlap
                            newRuns.Add(Text { Start = start; Length = newLen; Style = tr.Style })

                { buffer with
                    Rope = Rope.delete start len buffer.Rope
                    Runs = mergedRunList newRuns }

/// Get substring
let slice start len buffer =
    Rope.slice start len buffer.Rope

/// Get a buffer slice while preserving text styles and image metadata.
let sliceBuffer start len buffer =
    if len <= 0 then empty
    else
        let bufferLen = length buffer
        let sliceStart = max 0 (min start bufferLen)
        let sliceEnd = max sliceStart (min bufferLen (sliceStart + len))
        if sliceStart = 0 && sliceEnd = bufferLen then
            buffer
        else
            let newRuns = ResizeArray<Run>()
            for run in buffer.Runs do
                let runStart = Run.start run
                let runEnd = Run.endPos run
                let overlapStart = max runStart sliceStart
                let overlapEnd = min runEnd sliceEnd
                if overlapStart < overlapEnd then
                    match run with
                    | Text tr ->
                        newRuns.Add(Text { tr with Start = overlapStart - sliceStart; Length = overlapEnd - overlapStart })
                    | Image ir ->
                        newRuns.Add(Image { ir with Start = ir.Start - sliceStart })
            { Rope = Rope.sliceRope sliceStart (sliceEnd - sliceStart) buffer.Rope
              Runs = mergedRunList newRuns }

/// Append two buffers without flattening either side to a string.
let append left right =
    let leftLen = length left
    if leftLen = 0 then right
    elif length right = 0 then left
    else
        let rightLen = length right
        match left.Runs, right.Runs with
        | [ Text leftRun ], [ Text rightRun ] when leftRun.Start = 0 && rightRun.Start = 0 && leftRun.Length = leftLen && rightRun.Length = rightLen ->
            let runs =
                if leftRun.Style = rightRun.Style then
                    [ Text { leftRun with Length = leftLen + rightLen } ]
                else
                    [ Text leftRun; Text { rightRun with Start = leftLen } ]
            { Rope = Rope.balance left.Rope right.Rope
              Runs = runs }
        | _ ->
            let runs = ResizeArray<Run>()
            for run in left.Runs do
                runs.Add run
            for run in right.Runs do
                runs.Add(Run.shiftStart leftLen run)
            { Rope = Rope.balance left.Rope right.Rope
              Runs = mergedRunList runs }

let trySliceNoCopyV start len buffer =
    match buffer.Rope with
    | Leaf text when start = 0 && len = text.Length -> ValueSome text
    | _ -> ValueNone

let trySliceNoCopy start len buffer =
    match trySliceNoCopyV start len buffer with
    | ValueSome text -> Some text
    | ValueNone -> None

let private styleAt pos buffer =
    if pos < 0 || pos >= length buffer then
        defaultStyle
    else
        match buffer.Runs with
        | [ Text tr ] when pos >= tr.Start && pos < tr.Start + tr.Length -> tr.Style
        | _ ->
            let runs = List.toArray buffer.Runs
            let mutable lo = 0
            let mutable hi = runs.Length - 1
            let mutable found = defaultStyle
            let mutable searching = true
            while searching && lo <= hi do
                let mid = lo + ((hi - lo) / 2)
                let run = runs.[mid]
                let runStart = Run.start run
                if pos < runStart then
                    hi <- mid - 1
                else
                    let runEnd = runStart + Run.length run
                    if pos >= runEnd then
                        lo <- mid + 1
                    else
                        match run with
                        | Text tr -> found <- tr.Style
                        | Image _ -> ()
                        searching <- false
            found

/// Replace sorted, non-overlapping ranges in one pass using text already
/// materialized by the caller. Replacement text uses the same insertion-side
/// style as repeated delete+insert at each range start.
let replaceRangesWithText (originalText: string) (ranges: ResizeArray<struct (int * int)>) (replacement: string) buffer =
    if ranges.Count = 0 then
        buffer
    else
        let originalLength = originalText.Length
        let replacementLength = replacement.Length
        let mutable removedLength = 0
        let mutable sizingCursor = 0
        for i = 0 to ranges.Count - 1 do
            let struct (rawStart, rawFinish) = ranges.[i]
            let start = max sizingCursor (min originalLength rawStart)
            let finish = max start (min originalLength rawFinish)
            removedLength <- removedLength + (finish - start)
            sizingCursor <- finish
        let estimatedLength = int (min (int64 System.Int32.MaxValue) (max 0L (int64 originalLength - int64 removedLength + int64 ranges.Count * int64 replacementLength)))
        let text = System.Text.StringBuilder(estimatedLength)
        let appendTextSlice start finish =
            if finish > start then
#if FABLE_COMPILER
                text.Append(originalText.Substring(start, finish - start)) |> ignore
#else
                text.Append(originalText, start, finish - start) |> ignore
#endif
        match buffer.Runs with
        | [ Text tr ] when tr.Start = 0 && tr.Length = originalLength ->
            let mutable cursor = 0
            for i = 0 to ranges.Count - 1 do
                let struct (rawStart, rawFinish) = ranges.[i]
                let start = max cursor (min originalLength rawStart)
                let finish = max start (min originalLength rawFinish)
                appendTextSlice cursor start
                if replacementLength > 0 then
                    text.Append replacement |> ignore
                cursor <- finish
            appendTextSlice cursor originalLength
            { Rope = Rope.ofString (text.ToString())
              Runs = if estimatedLength > 0 then [ Text { tr with Length = estimatedLength } ] else [] }
        | _ ->
            let runs = buffer.Runs |> List.toArray
            let newRuns = ResizeArray<Run>(runs.Length + ranges.Count)
            let mutable copyRunIndex = 0
            let addCopiedRuns rangeStart rangeEnd outStart =
                if rangeEnd > rangeStart then
                    while copyRunIndex < runs.Length && Run.endPos runs.[copyRunIndex] <= rangeStart do
                        copyRunIndex <- copyRunIndex + 1
                    let mutable index = copyRunIndex
                    let mutable copying = true
                    while copying && index < runs.Length && Run.start runs.[index] < rangeEnd do
                        let run = runs.[index]
                        let runStart = Run.start run
                        let runEnd = Run.endPos run
                        let overlapStart = max runStart rangeStart
                        let overlapEnd = min runEnd rangeEnd
                        if overlapStart < overlapEnd then
                            let shiftedStart = outStart + (overlapStart - rangeStart)
                            match run with
                            | Text tr ->
                                if overlapStart = runStart && overlapEnd = runEnd && shiftedStart = tr.Start then
                                    newRuns.Add run
                                else
                                    newRuns.Add(Text { tr with Start = shiftedStart; Length = overlapEnd - overlapStart })
                            | Image ir ->
                                if overlapStart = runStart && shiftedStart = ir.Start then
                                    newRuns.Add run
                                else
                                    newRuns.Add(Image { ir with Start = shiftedStart })
                        if runEnd <= rangeEnd then
                            index <- index + 1
                        else
                            copying <- false
                    copyRunIndex <- index
            let mutable styleRunIndex = 0
            let styleAtAscending pos =
                while styleRunIndex < runs.Length && Run.endPos runs.[styleRunIndex] <= pos do
                    styleRunIndex <- styleRunIndex + 1
                if styleRunIndex < runs.Length && pos >= Run.start runs.[styleRunIndex] then
                    match runs.[styleRunIndex] with
                    | Text tr -> tr.Style
                    | Image _ -> defaultStyle
                else
                    defaultStyle
            let mutable cursor = 0
            let mutable outPos = 0
            for i = 0 to ranges.Count - 1 do
                let struct (rawStart, rawFinish) = ranges.[i]
                let start = max cursor (min originalLength rawStart)
                let finish = max start (min originalLength rawFinish)
                appendTextSlice cursor start
                addCopiedRuns cursor start outPos
                outPos <- outPos + (start - cursor)
                if replacementLength > 0 then
                    // Same rule as typing at that position: inside a run
                    // the replacement belongs to it, at a run edge it takes
                    // the formatting but not the annotation identity.
                    let style =
                        if start > 0 then
                            let before = styleAtAscending (start - 1)
                            if styleRunIndex < runs.Length && Run.endPos runs.[styleRunIndex] > start then
                                before
                            else
                                let after =
                                    if styleRunIndex + 1 < runs.Length then
                                        match runs.[styleRunIndex + 1] with
                                        | Text next when next.Start = start -> Some next.Style
                                        | _ -> None
                                    else None
                                boundaryInsertStyle before after
                        elif start < originalLength then boundaryInsertStyle (styleAtAscending start) None
                        else defaultStyle
                    text.Append replacement |> ignore
                    newRuns.Add(Text { Start = outPos; Length = replacementLength; Style = style })
                    outPos <- outPos + replacementLength
                cursor <- finish
            appendTextSlice cursor originalLength
            addCopiedRuns cursor originalLength outPos
            { Rope = Rope.ofString (text.ToString())
              Runs = mergedRunList newRuns }

/// Replace sorted, non-overlapping ranges in one pass. Replacement text uses
/// the same insertion-side style as repeated delete+insert at each range start.
let replaceRanges (ranges: ResizeArray<struct (int * int)>) (replacement: string) buffer =
    replaceRangesWithText (toString buffer) ranges replacement buffer

/// Object Replacement Character - used as placeholder for images
let objectReplacementChar = "\uFFFC"

/// Resize image at position
let resizeImage pos newWidth newHeight buffer =
    if pos < 0 || pos >= length buffer then
        buffer
    else
        let mutable found = false
        let mutable unchanged = false
        let mutable remaining = buffer.Runs
        while not found && not remaining.IsEmpty do
            match remaining with
            | Image imgRun :: tail when imgRun.Start = pos ->
                found <- true
                let widthUnchanged =
                    match imgRun.Width with
                    | Some width -> width = newWidth
                    | None -> false
                let heightUnchanged =
                    match imgRun.Height with
                    | Some height -> height = newHeight
                    | None -> false
                unchanged <- widthUnchanged && heightUnchanged
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()

        if not found || unchanged then
            buffer
        else
            let newRuns = ResizeArray<Run>()
            for run in buffer.Runs do
                match run with
                | Image imgRun when imgRun.Start = pos ->
                    newRuns.Add(Image { imgRun with Width = Some newWidth; Height = Some newHeight })
                | _ ->
                    newRuns.Add run
            { buffer with Runs = resizeArrayToList newRuns }

/// Insert an image run as it is (every field kept, an opaque object's XML
/// included) at a position.
let insertImageRun pos (image: ImageRun) buffer =
    let newRope = Rope.insert pos objectReplacementChar buffer.Rope
    let imageRun = Image { image with Start = pos }

    // Insert image run and shift subsequent runs
    let newRuns = ResizeArray<Run>()
    let mutable inserted = false
    for run in buffer.Runs do
        let runStart = Run.start run
        let runEnd = Run.endPos run
        if inserted then
            newRuns.Add(Run.shiftStart 1 run)
        elif pos < runStart then
            newRuns.Add imageRun
            newRuns.Add(Run.shiftStart 1 run)
            inserted <- true
        elif pos >= runStart && pos <= runEnd then
            match run with
            | Image _ ->
                if pos <= runStart then
                    newRuns.Add imageRun
                    newRuns.Add(Run.shiftStart 1 run)
                else
                    newRuns.Add run
                    newRuns.Add imageRun
                inserted <- true
            | Text tr ->
                let beforeLen = pos - tr.Start
                let afterLen = tr.Length - beforeLen
                if beforeLen > 0 then
                    newRuns.Add(Text { tr with Length = beforeLen })
                newRuns.Add imageRun
                if afterLen > 0 then
                    newRuns.Add(Text { Start = pos + 1; Length = afterLen; Style = tr.Style })
                inserted <- true
        else
            newRuns.Add run

    if not inserted then
        newRuns.Add imageRun

    { Rope = newRope; Runs = resizeArrayToList newRuns }

/// Insert image at position
let insertImageWithMetadata pos src linked width height alt name title floating wrap positionX positionY positionHRelativeFrom positionVRelativeFrom behindDoc allowOverlap layoutInCell distanceTop distanceBottom distanceLeft distanceRight relativeHeight locked simplePosition simplePositionX simplePositionY cropLeft cropTop cropRight cropBottom buffer =
    insertImageRun pos {
        Start = pos
        Src = src
        Linked = linked
        Width = width
        Height = height
        Alt = alt
        Name = name
        Title = title
        Floating = floating
        Wrap = wrap
        PositionX = positionX
        PositionY = positionY
        PositionHRelativeFrom = positionHRelativeFrom
        PositionVRelativeFrom = positionVRelativeFrom
        BehindDoc = behindDoc
        AllowOverlap = allowOverlap
        LayoutInCell = layoutInCell
        DistanceTop = distanceTop
        DistanceBottom = distanceBottom
        DistanceLeft = distanceLeft
        DistanceRight = distanceRight
        RelativeHeight = relativeHeight
        Locked = locked
        SimplePosition = simplePosition
        SimplePositionX = simplePositionX
        SimplePositionY = simplePositionY
        CropLeft = cropLeft
        CropTop = cropTop
        CropRight = cropRight
        CropBottom = cropBottom
        PositionHAlign = None
        PositionVAlign = None
        WrapText = None
        WrapPolygon = []
        Opaque = None
    } buffer

let insertImage pos src width height alt buffer =
    insertImageWithMetadata pos src false width height alt None None false None None None None None None None None None None None None None None None None None None None None None buffer

let tryMoveImageWithinBuffer offset insertOffset (img: ImageRun) buffer =
    let bufferAfterDeleteLen = max 0 (length buffer - 1)
    let insertOffset = max 0 (min bufferAfterDeleteLen insertOffset)
    if insertOffset = offset then
        Some buffer
    else
        let mutable canUseFastPath = true
        let newRuns = ResizeArray<Run>()
        let moved = Image { img with Start = insertOffset }
        let mutable inserted = false
        let mutable removed = false

        let addMovedIfBefore startAfterDelete =
            if canUseFastPath && not inserted && insertOffset <= startAfterDelete then
                newRuns.Add moved
                inserted <- true

        for run in buffer.Runs do
            match run with
            | Image image when image.Start = offset ->
                removed <- true
            | _ ->
                if canUseFastPath then
                    let start = Run.start run
                    let startAfterDelete = if start > offset then start - 1 else start
                    match run with
                    | Text tr ->
                        let endAfterDelete = startAfterDelete + tr.Length
                        if insertOffset > startAfterDelete && insertOffset < endAfterDelete then
                            canUseFastPath <- false
                    | Image _ -> ()
                    if canUseFastPath then
                        addMovedIfBefore startAfterDelete
                        let finalStart = if inserted && startAfterDelete >= insertOffset then startAfterDelete + 1 else startAfterDelete
                        newRuns.Add(Run.shiftStart (finalStart - start) run)

        if canUseFastPath && removed then
            if not inserted then
                newRuns.Add moved
            Some {
                Rope = buffer.Rope |> Rope.delete offset 1 |> Rope.insert insertOffset objectReplacementChar
                Runs = resizeArrayToList newRuns
            }
        else
            None

/// Apply style to range [start, start + len)
let applyStyle start len styleTransform buffer =
    if len <= 0 then buffer
    else
        let rangeEnd = start + len

        match buffer.Runs with
        | [ Text tr ] when tr.Start = 0 && tr.Length = length buffer ->
            let runStart = tr.Start
            let runEnd = tr.Start + tr.Length
            if runEnd <= start || runStart >= rangeEnd then
                buffer
            elif runStart >= start && runEnd <= rangeEnd then
                let transformedStyle = styleTransform tr.Style
                if sameTextStyle transformedStyle tr.Style then buffer
                else { buffer with Runs = [ Text { tr with Style = transformedStyle } ] }
            else
                let transformedStyle = styleTransform tr.Style
                if sameTextStyle transformedStyle tr.Style then
                    buffer
                else
                    let runs = ResizeArray<Run>(3)
                    let addText start length style =
                        if length > 0 then
                            runs.Add(Text { Start = start; Length = length; Style = style })
                    let overlapStart = max runStart start
                    let overlapEnd = min runEnd rangeEnd
                    addText runStart (overlapStart - runStart) tr.Style
                    addText overlapStart (overlapEnd - overlapStart) transformedStyle
                    addText overlapEnd (runEnd - overlapEnd) tr.Style
                    { buffer with Runs = mergedRunList runs }
        | _ ->
            let bufferLen = length buffer
            let fullRangeApply () =
                let newRuns = ResizeArray<Run>()
                let mutable changed = false
                let mutable hasPendingText = false
                let mutable pendingStart = 0
                let mutable pendingLength = 0
                let mutable pendingStyle = defaultStyle
                let flushPendingText () =
                    if hasPendingText then
                        newRuns.Add(Text { Start = pendingStart; Length = pendingLength; Style = pendingStyle })
                        hasPendingText <- false
                let addText start length style =
                    if length > 0 then
                        if hasPendingText && pendingStart + pendingLength = start && sameTextStyle pendingStyle style then
                            pendingLength <- pendingLength + length
                        else
                            flushPendingText ()
                            pendingStart <- start
                            pendingLength <- length
                            pendingStyle <- style
                            hasPendingText <- true
                for run in buffer.Runs do
                    match run with
                    | Text tr ->
                        let transformedStyle = styleTransform tr.Style
                        changed <- changed || not (sameTextStyle transformedStyle tr.Style)
                        addText tr.Start tr.Length transformedStyle
                    | Image _ ->
                        flushPendingText ()
                        newRuns.Add run
                if changed then
                    flushPendingText ()
                    { buffer with Runs = resizeArrayToList newRuns }
                else
                    buffer

            let partialRangeApply () =
                let mutable changed = false
                let mutable prefixCount = 0
                let mutable newRuns = Unchecked.defaultof<ResizeArray<Run>>
                let copyPrefix () =
                    let mutable copied = 0
                    let mutable prefix = buffer.Runs
                    while copied < prefixCount do
                        match prefix with
                        | run :: tail ->
                            newRuns.Add run
                            prefix <- tail
                            copied <- copied + 1
                        | [] -> copied <- prefixCount
                let ensureChanged () =
                    if not changed then
                        changed <- true
                        newRuns <- ResizeArray<Run>()
                        copyPrefix ()
                let addChangedText (tr: TextRun) runStart runEnd transformedStyle =
                    if runEnd <= start || runStart >= rangeEnd then
                        newRuns.Add(Text tr)
                    elif runStart >= start && runEnd <= rangeEnd then
                        newRuns.Add(Text { tr with Style = transformedStyle })
                    elif runStart < start && runEnd > rangeEnd then
                        let beforeLen = start - tr.Start
                        let rangeLen = len
                        let afterLen = runEnd - rangeEnd
                        newRuns.Add(Text { tr with Length = beforeLen })
                        newRuns.Add(Text { Start = start; Length = rangeLen; Style = transformedStyle })
                        newRuns.Add(Text { Start = rangeEnd; Length = afterLen; Style = tr.Style })
                    elif runStart < start then
                        let beforeLen = start - tr.Start
                        let overlapLen = runEnd - start
                        newRuns.Add(Text { tr with Length = beforeLen })
                        newRuns.Add(Text { Start = start; Length = overlapLen; Style = transformedStyle })
                    else
                        let overlapLen = rangeEnd - tr.Start
                        let afterLen = runEnd - rangeEnd
                        newRuns.Add(Text { Start = tr.Start; Length = overlapLen; Style = transformedStyle })
                        newRuns.Add(Text { Start = rangeEnd; Length = afterLen; Style = tr.Style })

                let mutable remaining = buffer.Runs
                while not remaining.IsEmpty do
                    match remaining with
                    | run :: _ when Run.start run >= rangeEnd ->
                        if changed then
                            let mutable rest = remaining
                            while not rest.IsEmpty do
                                match rest with
                                | tailRun :: tail ->
                                    newRuns.Add tailRun
                                    rest <- tail
                                | [] -> ()
                        remaining <- []
                    | Image _ as run :: tail ->
                        if changed then newRuns.Add run
                        else prefixCount <- prefixCount + 1
                        remaining <- tail
                    | Text tr as run :: tail ->
                        let runStart = tr.Start
                        let runEnd = tr.Start + tr.Length
                        if runEnd <= start || runStart >= rangeEnd then
                            if changed then newRuns.Add run
                            else prefixCount <- prefixCount + 1
                        else
                            let transformedStyle = styleTransform tr.Style
                            if sameTextStyle transformedStyle tr.Style then
                                if changed then newRuns.Add run
                                else prefixCount <- prefixCount + 1
                            else
                                ensureChanged ()
                                addChangedText tr runStart runEnd transformedStyle
                        remaining <- tail
                    | [] -> ()

                if changed then { buffer with Runs = mergedRunList newRuns }
                else buffer

            if start <= 0 && rangeEnd >= bufferLen then fullRangeApply ()
            else partialRangeApply ()

let private textRunsInRangeSatisfy start len predicate buffer =
    if len <= 0 then
        true
    else
        let rangeEnd = start + len
        let mutable allMatch = true
        let mutable remaining = buffer.Runs
        while allMatch && not remaining.IsEmpty do
            match remaining with
            | Text tr :: tail ->
                let runStart = tr.Start
                let runEnd = tr.Start + tr.Length
                if runEnd > start && runStart < rangeEnd then
                    allMatch <- predicate tr.Style
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()
        allMatch

let private textRunsInRangeHaveVertAlign start len align buffer =
    if len <= 0 then
        true
    else
        let rangeEnd = start + len
        let mutable allMatch = true
        let mutable remaining = buffer.Runs
        while allMatch && not remaining.IsEmpty do
            match remaining with
            | Text tr :: tail ->
                let runStart = tr.Start
                let runEnd = tr.Start + tr.Length
                if runEnd > start && runStart < rangeEnd then
                    allMatch <- tr.Style.VertAlign = align
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()
        allMatch

let private textRunsInRangeHaveColor start len color buffer =
    if len <= 0 then
        true
    else
        let rangeEnd = start + len
        let mutable allMatch = true
        let mutable remaining = buffer.Runs
        while allMatch && not remaining.IsEmpty do
            match remaining with
            | Text tr :: tail ->
                let runStart = tr.Start
                let runEnd = tr.Start + tr.Length
                if runEnd > start && runStart < rangeEnd then
                    allMatch <- tr.Style.Color = color
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()
        allMatch

let private textRunsInRangeHaveHighlight start len color buffer =
    if len <= 0 then
        true
    else
        let rangeEnd = start + len
        let mutable allMatch = true
        let mutable remaining = buffer.Runs
        while allMatch && not remaining.IsEmpty do
            match remaining with
            | Text tr :: tail ->
                let runStart = tr.Start
                let runEnd = tr.Start + tr.Length
                if runEnd > start && runStart < rangeEnd then
                    allMatch <- tr.Style.Highlight = color
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()
        allMatch

let private textRunsInRangeHaveHyperlink start len target buffer =
    if len <= 0 then
        true
    else
        let rangeEnd = start + len
        let mutable allMatch = true
        let mutable remaining = buffer.Runs
        while allMatch && not remaining.IsEmpty do
            match remaining with
            | Text tr :: tail ->
                let runStart = tr.Start
                let runEnd = tr.Start + tr.Length
                if runEnd > start && runStart < rangeEnd then
                    allMatch <- tr.Style.Hyperlink = target
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()
        allMatch

let private textRunsInRangeHaveCharacterStyle start len styleId buffer =
    if len <= 0 then
        true
    else
        let rangeEnd = start + len
        let mutable allMatch = true
        let mutable remaining = buffer.Runs
        while allMatch && not remaining.IsEmpty do
            match remaining with
            | Text tr :: tail ->
                let runStart = tr.Start
                let runEnd = tr.Start + tr.Length
                if runEnd > start && runStart < rangeEnd then
                    allMatch <- tr.Style.Style = styleId
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()
        allMatch

let private textRunsInRangeHaveFontSize start len size buffer =
    if len <= 0 then
        true
    else
        let rangeEnd = start + len
        let mutable allMatch = true
        let mutable remaining = buffer.Runs
        while allMatch && not remaining.IsEmpty do
            match remaining with
            | Text tr :: tail ->
                let runStart = tr.Start
                let runEnd = tr.Start + tr.Length
                if runEnd > start && runStart < rangeEnd then
                    allMatch <- tr.Style.FontSize = size && ExplicitFormat.has ExplicitFormat.FontSize tr.Style
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()
        allMatch

let private textRunsInRangeHaveFontFamily start len family buffer =
    if len <= 0 then
        true
    else
        let rangeEnd = start + len
        let mutable allMatch = true
        let mutable remaining = buffer.Runs
        while allMatch && not remaining.IsEmpty do
            match remaining with
            | Text tr :: tail ->
                let runStart = tr.Start
                let runEnd = tr.Start + tr.Length
                if runEnd > start && runStart < rangeEnd then
                    match tr.Style.FontFamily with
                    | Some existingFamily -> allMatch <- existingFamily = family
                    | None -> allMatch <- false
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()
        allMatch

/// True when every text run overlapping the range satisfies the predicate.
let rangeSatisfies start len (predicate: TextStyle -> bool) buffer =
    textRunsInRangeSatisfy start len predicate buffer

/// Toggle an on/off character property the way Word does: if the whole
/// range already has it, clear it; otherwise set it on the whole range.
/// (Flipping each run separately would turn "bold plain" into "plain bold".)
let private toggleProperty start len (isSet: TextStyle -> bool) (apply: bool -> TextStyle -> TextStyle) buffer =
    let turnOn = not (textRunsInRangeSatisfy start len isSet buffer)
    applyStyle start len (apply turnOn) buffer

/// Toggle bold for range
let toggleBold start len buffer =
    toggleProperty start len (fun style -> style.Bold) withExplicitBold buffer

/// Toggle italic for range
let toggleItalic start len buffer =
    toggleProperty start len (fun style -> style.Italic) withExplicitItalic buffer

/// Toggle underline for range
let toggleUnderline start len buffer =
    toggleProperty start len (fun style -> style.Underline) withExplicitUnderline buffer

/// Toggle all caps for range
let toggleAllCaps start len buffer =
    toggleProperty start len (fun style -> style.AllCaps) withExplicitAllCaps buffer

/// Toggle small caps for range
let toggleSmallCaps start len buffer =
    toggleProperty start len (fun style -> style.SmallCaps) withExplicitSmallCaps buffer

/// Toggle strikethrough for range
let toggleStrike start len buffer =
    toggleProperty start len (fun style -> style.Strike) withExplicitStrike buffer

/// Set vertical alignment for range
let setVertAlign start len align buffer =
    if textRunsInRangeHaveVertAlign start len align buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with VertAlign = align }) buffer

/// Set text color for range
let setColor start len color buffer =
    if textRunsInRangeHaveColor start len color buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with Color = color }) buffer

/// Set highlight color for range
let setHighlight start len color buffer =
    if textRunsInRangeHaveHighlight start len color buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with Highlight = color }) buffer

/// Set hyperlink target for range
let setHyperlink start len (target: string option) buffer =
    match target with
    // Only a target that may be followed becomes a link; anything else is kept
    // on the text as an inactive target.
    | Some value when not (isActiveHyperlinkTarget value) ->
        applyStyle start len (fun style -> { style with Hyperlink = None; InactiveHyperlink = target }) buffer
    | _ ->
        if textRunsInRangeHaveHyperlink start len target buffer then
            applyStyle start len (fun style -> if style.InactiveHyperlink.IsSome then { style with InactiveHyperlink = None } else style) buffer
        else
            applyStyle start len (fun style -> { style with Hyperlink = target; InactiveHyperlink = None }) buffer

/// Set comment metadata for range
let setComment start len commentId commentText author date buffer =
    if textRunsInRangeSatisfy start len (fun style ->
        style.CommentId = commentId
            && style.CommentText = commentText
            && style.CommentAuthor = author
            && style.CommentDate = date) buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with CommentId = commentId; CommentText = commentText; CommentAuthor = author; CommentDate = date }) buffer

/// Set footnote/endnote metadata for range
let setNote start len noteId noteType noteText buffer =
    if textRunsInRangeSatisfy start len (fun style ->
        style.NoteId = noteId
            && style.NoteType = noteType
            && style.NoteText = noteText) buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with NoteId = noteId; NoteType = noteType; NoteText = noteText }) buffer

/// Set bookmark markers for range
let setBookmark start len bookmarkId bookmarkName buffer =
    applyStyle start len (fun style -> { style with BookmarkStartId = Some bookmarkId; BookmarkName = Some bookmarkName; BookmarkEndId = Some bookmarkId }) buffer

/// Set named character style for range
let setCharacterStyle start len styleId buffer =
    if textRunsInRangeHaveCharacterStyle start len styleId buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with Style = styleId }) buffer

/// Set revision metadata for range
let setRevision start len revisionType revisionId author date buffer =
    if textRunsInRangeSatisfy start len (fun style ->
        style.RevisionType = Some revisionType
            && style.RevisionId = Some revisionId
            && style.RevisionAuthor = author
            && style.RevisionDate = date) buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with RevisionType = Some revisionType; RevisionId = Some revisionId; RevisionAuthor = author; RevisionDate = date; DeletedInsertion = None; MoveName = None }) buffer

/// Record the deletion of a range as a tracked change. Ordinary text is kept
/// and marked as deleted. Text that is itself a pending insertion by the same
/// author is simply removed: it was never part of the original, so marking it
/// "deleted" would make Reject All resurrect it. Text already marked deleted
/// keeps its original deletion record.
let trackDelete start len revisionId author date buffer =
    if len <= 0 then buffer
    else
        let rangeEnd = start + len
        let isOwnInsertion (style: TextStyle) =
            isAdditionRevision style && style.RevisionAuthor = author
        let removals = ResizeArray<struct (int * int)>()
        for run in buffer.Runs do
            match run with
            | Text tr when isOwnInsertion tr.Style ->
                let overlapStart = max tr.Start start
                let overlapEnd = min (tr.Start + tr.Length) rangeEnd
                if overlapStart < overlapEnd then
                    removals.Add(struct (overlapStart, overlapEnd - overlapStart))
            | _ -> ()
        let mutable result =
            applyStyle start len (fun style ->
                if isOwnInsertion style || isRemovalRevision style then style
                elif isAdditionRevision style then
                    // Another author's pending insertion: it stays an insertion and is
                    // deleted inside it (w:del in w:ins). Relabelling it a plain deletion
                    // made Reject All bring the text back as if it had always been there.
                    { style with
                        RevisionType = Some "delete"; RevisionId = Some revisionId; RevisionAuthor = author; RevisionDate = date
                        DeletedInsertion = Some { Id = style.RevisionId; Author = style.RevisionAuthor; Date = style.RevisionDate }
                        MoveName = None }
                else { style with RevisionType = Some "delete"; RevisionId = Some revisionId; RevisionAuthor = author; RevisionDate = date }) buffer
        for index = removals.Count - 1 downto 0 do
            let struct (removeStart, removeLength) = removals.[index]
            result <- delete removeStart removeLength result
        result

/// Clear revision metadata for range
let clearRevision start len buffer =
    if textRunsInRangeSatisfy start len (fun style ->
        Option.isNone style.RevisionType
            && Option.isNone style.RevisionId
            && Option.isNone style.RevisionAuthor
            && Option.isNone style.RevisionDate) buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with RevisionType = None; RevisionId = None; RevisionAuthor = None; RevisionDate = None; DeletedInsertion = None; MoveName = None }) buffer

let private clearRevisionStyle style =
    { style with RevisionType = None; RevisionId = None; RevisionAuthor = None; RevisionDate = None; DeletedInsertion = None; MoveName = None }

/// Resolve revision markup in one pass. Accepted insertions / rejected deletions
/// become normal text; accepted deletions / rejected insertions are removed.
let resolveRevisionsInRange accept selStart selEnd isCollapsed buffer =
    let intersects start endPos =
        if isCollapsed then selStart >= start && selStart <= endPos
        else start < selEnd && endPos > selStart
    // Accepting removes what was deleted or moved away, rejecting what was
    // inserted or moved in. A deleted insertion goes either way: accepted, the
    // deletion stands; rejected, the insertion never happened.
    let shouldDeleteRun (style: TextStyle) =
        style.DeletedInsertion.IsSome
        || (if accept then isRemovalRevision style else isAdditionRevision style)
    let mutable changed = false
    let mutable firstStart = None
    let mutable deletesText = false
    let mutable keepsContent = false
    let mutable remaining = buffer.Runs
    while not remaining.IsEmpty do
        match remaining with
        | Text tr :: tail ->
            match tr.Style.RevisionType with
            | Some _ when intersects tr.Start (tr.Start + tr.Length) ->
                if not changed then
                    changed <- true
                    firstStart <- Some tr.Start
                if shouldDeleteRun tr.Style then deletesText <- true
                else keepsContent <- keepsContent || tr.Length > 0
                remaining <- tail
            | _ ->
                keepsContent <- keepsContent || tr.Length > 0
                remaining <- tail
        | Image _ :: tail ->
            keepsContent <- true
            remaining <- tail
        | [] -> ()
    if not changed then
        struct (buffer, None, false)
    elif deletesText && not keepsContent then
        struct (empty, firstStart, true)
    elif not deletesText then
        let clearedStyles = ResizeArray<TextStyle * TextStyle>()
        let clearCached style =
            let mutable found = ValueNone
            let mutable index = 0
            while ValueOption.isNone found && index < clearedStyles.Count do
                let original, cleared = clearedStyles.[index]
                if original = style then
                    found <- ValueSome cleared
                index <- index + 1
            match found with
            | ValueSome cleared -> cleared
            | ValueNone ->
                let cleared = clearRevisionStyle style
                clearedStyles.Add(style, cleared)
                cleared
        let newRuns = ResizeArray<Run>()
        let mutable hasPendingText = false
        let mutable pendingStart = 0
        let mutable pendingLength = 0
        let mutable pendingStyle = defaultStyle
        let sameStyle left right = obj.ReferenceEquals(left, right) || left = right
        let flushPendingText () =
            if hasPendingText then
                newRuns.Add(Text { Start = pendingStart; Length = pendingLength; Style = pendingStyle })
                hasPendingText <- false
        let addText start length style =
            if length > 0 then
                if hasPendingText && pendingStart + pendingLength = start && sameStyle pendingStyle style then
                    pendingLength <- pendingLength + length
                else
                    flushPendingText ()
                    pendingStart <- start
                    pendingLength <- length
                    pendingStyle <- style
                    hasPendingText <- true
        let addRun run =
            match run with
            | Text tr -> addText tr.Start tr.Length tr.Style
            | Image _ ->
                flushPendingText ()
                newRuns.Add run
        for run in buffer.Runs do
            match run with
            | Text tr ->
                match tr.Style.RevisionType with
                | Some _ when intersects tr.Start (tr.Start + tr.Length) ->
                    addText tr.Start tr.Length (clearCached tr.Style)
                | _ ->
                    addRun run
            | Image _ ->
                addRun run
        flushPendingText ()
        struct ({ buffer with Runs = resizeArrayToList newRuns }, firstStart, true)
    else
        let clearedStyles = ResizeArray<TextStyle * TextStyle>()
        let clearCached style =
            let mutable found = ValueNone
            let mutable index = 0
            while ValueOption.isNone found && index < clearedStyles.Count do
                let original, cleared = clearedStyles.[index]
                if original = style then
                    found <- ValueSome cleared
                index <- index + 1
            match found with
            | ValueSome cleared -> cleared
            | ValueNone ->
                let cleared = clearRevisionStyle style
                clearedStyles.Add(style, cleared)
                cleared
        let originalText = toString buffer
        let text = System.Text.StringBuilder(originalText.Length)
        let newRuns = ResizeArray<Run>()
        let mutable outPos = 0
        let mutable hasPendingText = false
        let mutable pendingStart = 0
        let mutable pendingLength = 0
        let mutable pendingStyle = defaultStyle
        let sameStyle left right = obj.ReferenceEquals(left, right) || left = right
        let flushPendingText () =
            if hasPendingText then
                newRuns.Add(Text { Start = pendingStart; Length = pendingLength; Style = pendingStyle })
                hasPendingText <- false
        let addBufferedText start length style =
            if length > 0 then
                if hasPendingText && pendingStart + pendingLength = start && sameStyle pendingStyle style then
                    pendingLength <- pendingLength + length
                else
                    flushPendingText ()
                    pendingStart <- start
                    pendingLength <- length
                    pendingStyle <- style
                    hasPendingText <- true
        let appendOriginal start len =
            if len > 0 then
#if FABLE_COMPILER
                text.Append(originalText.Substring(start, len)) |> ignore
#else
                text.Append(originalText, start, len) |> ignore
#endif
        let addTextRun (tr: TextRun) style =
            if tr.Length > 0 then
                appendOriginal tr.Start tr.Length
                addBufferedText outPos tr.Length style
                outPos <- outPos + tr.Length
        let addImageRun (ir: ImageRun) =
            flushPendingText ()
            text.Append(objectReplacementChar) |> ignore
            newRuns.Add(Image { ir with Start = outPos })
            outPos <- outPos + 1
        for run in buffer.Runs do
            match run with
            | Text tr ->
                match tr.Style.RevisionType with
                | Some _ when intersects tr.Start (tr.Start + tr.Length) ->
                    if firstStart.IsNone then
                        firstStart <- Some tr.Start
                    if not (shouldDeleteRun tr.Style) then
                        addTextRun tr (clearCached tr.Style)
                | _ ->
                    addTextRun tr tr.Style
            | Image ir ->
                addImageRun ir
        flushPendingText ()
        struct ({ Rope = Rope.ofString (text.ToString()); Runs = resizeArrayToList newRuns }, firstStart, true)

/// Set font size for range
let setFontSize start len size buffer =
    if textRunsInRangeHaveFontSize start len size buffer then
        buffer
    else
        applyStyle start len (withExplicitFontSize size) buffer

/// Set font family for range
let setFontFamily start len family buffer =
    if textRunsInRangeHaveFontFamily start len family buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with FontFamily = Some family }) buffer

/// Clear direct character formatting for range while preserving annotation targets.
let clearFormatting start len buffer =
    applyStyle start len (fun style ->
        {
            defaultStyle with
                Hyperlink = style.Hyperlink
                InactiveHyperlink = style.InactiveHyperlink
                CommentId = style.CommentId
                CommentText = style.CommentText
                CommentAuthor = style.CommentAuthor
                CommentDate = style.CommentDate
                NoteId = style.NoteId
                NoteType = style.NoteType
                NoteText = style.NoteText
                BookmarkStartId = style.BookmarkStartId
                BookmarkName = style.BookmarkName
                BookmarkEndId = style.BookmarkEndId
                FieldCode = style.FieldCode
                FieldResult = style.FieldResult
                RenderedPageBreakBefore = style.RenderedPageBreakBefore
                EquationXml = style.EquationXml
                RevisionId = style.RevisionId
                RevisionType = style.RevisionType
                RevisionAuthor = style.RevisionAuthor
                RevisionDate = style.RevisionDate
                DeletedInsertion = style.DeletedInsertion
                MoveName = style.MoveName
        }) buffer

/// Get style at position
let getStyleAt pos buffer =
    styleAt pos buffer

/// Convert position to line/column
let posToLineCol pos buffer =
    Rope.positionToLineCol pos buffer.Rope

/// Convert line/column to position
let lineColToPos line col buffer =
    Rope.lineColToPosition line col buffer.Rope
