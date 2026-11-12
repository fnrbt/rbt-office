/// Values the editor's package hands the writer, read into what they say: the bookmarks of a run (typed),
/// a run's text with its tabs and breaks, link targets, picture sources, colours, the keys of numbering
/// levels. (XML is not among them: that is XmlPatch's.)
module ModelText

// ---- bookmarks ----
// A run has all the bookmarks that start before it and end after it, in the order of the file: the reader
// reads them into these, the editor's package carries them as arrays of objects (a run's BookmarkStarts and
// BookmarkEnds), and the writer writes them from those. (Typed: no text to take apart.)

/// A bookmark that starts before a run (its name "" when the file gives none); `EndsHere`: it ends where it
/// starts, before the run too.
type BookmarkStart = { Id: string; Name: string; EndsHere: bool }

/// What ends after a run: a bookmark's end, or a bookmark opened there - after the last run of a paragraph
/// (`EndsHere`: it ends there too; otherwise further on).
type BookmarkEnd =
    | Ends of id: string
    | Opens of BookmarkStart

/// Zero-width Word complex-field tokens, in document order. Unlike a simple
/// field, a complex field can enclose several paragraphs and other fields.
type FieldBoundary =
    | Begin
    | Instruction of string
    | Separate
    | End

// ---- a run's text ----

/// A piece of a run's text: text, a tab ("\t"), a break ("\n", "\r", "\r\n").
type TextPiece =
    | Text of string
    | Tab
    | Break

/// A run's text in its pieces.
let textPieces (text: string) : TextPiece list =
    let pieces = ResizeArray<TextPiece>()
    let current = System.Text.StringBuilder()
    let flush () =
        if current.Length > 0 then
            pieces.Add (Text (current.ToString()))
            current.Clear() |> ignore
    let mutable i = 0
    while i < text.Length do
        match text.[i] with
        | '\t' -> flush (); pieces.Add Tab
        | '\n' -> flush (); pieces.Add Break
        | '\r' ->
            flush ()
            pieces.Add Break
            if i + 1 < text.Length && text.[i + 1] = '\n' then i <- i + 1
        | c -> current.Append c |> ignore
        i <- i + 1
    flush ()
    List.ofSeq pieces

/// Whether a run's text is what the editor shows of a field without a result ("[Field:<code>]").
let isFieldPlaceholder (text: string) : bool = text.StartsWith "[Field:" && text.EndsWith "]"

/// A text held in pieces (`pieces`, which together are `was`) changed to `now`: the piece the change is in
/// and what that piece is now - None where the change is not within one piece (or the pieces are not `was`).
let changeInPieces (pieces: string list) (was: string) (now: string) : (int * string) option =
    if String.concat "" pieces <> was then None
    else
        // (What is the same at the start and at the end; the change is between.)
        let common = Seq.zip was now |> Seq.takeWhile (fun (a, b) -> a = b) |> Seq.length
        let room = min (was.Length - common) (now.Length - common)
        let same = Seq.zip (Seq.rev was) (Seq.rev now) |> Seq.truncate room |> Seq.takeWhile (fun (a, b) -> a = b) |> Seq.length
        let changedFrom, changedTo = common, was.Length - same
        let inserted = now.Substring(common, now.Length - same - common)
        let starts = pieces |> List.scan (fun at (piece: string) -> at + piece.Length) 0
        List.zip pieces (List.take pieces.Length starts)
        |> List.mapi (fun index (piece, start) -> index, piece, start)
        // (The piece the change is in, from its start to its end; an insertion where two meet is the first's.)
        |> List.tryFind (fun (_, piece, start) -> changedFrom >= start && changedTo <= start + piece.Length && (changedFrom < start + piece.Length || changedFrom = changedTo))
        |> Option.map (fun (index, piece, start) -> index, piece.Substring(0, changedFrom - start) + inserted + piece.Substring(changedTo - start))

/// The text put in before a text (`now` is it and then `was`), when that is all that changed.
let insertedBefore (was: string) (now: string) : string option =
    if now.Length > was.Length && now.EndsWith was then Some (now.Substring(0, now.Length - was.Length)) else None

// ---- links, pictures, colours, lists ----

/// The bookmark a link names in the document itself ("#name" -> "name"); None for a link elsewhere.
let anchorOf (target: string) : string option =
    if target.Length > 1 && target.[0] = '#' then Some (target.Substring 1) else None

/// Whether a link target is a place in the document itself ("#name").
let isAnchor (target: string) : bool = target.Length > 0 && target.[0] = '#'

/// A picture's source as a data URL: its media type, and its data when it is base64.
type DataUrl = { MediaType: string; Base64: string option }

let dataUrl (src: string) : DataUrl option =
    if not (src.StartsWith "data:") then None
    else
        match src.Split([| ',' |], 2) with
        | [| meta; data |] ->
            let parameters = (meta.Split([| ':' |], 2) |> Array.last).Split(';') |> Array.toList
            let mediaType = List.tryHead parameters |> Option.defaultValue ""
            Some { MediaType = mediaType; Base64 = if parameters |> List.exists (fun p -> p.ToLowerInvariant() = "base64") then Some data else None }
        | _ -> None

/// Whether a picture's source is a file elsewhere (a URL, not data in the page).
let isLinkedSource (src: string) : bool =
    src.Contains "://" && not (src.StartsWith "data:") && not (src.StartsWith "blob:")

/// A CSS colour (#RRGGBB) as WordprocessingML writes one (RRGGBB).
let wordColor (color: string) : string = color.TrimStart('#').ToUpper()

/// A numbering level the editor names ("list id/level", levels from one): the list's id and the level
/// (from zero; -1 for a key with no level).
let numberingLevel (key: string) : string * int =
    match key.Split('/') |> Array.toList |> List.rev with
    | level :: (_ :: _ as list) ->
        let numId = list |> List.rev |> String.concat "/"
        numId, (match System.Int32.TryParse level with | true, value -> value - 1 | _ -> -1)
    | _ -> key, -1

// ---- names of attributes ----

/// The prefix of a qualified name ("" for none).
let prefixOf (name: string) : string =
    match name.Split(':') with
    | [| prefix; _ |] -> prefix
    | _ -> ""

/// Whether an attribute is a revision id (w:rsidR, w:rsidRPr ...).
let isRevisionId (name: string) : bool = name.StartsWith "w:rsid"

/// Whether an attribute says in other units what another says (w:afterLines beside w:after, w:leftChars beside w:left).
let isInOtherUnits (name: string) : bool = name.EndsWith "Lines" || name.EndsWith "Chars"
