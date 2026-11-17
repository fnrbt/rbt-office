/// The fonts the package carries (`fonts/` beside the bundles): free faces with the widths of the Office
/// fonts, so a PDF made in the browser has its text when the host supplies no face of its own. Carlito
/// stands in for Calibri, Caladea for Cambria, Liberation Sans / Serif / Mono for Arial, Times New Roman
/// and Courier New (all SIL Open Font License 1.1, `fonts/README.md`). A face the host supplies is always
/// used before these; one of these is fetched only for a face the text needs and the host has not got.
module DefaultFonts

open System

/// One face to fetch: the requested family and style it is set in for, its file under `fonts/`, and the
/// name the PDF gives it.
type PlannedFace = {
    Family: string
    Bold: bool
    Italic: bool
    FaceName: string
    File: string
}

/// The family of `fonts/` that sets text asked for in `family`; None for a symbol font (Symbol,
/// Wingdings, ...), whose characters no text face draws.
let substituteFamily (family: string) : string option =
    let name = if isNull family then "" else family.Trim().ToLowerInvariant()
    match name with
    | "symbol" | "webdings" | "marlett" | "ms outlook" | "bookshelf symbol 7" | "mt extra" -> None
    | _ when name.StartsWith "wingdings" -> None
    | "calibri" | "calibri light" | "carlito" -> Some "Carlito"
    | "cambria" | "cambria math" | "caladea" -> Some "Caladea"
    | "times new roman" | "times" | "tinos" | "liberation serif" | "georgia" | "garamond" | "book antiqua"
    | "palatino" | "palatino linotype" | "constantia" | "century" | "century schoolbook" | "bookman old style"
    | "baskerville old face" | "perpetua" | "sylfaen" | "serif" -> Some "Liberation Serif"
    | "courier new" | "courier" | "cousine" | "liberation mono" | "consolas" | "lucida console"
    | "lucida sans typewriter" | "monospace" -> Some "Liberation Mono"
    | _ -> Some "Liberation Sans"

let private stem (family: string) = family.Replace(" ", "")

let private styleName bold italic =
    if bold && italic then "BoldItalic" elif bold then "Bold" elif italic then "Italic" else "Regular"

/// The face of `fonts/` for text in `family` with that style.
let tryFace (family: string) (bold: bool) (italic: bool) : PlannedFace option =
    substituteFamily family
    |> Option.map (fun substitute ->
        let faceName = stem substitute + "-" + styleName bold italic
        { Family = family; Bold = bold; Italic = italic; FaceName = faceName; File = faceName + ".ttf" })

/// The faces to fetch for text set in `required` ((family, bold, italic), with the default family's
/// regular face always among them: the writers fall back to it), leaving out those `hostHas`.
let plan (hostHas: string -> bool -> bool -> bool) (defaultFamily: string) (required: (string * bool * bool) seq) : PlannedFace list =
    let seen = Collections.Generic.HashSet<string * bool * bool>()
    [ for family, bold, italic in Seq.append [ defaultFamily, false, false ] required do
        if not (String.IsNullOrWhiteSpace family)
           && seen.Add((family.Trim().ToLowerInvariant(), bold, italic))
           && not (hostHas family bold italic) then
            match tryFace family bold italic with
            | Some face -> yield face
            | None -> () ]

/// Whether `faces` ((family, bold, italic)) has that face; families compare without case.
let hasFace (faces: (string * bool * bool) seq) (family: string) (bold: bool) (italic: bool) =
    faces |> Seq.exists (fun (f, b, i) -> b = bold && i = italic && String.Equals(f, family, StringComparison.OrdinalIgnoreCase))

#if FABLE_COMPILER
open Fable.Core

[<Emit("fetch(new URL($1, $0)).then(response => { if (!response.ok) throw new Error('The default font ' + $1 + ' could not be loaded (' + response.status + ')'); return response.arrayBuffer(); }).then(buffer => new Uint8Array(buffer))")>]
let private fetchFont (_baseUrl: string) (_file: string) : JS.Promise<byte[]> = jsNative

[<Emit("Promise.all($0)")>]
let private all (_: JS.Promise<'T> array) : JS.Promise<'T array> = jsNative

[<Emit("$0.then($1)")>]
let private thenMap (_: JS.Promise<'T>) (_: 'T -> 'U) : JS.Promise<'U> = jsNative

[<Emit("$0.catch(error => { $1(); throw error; })")>]
let private onFailure (_: JS.Promise<'T>) (_: unit -> unit) : JS.Promise<'T> = jsNative

// One fetch per file and base: every export after the first reuses the bytes (and the browser's cache
// holds them across page loads). A failed fetch is forgotten, so the next export tries again.
let private loads = Collections.Generic.Dictionary<string, JS.Promise<byte[]>>()

let private load (baseUrl: string) (file: string) =
    let key = baseUrl + "\n" + file
    match loads.TryGetValue key with
    | true, pending -> pending
    | _ ->
        let pending = onFailure (fetchFont baseUrl file) (fun () -> loads.Remove key |> ignore)
        loads.[key] <- pending
        pending

/// Fetch the planned faces from `baseUrl` (the URL of `fonts/`): each with its bytes.
let fetchFaces (baseUrl: string) (faces: PlannedFace list) : JS.Promise<(PlannedFace * byte[]) array> =
    let faces = List.toArray faces
    thenMap (all (faces |> Array.map (fun face -> load baseUrl face.File))) (fun bytes -> Array.zip faces bytes)
#endif
