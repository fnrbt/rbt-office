/// Paths of the parts of a package ("word/media/image1.png"), by their segments.
module PartPath

let private segments (path: string) = path.Split('/') |> Array.toList

/// The directory of a part ("word/media/image1.png" -> "word/media"; "" at the root).
let directory (path: string) : string =
    match List.rev (segments path) with
    | _ :: rest -> rest |> List.rev |> String.concat "/"
    | [] -> ""

/// The name of a part, without its directory ("word/media/image1.png" -> "image1.png").
let fileName (path: string) : string = segments path |> List.last

/// The extension of a part's name, lower case, without its dot ("" for none).
let extension (path: string) : string =
    match (fileName path).Split('.') with
    | parts when parts.Length > 1 -> parts.[parts.Length - 1].ToLowerInvariant()
    | _ -> ""

/// The name of a part without its extension ("document.xml.rels" -> "document.xml").
let private withoutExtension (name: string) =
    match name.Split('.') |> Array.toList |> List.rev with
    | _ :: (_ :: _ as rest) -> rest |> List.rev |> String.concat "."
    | _ -> name

/// The part a relationships part is of ("word/_rels/document.xml.rels" -> "word/document.xml"; "_rels/.rels"
/// -> ""): None when the path is not of a relationships part.
let ofRelationships (path: string) : string option =
    match List.rev (segments path) with
    | name :: "_rels" :: directory when extension name = "rels" ->
        let owner = withoutExtension name
        Some ((List.rev directory @ (if owner = "" then [] else [ owner ])) |> String.concat "/")
    | _ -> None

/// The relationships part of a part ("word/document.xml" -> "word/_rels/document.xml.rels").
let relationshipsOf (path: string) : string =
    let dir = directory path
    (if dir = "" then "" else dir + "/") + "_rels/" + fileName path + ".rels"

/// Whether a part is in a directory, at any depth ("xl/media/image1.png" is in "xl").
let isIn (directory: string) (path: string) : bool =
    let inside = segments directory
    let all = segments path
    all.Length > inside.Length && List.take inside.Length all = inside

/// A part's path relative to a directory: the target that names it from a part there ("xl/worksheets",
/// "xl/comments1.xml" -> "../comments1.xml").
let relative (fromDirectory: string) (path: string) : string =
    let from = if fromDirectory = "" then [] else segments fromDirectory
    let all = segments path
    let rec common (a: string list) (b: string list) = match a, b with x :: xs, y :: ys when x = y -> common xs ys | _ -> a, b
    let up, down = common from all
    (List.replicate up.Length ".." @ down) |> String.concat "/"

/// The number of a part named by its kind and a number ("xl/tables/table12.xml", "table" -> 12): None for
/// a part not named so.
let number (stem: string) (path: string) : int option =
    let name = withoutExtension (fileName path)
    if name.Length > stem.Length && name.StartsWith stem then
        match System.Int32.TryParse (name.Substring stem.Length) with
        | true, number -> Some number
        | _ -> None
    else None
