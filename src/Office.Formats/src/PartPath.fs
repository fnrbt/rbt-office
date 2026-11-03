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
