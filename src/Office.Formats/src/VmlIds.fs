/// The ids of the shapes of a VML drawing (`_x0000_s1025`: shape 1025) and of its block of ids
/// (`<o:idmap data="1">`: shapes 1025 to 2048).
module VmlIds

let private shapePrefix = "_x0000_s"

/// The number of a shape id ("_x0000_s1025" -> 1025); None for another value.
let shapeNumber (value: string) : int option =
    if value.Length > shapePrefix.Length && value.StartsWith shapePrefix then
        match System.Int32.TryParse (value.Substring shapePrefix.Length) with
        | true, number -> Some number
        | _ -> None
    else None

/// The first block of an idmap's `data` ("1", "1,5" -> 1); None for none.
let firstBlock (data: string) : int option =
    let digits = data |> Seq.skipWhile System.Char.IsWhiteSpace |> Seq.takeWhile System.Char.IsDigit |> Seq.toArray |> System.String
    match System.Int32.TryParse digits with
    | true, number -> Some number
    | _ -> None
