/// Colours of the model as text ("#FF0000", "theme:4:0.4", "indexed:10") and as the file format writes them.
module ColorText

/// A colour of the model, read.
type Color =
    /// An RGB colour (six hex digits, without the "#").
    | Rgb of string
    /// A colour of the theme, with its tint.
    | Theme of index: string * tint: string option
    /// A colour of the legacy palette.
    | Indexed of string

/// A colour of the model ("#RRGGBB", "theme:index[:tint]", "indexed:n"); None for another text.
let parse (color: string) : Color option =
    if color.StartsWith "#" && color.Length = 7 then Some (Rgb (color.Substring 1))
    elif color.StartsWith "theme:" then
        match color.Split(':') with
        | [| _; index |] -> Some (Theme (index, None))
        | [| _; index; tint |] -> Some (Theme (index, Some tint))
        | _ -> None
    elif color.StartsWith "indexed:" then Some (Indexed (color.Substring 8))
    else None

/// An RGB colour of the model as the file's ARGB ("#ff0000" -> "FFFF0000"; "ff0000" too).
let argb (color: string) : string =
    if color.StartsWith "#" then "FF" + color.Substring(1).ToUpper() else "FF" + color.ToUpper()
