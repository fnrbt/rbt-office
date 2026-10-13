/// Translate X11 keycodes to character strings
module KeyMap

/// Translate X11 keycode + shift state to a key string.
/// Returns single-char strings for printable keys, or names like "Enter", "Backspace".
let translate (keycode: int) (shift: bool) : string option =
    match keycode with
    // Letters a-z (keycodes 24-33 = qwertyuiop, 38-46 = asdfghjkl, 52-58 = zxcvbnm)
    | 24 -> Some (if shift then "Q" else "q")
    | 25 -> Some (if shift then "W" else "w")
    | 26 -> Some (if shift then "E" else "e")
    | 27 -> Some (if shift then "R" else "r")
    | 28 -> Some (if shift then "T" else "t")
    | 29 -> Some (if shift then "Y" else "y")
    | 30 -> Some (if shift then "U" else "u")
    | 31 -> Some (if shift then "I" else "i")
    | 32 -> Some (if shift then "O" else "o")
    | 33 -> Some (if shift then "P" else "p")
    | 38 -> Some (if shift then "A" else "a")
    | 39 -> Some (if shift then "S" else "s")
    | 40 -> Some (if shift then "D" else "d")
    | 41 -> Some (if shift then "F" else "f")
    | 42 -> Some (if shift then "G" else "g")
    | 43 -> Some (if shift then "H" else "h")
    | 44 -> Some (if shift then "J" else "j")
    | 45 -> Some (if shift then "K" else "k")
    | 46 -> Some (if shift then "L" else "l")
    | 52 -> Some (if shift then "Z" else "z")
    | 53 -> Some (if shift then "X" else "x")
    | 54 -> Some (if shift then "C" else "c")
    | 55 -> Some (if shift then "V" else "v")
    | 56 -> Some (if shift then "B" else "b")
    | 57 -> Some (if shift then "N" else "n")
    | 58 -> Some (if shift then "M" else "m")
    // Numbers 1-9,0 (keycodes 10-19)
    | 10 -> Some (if shift then "!" else "1")
    | 11 -> Some (if shift then "@" else "2")
    | 12 -> Some (if shift then "#" else "3")
    | 13 -> Some (if shift then "$" else "4")
    | 14 -> Some (if shift then "%" else "5")
    | 15 -> Some (if shift then "^" else "6")
    | 16 -> Some (if shift then "&" else "7")
    | 17 -> Some (if shift then "*" else "8")
    | 18 -> Some (if shift then "(" else "9")
    | 19 -> Some (if shift then ")" else "0")
    // Special keys
    | 22 -> Some "Backspace"
    | 23 -> Some "Tab"
    | 36 -> Some "Enter"
    | 9  -> Some "Escape"
    | 65 -> Some " "
    // Arrow keys
    | 113 -> Some "ArrowLeft"
    | 114 -> Some "ArrowRight"
    | 111 -> Some "ArrowUp"
    | 116 -> Some "ArrowDown"
    // Editing
    | 119 -> Some "Delete"
    | 110 -> Some "Home"
    | 115 -> Some "End"
    // Page navigation
    | 112 -> Some "PageUp"
    | 117 -> Some "PageDown"
    // Punctuation
    | 20 -> Some (if shift then "_" else "-")
    | 21 -> Some (if shift then "+" else "=")
    | 34 -> Some (if shift then "{" else "[")
    | 35 -> Some (if shift then "}" else "]")
    | 47 -> Some (if shift then ":" else ";")
    | 48 -> Some (if shift then "\"" else "'")
    | 49 -> Some (if shift then "~" else "`")
    | 51 -> Some (if shift then "|" else "\\")
    | 59 -> Some (if shift then "<" else ",")
    | 60 -> Some (if shift then ">" else ".")
    | 61 -> Some (if shift then "?" else "/")
    | _ -> None

/// Detect if a keycode is a modifier key (shift, ctrl, alt, super)
let isModifier (keycode: int) : bool =
    match keycode with
    | 50 | 62 -> true   // Left/Right Shift
    | 37 | 105 -> true  // Left/Right Ctrl
    | 64 | 108 -> true  // Left/Right Alt
    | 133 | 134 -> true // Super
    | _ -> false
