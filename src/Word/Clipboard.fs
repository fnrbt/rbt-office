module RedBat.Word.Clipboard

#if FABLE_COMPILER
open Fable.Core
open Fable.Core.JsInterop

/// Copy text to clipboard
[<Emit("navigator.clipboard.writeText($0)")>]
let copyText (text: string) : JS.Promise<unit> = jsNative

/// Read text from clipboard
[<Emit("navigator.clipboard.readText()")>]
let pasteText () : JS.Promise<string> = jsNative

#else
// Server-side stubs
let copyText (text: string) = async { return () }
let pasteText () = async { return "" }
#endif
