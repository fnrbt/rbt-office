/// A DOCX file read whole as the editor package (Main's `processors.forEditorPackage`): a module of its
/// own, apart from Main (which loads the XLSX reader and the writers too), so that an editor that opens a
/// document loads the DOCX reader and no more (WordHost's dynamic import).
module WordPackageOpen

#if FABLE_COMPILER
open Fable.Core
open Fable.Core.JsInterop
open ZipInterop

/// A copy of the file a package was opened from, when it was given as bytes: kept with what
/// the writer needs (Preserved.OpenedBytes), it is what an unchanged document is saved as
/// (see DocumentWriter.writeFromEditorPackage).
[<Emit("($0 instanceof Uint8Array ? $0.slice() : ($0 instanceof ArrayBuffer ? new Uint8Array($0.slice(0)) : null))")>]
let openedBytesOf (_fileData: obj) : obj = jsNative

let forEditorPackage (fileData: obj) : Fable.Core.JS.Promise<obj> =
    promise {
        let! archive = loadArchive fileData
        let! docOpt = DocxReader.parseDocument archive

        match docOpt with
        | Some doc ->
            let typed = WordEditorPackageReader.ofDocument doc
            let opened = openedBytesOf fileData
            let typed =
                if isNull opened then typed
                else { typed with
                         Preserved = typed.Preserved |> Option.map (fun preserved ->
                             { preserved with OpenedBytes = Some (unbox<byte[]> opened) }) }
            return WordEditorPackage.packageToEditorObject typed
        | None ->
            return createObj [
                "Paragraphs" ==> [||]
                "Styles" ==> [||]
            ]
    }
#endif
