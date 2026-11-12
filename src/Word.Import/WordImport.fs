module Rbt.Word.Import

/// Import an already parsed document using the shared editor-package conversion.
let importDoc (document: DocxReader.WordDocument) : Rbt.Word.Document.Document =
    document
    |> Main.Processors.forTypedEditorPackage
    |> Rbt.Word.TypedDocxImport.importPackage
