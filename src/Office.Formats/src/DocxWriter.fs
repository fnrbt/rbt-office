module DocxWriter

open Fable.Core.JsInterop
open XmlGenerator
open ZipWriter
open ZipInterop
open System.Text

let private appendEscaped (sb: StringBuilder) (text: string) =
    for ch in text do
        match ch with
        | '&' -> sb.Append("&amp;") |> ignore
        | '<' -> sb.Append("&lt;") |> ignore
        | '>' -> sb.Append("&gt;") |> ignore
        | '"' -> sb.Append("&quot;") |> ignore
        | '\'' -> sb.Append("&apos;") |> ignore
        | _ -> sb.Append(ch) |> ignore

/// Generate minimal document.xml with plain text paragraphs
let generateDocumentXml (paragraphs: string list) : string =
    let sb = StringBuilder()
    sb.Append(declaration).Append('\n') |> ignore
    sb.Append("<w:document xmlns:w=\"").Append(Namespaces.word)
        .Append("\" xmlns:r=\"").Append(Namespaces.relationships).Append("\">") |> ignore
    sb.Append("<w:body>") |> ignore
    for text in paragraphs do
        sb.Append("<w:p><w:r><w:t xml:space=\"preserve\">") |> ignore
        appendEscaped sb text
        sb.Append("</w:t></w:r></w:p>") |> ignore
    sb.Append("</w:body></w:document>") |> ignore
    sb.ToString()

/// Generate [Content_Types].xml
let generateContentTypes () : string =
    let children = [
        selfClosing "Default" [("Extension", "rels"); ("ContentType", "application/vnd.openxmlformats-package.relationships+xml")]
        selfClosing "Default" [("Extension", "xml"); ("ContentType", "application/xml")]
        selfClosing "Override" [("PartName", "/word/document.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")]
        selfClosing "Override" [("PartName", "/word/styles.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")]
    ]
    let root = element "Types" [("xmlns", Namespaces.contentTypes)] children
    document root

/// Generate _rels/.rels
let generateRootRels () : string =
    let rel = selfClosing "Relationship" [("Id", "rId1"); ("Type", RelationshipTypes.officeDocument); ("Target", "word/document.xml")]
    let root = element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] [rel]
    document root

/// Generate word/_rels/document.xml.rels
let generateDocumentRels () : string =
    let rel = selfClosing "Relationship" [("Id", "rId1"); ("Type", RelationshipTypes.styles); ("Target", "styles.xml")]
    let root = element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] [rel]
    document root

/// Generate minimal styles.xml
let generateDefaultStyles () : string =
    let rPr = element "w:rPr" [] [
        selfClosing "w:rFonts" [("w:ascii", "Calibri"); ("w:hAnsi", "Calibri"); ("w:cs", "Calibri")]
        selfClosing "w:sz" [("w:val", "22")]
        selfClosing "w:szCs" [("w:val", "22")]
    ]
    let rPrDefault = element "w:rPrDefault" [] [rPr]
    let pPrDefault = element "w:pPrDefault" [] []
    let docDefaults = element "w:docDefaults" [] [rPrDefault; pPrDefault]
    let root = element "w:styles" [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships)] [docDefaults]
    document root

/// Create minimal DOCX from plain text paragraphs
let writeMinimalDocx (paragraphs: string list) : Fable.Core.JS.Promise<obj> =
    let files = [
        { Path = "[Content_Types].xml"; Content = generateContentTypes() }
        { Path = "_rels/.rels"; Content = generateRootRels() }
        { Path = "word/document.xml"; Content = generateDocumentXml paragraphs }
        { Path = "word/_rels/document.xml.rels"; Content = generateDocumentRels() }
        { Path = "word/styles.xml"; Content = generateDefaultStyles() }
    ]

    createPackage files
