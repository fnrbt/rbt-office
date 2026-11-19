/// Types of the DOCX reader (DocxReader) that the editor package (WordEditorPackage) and the Word editor
/// hold: a module of their own, so that what holds them does not load the reader - the editor loads it
/// when it opens a document (WordPackageOpen). DocxReader names them as well (DocxReader.Theme ...).
module DocxTypes

/// An element kept verbatim. Xml is self-contained: the namespace declarations
/// it relies on are on its root element.
type OpaqueObject = {
    Xml: string
    /// chart | diagram | ole | shape | drawing
    Kind: string
    /// A child of w:p (it holds its own runs) rather than of w:r.
    ParagraphLevel: bool
    /// The body's shape import models this one as a text box / shape.
    ShapeCandidate: bool
}

/// Id, author and date of a tracked change.
type RevisionMark = {
    Id: string option
    Author: string option
    Date: string option
}

/// The theme's fonts and colours.
type Theme = {
    MajorFont: string option
    MinorFont: string option
    Colors: Map<string, string>
}
