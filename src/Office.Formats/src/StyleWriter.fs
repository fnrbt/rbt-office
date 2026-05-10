module StyleWriter

open XmlGenerator

/// Style fingerprint for deduplication
type StyleFingerprint = {
    Bold: bool
    Italic: bool
    Underline: bool
    Strike: bool
    FontSize: string
    FontFamily: string
    Color: string option
}

/// Excel cell style fingerprint
type CellStyleFingerprint = {
    Bold: bool option
    Italic: bool option
    Underline: bool option
    FontSize: float option
    FontFamily: string option
    FontColor: string option
    FillColor: string option
}

/// Style manager for deduplication
type StyleManager = {
    mutable Styles: Map<StyleFingerprint, string>  // fingerprint -> style ID
    mutable NextId: int
}

/// Create a new style manager
let createStyleManager () : StyleManager = {
    Styles = Map.empty
    NextId = 1
}

/// Get or create style ID for a fingerprint
let getOrCreateStyleId (manager: StyleManager) (fingerprint: StyleFingerprint) : string =
    match Map.tryFind fingerprint manager.Styles with
    | Some styleId -> styleId
    | None ->
        let styleId = sprintf "style%d" manager.NextId
        manager.Styles <- Map.add fingerprint styleId manager.Styles
        manager.NextId <- manager.NextId + 1
        styleId

/// Generate Word style definition XML
let generateWordStyle (styleId: string) (fingerprint: StyleFingerprint) : string =
    let rPr = []

    let rPr = if fingerprint.Bold then (selfClosing "w:b" []) :: rPr else rPr
    let rPr = if fingerprint.Italic then (selfClosing "w:i" []) :: rPr else rPr
    let rPr = if fingerprint.Underline then (selfClosing "w:u" [("w:val", "single")]) :: rPr else rPr
    let rPr = if fingerprint.Strike then (selfClosing "w:strike" []) :: rPr else rPr

    let rPr =
        (selfClosing "w:sz" [("w:val", fingerprint.FontSize)]) ::
        (selfClosing "w:szCs" [("w:val", fingerprint.FontSize)]) :: rPr

    let rPr =
        (selfClosing "w:rFonts" [("w:ascii", fingerprint.FontFamily); ("w:hAnsi", fingerprint.FontFamily); ("w:cs", fingerprint.FontFamily)]) :: rPr

    let rPr =
        match fingerprint.Color with
        | Some color -> (selfClosing "w:color" [("w:val", color)]) :: rPr
        | None -> rPr

    let rPrElement = element "w:rPr" [] (List.rev rPr)

    element "w:style" [("w:type", "character"); ("w:styleId", styleId)] [
        element "w:name" [("w:val", styleId)] []
        rPrElement
    ]

/// Generate Word styles.xml with deduplicated styles
let generateStylesXml (manager: StyleManager) : string =
    let styleElements =
        manager.Styles
        |> Map.toList
        |> List.map (fun (fingerprint, styleId) -> generateWordStyle styleId fingerprint)
        |> String.concat ""

    let rPr = element "w:rPr" [] [
        selfClosing "w:rFonts" [("w:ascii", "Calibri"); ("w:hAnsi", "Calibri"); ("w:cs", "Calibri")]
        selfClosing "w:sz" [("w:val", "22")]
        selfClosing "w:szCs" [("w:val", "22")]
    ]
    let rPrDefault = element "w:rPrDefault" [] [rPr]
    let pPrDefault = element "w:pPrDefault" [] []
    let docDefaults = element "w:docDefaults" [] [rPrDefault; pPrDefault]

    let root = element "w:styles" [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships)] [docDefaults; styleElements]
    document root

/// Excel style manager
type ExcelStyleManager = {
    mutable Styles: Map<CellStyleFingerprint, int>  // fingerprint -> style index
    mutable SharedStrings: Map<string, int>  // string -> index
    mutable NextStyleIndex: int
    mutable NextStringIndex: int
}

/// Create a new Excel style manager
let createExcelStyleManager () : ExcelStyleManager = {
    Styles = Map.empty
    SharedStrings = Map.empty
    NextStyleIndex = 1
    NextStringIndex = 0
}

/// Get or create style index for a cell style
let getOrCreateCellStyleIndex (manager: ExcelStyleManager) (fingerprint: CellStyleFingerprint) : int option =
    // If all properties are None, return None (no custom style)
    let hasStyle =
        fingerprint.Bold.IsSome || fingerprint.Italic.IsSome || fingerprint.Underline.IsSome ||
        fingerprint.FontSize.IsSome || fingerprint.FontFamily.IsSome ||
        fingerprint.FontColor.IsSome || fingerprint.FillColor.IsSome

    if not hasStyle then
        None
    else
        match Map.tryFind fingerprint manager.Styles with
        | Some styleIndex -> Some styleIndex
        | None ->
            let styleIndex = manager.NextStyleIndex
            manager.Styles <- Map.add fingerprint styleIndex manager.Styles
            manager.NextStyleIndex <- manager.NextStyleIndex + 1
            Some styleIndex

/// Get or create shared string index
let getOrCreateSharedString (manager: ExcelStyleManager) (text: string) : int =
    match Map.tryFind text manager.SharedStrings with
    | Some index -> index
    | None ->
        let index = manager.NextStringIndex
        manager.SharedStrings <- Map.add text index manager.SharedStrings
        manager.NextStringIndex <- manager.NextStringIndex + 1
        index

/// Generate Excel styles.xml
let generateExcelStylesXml (manager: ExcelStyleManager) : string =
    // For Phase 3, we'll implement basic style generation
    // This is a placeholder that will be expanded
    let root = element "styleSheet" [("xmlns", Namespaces.spreadsheet)] [
        element "fonts" [("count", "1")] [
            element "font" [] [
                selfClosing "sz" [("val", "11")]
                selfClosing "name" [("val", "Calibri")]
            ]
        ]
        element "fills" [("count", "1")] [
            element "fill" [] [
                selfClosing "patternFill" [("patternType", "none")]
            ]
        ]
        element "borders" [("count", "1")] [
            element "border" [] []
        ]
        element "cellXfs" [("count", "1")] [
            selfClosing "xf" [("numFmtId", "0"); ("fontId", "0"); ("fillId", "0"); ("borderId", "0")]
        ]
    ]
    document root

/// Generate Excel sharedStrings.xml
let generateSharedStringsXml (manager: ExcelStyleManager) : string =
    let count = Map.count manager.SharedStrings

    if count = 0 then
        let root = element "sst" [("xmlns", Namespaces.spreadsheet); ("count", "0"); ("uniqueCount", "0")] []
        document root
    else
        let stringItems =
            manager.SharedStrings
            |> Map.toList
            |> List.sortBy snd
            |> List.map (fun (text, _) ->
                element "si" [] [
                    elementWithText "t" [] text
                ]
            )
            |> String.concat ""

        let root = element "sst" [("xmlns", Namespaces.spreadsheet); ("count", string count); ("uniqueCount", string count)] [stringItems]
        document root
