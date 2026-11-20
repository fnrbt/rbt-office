/// The fonts a Word document embeds: word/fontTable.xml's w:embedRegular, w:embedBold, w:embedItalic and
/// w:embedBoldItalic, each naming an obfuscated font part (word/fonts/*.odttf) and the key it was
/// obfuscated with (w:fontKey). ECMA-376 Part 2 / Word obfuscate a font by XORing its first 32 bytes with
/// the 16 bytes of the key's GUID, last byte first; the same XOR gives the font back.
///
/// What Word does with them (docs/layout-parity-flow.md, fixtures embedded-fonts, Word 2019):
/// - it sets text in an embedded face when the machine has no font of that family, whatever
///   w:embedTrueTypeFonts says; a family the machine has (Calibri) is set in the machine's font;
/// - the family is the font table's w:name, not the name the face gives itself;
/// - a face whose OS/2 fsType allows restricted-licence embedding only is not used (fsType 2); preview
///   and print (4) and editable (8) faces are;
/// - a family with some faces sets its other styles in them (bold of the regular face: the regular's
///   advances; regular of the bold face: the bold's) - see ShaperText's face choice;
/// - a subset face sets a character it has no glyph of as its .notdef, not in another font.
///
/// The parts are read from the package as it was opened (PreservedPackage: they are kept and written
/// back as they were on save); nothing is read until a host or the layout asks.
module EmbeddedFonts

open XmlDsl

/// One embedded face: the family it is set in for (the font table's name), its style, and the font
/// file (de-obfuscated: a TrueType / OpenType font).
type EmbeddedFace = {
    Family: string
    Bold: bool
    Italic: bool
    Data: byte[]
    /// w:subsetted: the face has only the glyphs of the characters the document used.
    Subsetted: bool
}

let private localName (name: string) =
    match name.IndexOf ':' with
    | -1 -> name
    | i -> name.Substring(i + 1)

let private attrLocal (local: string) (node: XmlNode) =
    match node with
    | Element (_, attrs, _, _) ->
        attrs |> Map.toSeq |> Seq.tryPick (fun (key, value) -> if localName key = local then Some value else None)
    | _ -> None

let private elements (node: XmlNode) =
    Xml.children node |> List.filter (function Element _ -> true | _ -> false)

let private text (data: byte[]) =
    let s = System.Text.Encoding.UTF8.GetString data
    if s.Length > 0 && s.[0] = '﻿' then s.Substring 1 else s

/// A relationship's target from a part in `directory`, as a package path.
let private resolve (directory: string) (target: string) =
    let joined = if target.StartsWith "/" then target.Substring 1 elif directory = "" then target else directory + "/" + target
    let parts = System.Collections.Generic.List<string>()
    for segment in joined.Split('/') do
        if segment = ".." then (if parts.Count > 0 then parts.RemoveAt(parts.Count - 1))
        elif segment <> "." && segment <> "" then parts.Add segment
    String.concat "/" parts

/// The 16 key bytes of a w:fontKey GUID ("{3EEE3167-E5B8-4798-AE48-EA6B71E31D4D}"), last first.
let private keyBytes (guid: string) : byte[] option =
    let hex = guid.Replace("{", "").Replace("}", "").Replace("-", "")
    if hex.Length <> 32 then None
    else
        try
            Some (Array.init 16 (fun i -> System.Convert.ToByte(hex.Substring((15 - i) * 2, 2), 16)))
        with _ -> None

/// The font file of an obfuscated part (or the part of a font, obfuscated: the XOR is its own inverse).
let deobfuscate (guid: string) (data: byte[]) : byte[] option =
    match keyBytes guid with
    | Some key when data.Length >= 32 ->
        let font = Array.copy data
        for i in 0 .. 31 do
            font.[i] <- font.[i] ^^^ key.[i % 16]
        Some font
    | _ -> None

let private u16 (data: byte[]) (at: int) = (int data.[at] <<< 8) ||| int data.[at + 1]
let private u32 (data: byte[]) (at: int) = (u16 data at <<< 16) ||| u16 data (at + 2)

/// Whether a font file is one (TrueType, OpenType or a collection) and its OS/2 fsType (0 without one).
let private fontInfo (data: byte[]) : (bool * int) =
    if data.Length < 12 then false, 0
    else
        let tag = u32 data 0
        let isFont = tag = 0x00010000 || tag = 0x4F54544F (* OTTO *) || tag = 0x74727565 (* true *) || tag = 0x74746366 (* ttcf *)
        if not isFont || tag = 0x74746366 then isFont, 0
        else
            let tables = u16 data 4
            let mutable fsType = 0
            for i in 0 .. tables - 1 do
                let record = 12 + i * 16
                if record + 16 <= data.Length && u32 data record = 0x4F532F32 (* OS/2 *) then
                    let offset = u32 data (record + 8)
                    if offset + 10 <= data.Length then fsType <- u16 data (offset + 8)
            true, fsType

/// Whether Word sets text in a face whose OS/2 fsType is this: not when it allows restricted-licence
/// embedding only (bit 1 alone of bits 1 - 3; fixture restricted).
let usableFsType (fsType: int) = (fsType &&& 0x000E) <> 0x0002

/// The faces embedded in a package's parts (`parts`: package path -> bytes), given the path of its main
/// document part: those Word would set text in (see usableFsType). [] when there are none.
let ofParts (mainPath: string) (part: string -> byte[] option) : EmbeddedFace list =
    let mainDirectory = PartPath.directory mainPath
    let relationshipsOf (path: string) =
        match part (PartPath.relationshipsOf path) with
        | None -> []
        | Some data ->
            try
                parse (text data)
                |> elements
                |> List.choose (fun rel ->
                    match attrLocal "Id" rel, attrLocal "Type" rel, attrLocal "Target" rel with
                    | Some id, Some kind, Some target when attrLocal "TargetMode" rel <> Some "External" ->
                        Some (id, kind, resolve (PartPath.directory path) target)
                    | _ -> None)
            with _ -> []
    let fontTable =
        relationshipsOf mainPath
        |> List.tryPick (fun (_, kind, target) -> if kind.EndsWith "/fontTable" then Some target else None)
    match fontTable |> Option.bind (fun path -> part path |> Option.map (fun data -> path, data)) with
    | None -> []
    | Some (tablePath, tableData) ->
        let targets = relationshipsOf tablePath |> List.map (fun (id, _, target) -> id, target) |> Map.ofList
        let root = try parse (text tableData) with _ -> Empty
        [ for font in elements root do
            if localName (defaultArg (Xml.name font) "") = "font" then
                match attrLocal "name" font with
                | Some family when family.Trim() <> "" ->
                    for embed in elements font do
                        let style =
                            match localName (defaultArg (Xml.name embed) "") with
                            | "embedRegular" -> Some (false, false)
                            | "embedBold" -> Some (true, false)
                            | "embedItalic" -> Some (false, true)
                            | "embedBoldItalic" -> Some (true, true)
                            | _ -> None
                        match style, attrLocal "id" embed, attrLocal "fontKey" embed with
                        | Some (bold, italic), Some id, key ->
                            let data =
                                Map.tryFind id targets
                                |> Option.bind part
                                |> Option.bind (fun data ->
                                    match key with
                                    | Some guid when guid <> "" -> deobfuscate guid data
                                    | _ -> Some data)
                            match data with
                            | Some data ->
                                let isFont, fsType = fontInfo data
                                if isFont && usableFsType fsType then
                                    let subsetted = attrLocal "subsetted" embed |> Option.map (fun v -> v = "1" || v = "true" || v = "on")
                                    yield { Family = family.Trim(); Bold = bold; Italic = italic; Data = data
                                            Subsetted = defaultArg subsetted false }
                            | None -> ()
                        | _ -> ()
                | _ -> () ]

#if FABLE_COMPILER
[<Fable.Core.Emit("new WeakMap()")>]
let private newCache () : obj = Fable.Core.Util.jsNative
[<Fable.Core.Emit("$0.get($1)")>]
let private cacheGet (_cache: obj) (_key: obj) : EmbeddedFace list = Fable.Core.Util.jsNative
[<Fable.Core.Emit("$0.set($1, $2)")>]
let private cacheSet (_cache: obj) (_key: obj) (_value: EmbeddedFace list) : unit = Fable.Core.Util.jsNative
let private cache = newCache ()
let private cached (key: obj) (read: unit -> EmbeddedFace list) =
    let known = cacheGet cache key
    if isNull (box known) then
        let value = read ()
        cacheSet cache key value
        value
    else known
#else
let private cache = System.Runtime.CompilerServices.ConditionalWeakTable<obj, EmbeddedFace list>()
let private cached (key: obj) (read: unit -> EmbeddedFace list) =
    match cache.TryGetValue key with
    | true, known -> known
    | _ ->
        let value = read ()
        lock cache (fun () -> match cache.TryGetValue key with | true, _ -> () | _ -> cache.Add(key, value))
        value
#endif

/// The faces a package as opened embeds (its kept parts), read once per package.
let ofPreserved (preserved: WordEditorPackage.PreservedPackage) : EmbeddedFace list =
    cached (box preserved) (fun () ->
        let mainPath =
            preserved.Overrides
            |> List.tryPick (fun (name, contentType) -> if Some contentType = preserved.MainContentType then Some (name.TrimStart '/') else None)
            |> Option.defaultValue "word/document.xml"
        let parts = preserved.Parts |> List.map (fun part -> part.Path, part.Data) |> Map.ofList
        // (The main part's relationships are not kept as a part: the document's kept relationships are
        // PartRelationships.["document"], its source ones SourceRelationships.["document"].)
        let documentRelationships =
            (Map.tryFind "document" preserved.PartRelationships |> Option.defaultValue [])
            @ (Map.tryFind "document" preserved.SourceRelationships |> Option.defaultValue [])
        let mainRels = PartPath.relationshipsOf mainPath
        let relsXml =
            "<Relationships>"
            + (documentRelationships
               |> List.filter (fun rel -> not rel.External)
               |> List.map (fun rel ->
                   let escape (s: string) = s.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;")
                   sprintf "<Relationship Id=\"%s\" Type=\"%s\" Target=\"%s\"/>" (escape rel.Id) (escape rel.Type) (escape rel.Target))
               |> String.concat "")
            + "</Relationships>"
        ofParts mainPath (fun path ->
            if path = mainRels then Some (System.Text.Encoding.UTF8.GetBytes relsXml)
            else Map.tryFind path parts))
