/// What the Word editor knows of a document's VBA macros without loading them: whether its package keeps a
/// VBA project (a .docm's word/vbaProject.bin), the project's bytes, and the interface of a loaded project
/// (`WordMacroProject.fs`, which brings the VBA engine and Word's object model, `WordMacroHost.fs`). In the
/// browser that module is fetched only when macros are enabled, run or edited; nothing here refers to it.
module Rbt.Word.WordMacroApi

open System
open Rbt.Word.Document

/// What a host gives a run: MsgBox / InputBox / Debug.Print, and the document's file name.
type MacroHooks = {
    MessageBox: string -> int -> string -> int
    InputBox: string -> string -> string -> string option
    Output: string -> unit
    FileName: string
}

let defaultHooks = {
    MessageBox = fun _ _ _ -> 1
    InputBox = fun _ _ _ -> None
    Output = ignore
    FileName = "Document1.docm"
}

/// Why a macro stopped or a project does not compile: Word's number and description, and where.
type MacroFailure = {
    Number: int
    Description: string
    Module: string
    Procedure: string
    Line: int
    Column: int
    /// The whole message, as the VBA editor shows it ("Run-time error '5941': ... (Module1.Main, line 3)").
    Text: string
}

/// How a run ended: the document as the macro left it (its selection where the macro's Selection ended),
/// or why it stopped (and then nothing of it counts).
type MacroOutcome =
    | MacroDone of Document
    | MacroFailed of MacroFailure

/// A module of a project: "standard", "class" or "document" (ThisDocument), its source with the Attribute
/// lines, and its code as the VBA editor shows it (without them).
type MacroModuleInfo = { Name: string; Kind: string; Source: string; Code: string }

/// A document's VBA project, compiled against Word's object model; its module state lasts between runs.
type IWordMacroProject =
    abstract Modules: MacroModuleInfo list
    /// The macros a user can run: public Subs without required arguments (ThisDocument's as `ThisDocument.Name`).
    abstract Macros: string list
    /// What keeps it from compiling (none when it compiles).
    abstract Problems: MacroFailure list
    /// Whether it has that procedure (`Name` or `Module.Name`).
    abstract Has: string -> bool
    abstract Run: document: Document * hooks: MacroHooks * name: string -> MacroOutcome
    /// An Immediate-window line (`?expression`, `Print expression`, statements); a value goes to Output.
    abstract Immediate: document: Document * hooks: MacroHooks * line: string -> MacroOutcome
    /// The project with a module's code replaced (a module it does not have is added, as a standard module).
    abstract WithModule: name: string * code: string -> IWordMacroProject
    /// The project as vbaProject.bin, from its modules (for a save after an edit).
    abstract ProjectBytes: unit -> byte[]

// ---- what an opened file tells without the VBA engine ------------------------------------------------

let private vbaRelationshipType = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
let private vbaContentType = "application/vnd.ms-office.vbaProject"
let private macroEnabledMain = "application/vnd.ms-word.document.macroEnabled.main+xml"

let private isVbaRelationship (relType: string) = relType.EndsWith "/vbaProject"

/// The path of the main document part (word/document.xml) of a package. (Its relationships are kept under
/// the key "document" in PreservedPackage.PartRelationships.)
let private mainPartPath (preserved: WordEditorPackage.PreservedPackage) =
    preserved.RootRelationships
    |> List.tryPick (fun r -> if r.Type.EndsWith "/officeDocument" && not r.External then Some (r.Target.TrimStart '/') else None)
    |> Option.defaultValue "word/document.xml"

let private resolveTarget (fromPart: string) (target: string) =
    if target.StartsWith "/" then target.TrimStart '/'
    else
        let folder = match fromPart.LastIndexOf '/' with -1 -> "" | at -> fromPart.Substring(0, at + 1)
        let parts = ResizeArray<string>()
        for piece in (folder + target).Split('/') do
            if piece = ".." then (if parts.Count > 0 then Portable.Stack.dropLast parts)
            elif piece <> "." && piece <> "" then parts.Add piece
        String.Join("/", parts)

/// The path of the VBA project part an opened document keeps (word/vbaProject.bin), when it keeps one.
let projectPartPath (doc: Document) : string option =
    match doc.Package.Preserved with
    | None -> None
    | Some preserved ->
        let main = mainPartPath preserved
        let fromRelationship =
            preserved.PartRelationships
            |> Map.tryFind "document" |> Option.orElse (preserved.PartRelationships |> Map.tryFind main)
            |> Option.bind (List.tryPick (fun r -> if isVbaRelationship r.Type && not r.External then Some (resolveTarget main r.Target) else None))
        let exists path = preserved.Parts |> List.exists (fun p -> p.Path.TrimStart '/' = path)
        match fromRelationship with
        | Some path when exists path -> Some path
        | _ ->
            preserved.Parts
            |> List.tryPick (fun p -> let path = p.Path.TrimStart '/' in if path.EndsWith "vbaProject.bin" then Some path else None)

/// The bytes of the VBA project an opened document keeps.
let projectBytes (doc: Document) : byte[] option =
    match projectPartPath doc, doc.Package.Preserved with
    | Some path, Some preserved -> preserved.Parts |> List.tryPick (fun p -> if p.Path.TrimStart '/' = path then Some p.Data else None)
    | _ -> None

/// Whether an opened document keeps a VBA project.
let hasProject (doc: Document) = (projectBytes doc).IsSome

/// The document with `bytes` as its VBA project (for a save after the project was edited): the part is
/// replaced, or added with its relationship and content type (and the document becomes macro-enabled).
/// A document made from scratch (no package) is given back as it is.
let withProjectBytes (bytes: byte[]) (doc: Document) : Document =
    match doc.Package.Preserved with
    | None -> doc
    | Some preserved ->
        let next =
            match projectPartPath doc with
            | Some path ->
                { preserved with Parts = preserved.Parts |> List.map (fun p -> if p.Path.TrimStart '/' = path then { p with Data = bytes } else p) }
            | None ->
                let main = mainPartPath preserved
                let path = "word/vbaProject.bin"
                let relationships = preserved.PartRelationships |> Map.tryFind "document" |> Option.defaultValue []
                let id =
                    Seq.initInfinite (fun i -> sprintf "rIdVba%d" (i + 1))
                    |> Seq.find (fun candidate -> relationships |> List.forall (fun r -> r.Id <> candidate))
                let relationship : WordEditorPackage.Relationship = { Id = id; Type = vbaRelationshipType; Target = "vbaProject.bin"; External = false }
                { preserved with
                    Parts = preserved.Parts @ [ { Path = path; Data = bytes } ]
                    PartRelationships = preserved.PartRelationships |> Map.add "document" (relationships @ [ relationship ])
                    Overrides =
                        (preserved.Overrides |> List.filter (fun (part, _) -> part.TrimStart '/' <> path))
                        @ [ "/" + path, vbaContentType ]
                    MainContentType = Some macroEnabledMain }
        { doc with Package = { doc.Package with Preserved = Some next } }
