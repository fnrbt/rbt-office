/// What the grid knows of a workbook's VBA macros without loading them: whether it has a project, the
/// code names of its sheets, and the interface of a loaded project (`MacroProject.fs`, which brings the
/// VBA engine and Excel's object model). In the browser that module is fetched only when macros are
/// enabled, run or edited; nothing here refers to it.
module Rbt.Spreadsheet.MacroApi

open System
open Rbt.Spreadsheet.Workbook

/// What a host gives a run: MsgBox / InputBox / Debug.Print, and the workbook's file name.
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
    FileName = "Book1.xlsx"
}

/// Why a macro stopped or a project does not compile: Excel's number and description, and where.
type MacroFailure = {
    Number: int
    Description: string
    Module: string
    Procedure: string
    Line: int
    Column: int
    /// The whole message, as the VBA editor shows it ("Run-time error '9': ... (Module1.Main, line 3)").
    Text: string
}

/// How a run ended: the workbook as the macro left it (calculated) and the cells it selected
/// (sheet, top, left, bottom, right), or why it stopped (and then nothing of it counts).
type MacroOutcome =
    | MacroDone of workbook: Workbook * selection: (SheetId * int * int * int * int) option
    | MacroFailed of MacroFailure

/// A macro run a slice at a time, so a host can show what it has done so far between slices (as Excel
/// repaints while a macro runs) and stop it.
type IMacroRun =
    /// Run up to `maxInstructions` more: None while it is still running, then how it ended.
    abstract Step: maxInstructions: int -> MacroOutcome option
    /// The workbook as the macro has left it so far (calculated) when it changed since the last frame and the
    /// macro has not switched the screen off (Application.ScreenUpdating = False); None otherwise.
    abstract Frame: unit -> Workbook option
    /// How much longer the macro asked to wait (Application.Wait), in milliseconds; 0 when it is not waiting.
    abstract WaitMs: unit -> float
    /// Stop it: the next Step ends it (error 18, as Excel's Esc / Ctrl+Break).
    abstract Cancel: unit -> unit

/// A module of a project: "standard", "class" or "document" (ThisWorkbook, a sheet's module), its
/// source with the Attribute lines, and its code as the VBA editor shows it (without them).
type MacroModuleInfo = { Name: string; Kind: string; Source: string; Code: string }

/// A workbook's VBA project, compiled against Excel's object model; its module state lasts between runs.
type IMacroProject =
    abstract Modules: MacroModuleInfo list
    /// The macros a user can run: public Subs without required arguments (a sheet's as `Sheet1.Name`).
    abstract Macros: string list
    /// What keeps it from compiling (none when it compiles).
    abstract Problems: MacroFailure list
    /// Whether it has that procedure (`Name` or `Module.Name`).
    abstract Has: string -> bool
    abstract Run: workbook: Workbook * hooks: MacroHooks * name: string -> MacroOutcome
    /// Start `name` on the workbook, to be run a slice at a time (no limit on its instructions: the host
    /// stops it with Cancel).
    abstract Start: workbook: Workbook * hooks: MacroHooks * name: string -> IMacroRun
    /// An Immediate-window line (`?expression`, `Print expression`, statements); a value goes to Output.
    abstract Immediate: workbook: Workbook * hooks: MacroHooks * line: string -> MacroOutcome
    /// An event handler with an optional Target range; None when the project has no such handler.
    abstract RunEvent: workbook: Workbook * hooks: MacroHooks * handler: string * target: (SheetId * int * int * int * int) option -> MacroOutcome option
    /// The project with a module's code replaced (a module it does not have is added, as a standard module).
    abstract WithModule: name: string * code: string -> IMacroProject
    /// The project as vbaProject.bin, from its modules (for a save after an edit).
    abstract ProjectBytes: unit -> byte[]

// ---- what an opened file tells without the VBA engine ------------------------------------------------

/// Whether an opened workbook keeps a VBA project.
let hasProject (preserved: XlsxReader.WorkbookPreservation) =
    preserved.WorkbookRelationships
    |> List.exists (fun (_, relType, _, isExternal) -> not isExternal && OpcXml.relationshipKind relType = "vbaProject")

/// The code name of each sheet of an opened workbook (`<sheetPr codeName>`, the name of its module), by
/// sheet name; Excel's default (Sheet1, Sheet2 ... in order) for a sheet without one.
let sheetCodeNames (source: XlsxReader.Workbook) : Map<string, string> =
    let pattern = Text.RegularExpressions.Regex("<sheetPr\\b[^>]*\\bcodeName=\"([^\"]+)\"")
    let texts = source.Preserved.SourceParts |> Map.ofList
    source.Sheets
    |> List.mapi (fun index sheet ->
        let codeName =
            sheet.Extras.SourcePart
            |> Option.bind (fun (path, _) -> Map.tryFind path texts)
            |> Option.bind (fun bytes ->
                let text = Text.Encoding.UTF8.GetString(bytes, 0, min bytes.Length 4096)
                let m = pattern.Match text
                if m.Success then Some m.Groups.[1].Value else None)
        sheet.Name, codeName |> Option.defaultValue (sprintf "Sheet%d" (index + 1)))
    |> Map.ofList

/// The code name of the workbook (`<workbookPr codeName>`): its module, ThisWorkbook by default.
let workbookCodeName (source: XlsxReader.Workbook) : string =
    source.Preserved.WorkbookPart
    |> Option.bind (fun path -> source.Preserved.SourceParts |> List.tryPick (fun (part, data) -> if part = path then Some data else None))
    |> Option.bind (fun bytes ->
        let m = Text.RegularExpressions.Regex("<workbookPr\\b[^>]*\\bcodeName=\"([^\"]+)\"").Match(Text.Encoding.UTF8.GetString bytes)
        if m.Success then Some m.Groups.[1].Value else None)
    |> Option.defaultValue "ThisWorkbook"
