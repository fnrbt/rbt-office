/// A document's VBA project for the Word editor (WordMacroApi.IWordMacroProject): its modules compiled with
/// the VBA engine against Word's object model (WordMacroHost). In the browser the editor imports this
/// module only when macros are enabled, run or edited, so a document that never runs a macro never
/// loads the engine.
module Rbt.Word.WordMacroProject

open Rbt.Vb.Api
open Rbt.Word.WordMacroApi
module Host = Rbt.Word.WordMacroHost

let private failureOf (f: VbFailure) : MacroFailure =
    { Number = f.Number; Description = f.Description; Module = f.Module; Procedure = f.Procedure
      Line = f.Line; Column = f.Column; Text = VbFailure.format f }

let private outcomeOf (result: Host.MacroResult) =
    match result with
    | Host.MacroDone doc -> MacroDone doc
    | Host.MacroFailed failure -> MacroFailed (failureOf failure)

let private kindName (kind: ModuleKind) =
    match kind with
    | Standard -> "standard"
    | Class -> "class"
    | ModuleKind.Document -> "document"

type private Project(modules: ModuleSource list) =
    let program = Host.MacroProgram.Compile modules
    let compileFailure () =
        match program with
        | Result.Error failures -> MacroFailed (failureOf (List.head failures))
        | Result.Ok _ -> MacroFailed { Number = 0; Description = "The project compiled"; Module = ""; Procedure = ""; Line = 0; Column = 0; Text = "" }
    interface IWordMacroProject with
        member _.Modules =
            [ for m in modules -> { Name = m.Name; Kind = kindName m.Kind; Source = m.Source; Code = snd (Host.splitHeader m.Source) } ]
        member _.Macros = match program with Result.Ok p -> p.Macros | Result.Error _ -> []
        member _.Problems = match program with Result.Error failures -> failures |> List.map failureOf | Result.Ok _ -> []
        member _.Has name = match program with Result.Ok p -> p.Has name | Result.Error _ -> false
        member _.Run(doc, hooks, name) =
            match program with
            | Result.Ok p -> outcomeOf (p.Run(doc, hooks, name))
            | Result.Error _ -> compileFailure ()
        member _.Immediate(doc, hooks, line) =
            match program with
            | Result.Ok p -> outcomeOf (p.Immediate(doc, hooks, line))
            | Result.Error _ -> compileFailure ()
        member _.WithModule(name, code) =
            let modules =
                if modules |> List.exists (fun m -> m.Name = name) then modules
                else modules @ [ ModuleSource.standard name "" ]
            let edited =
                modules |> List.map (fun m ->
                    if m.Name <> name then m
                    else { m with Source = Host.joinHeader (fst (Host.splitHeader m.Source)) code })
            Project(edited) :> IWordMacroProject
        member _.ProjectBytes() = Host.modulesBytes modules

/// The project an opened document keeps (a document without one gets an empty project, in which the
/// Immediate window works and modules can be added).
let openProject (doc: Rbt.Word.Document.Document) : IWordMacroProject =
    Project(Host.macroModulesOf doc) :> IWordMacroProject

/// A project of the given modules (name, code), as standard modules: for hosts and tests.
let ofModules (modules: (string * string) list) : IWordMacroProject =
    Project([ for name, code in modules -> ModuleSource.standard name code ]) :> IWordMacroProject
