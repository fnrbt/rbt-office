/// A workbook's VBA project for the grid (MacroApi.IMacroProject): its modules compiled with the VBA
/// engine against Excel's object model (MacroHost). In the browser the grid imports this module only
/// when macros are enabled, run or edited, so a workbook that never runs a macro never loads the engine.
module Rbt.Spreadsheet.MacroProject

open Rbt.Vb.Api
open Rbt.Spreadsheet.MacroApi
module Host = Rbt.Spreadsheet.MacroHost

let private failureOf (f: VbFailure) : MacroFailure =
    { Number = f.Number; Description = f.Description; Module = f.Module; Procedure = f.Procedure
      Line = f.Line; Column = f.Column; Text = VbFailure.format f }

let private outcomeOf (result: Host.MacroResult) =
    match result with
    | Host.MacroDone (workbook, selection) -> MacroDone (workbook, selection)
    | Host.MacroFailed failure -> MacroFailed (failureOf failure)

let private kindName (kind: ModuleKind) =
    match kind with
    | Standard -> "standard"
    | Class -> "class"
    | Document -> "document"

type private Project(modules: ModuleSource list) =
    let program = Host.MacroProgram.Compile modules
    let compileFailure () =
        match program with
        | Result.Error failures -> Some (MacroFailed (failureOf (List.head failures)))
        | Result.Ok _ -> None
    interface IMacroProject with
        member _.Modules =
            [ for m in modules -> { Name = m.Name; Kind = kindName m.Kind; Source = m.Source; Code = snd (Host.splitHeader m.Source) } ]
        member _.Macros = match program with Result.Ok p -> p.Macros | Result.Error _ -> []
        member _.Problems = match program with Result.Error failures -> failures |> List.map failureOf | Result.Ok _ -> []
        member _.Has name = match program with Result.Ok p -> p.Has name | Result.Error _ -> false
        member _.Run(workbook, hooks, name) =
            match program with
            | Result.Ok p -> outcomeOf (p.Run(workbook, hooks, name))
            | Result.Error _ -> (compileFailure ()).Value
        member _.Start(workbook, hooks, name) =
            match program with
            | Result.Error _ ->
                let failure = (compileFailure ()).Value
                { new IMacroRun with
                    member _.Step _ = Some failure
                    member _.Frame() = None
                    member _.WaitMs() = 0.0
                    member _.Cancel() = () }
            | Result.Ok p ->
                let begun = p.Begin(workbook, hooks, name, [])
                let mutable ended : MacroOutcome option =
                    match begun with
                    | Result.Error failure -> Some (MacroFailed (failureOf failure))
                    | Result.Ok () -> None
                { new IMacroRun with
                    member _.Step maxInstructions =
                        match ended with
                        | Some outcome -> Some outcome
                        | None ->
                            match p.Step maxInstructions with
                            | Running -> None
                            | Paused _ -> None   // (Stop: no debugger is attached; the next slice carries on.)
                            | status ->
                                ended <- Some (outcomeOf (p.Result status))
                                ended
                    member _.Frame() = if ended.IsSome then None else p.Frame()
                    member _.WaitMs() = if ended.IsSome then 0.0 else p.WaitMs()
                    member _.Cancel() = p.Cancel() }
        member _.Immediate(workbook, hooks, line) =
            match program with
            | Result.Ok p -> outcomeOf (p.Immediate(workbook, hooks, line))
            | Result.Error _ -> (compileFailure ()).Value
        member _.RunEvent(workbook, hooks, handler, target) =
            match program with
            | Result.Ok p -> p.RunEvent(workbook, hooks, handler, target) |> Option.map outcomeOf
            | Result.Error _ -> None
        member _.WithModule(name, code) =
            // (A module it does not have is added: a class when the code has a class's Attribute lines.)
            let edited =
                if modules |> List.exists (fun m -> m.Name = name) then
                    modules |> List.map (fun m ->
                        if m.Name <> name then m
                        else { m with Source = Host.joinHeader (fst (Host.splitHeader m.Source)) code })
                else modules @ [ { Host.macroModule name code with Name = name } ]
            Project(edited) :> IMacroProject
        member _.ProjectBytes() = Host.projectBytes modules

/// The project an opened workbook keeps (None: a workbook without one gets an empty project, in which the
/// Immediate window works and modules can be added).
let openProject (preserved: XlsxReader.WorkbookPreservation option) : IMacroProject =
    let modules = preserved |> Option.map Host.macroModulesOf |> Option.defaultValue []
    Project(modules) :> IMacroProject
