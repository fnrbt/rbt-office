namespace Rbt.Office

/// Shared dispatch for browser and native hosts. Based on the parked UI
/// foundation's command contract; values travel with an invocation rather
/// than being captured in a second set of toolbar-only handlers.
module Commands =
    open System
    open System.Collections.Generic

    [<Struct>]
    type KeyChord = {
        Key: string
        Ctrl: bool
        Shift: bool
        Alt: bool
    }

    let chordOf (key: string) ctrl shift alt =
        { Key = if isNull key then "" else key.ToLowerInvariant()
          Ctrl = ctrl; Shift = shift; Alt = alt }

    let parseChord (text: string) =
        if String.IsNullOrWhiteSpace text then None
        else
            let parts = text.Split('+') |> Array.filter (String.IsNullOrWhiteSpace >> not)
            let parts = if text.EndsWith "++" then Array.append parts [| "+" |] else parts
            if parts.Length = 0 then None
            else
                let mutable ctrl = false
                let mutable shift = false
                let mutable alt = false
                let mutable valid = true
                for i = 0 to parts.Length - 2 do
                    match parts.[i].Trim().ToLowerInvariant() with
                    | "ctrl" | "cmd" | "meta" -> ctrl <- true
                    | "shift" -> shift <- true
                    | "alt" | "option" -> alt <- true
                    | _ -> valid <- false
                if valid then Some (chordOf (parts.[parts.Length - 1].Trim()) ctrl shift alt)
                else None

    let chordMatches shortcut chord =
        shortcut.Key = chord.Key && shortcut.Ctrl = chord.Ctrl
        && shortcut.Shift = chord.Shift && shortcut.Alt = chord.Alt

    type CommandState = {
        Enabled: bool
        Checked: bool option
        Mixed: bool
    }

    type Command = {
        Id: string
        Label: string
        Shortcut: string option
        IsEnabled: unit -> bool
        /// None means not a toggle. A toggle returning None is mixed.
        IsChecked: (unit -> bool option) option
        Run: obj -> unit
    }

    let command id label shortcut isEnabled run =
        { Id = id; Label = label; Shortcut = shortcut
          IsEnabled = isEnabled; IsChecked = None; Run = run }

    let toggle id label shortcut isEnabled isChecked run =
        { command id label shortcut isEnabled run with IsChecked = Some isChecked }

    type CommandRegistry() =
        let order = ResizeArray<string>()
        let commands = Dictionary<string, Command>()
        let chords = ResizeArray<KeyChord * string>()
        let listeners = ResizeArray<unit -> unit>()

        member _.Add(command: Command) =
            if not (commands.ContainsKey command.Id) then order.Add command.Id
            commands.[command.Id] <- command
            chords.RemoveAll(fun (_, id) -> id = command.Id) |> ignore
            match command.Shortcut |> Option.bind parseChord with
            | Some chord -> chords.Add(chord, command.Id)
            | None -> ()

        member this.AddAll(commands: Command list) =
            for command in commands do this.Add command

        member _.TryFind(id: string) =
            match commands.TryGetValue id with
            | true, command -> Some command
            | _ -> None

        member _.All = [ for id in order -> commands.[id] ]

        member this.State(id: string) =
            this.TryFind id |> Option.map (fun command ->
                let checkedState = command.IsChecked |> Option.bind (fun read -> read ())
                { Enabled = command.IsEnabled ()
                  Checked = checkedState
                  Mixed = command.IsChecked.IsSome && checkedState.IsNone })

        /// A disabled command leaves its invocation unhandled. Unknown IDs
        /// are programming errors. Handlers retain the engine's undo policy.
        member this.Run(id: string, ?argument: obj) =
            match this.TryFind id with
            | None -> invalidArg "id" (sprintf "Unknown command '%s'" id)
            | Some command when command.IsEnabled () ->
                command.Run (defaultArg argument null)
                true
            | Some _ -> false

        member this.ForChord(chord: KeyChord) =
            let mutable found = None
            let mutable i = 0
            while found.IsNone && i < chords.Count do
                let shortcut, id = chords.[i]
                if chordMatches shortcut chord then found <- this.TryFind id
                i <- i + 1
            found

        member this.HandleKey(chord: KeyChord) =
            match this.ForChord chord with
            | Some command -> this.Run command.Id
            | None -> false

        member _.OnChange(listener: unit -> unit) =
            listeners.Add listener
            fun () -> listeners.Remove listener |> ignore

        /// Snapshot permits listeners to unsubscribe during notification.
        member _.NotifyChanged() =
            for listener in listeners.ToArray() do listener ()
