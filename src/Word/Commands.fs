module RedBat.Word.Commands

open RedBat.Word.Document
open RedBat.Word.TextBuffer

/// Command interface for undo/redo
type ICommand =
    abstract Execute: Document -> Document
    abstract Undo: Document -> Document
    abstract Description: string

/// Command history for undo/redo
type CommandHistory = {
    Undos: ICommand list
    Redos: ICommand list
    MaxHistory: int
}

/// Create empty command history
let empty = {
    Undos = []
    Redos = []
    MaxHistory = 100
}

/// Execute a command and add to history
let execute (cmd: ICommand) (doc: Document) (history: CommandHistory) =
    let newDoc = cmd.Execute doc
    let newHistory = {
        Undos = cmd :: (history.Undos |> List.truncate (history.MaxHistory - 1))
        Redos = []  // Clear redo stack when new command is executed
        MaxHistory = history.MaxHistory
    }
    (newDoc, newHistory)

/// Undo last command
let undo (doc: Document) (history: CommandHistory) =
    match history.Undos with
    | [] -> (doc, history)  // Nothing to undo
    | cmd :: rest ->
        let newDoc = cmd.Undo doc
        let newHistory = {
            Undos = rest
            Redos = cmd :: history.Redos
            MaxHistory = history.MaxHistory
        }
        (newDoc, newHistory)

/// Redo last undone command
let redo (doc: Document) (history: CommandHistory) =
    match history.Redos with
    | [] -> (doc, history)  // Nothing to redo
    | cmd :: rest ->
        let newDoc = cmd.Execute doc
        let newHistory = {
            Undos = cmd :: history.Undos
            Redos = rest
            MaxHistory = history.MaxHistory
        }
        (newDoc, newHistory)

/// Insert text command
type InsertTextCommand(text: string, position: Position, beforeDoc: Document) =
    interface ICommand with
        member _.Execute doc =
            // Set cursor to position and insert text
            let doc' = setCursor position.BlockId position.Offset doc
            insertText text doc'

        member _.Undo doc =
            beforeDoc

        member _.Description = sprintf "Insert '%s'" text

/// Delete text command (backspace/delete)
type DeleteTextCommand(beforeDoc: Document, afterDoc: Document, description: string) =
    interface ICommand with
        member _.Execute doc = afterDoc
        member _.Undo doc = beforeDoc
        member _.Description = description

/// Split block command (Enter key)
type SplitBlockCommand(beforeDoc: Document, afterDoc: Document) =
    interface ICommand with
        member _.Execute doc = afterDoc
        member _.Undo doc = beforeDoc
        member _.Description = "Split paragraph"

/// Format text command (bold/italic/underline)
type FormatTextCommand(beforeDoc: Document, afterDoc: Document, formatType: string) =
    interface ICommand with
        member _.Execute doc = afterDoc
        member _.Undo doc = beforeDoc
        member _.Description = sprintf "Format: %s" formatType

/// Composite command for grouping multiple commands
type CompositeCommand(commands: ICommand list, description: string) =
    interface ICommand with
        member _.Execute doc =
            commands |> List.fold (fun d cmd -> cmd.Execute d) doc

        member _.Undo doc =
            commands |> List.rev |> List.fold (fun d cmd -> cmd.Undo d) doc

        member _.Description = description
