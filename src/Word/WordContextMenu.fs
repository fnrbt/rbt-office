namespace Rbt.Word

#if FABLE_COMPILER
module internal ContextMenu =
    open Fable.Core
    open Fable.Core.JsInterop
    open Browser.Dom
    open Browser.Types
    open Rbt.Office.Commands
    module Word = Rbt.Word.API
    module Doc = Rbt.Word.Document

    [<Emit("$0 && $0.focus({preventScroll:true})")>]
    let private focus (_: obj) : unit = jsNative
    [<Emit("!!($0 && !$0.readOnly && !$0.disabled)")>]
    let private editable (_: obj) : bool = jsNative
    [<Emit("!!(navigator.clipboard && (navigator.clipboard.write || navigator.clipboard.writeText))")>]
    let private canWrite () : bool = jsNative
    [<Emit("!!(navigator.clipboard && (navigator.clipboard.read || navigator.clipboard.readText))")>]
    let private canRead () : bool = jsNative
    [<Emit("(async function(p){if(navigator.clipboard.write && typeof ClipboardItem !== 'undefined'){await navigator.clipboard.write([new ClipboardItem({'text/plain':new Blob([p.Text],{type:'text/plain'}),'text/html':new Blob([p.Html],{type:'text/html'})})]);}else{await navigator.clipboard.writeText(p.Text);}})($0)")>]
    let private writeClipboard (_: obj) : JS.Promise<unit> = jsNative
    [<Emit("(async function(){if(!navigator.clipboard.read)return {Text:await navigator.clipboard.readText(),Html:''};const items=await navigator.clipboard.read();for(const item of items){if(item.types.includes('text/plain') || item.types.includes('text/html'))return {Text:item.types.includes('text/plain')?await (await item.getType('text/plain')).text():'',Html:item.types.includes('text/html')?await (await item.getType('text/html')).text():''};}throw new Error('The clipboard contains no text.');})()")>]
    let private readClipboard () : JS.Promise<obj> = jsNative
    [<Emit("$0.isComposing || $0.keyCode === 229")>]
    let private composing (_: KeyboardEvent) : bool = jsNative
    [<Emit("Date.now()")>]
    let private now () : float = jsNative

    // The shell owns this adapter, not the canvas widget. A pending clipboard
    // read/write must never modify a replacement document or a moved selection.
    let attach (root: Element) (host: Host.Host) (registry: CommandRegistry) =
        let mutable disposed = false
        let mutable popup: HTMLDivElement = null
        let mutable stillCurrent: (unit -> bool) = fun () -> true
        let mutable keyboardOpened = -1000.0
        let editorInput () = host.Editor |> Option.map Word.getInputElement |> Option.defaultValue null
        let close restore =
            if not (isNull popup) then
                popup.remove()
                popup <- null
                if restore && not disposed then focus (editorInput())
        let capture (editor: Word.Editor) =
            let doc = editor.Doc
            let generation = host.OpenGeneration
            let index = editor.Handler.Session.Epoch
            fun () ->
                not disposed && generation = host.OpenGeneration
                && (host.Editor |> Option.exists (fun current ->
                    obj.ReferenceEquals(current, editor) && obj.ReferenceEquals(current.Doc, doc)
                    && obj.ReferenceEquals(current.Handler.Session.Epoch, index)))
        let isComposing () = host.Editor |> Option.exists (fun editor -> editor.Handler.Composition.IsSome)
        let canEdit () = host.Editor |> Option.exists (fun editor -> editable (Word.getInputElement editor) && editor.Doc.Selection.IsSome)
        let hasSelection () = host.Editor |> Option.exists (fun editor -> editor.Doc.Selection |> Option.exists (fun selection -> not selection.IsCollapsed))
        let clipboard cut paste =
            match host.Editor with
            | None -> ()
            | Some editor ->
                let current = capture editor
                let payload = if paste then null else Word.copySelection editor
                if paste || not (isNullOrUndefined payload) then
                    async {
                        try
                            if paste then
                                let! contents = readClipboard() |> Async.AwaitPromise
                                if current() && canEdit() then
                                    Host.runCommand host (fun target -> Word.pasteClipboard target contents?Text contents?Html)
                                    host.Status.Value <- "Pasted"
                                elif not disposed then host.Status.Value <- "Paste cancelled: the document or selection changed."
                            else
                                do! writeClipboard payload |> Async.AwaitPromise
                                if cut then
                                    if current() && canEdit() then
                                        Host.runCommand host (fun target -> Word.cutSelection target |> ignore)
                                        host.Status.Value <- "Cut"
                                    elif not disposed then host.Status.Value <- "Copied, but not cut: the document or selection changed."
                                elif current() then host.Status.Value <- "Copied"
                        with error ->
                            if current() then host.Status.Value <- "Clipboard failed: " + error.Message
                    } |> Async.StartImmediate
        registry.Add(command "edit.copy" "Copy" None (fun () -> canWrite() && hasSelection()) (fun _ -> clipboard false false))
        registry.Add(command "edit.cut" "Cut" None (fun () -> canWrite() && canEdit() && hasSelection()) (fun _ -> clipboard true false))
        registry.Add(command "edit.paste" "Paste" None (fun () -> canRead() && canEdit()) (fun _ -> clipboard false true))
        registry.Add(command "link.remove" "Remove link" None
            (fun () -> canEdit() && (host.Editor |> Option.exists (Word.getHyperlink >> isNullOrUndefined >> not)))
            (fun _ -> Host.runCommand host (fun editor -> Word.setHyperlinkTarget editor "")))
        let openAt (editor: Word.Editor) x y =
            close false
            let current = capture editor
            stillCurrent <- current
            let menu = document.createElement("div") :?> HTMLDivElement
            menu.className <- "rb-word-menu rb-word-context-menu"
            menu.setAttribute("role", "menu")
            menu.setAttribute("aria-label", "Document actions")
            menu.setAttribute("data-testid", "word-context-menu")
            menu.setAttribute("style", "position:fixed;z-index:40;left:0;top:0")
            let add id label =
                match registry.State id with
                | None -> ()
                | Some state ->
                    let item = document.createElement("button") :?> HTMLButtonElement
                    item.className <- "rb-word-menu-item"
                    item.textContent <- label
                    item.setAttribute("data-command", id)
                    item.setAttribute("role", if state.Checked.IsSome || state.Mixed then "menuitemcheckbox" else "menuitem")
                    if state.Checked.IsSome || state.Mixed then
                        item.setAttribute("aria-checked", if state.Mixed then "mixed" elif state.Checked = Some true then "true" else "false")
                    // Existing Undo/Redo and paragraph commands predate read-only
                    // shell adapters; enforce that boundary here as well.
                    let allowed () = current() && (id = "edit.copy" || canEdit())
                    item.disabled <- not state.Enabled || not (allowed())
                    item.addEventListener("click", fun _ ->
                        if obj.ReferenceEquals(popup, menu) then
                            let valid = allowed()
                            close true
                            if valid then registry.Run id |> ignore)
                    menu.appendChild item |> ignore
            add "edit.undo" "Undo"
            add "edit.redo" "Redo"
            add "edit.cut" "Cut"
            add "edit.copy" "Copy"
            add "edit.paste" "Paste"
            add "format.bold" "Bold"
            add "format.italic" "Italic"
            add "format.underline" "Underline"
            add "format.fontDialog" "Font…"
            add "format.paragraphDialog" "Paragraph…"
            let link = Word.getHyperlink editor
            add "insert.link" (if isNullOrUndefined link then "Insert link…" else "Edit link…")
            if not (isNullOrUndefined link) then add "link.remove" "Remove link"
            if (Doc.caretContext editor.Doc).InTable then
                for id, label in [
                    "table.insertRow", "Insert row"
                    "table.insertColumn", "Insert column"
                    "table.deleteRow", "Delete row"
                    "table.deleteColumn", "Delete column"
                    "table.cellStyle", "Cell formatting…"
                    "table.margins", "Cell margins…"
                ] do add id label
            popup <- menu
            root.appendChild menu |> ignore
            let bounds = menu.getBoundingClientRect()
            let left = max 8.0 (min x (window.innerWidth - bounds.width - 8.0))
            let top = max 8.0 (min y (window.innerHeight - bounds.height - 8.0))
            menu.setAttribute("style", sprintf "position:fixed;z-index:40;left:%gpx;top:%gpx" left top)
            menu.tabIndex <- -1
            let first = menu.querySelector("button:not(:disabled)")
            focus (if isNull first then box menu else box first)
        let ownsEditorTarget (event: Event) =
            obj.ReferenceEquals(event.target, editorInput())
            || (host.Canvas |> Option.exists (fun canvas -> obj.ReferenceEquals(event.target, canvas)))
        let onContext (event: Event) =
            if ownsEditorTarget event && not (isComposing()) then
                event.preventDefault()
                if now() - keyboardOpened > 500.0 then
                    match host.Editor, host.Canvas with
                    | Some editor, Some canvas ->
                        let mouse = event :?> MouseEvent
                        let rect = canvas.getBoundingClientRect()
                        let hit = Word.hitTestViewport editor (mouse.clientX - rect.left) (mouse.clientY - rect.top)
                        if not (isNullOrUndefined hit) then
                            let blockId: string = hit?BlockId
                            let offset: int = hit?Offset
                            let mutable inside = false
                            Doc.forEachSelectedParagraphRange editor.Doc (fun paragraph first last ->
                                if paragraph.Id = blockId && offset >= first && offset <= last then inside <- true
                                not inside) |> ignore
                            if inside || Word.setSelection editor blockId offset blockId offset then
                                openAt editor mouse.clientX mouse.clientY
                    | _ -> ()
        let onKey (event: Event) =
            let key = event :?> KeyboardEvent
            if not (composing key) && not (isComposing()) then
                if not (isNull popup) && popup.contains(event.target :?> Node) then
                    match key.key with
                    | "Escape" | "Tab" -> close true; key.preventDefault(); key.stopPropagation()
                    | "ArrowDown" | "ArrowUp" | "Home" | "End" ->
                        let items = popup.querySelectorAll("button:not(:disabled)")
                        if items.length > 0 then
                            let mutable index = 0
                            for i = 0 to items.length - 1 do
                                if obj.ReferenceEquals(items.[i], document.activeElement) then index <- i
                            let next =
                                match key.key with
                                | "Home" -> 0 | "End" -> items.length - 1
                                | "ArrowUp" -> (index + items.length - 1) % items.length
                                | _ -> (index + 1) % items.length
                            focus items.[next]
                        key.preventDefault(); key.stopPropagation()
                    | _ -> ()
                elif ownsEditorTarget event && (key.key = "ContextMenu" || (key.key = "F10" && key.shiftKey && not key.ctrlKey && not key.altKey && not key.metaKey)) then
                    key.preventDefault(); key.stopPropagation()
                    if not key.repeat then
                        match host.Editor, host.Canvas with
                        | Some editor, Some canvas ->
                            let caret = Word.getCaretRect editor
                            if not (isNullOrUndefined caret) then
                                let rect = canvas.getBoundingClientRect()
                                keyboardOpened <- now()
                                openAt editor (rect.left + unbox<float> caret?ViewportX) (rect.top + unbox<float> caret?ViewportY + unbox<float> caret?ViewportHeight)
                        | _ -> ()
        let onOutside (event: Event) =
            if not (isNull popup) && not (popup.contains(event.target :?> Node)) then close false
        let onResize (_: Event) = close false
        document.addEventListener("contextmenu", onContext)
        document.addEventListener("keydown", onKey, true)
        document.addEventListener("pointerdown", onOutside, true)
        document.addEventListener("focusin", onOutside, true)
        document.addEventListener("scroll", onOutside, true)
        window.addEventListener("resize", onResize)
        let unsubscribe = registry.OnChange(fun () -> if not (stillCurrent()) then close false)
        fun () ->
            disposed <- true
            close false
            unsubscribe()
            document.removeEventListener("contextmenu", onContext)
            document.removeEventListener("keydown", onKey, true)
            document.removeEventListener("pointerdown", onOutside, true)
            document.removeEventListener("focusin", onOutside, true)
            document.removeEventListener("scroll", onOutside, true)
            window.removeEventListener("resize", onResize)
#endif
