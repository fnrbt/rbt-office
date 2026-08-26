module Rbt.Spreadsheet.Tabs

open Fable.Core
open Fable.Core.JsInterop
open Rbt.Grid.Dom
open Rbt.Spreadsheet.Workbook

/// JavaScript helper for prompts
[<Emit("prompt($0, $1)")>]
let private prompt (message: string) (defaultValue: string): string = jsNative

/// Tab event handlers
type TabHandlers = {
    OnTabClick: SheetId -> unit
    OnTabAdd: unit -> unit
    OnTabRename: SheetId * string -> unit
    OnTabDelete: SheetId -> unit
    OnTabReorder: SheetId list -> unit
}

/// Tabs component for sheet navigation
type Tabs(workbook: Workbook, handlers: TabHandlers) =
    let mutable tabsElement: Element option = None
    let mutable currentWorkbook = workbook

    let createTab (sheetId: SheetId) (sheetName: string) (isActive: bool): Element =
        let tab = createElement "div"
        setClassName tab (if isActive then "sheet-tab active" else "sheet-tab")
        setStyles tab [
            "display", "inline-block"
            "padding", "8px 16px"
            "border", "1px solid #d0d0d0"
            "border-bottom", if isActive then "none" else "1px solid #d0d0d0"
            "background", if isActive then "#fff" else "#f0f0f0"
            "cursor", "pointer"
            "margin-right", "2px"
            "border-radius", "4px 4px 0 0"
            "user-select", "none"
            "font-size", "12px"
            "position", "relative"
            "min-width", "80px"
        ]

        // Tab name
        let nameSpan = createElement "span"
        setClassName nameSpan "tab-name"
        setTextContent nameSpan sheetName
        appendChild tab nameSpan

        // Close button (if more than one sheet)
        if Map.count currentWorkbook.Sheets > 1 then
            let closeBtn = createElement "span"
            setClassName closeBtn "tab-close"
            setTextContent closeBtn "×"
            setStyles closeBtn [
                "margin-left", "8px"
                "font-weight", "bold"
                "color", "#999"
                "cursor", "pointer"
            ]

            addEventListener closeBtn "mouseenter" (fun _ ->
                setStyle closeBtn "color" "#e74c3c"
            )

            addEventListener closeBtn "mouseleave" (fun _ ->
                setStyle closeBtn "color" "#999"
            )

            addEventListener closeBtn "click" (fun e ->
                stopPropagation e
                preventDefault e
                handlers.OnTabDelete sheetId
            )

            appendChild tab closeBtn

        // Tab click
        addEventListener tab "click" (fun _ ->
            handlers.OnTabClick sheetId
        )

        // Hover effect
        if not isActive then
            addEventListener tab "mouseenter" (fun _ ->
                setStyle tab "background" "#e8e8e8"
            )
            addEventListener tab "mouseleave" (fun _ ->
                setStyle tab "background" "#f0f0f0"
            )

        // Double-click to rename
        addEventListener tab "dblclick" (fun e ->
            stopPropagation e
            preventDefault e
            let newName = prompt "Enter new name:" sheetName
            if not (System.String.IsNullOrWhiteSpace newName) then
                handlers.OnTabRename (sheetId, newName)
        )

        tab

    let createAddButton (): Element =
        let btn = createElement "button"
        setClassName btn "add-sheet-btn"
        setTextContent btn "+"
        setStyles btn [
            "display", "inline-block"
            "padding", "8px 16px"
            "border", "1px solid #d0d0d0"
            "background", "#f8f8f8"
            "cursor", "pointer"
            "margin-left", "8px"
            "border-radius", "4px"
            "font-size", "16px"
            "font-weight", "bold"
            "color", "#666"
            "min-width", "40px"
        ]

        addEventListener btn "click" (fun _ ->
            handlers.OnTabAdd ()
        )

        addEventListener btn "mouseenter" (fun _ ->
            setStyle btn "background" "#e0e0e0"
        )

        addEventListener btn "mouseleave" (fun _ ->
            setStyle btn "background" "#f8f8f8"
        )

        btn

    let render (): Element =
        let container = createElement "div"
        setClassName container "sheet-tabs"
        setStyles container [
            "display", "flex"
            "align-items", "flex-end"
            "padding", "4px"
            "background", "#f8f8f8"
            "border-top", "1px solid #d0d0d0"
            "overflow-x", "auto"
            "white-space", "nowrap"
        ]

        // Render tabs in order
        for sheetId in currentWorkbook.SheetOrder do
            match Map.tryFind sheetId currentWorkbook.Sheets with
            | Some sheet ->
                let tab = createTab sheetId sheet.Name (sheetId = currentWorkbook.ActiveSheet)
                appendChild container tab
            | None -> ()

        // Add button
        let addBtn = createAddButton()
        appendChild container addBtn

        container

    member _.Mount(container: Element) =
        let tabs = render()
        appendChild container tabs
        tabsElement <- Some tabs

    member _.Update(workbook: Workbook) =
        currentWorkbook <- workbook
        match tabsElement with
        | Some tabs ->
            // Re-render tabs
            setInnerHTML tabs ""
            let newTabs = render()
            let children = getChildren newTabs
            for child in children do
                appendChild tabs child
        | None -> ()

    member _.GetElement() = tabsElement

/// Export for JavaScript
[<ExportDefault>]
let createTabs (workbook: Workbook) (handlers: TabHandlers): Tabs =
    new Tabs(workbook, handlers)
