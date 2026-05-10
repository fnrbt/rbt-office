module RedBat.FableUI.ParagraphStyleDialogTest

open Browser.Dom
open Browser.Types
open RedBat.FableUI.Reactive
open RedBat.FableUI.Dom
open RedBat.FableUI.UI

let private selectValue (event: Event) =
    (event.target :?> HTMLSelectElement).value

let app () =
    let isOpen = signal false
    let alignment = signal "left"
    let spacingAfter = signal "10"
    let fontSize = signal "11"
    let result = signal "Alignment: left; spacing after: 10pt; font size: 11pt"

    let apply () =
        result.Value <- sprintf "Alignment: %s; spacing after: %spt; font size: %spt" alignment.Value spacingAfter.Value fontSize.Value
        isOpen.Value <- false

    div [ cls "dialog-test-shell" ] [
        h2 [] [ text "Paragraph Style Dialog" ]
        p [ attr "data-testid" "paragraph-style-summary" ] [ textR (fun () -> result.Value) ]
        Button [ attr "data-testid" "open-paragraph-style"; onClick (fun _ -> isOpen.Value <- true) ] [ text "Paragraph style" ]

        DialogWithTestIds "paragraph-dialog-overlay" "paragraph-dialog" isOpen [
            DialogHeader [] [ DialogTitle [] [ text "Paragraph style" ] ]
            DialogBody [] [
                label [] [
                    span [] [ text "Alignment" ]
                    select [ attr "data-testid" "paragraph-align"; propS "value" alignment; onChange (fun event -> alignment.Value <- selectValue event) ] [
                        option [ attr "value" "left" ] [ text "Left" ]
                        option [ attr "value" "center" ] [ text "Center" ]
                        option [ attr "value" "right" ] [ text "Right" ]
                        option [ attr "value" "justify" ] [ text "Justify" ]
                    ]
                ]
                label [] [
                    span [] [ text "Spacing after" ]
                    Input [ typ "number"; attr "data-testid" "paragraph-spacing-after"; valueS spacingAfter; onInput (fun value -> spacingAfter.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Font size" ]
                    Input [ typ "number"; attr "data-testid" "paragraph-font-size"; valueS fontSize; onInput (fun value -> fontSize.Value <- value) ]
                ]
            ]
            DialogFooter [] [
                Button [ attr "data-testid" "cancel-paragraph-style"; onClick (fun _ -> isOpen.Value <- false) ] [ text "Cancel" ]
                Button [ attr "data-testid" "apply-paragraph-style"; onClick (fun _ -> apply()) ] [ text "Apply" ]
            ]
        ]
    ]

let mount selector = mountApp selector (app ())
