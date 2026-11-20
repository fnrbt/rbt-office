namespace Rbt.Word
open Rbt.Vb

module Vba =
    open Values
    open Ast

#if FABLE_COMPILER
    open Fable.Core.JsInterop

    module Doc = Document
    module TextBuf = TextBuffer
    module WordApi = API

    let private eq (left: string) (right: string) =
        left.ToLowerInvariant() = right.ToLowerInvariant()

    type private HostObject(typeName: string, getMember: string -> Value option, setMember: string -> Value -> bool, callMethod: string -> Value array -> Value option) =
        interface IHostObject with
            member _.TypeName = typeName
            member _.GetMember name = getMember name
            member _.SetMember name value = setMember name value
            member _.CallMethod name args = callMethod name args
            member _.CallNamed name args _ = callMethod name args
            member _.Enumerate () = None

    let private host (typeName: string) (getMember: string -> Value option) (callMethod: string -> Value array -> Value option) =
        VHostObject (HostObject(typeName, getMember, (fun _ _ -> false), callMethod) :> IHostObject)

    let private isDefaultMemberName (name: string) =
        name = "" || eq name "_Default" || eq name "Item"

    let private valueToInt = function
        | VInteger i -> i
        | VDouble d -> int d
        | VBoolean b -> if b then 1 else 0
        | VString s -> match System.Int32.TryParse(s) with true, i -> i | _ -> 0
        | _ -> 0

    let private valueToDouble = function
        | VInteger i -> float i
        | VDouble d -> d
        | VBoolean b -> if b then 1.0 else 0.0
        | VString s -> match System.Double.TryParse(s) with true, d -> d | _ -> 0.0
        | _ -> 0.0

    let private valueToString = function
        | VString s -> s
        | VInteger i -> string i
        | VDouble d -> string d
        | VBoolean b -> if b then "True" else "False"
        | _ -> ""

    let private jsValue (value: obj) =
        if isNull value then VEmpty
        else
            match value with
            | :? bool as b -> VBoolean b
            | :? int as i -> VInteger i
            | :? float as f ->
                if System.Double.IsNaN f then VEmpty
                elif abs (f - float (int f)) < 0.000001 then VInteger (int f)
                else VDouble f
            | :? string as s -> VString s
            | _ -> VString (string value)

    let private propValue name (source: obj) =
        jsValue source?(name)

    let private paragraphsInOrder (doc: Doc.Document) =
        doc.Blocks
        |> List.collect (function
            | Doc.Paragraph p
            | Doc.Heading (_, p) -> [p]
            | Doc.Table t ->
                t.Rows
                |> List.collect (fun row ->
                    row.Cells
                    |> List.collect (fun cell -> cell.Content))
            | Doc.TextBox t -> t.Content)

    let private resolvePosition (doc: Doc.Document) (offset: int) =
        let mutable remaining = max 0 offset
        let mutable found = None
        for para in paragraphsInOrder doc do
            if found.IsNone then
                let len = TextBuf.length para.Buffer
                if remaining <= len then
                    found <- Some (para.Id, remaining)
                else
                    // Match Document.getText's paragraph separator.
                    remaining <- remaining - len - 1
        match found with
        | Some pos -> pos
        | None ->
            match paragraphsInOrder doc |> List.rev with
            | last :: _ -> last.Id, TextBuf.length last.Buffer
            | [] -> "", 0

    let private resolveRange (editor: WordApi.Editor) startOffset endOffset =
        let blockId, localStart = resolvePosition editor.Doc startOffset
        let _, localEnd = resolvePosition editor.Doc endOffset
        blockId, localStart, max localStart localEnd

    let private infoByConstant constant info =
        match constant with
        | 1 -> propValue "wdActiveEndAdjustedPageNumber" info
        | 2 -> propValue "wdActiveEndSectionNumber" info
        | 3 -> propValue "wdActiveEndPageNumber" info
        | 4 -> propValue "wdNumberOfPagesInDocument" info
        | 5 -> propValue "wdHorizontalPositionRelativeToPage" info
        | 6 -> propValue "wdVerticalPositionRelativeToPage" info
        | 7 -> propValue "wdHorizontalPositionRelativeToTextBoundary" info
        | 8 -> propValue "wdVerticalPositionRelativeToTextBoundary" info
        | 9 -> propValue "wdFirstCharacterColumnNumber" info
        | 10 -> propValue "wdFirstCharacterLineNumber" info
        | 12 -> propValue "wdWithInTable" info
        | 13 -> propValue "wdStartOfRangeRowNumber" info
        | 14 -> propValue "wdEndOfRangeRowNumber" info
        | 15 -> propValue "wdMaximumNumberOfRows" info
        | 16 -> propValue "wdStartOfRangeColumnNumber" info
        | 17 -> propValue "wdEndOfRangeColumnNumber" info
        | 18 -> propValue "wdMaximumNumberOfColumns" info
        | 31 -> propValue "wdAtEndOfRowMarker" info
        | 44 -> propValue "wdInFieldCode" info
        | 45 -> propValue "wdInFieldResult" info
        | 46 -> propValue "wdInContentControl" info
        | _ -> VEmpty

    let private statisticByConstant constant stats =
        match constant with
        | 0 -> propValue "Words" stats
        | 1 -> propValue "Lines" stats
        | 2 -> propValue "Pages" stats
        | 3 -> propValue "Characters" stats
        | 4 -> propValue "Paragraphs" stats
        | 5 -> propValue "CharactersWithSpaces" stats
        | _ -> VEmpty

    let rec private rangeObject (editor: WordApi.Editor) blockId startOffset endOffset =
        let getMember (name: string) =
            if eq name "Start" then Some (VInteger startOffset)
            elif eq name "End" then Some (VInteger endOffset)
            elif eq name "Text" then
                match Doc.findParagraph editor.Doc blockId with
                | Some para ->
                    let text = TextBuf.toString para.Buffer
                    let start = max 0 (min startOffset text.Length)
                    let finish = max start (min endOffset text.Length)
                    Some (VString (text.Substring(start, finish - start)))
                | None -> Some VEmpty
            elif eq name "Duplicate" then Some (rangeObject editor blockId startOffset endOffset)
            elif eq name "Document" then Some (documentObject editor)
            elif eq name "Words" then
                Some (host "Word.Words"
                    (fun field ->
                        if eq field "Count" then
                            WordApi.computeRangeStatistics editor blockId startOffset endOffset
                            |> statisticByConstant 0
                            |> Some
                        else None)
                    (fun _ _ -> None))
            elif eq name "Characters" then
                Some (host "Word.Characters"
                    (fun field ->
                        if eq field "Count" then
                            WordApi.computeRangeStatistics editor blockId startOffset endOffset
                            |> statisticByConstant 3
                            |> Some
                        else None)
                    (fun _ _ -> None))
            elif eq name "Paragraphs" then
                Some (host "Word.Paragraphs"
                    (fun field ->
                        if eq field "Count" then
                            WordApi.computeRangeStatistics editor blockId startOffset endOffset
                            |> statisticByConstant 4
                            |> Some
                        else None)
                    (fun _ _ -> None))
            else None

        let callMethod (name: string) (args: Value array) =
            if eq name "Information" && args.Length >= 1 then
                WordApi.getRangeInformation editor blockId startOffset endOffset
                |> infoByConstant (valueToInt args.[0])
                |> Some
            elif eq name "ComputeStatistics" && args.Length >= 1 then
                WordApi.computeRangeStatistics editor blockId startOffset endOffset
                |> statisticByConstant (valueToInt args.[0])
                |> Some
            elif eq name "Bounds" then
                Some (host "Word.RangeBounds"
                    (fun (field: string) ->
                        WordApi.getRangeBounds editor blockId startOffset endOffset
                        |> propValue field
                        |> Some)
                    (fun (_: string) (_: Value array) -> None))
            else None

        host "Word.Range" getMember callMethod

    and private selectionObject (editor: WordApi.Editor) =
        let currentRange () =
            match editor.Doc.Selection with
            | Some sel -> rangeObject editor sel.Focus.BlockId sel.Focus.Offset sel.Focus.Offset
            | None ->
                let blockId, offset = resolvePosition editor.Doc 0
                rangeObject editor blockId offset offset

        host "Word.Selection"
            (fun (name: string) ->
                if eq name "Range" then Some (currentRange ())
                else None)
            (fun (name: string) (args: Value array) ->
                if eq name "Information" && args.Length >= 1 then
                    WordApi.getSelectionInformation editor
                    |> infoByConstant (valueToInt args.[0])
                    |> Some
                else None)

    and private documentObject (editor: WordApi.Editor) =
        host "Word.Document"
            (fun (name: string) ->
                if eq name "Application" then None
                elif eq name "Content" then
                    let text = Doc.getText editor.Doc
                    let blockId, startOffset = resolvePosition editor.Doc 0
                    Some (rangeObject editor blockId startOffset text.Length)
                elif eq name "Words" then
                    Some (host "Word.Words"
                        (fun field ->
                            if eq field "Count" then
                                (Doc.getText editor.Doc).Split([|' '; '\t'; '\r'; '\n'|], System.StringSplitOptions.RemoveEmptyEntries).Length
                                |> VInteger
                                |> Some
                            else None)
                        (fun _ _ -> None))
                elif eq name "Characters" then
                    Some (host "Word.Characters"
                        (fun field ->
                            if eq field "Count" then Some (VInteger (Doc.getText editor.Doc).Length)
                            else None)
                        (fun _ _ -> None))
                elif eq name "Paragraphs" then
                    Some (host "Word.Paragraphs"
                        (fun field ->
                            if eq field "Count" then Some (VInteger (paragraphsInOrder editor.Doc).Length)
                            else None)
                        (fun _ _ -> None))
                else None)
            (fun (name: string) (args: Value array) ->
                if eq name "Range" then
                    let startOffset = if args.Length >= 1 then valueToInt args.[0] else 0
                    let endOffset = if args.Length >= 2 then valueToInt args.[1] else startOffset
                    let blockId, localStart, localEnd = resolveRange editor startOffset endOffset
                    Some (rangeObject editor blockId localStart localEnd)
                elif eq name "ComputeStatistics" && args.Length >= 1 then
                    let text = Doc.getText editor.Doc
                    let blockId, localStart, localEnd = resolveRange editor 0 text.Length
                    WordApi.computeRangeStatistics editor blockId localStart localEnd
                    |> statisticByConstant (valueToInt args.[0])
                    |> Some
                else None)

    let private applicationObject (editor: WordApi.Editor) =
        host "Word.Application"
            (fun (name: string) ->
                if eq name "ActiveDocument" then Some (documentObject editor)
                elif eq name "Selection" then Some (selectionObject editor)
                elif eq name "Version" then Some (VString "15.0")
                else None)
            (fun (name: string) (args: Value array) ->
                if eq name "InchesToPoints" && args.Length >= 1 then Some (VDouble (valueToDouble args.[0] * 72.0))
                elif eq name "CentimetersToPoints" && args.Length >= 1 then Some (VDouble (valueToDouble args.[0] * 28.3464567))
                else None)

    let private createObjectFunction (editor: WordApi.Editor) =
        host "VBScript.CreateObject"
            (fun (_: string) -> None)
            (fun (name: string) (args: Value array) ->
                if isDefaultMemberName name && args.Length >= 1 then
                    let progId = valueToString args.[0]
                    if eq progId "Word.Application" then Some (applicationObject editor)
                    else Some VNothing
                else None)

    let constants =
        [
            "wdActiveEndAdjustedPageNumber", VInteger 1
            "wdActiveEndSectionNumber", VInteger 2
            "wdActiveEndPageNumber", VInteger 3
            "wdNumberOfPagesInDocument", VInteger 4
            "wdHorizontalPositionRelativeToPage", VInteger 5
            "wdVerticalPositionRelativeToPage", VInteger 6
            "wdHorizontalPositionRelativeToTextBoundary", VInteger 7
            "wdVerticalPositionRelativeToTextBoundary", VInteger 8
            "wdFirstCharacterColumnNumber", VInteger 9
            "wdFirstCharacterLineNumber", VInteger 10
            "wdWithInTable", VInteger 12
            "wdStartOfRangeRowNumber", VInteger 13
            "wdEndOfRangeRowNumber", VInteger 14
            "wdMaximumNumberOfRows", VInteger 15
            "wdStartOfRangeColumnNumber", VInteger 16
            "wdEndOfRangeColumnNumber", VInteger 17
            "wdMaximumNumberOfColumns", VInteger 18
            "wdAtEndOfRowMarker", VInteger 31
            "wdInFieldCode", VInteger 44
            "wdInFieldResult", VInteger 45
            "wdInContentControl", VInteger 46
            "wdStatisticWords", VInteger 0
            "wdStatisticLines", VInteger 1
            "wdStatisticPages", VInteger 2
            "wdStatisticCharacters", VInteger 3
            "wdStatisticParagraphs", VInteger 4
            "wdStatisticCharactersWithSpaces", VInteger 5
        ]

    let hostGlobals (editor: WordApi.Editor) =
        [
            "Application", applicationObject editor
            "ActiveDocument", documentObject editor
            "Selection", selectionObject editor
        ] @ constants
        |> Map.ofList

    let hostGlobalsVbs (editor: WordApi.Editor) =
        hostGlobals editor
        |> Map.add "CreateObject" (createObjectFunction editor)

    let run (editor: WordApi.Editor) (source: string) =
        let options =
            Api.RunOptions.vba
            |> Api.RunOptions.withHostGlobals (hostGlobals editor)
        Api.runSource options source

    let runVbs (editor: WordApi.Editor) (source: string) =
        let options =
            Api.RunOptions.vbScript
            |> Api.RunOptions.withHostGlobals (hostGlobalsVbs editor)
        Api.runSource options source
#else
    module Doc = Document
    module TextBuf = TextBuffer
    module WordLayout = Layout

    let private eq (left: string) (right: string) =
        left.ToLowerInvariant() = right.ToLowerInvariant()

    type private HostObject(typeName: string, getMember: string -> Value option, setMember: string -> Value -> bool, callMethod: string -> Value array -> Value option) =
        interface IHostObject with
            member _.TypeName = typeName
            member _.GetMember name = getMember name
            member _.SetMember name value = setMember name value
            member _.CallMethod name args = callMethod name args
            member _.CallNamed name args _ = callMethod name args
            member _.Enumerate () = None

    let private host (typeName: string) (getMember: string -> Value option) (callMethod: string -> Value array -> Value option) =
        VHostObject (HostObject(typeName, getMember, (fun _ _ -> false), callMethod) :> IHostObject)

    let private isDefaultMemberName (name: string) =
        name = "" || eq name "_Default" || eq name "Item"

    let private valueToInt = function
        | VInteger i -> i
        | VDouble d -> int d
        | VBoolean b -> if b then 1 else 0
        | VString s -> match System.Int32.TryParse(s) with true, i -> i | _ -> 0
        | _ -> 0

    let private valueToDouble = function
        | VInteger i -> float i
        | VDouble d -> d
        | VBoolean b -> if b then 1.0 else 0.0
        | VString s -> match System.Double.TryParse(s) with true, d -> d | _ -> 0.0
        | _ -> 0.0

    let private valueToString = function
        | VString s -> s
        | VInteger i -> string i
        | VDouble d -> string d
        | VBoolean b -> if b then "True" else "False"
        | _ -> ""

    let private countWords (text: string) =
        text.Split([|' '; '\t'; '\r'; '\n'|], System.StringSplitOptions.RemoveEmptyEntries).Length

    let private paragraphsInOrder (doc: Doc.Document) =
        doc.Blocks
        |> List.collect (function
            | Doc.Paragraph p
            | Doc.Heading (_, p) -> [p]
            | Doc.Table t ->
                t.Rows
                |> List.collect (fun row ->
                    row.Cells
                    |> List.collect (fun cell -> cell.Content))
            | Doc.TextBox t -> t.Content)

    let private resolvePosition (doc: Doc.Document) (offset: int) =
        let mutable remaining = max 0 offset
        let mutable found = None
        for para in paragraphsInOrder doc do
            if found.IsNone then
                let len = TextBuf.length para.Buffer
                if remaining <= len then
                    found <- Some (para.Id, remaining)
                else
                    remaining <- remaining - len - 1
        match found with
        | Some pos -> pos
        | None ->
            match paragraphsInOrder doc |> List.rev with
            | last :: _ -> last.Id, TextBuf.length last.Buffer
            | [] -> "", 0

    let private nativeSettings (doc: Doc.Document) : WordLayout.LayoutSettings = {
        Width = doc.PageSetup.PageWidth
        PageHeight = doc.PageSetup.PageHeight
        FontSize = 16.0
        LineHeight = 22.0
        PaddingLeft = doc.PageSetup.MarginLeft
        PaddingTop = doc.PageSetup.MarginTop
        MarginRight = doc.PageSetup.MarginRight
        MarginBottom = doc.PageSetup.MarginBottom
        InTableCell = false
        TableContinuationHeaderHeight = 0.0
        LongTrailingSpaceFitSafety = 0.0
        DefaultTabStop = doc.PageSetup.DefaultTabStop
        CompatibilityMode = doc.PageSetup.CompatibilityMode
        HangingPunctuation = Layout.documentHangingPunctuation doc.PageSetup
        SpaceShrink = Layout.documentSpaceShrink doc.PageSetup
    }

    let private measureText (text: string) (style: TextBuf.TextStyle) =
        let size = if style.FontSize > 0.0 then style.FontSize else 11.0
        let weight = if style.Bold then 0.58 else 0.52
        float text.Length * size * weight

    let private measureRange = WordLayout.rangeMeasurerFromTextMeasurer measureText

    let private layoutsFor doc =
        WordLayout.layoutDocumentWithRange (nativeSettings doc) measureText measureRange doc

    let private pageForY (doc: Doc.Document) y =
        int (floor (max 0.0 y / (doc.PageSetup.PageHeight + WordLayout.pageGap))) + 1

    let private pageTop (doc: Doc.Document) page =
        float (page - 1) * (doc.PageSetup.PageHeight + WordLayout.pageGap)

    let private rangeText (doc: Doc.Document) startOffset endOffset =
        let text = Doc.getText doc
        let start = max 0 (min startOffset text.Length)
        let finish = max start (min endOffset text.Length)
        text.Substring(start, finish - start)

    let private statisticsByConstant (doc: Doc.Document) startOffset endOffset constant =
        let text = rangeText doc startOffset endOffset
        let blockId, localStart = resolvePosition doc startOffset
        let _, localEnd = resolvePosition doc endOffset
        let layouts = layoutsFor doc
        let lines =
            layouts
            |> List.filter (fun layout -> layout.BlockId = blockId)
            |> List.collect (fun layout -> layout.Lines)
            |> List.filter (fun line ->
                if localStart = localEnd then localStart >= line.StartOffset && localStart <= line.EndOffset
                else localStart < line.EndOffset && localEnd > line.StartOffset)
        let pages = lines |> List.map (fun line -> pageForY doc line.Y) |> List.distinct
        match constant with
        | 0 -> VInteger (countWords text)
        | 1 -> VInteger lines.Length
        | 2 -> VInteger pages.Length
        | 3 -> VInteger text.Length
        | 4 -> VInteger (max 1 (text.Split([|'\n'|]).Length))
        | 5 -> VInteger text.Length
        | _ -> VEmpty

    let private rangeInfoByConstant (doc: Doc.Document) startOffset endOffset constant =
        let blockId, localOffset = resolvePosition doc startOffset
        let layouts = layoutsFor doc
        let pageCount =
            layouts
            |> List.collect (fun layout -> layout.Lines)
            |> List.map (fun line -> pageForY doc (line.Y + line.Height))
            |> List.fold max 1
        match layouts |> List.tryFind (fun layout -> layout.BlockId = blockId) with
        | None -> VEmpty
        | Some layout ->
            let line = WordLayout.lineForOffset layout localOffset
            let alignOffset = WordLayout.alignmentOffset layout line
            let x = doc.PageSetup.MarginLeft + alignOffset + WordLayout.xForOffsetWithRange measureRange line localOffset
            let page = pageForY doc line.Y
            let pageY = line.Y - pageTop doc page
            match constant with
            | 1 -> VInteger page
            | 2 -> VInteger 1
            | 3 -> VInteger page
            | 4 -> VInteger pageCount
            | 5 -> VDouble x
            | 6 -> VDouble pageY
            | 7 -> VDouble (x - doc.PageSetup.MarginLeft)
            | 8 -> VDouble (max 0.0 (line.Y - doc.PageSetup.MarginTop))
            | 9 -> VInteger (max 1 (localOffset - line.StartOffset + 1))
            | 10 -> VInteger 1
            | 12 -> VBoolean false
            | 31 -> VBoolean false
            | 44 -> VBoolean false
            | 45 -> VBoolean false
            | 46 -> VBoolean false
            | _ -> VEmpty

    let private rangeBoundsObject (doc: Doc.Document) startOffset endOffset =
        host "Word.RangeBounds"
            (fun (field: string) ->
                let blockId, localStart = resolvePosition doc startOffset
                let _, localEnd = resolvePosition doc endOffset
                let layouts = layoutsFor doc
                match layouts |> List.tryFind (fun layout -> layout.BlockId = blockId) with
                | None -> Some VEmpty
                | Some layout ->
                    let line = WordLayout.lineForOffset layout localStart
                    let alignOffset = WordLayout.alignmentOffset layout line
                    let x1 = doc.PageSetup.MarginLeft + alignOffset + WordLayout.xForOffsetWithRange measureRange line localStart
                    let x2 = doc.PageSetup.MarginLeft + alignOffset + WordLayout.xForOffsetWithRange measureRange line (max localStart localEnd)
                    let page = pageForY doc line.Y
                    if eq field "X" then Some (VDouble x1)
                    elif eq field "Y" then Some (VDouble line.Y)
                    elif eq field "PageY" then Some (VDouble (line.Y - pageTop doc page))
                    elif eq field "Width" then Some (VDouble (max 0.0 (x2 - x1)))
                    elif eq field "Height" then Some (VDouble line.Height)
                    else Some VEmpty)
            (fun (_: string) (_: Value array) -> None)

    let rec private rangeObject (doc: Doc.Document) startOffset endOffset =
        host "Word.Range"
            (fun (name: string) ->
                if eq name "Start" then Some (VInteger startOffset)
                elif eq name "End" then Some (VInteger endOffset)
                elif eq name "Text" then Some (VString (rangeText doc startOffset endOffset))
                elif eq name "Duplicate" then Some (rangeObject doc startOffset endOffset)
                elif eq name "Document" then Some (documentObject doc)
                elif eq name "Words" then
                    Some (host "Word.Words"
                        (fun field -> if eq field "Count" then Some (statisticsByConstant doc startOffset endOffset 0) else None)
                        (fun _ _ -> None))
                elif eq name "Characters" then
                    Some (host "Word.Characters"
                        (fun field -> if eq field "Count" then Some (statisticsByConstant doc startOffset endOffset 3) else None)
                        (fun _ _ -> None))
                elif eq name "Paragraphs" then
                    Some (host "Word.Paragraphs"
                        (fun field -> if eq field "Count" then Some (statisticsByConstant doc startOffset endOffset 4) else None)
                        (fun _ _ -> None))
                else None)
            (fun (name: string) (args: Value array) ->
                if eq name "Information" && args.Length >= 1 then Some (rangeInfoByConstant doc startOffset endOffset (valueToInt args.[0]))
                elif eq name "ComputeStatistics" && args.Length >= 1 then Some (statisticsByConstant doc startOffset endOffset (valueToInt args.[0]))
                elif eq name "Bounds" then Some (rangeBoundsObject doc startOffset endOffset)
                else None)

    and private selectionObject (doc: Doc.Document) =
        let currentRange () =
            match doc.Selection with
            | Some sel ->
                let absolute =
                    let mutable acc = 0
                    let mutable result = sel.Focus.Offset
                    for para in paragraphsInOrder doc do
                        if para.Id = sel.Focus.BlockId then
                            result <- acc + sel.Focus.Offset
                        acc <- acc + TextBuf.length para.Buffer + 1
                    result
                rangeObject doc absolute absolute
            | None -> rangeObject doc 0 0
        host "Word.Selection"
            (fun (name: string) ->
                if eq name "Range" then Some (currentRange ())
                elif eq name "Start" then Some (VInteger 0)
                elif eq name "End" then Some (VInteger 0)
                else None)
            (fun (name: string) (args: Value array) ->
                if eq name "Information" && args.Length >= 1 then Some (rangeInfoByConstant doc 0 0 (valueToInt args.[0]))
                else None)

    and private documentObject (doc: Doc.Document) =
        host "Word.Document"
            (fun (name: string) ->
                if eq name "Application" then None
                elif eq name "Content" then Some (rangeObject doc 0 (Doc.getText doc).Length)
                elif eq name "Words" then
                    Some (host "Word.Words"
                        (fun field -> if eq field "Count" then Some (statisticsByConstant doc 0 (Doc.getText doc).Length 0) else None)
                        (fun _ _ -> None))
                elif eq name "Characters" then
                    Some (host "Word.Characters"
                        (fun field -> if eq field "Count" then Some (statisticsByConstant doc 0 (Doc.getText doc).Length 3) else None)
                        (fun _ _ -> None))
                elif eq name "Paragraphs" then
                    Some (host "Word.Paragraphs"
                        (fun field -> if eq field "Count" then Some (VInteger (paragraphsInOrder doc).Length) else None)
                        (fun _ _ -> None))
                else None)
            (fun (name: string) (args: Value array) ->
                if eq name "Range" then
                    let startOffset = if args.Length >= 1 then valueToInt args.[0] else 0
                    let endOffset = if args.Length >= 2 then valueToInt args.[1] else startOffset
                    Some (rangeObject doc startOffset endOffset)
                elif eq name "ComputeStatistics" && args.Length >= 1 then
                    Some (statisticsByConstant doc 0 (Doc.getText doc).Length (valueToInt args.[0]))
                else None)

    let private applicationObject (doc: Doc.Document) =
        host "Word.Application"
            (fun (name: string) ->
                if eq name "ActiveDocument" then Some (documentObject doc)
                elif eq name "Selection" then Some (selectionObject doc)
                elif eq name "Version" then Some (VString "15.0")
                else None)
            (fun (name: string) (args: Value array) ->
                if eq name "InchesToPoints" && args.Length >= 1 then Some (VDouble (valueToDouble args.[0] * 72.0))
                elif eq name "CentimetersToPoints" && args.Length >= 1 then Some (VDouble (valueToDouble args.[0] * 28.3464567))
                else None)

    let private createObjectFunction (doc: Doc.Document) =
        host "VBScript.CreateObject"
            (fun (_: string) -> None)
            (fun (name: string) (args: Value array) ->
                if isDefaultMemberName name && args.Length >= 1 then
                    let progId = valueToString args.[0]
                    if eq progId "Word.Application" then Some (applicationObject doc)
                    else Some VNothing
                else None)

    let private constants =
        [
            "wdActiveEndAdjustedPageNumber", VInteger 1
            "wdActiveEndSectionNumber", VInteger 2
            "wdActiveEndPageNumber", VInteger 3
            "wdNumberOfPagesInDocument", VInteger 4
            "wdHorizontalPositionRelativeToPage", VInteger 5
            "wdVerticalPositionRelativeToPage", VInteger 6
            "wdHorizontalPositionRelativeToTextBoundary", VInteger 7
            "wdVerticalPositionRelativeToTextBoundary", VInteger 8
            "wdFirstCharacterColumnNumber", VInteger 9
            "wdFirstCharacterLineNumber", VInteger 10
            "wdWithInTable", VInteger 12
            "wdAtEndOfRowMarker", VInteger 31
            "wdInFieldCode", VInteger 44
            "wdInFieldResult", VInteger 45
            "wdInContentControl", VInteger 46
            "wdStatisticWords", VInteger 0
            "wdStatisticLines", VInteger 1
            "wdStatisticPages", VInteger 2
            "wdStatisticCharacters", VInteger 3
            "wdStatisticParagraphs", VInteger 4
            "wdStatisticCharactersWithSpaces", VInteger 5
        ]

    let hostGlobalsDocument (doc: Doc.Document) =
        [
            "Application", applicationObject doc
            "ActiveDocument", documentObject doc
            "Selection", selectionObject doc
        ] @ constants
        |> Map.ofList

    let hostGlobalsDocumentVbs (doc: Doc.Document) =
        hostGlobalsDocument doc
        |> Map.add "CreateObject" (createObjectFunction doc)

    let runDocument (doc: Doc.Document) (source: string) =
        let options =
            Api.RunOptions.vba
            |> Api.RunOptions.withHostGlobals (hostGlobalsDocument doc)
        Api.runSource options source

    let runDocumentVbs (doc: Doc.Document) (source: string) =
        let options =
            Api.RunOptions.vbScript
            |> Api.RunOptions.withHostGlobals (hostGlobalsDocumentVbs doc)
        Api.runSource options source

    let hostGlobals (_: API.Editor) : Map<string, Value> = Map.empty
    let run (_: API.Editor) (_: string) =
        Error (Api.RuntimeError "Use Word.Vba.runDocument for native Word VBA host execution")
    let runVbs (_: API.Editor) (_: string) =
        Error (Api.RuntimeError "Use Word.Vba.runDocumentVbs for native Word VBS host execution")
#endif
