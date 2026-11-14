namespace Rbt

module BorderGrid =
    open System.Collections.Generic

    type Stroke = {
        Color: string
        Width: float
        Style: string option
        Priority: int
    }

    type Collapse =
        | Separate
        | Collapse

    type Edges = {
        Top: Stroke option
        Right: Stroke option
        Bottom: Stroke option
        Left: Stroke option
    }

    type Cell = {
        Row: int
        Column: int
        RowSpan: int
        ColumnSpan: int
        X: float
        Y: float
        Width: float
        Height: float
        Edges: Edges
    }

    type Segment = {
        X1: float
        Y1: float
        X2: float
        Y2: float
        Stroke: Stroke
    }

    type private Candidate = {
        Fixed: float
        Start: float
        Finish: float
        Horizontal: bool
        Stroke: Stroke
    }

    let private resizeArrayToList (items: ResizeArray<'T>) =
        let mutable result = []
        for i = items.Count - 1 downto 0 do
            result <- items.[i] :: result
        result

    let normalizeStyle (style: string option) =
        match style |> Option.map (fun value -> value.Trim()) with
        | Some value when value <> "" -> Some value
        | _ -> Some "single"

    let private equalsIgnoreCase (left: string) (right: string) =
        if left.Length <> right.Length then
            false
        else
            let mutable index = 0
            let mutable equal = true
            while equal && index < left.Length do
                let l = left.[index]
                let r = right.[index]
                if l <> r && System.Char.ToUpperInvariant l <> System.Char.ToUpperInvariant r then
                    equal <- false
                index <- index + 1
            equal

    let styleKey (style: string option) =
        let value = normalizeStyle style |> Option.defaultValue "single"
        if equalsIgnoreCase value "double" then "double"
        elif equalsIgnoreCase value "triple" then "triple"
        elif equalsIgnoreCase value "thinthicksmallgap" then "thinthicksmallgap"
        elif equalsIgnoreCase value "thickthinsmallgap" then "thickthinsmallgap"
        elif equalsIgnoreCase value "thinthickthinsmallgap" then "thinthickthinsmallgap"
        elif equalsIgnoreCase value "thinthickmediumgap" then "thinthickmediumgap"
        elif equalsIgnoreCase value "thickthinmediumgap" then "thickthinmediumgap"
        elif equalsIgnoreCase value "thinthickthinmediumgap" then "thinthickthinmediumgap"
        elif equalsIgnoreCase value "thinthicklargegap" then "thinthicklargegap"
        elif equalsIgnoreCase value "thickthinlargegap" then "thickthinlargegap"
        elif equalsIgnoreCase value "thinthickthinlargegap" then "thinthickthinlargegap"
        elif equalsIgnoreCase value "hidden" then "hidden"
        elif equalsIgnoreCase value "thick" then "thick"
        elif equalsIgnoreCase value "single" then "single"
        elif equalsIgnoreCase value "solid" then "solid"
        elif equalsIgnoreCase value "doublewave" then "doublewave"
        elif equalsIgnoreCase value "wave" then "wave"
        elif equalsIgnoreCase value "dashed" then "dashed"
        elif equalsIgnoreCase value "dashsmallgap" then "dashsmallgap"
        elif equalsIgnoreCase value "dotdash" then "dotdash"
        elif equalsIgnoreCase value "dashdotstroked" then "dashdotstroked"
        elif equalsIgnoreCase value "dotted" then "dotted"
        elif equalsIgnoreCase value "dotdotdash" then "dotdotdash"
        elif equalsIgnoreCase value "threedemboss" then "threedemboss"
        elif equalsIgnoreCase value "threedengrave" then "threedengrave"
        elif equalsIgnoreCase value "inset" then "inset"
        elif equalsIgnoreCase value "outset" then "outset"
        elif equalsIgnoreCase value "none" then "none"
        elif equalsIgnoreCase value "nil" then "nil"
        else value

    let strokePattern style =
        match styleKey style with
        | "double" -> [ -1.0, 1.0; 1.0, 1.0 ]
        | "triple" -> [ -2.0, 1.0; 0.0, 1.0; 2.0, 1.0 ]
        | "thinthicksmallgap" -> [ -1.5, 1.0; 1.0, 2.0 ]
        | "thickthinsmallgap" -> [ -1.0, 2.0; 1.5, 1.0 ]
        | "thinthickthinsmallgap" -> [ -2.0, 1.0; 0.0, 2.0; 2.0, 1.0 ]
        | "thinthickmediumgap" -> [ -2.0, 1.0; 1.0, 2.0 ]
        | "thickthinmediumgap" -> [ -1.0, 2.0; 2.0, 1.0 ]
        | "thinthickthinmediumgap" -> [ -3.0, 1.0; 0.0, 2.0; 3.0, 1.0 ]
        | "thinthicklargegap" -> [ -2.5, 1.0; 1.5, 2.0 ]
        | "thickthinlargegap" -> [ -1.5, 2.0; 2.5, 1.0 ]
        | "thinthickthinlargegap" -> [ -4.0, 1.0; 0.0, 2.0; 4.0, 1.0 ]
        | _ -> [0.0, 1.0]

    let private normalizedStyle stroke = styleKey stroke.Style

    let private styleRank = function
        | "hidden" -> 100
        | "double"
        | "triple"
        | "thinthicksmallgap"
        | "thickthinsmallgap"
        | "thinthickthinsmallgap"
        | "thinthickmediumgap"
        | "thickthinmediumgap"
        | "thinthickthinmediumgap"
        | "thinthicklargegap"
        | "thickthinlargegap"
        | "thinthickthinlargegap" -> 90
        | "thick" -> 80
        | "single"
        | "solid" -> 70
        | "doublewave"
        | "wave" -> 65
        | "dashed" -> 60
        | "dashsmallgap"
        | "dotdash"
        | "dashdotstroked" -> 55
        | "dotted"
        | "dotdotdash" -> 50
        | "threedemboss"
        | "threedengrave"
        | "inset"
        | "outset" -> 40
        | "none"
        | "nil" -> 0
        | _ -> 70

    let private isVisible stroke =
        let style = normalizedStyle stroke
        style <> "none" && style <> "nil" && style <> "hidden" && stroke.Width > 0.0

    let private stronger first second =
        let firstStyle = normalizedStyle first.Stroke
        let secondStyle = normalizedStyle second.Stroke
        if second.Stroke.Priority > first.Stroke.Priority then second
        elif second.Stroke.Priority < first.Stroke.Priority then first
        elif styleRank secondStyle > styleRank firstStyle then second
        elif styleRank secondStyle < styleRank firstStyle then first
        elif second.Stroke.Width > first.Stroke.Width then second
        else first

    let private toCandidate horizontal fixedCoord start finish stroke =
        let a = min start finish
        let b = max start finish
        if a = b then None else Some { Fixed = fixedCoord; Start = a; Finish = b; Horizontal = horizontal; Stroke = stroke }

    let private addCandidate (result: ResizeArray<Candidate>) horizontal fixedCoord start finish stroke =
        match stroke |> Option.bind (toCandidate horizontal fixedCoord start finish) with
        | Some candidate -> result.Add candidate
        | None -> ()

    let private candidates cells =
        let result = ResizeArray<Candidate>()
        for cell in cells do
            addCandidate result true cell.Y cell.X (cell.X + cell.Width) cell.Edges.Top
            addCandidate result false (cell.X + cell.Width) cell.Y (cell.Y + cell.Height) cell.Edges.Right
            addCandidate result true (cell.Y + cell.Height) cell.X (cell.X + cell.Width) cell.Edges.Bottom
            addCandidate result false cell.X cell.Y (cell.Y + cell.Height) cell.Edges.Left
        result

    let private resolveAxis horizontal (candidates: ResizeArray<Candidate>) =
        let groups = Dictionary<float, ResizeArray<Candidate>>()
        let groupOrder = ResizeArray<float>()
        for candidate in candidates do
            if candidate.Horizontal = horizontal then
                match groups.TryGetValue candidate.Fixed with
                | true, group -> group.Add candidate
                | false, _ ->
                    let group = ResizeArray<Candidate>()
                    group.Add candidate
                    groups.[candidate.Fixed] <- group
                    groupOrder.Add candidate.Fixed

        let segments = ResizeArray<Segment>()
        for fixedCoord in groupOrder do
            let axisCandidates = groups.[fixedCoord]
            let points = ResizeArray<float>(axisCandidates.Count * 2)
            let seenPoints = HashSet<float>()
            for candidate in axisCandidates do
                if seenPoints.Add candidate.Start then points.Add candidate.Start
                if seenPoints.Add candidate.Finish then points.Add candidate.Finish
            let sortedPoints = points.ToArray()
            Array.sortInPlace sortedPoints

            let intervalCount = max 0 (sortedPoints.Length - 1)
            let winners = Array.zeroCreate<Candidate option> intervalCount
            let pointIndex = Dictionary<float, int>(sortedPoints.Length)
            for i = 0 to sortedPoints.Length - 1 do
                pointIndex.[sortedPoints.[i]] <- i

            for candidate in axisCandidates do
                let mutable startIndex = 0
                let mutable finishIndex = 0
                if pointIndex.TryGetValue(candidate.Start, &startIndex)
                   && pointIndex.TryGetValue(candidate.Finish, &finishIndex)
                   && finishIndex > startIndex then
                    for i = startIndex to finishIndex - 1 do
                        winners.[i] <-
                            match winners.[i] with
                            | None -> Some candidate
                            | Some current -> Some (stronger current candidate)

            for i = 0 to intervalCount - 1 do
                match winners.[i] with
                | Some candidate when isVisible candidate.Stroke ->
                    let startCoord = sortedPoints.[i]
                    let finishCoord = sortedPoints.[i + 1]
                    if horizontal then
                        segments.Add { X1 = startCoord; Y1 = fixedCoord; X2 = finishCoord; Y2 = fixedCoord; Stroke = candidate.Stroke }
                    else
                        segments.Add { X1 = fixedCoord; Y1 = startCoord; X2 = fixedCoord; Y2 = finishCoord; Stroke = candidate.Stroke }
                | _ -> ()
        segments

    let private resolveCandidates collapse candidates =
        match collapse with
        | Separate ->
            let segments = ResizeArray<Segment>()
            for candidate in candidates do
                if isVisible candidate.Stroke then
                    if candidate.Horizontal then
                        segments.Add { X1 = candidate.Start; Y1 = candidate.Fixed; X2 = candidate.Finish; Y2 = candidate.Fixed; Stroke = candidate.Stroke }
                    else
                        segments.Add { X1 = candidate.Fixed; Y1 = candidate.Start; X2 = candidate.Fixed; Y2 = candidate.Finish; Stroke = candidate.Stroke }
            resizeArrayToList segments
        | Collapse ->
            let segments = ResizeArray<Segment>()
            segments.AddRange(resolveAxis true candidates)
            segments.AddRange(resolveAxis false candidates)
            resizeArrayToList segments

    let resolveWith collapse (cells: Cell list) =
        candidates cells |> resolveCandidates collapse

    let resolveResizeArrayWith collapse (cells: ResizeArray<Cell>) =
        candidates cells |> resolveCandidates collapse

    let resolve cells = resolveWith Collapse cells

    let resolveResizeArray cells = resolveResizeArrayWith Collapse cells
