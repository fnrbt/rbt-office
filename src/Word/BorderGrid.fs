namespace RedBat

module BorderGrid =
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

    let normalizeStyle (style: string option) =
        match style |> Option.map (fun value -> value.Trim()) with
        | Some value when value <> "" -> Some value
        | _ -> Some "single"

    let styleKey (style: string option) =
        normalizeStyle style
        |> Option.map (fun value -> value.ToLowerInvariant())
        |> Option.defaultValue "single"

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

    let private candidates cells =
        cells
        |> List.collect (fun cell ->
            [ cell.Edges.Top |> Option.bind (toCandidate true cell.Y cell.X (cell.X + cell.Width))
              cell.Edges.Right |> Option.bind (toCandidate false (cell.X + cell.Width) cell.Y (cell.Y + cell.Height))
              cell.Edges.Bottom |> Option.bind (toCandidate true (cell.Y + cell.Height) cell.X (cell.X + cell.Width))
              cell.Edges.Left |> Option.bind (toCandidate false cell.X cell.Y (cell.Y + cell.Height)) ])
        |> List.choose id

    let private resolveAxis horizontal candidates =
        candidates
        |> List.filter (fun candidate -> candidate.Horizontal = horizontal)
        |> List.groupBy _.Fixed
        |> List.collect (fun (fixedCoord, axisCandidates) ->
            let points =
                axisCandidates
                |> List.collect (fun candidate -> [candidate.Start; candidate.Finish])
                |> List.distinct
                |> List.sort

            points
            |> List.pairwise
            |> List.choose (fun (startCoord, finishCoord) ->
                let covering =
                    axisCandidates
                    |> List.filter (fun candidate -> candidate.Start <= startCoord && candidate.Finish >= finishCoord)
                match covering with
                | [] -> None
                | first :: rest ->
                    let winner = rest |> List.fold stronger first
                    if not (isVisible winner.Stroke) then None
                    elif horizontal then Some { X1 = startCoord; Y1 = fixedCoord; X2 = finishCoord; Y2 = fixedCoord; Stroke = winner.Stroke }
                    else Some { X1 = fixedCoord; Y1 = startCoord; X2 = fixedCoord; Y2 = finishCoord; Stroke = winner.Stroke }))

    let resolveWith collapse (cells: Cell list) =
        let candidates = candidates cells
        match collapse with
        | Separate ->
            candidates
            |> List.choose (fun candidate ->
                if not (isVisible candidate.Stroke) then None
                elif candidate.Horizontal then Some { X1 = candidate.Start; Y1 = candidate.Fixed; X2 = candidate.Finish; Y2 = candidate.Fixed; Stroke = candidate.Stroke }
                else Some { X1 = candidate.Fixed; Y1 = candidate.Start; X2 = candidate.Fixed; Y2 = candidate.Finish; Stroke = candidate.Stroke })
        | Collapse ->
            (resolveAxis true candidates) @ (resolveAxis false candidates)

    let resolve cells = resolveWith Collapse cells
