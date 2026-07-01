module RedBat.Word.Shapes

open System
open System.Collections.Generic
open RedBat.Word.Document

// Vector geometry for OOXML preset/custom shapes. `presetPath` maps a preset
// name (a:prstGeom@prst) to a path in the shape's own pixel box (0,0)-(w,h).
// Unmapped presets fall back to a rectangle (callers may log via `isKnownPreset`).
//
// Arcs are emitted as cubic beziers (kappa approximation), so a single DrawShape
// path primitive covers every renderer (canvas / PDF / native).

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private prependArcSegs cx cy rx ry a0 a1 (tail: PathSeg list) =
    let total = a1 - a0
    let steps = max 1 (int (ceil (abs total / (Math.PI / 2.0))))
    let dq = total / float steps
    let mutable result = tail
    for i = steps - 1 downto 0 do
        let s0 = a0 + dq * float i
        let s1 = s0 + dq
        let k = (4.0 / 3.0) * tan ((s1 - s0) / 4.0)
        let p0x = cx + rx * cos s0
        let p0y = cy + ry * sin s0
        let p1x = cx + rx * cos s1
        let p1y = cy + ry * sin s1
        result <-
            PCubic(p0x - k * rx * sin s0, p0y + k * ry * cos s0,
                   p1x + k * rx * sin s1, p1y - k * ry * cos s1,
                   p1x, p1y) :: result
    result

/// Full ellipse path inside the box.
let private ellipsePath w h : PathSeg list =
    let cx, cy, rx, ry = w / 2.0, h / 2.0, w / 2.0, h / 2.0
    PMove(cx + rx, cy) :: prependArcSegs cx cy rx ry 0.0 (2.0 * Math.PI) [ PClose ]

/// Rounded rectangle with corner radius r.
let private roundRectPath w h r : PathSeg list =
    let r = max 0.0 (min r (min w h / 2.0))
    let tail = [ PClose ]
    let tail = prependArcSegs r r r r Math.PI (3.0 * Math.PI / 2.0) tail
    let tail = PLine(0.0, r) :: tail
    let tail = prependArcSegs r (h - r) r r (Math.PI / 2.0) Math.PI tail
    let tail = PLine(r, h) :: tail
    let tail = prependArcSegs (w - r) (h - r) r r 0.0 (Math.PI / 2.0) tail
    let tail = PLine(w, h - r) :: tail
    let tail = prependArcSegs (w - r) r r r (-Math.PI / 2.0) 0.0 tail
    PMove(r, 0.0) :: PLine(w - r, 0.0) :: tail

let private rectPath w h : PathSeg list =
    [ PMove(0.0, 0.0); PLine(w, 0.0); PLine(w, h); PLine(0.0, h); PClose ]

let private poly (pts: (float * float) list) : PathSeg list =
    match pts with
    | [] -> []
    | (x0, y0) :: rest ->
        let path = ResizeArray<PathSeg>(pts.Length + 1)
        path.Add(PMove(x0, y0))
        for x, y in rest do
            path.Add(PLine(x, y))
        path.Add PClose
        resizeArrayToList path

/// Regular n-gon inscribed in the box, first vertex at `startAngle`.
let private regularPolygon n w h startAngle : PathSeg list =
    let cx, cy, rx, ry = w / 2.0, h / 2.0, w / 2.0, h / 2.0
    if n <= 0 then []
    else
        let mutable path = [ PClose ]
        for i = n - 1 downto 0 do
            let a = startAngle + 2.0 * Math.PI * float i / float n
            let x, y = cx + rx * cos a, cy + ry * sin a
            path <- (if i = 0 then PMove(x, y) else PLine(x, y)) :: path
        path

/// n-point star, alternating outer/inner radius (innerRatio of the box radius).
let private starPath points w h innerRatio : PathSeg list =
    let cx, cy, rx, ry = w / 2.0, h / 2.0, w / 2.0, h / 2.0
    let start = -Math.PI / 2.0
    let pointCount = 2 * points
    if pointCount <= 0 then []
    else
        let mutable path = [ PClose ]
        for i = pointCount - 1 downto 0 do
            let a = start + Math.PI * float i / float points
            let r = if i % 2 = 0 then 1.0 else innerRatio
            let x, y = cx + rx * r * cos a, cy + ry * r * sin a
            path <- (if i = 0 then PMove(x, y) else PLine(x, y)) :: path
        path

// Block arrows ---------------------------------------------------------------
let private rightArrow w h : PathSeg list =
    let neck = h * 0.25           // tail half-thickness offset from edges
    let headW = w * 0.5
    poly [ 0.0, neck; w - headW, neck; w - headW, 0.0; w, h / 2.0
           w - headW, h; w - headW, h - neck; 0.0, h - neck ]
let private leftArrow w h : PathSeg list =
    let neck = h * 0.25
    let headW = w * 0.5
    poly [ w, neck; headW, neck; headW, 0.0; 0.0, h / 2.0
           headW, h; headW, h - neck; w, h - neck ]
let private upArrow w h : PathSeg list =
    let neck = w * 0.25
    let headH = h * 0.5
    poly [ neck, h; neck, headH; 0.0, headH; w / 2.0, 0.0
           w, headH; w - neck, headH; w - neck, h ]
let private downArrow w h : PathSeg list =
    let neck = w * 0.25
    let headH = h * 0.5
    poly [ neck, 0.0; neck, h - headH; 0.0, h - headH; w / 2.0, h
           w, h - headH; w - neck, h - headH; w - neck, 0.0 ]
let private leftRightArrow w h : PathSeg list =
    let neck = h * 0.25
    let headW = w * 0.25
    poly [ 0.0, h / 2.0; headW, 0.0; headW, neck; w - headW, neck; w - headW, 0.0
           w, h / 2.0; w - headW, h; w - headW, h - neck; headW, h - neck; headW, h ]
let private upDownArrow w h : PathSeg list =
    let neck = w * 0.25
    let headH = h * 0.25
    poly [ w / 2.0, 0.0; w, headH; w - neck, headH; w - neck, h - headH; w, h - headH
           w / 2.0, h; 0.0, h - headH; neck, h - headH; neck, headH; 0.0, headH ]
let private chevron w h : PathSeg list =
    let notch = w * 0.25
    poly [ 0.0, 0.0; w - notch, 0.0; w, h / 2.0; w - notch, h; 0.0, h; notch, h / 2.0 ]
let private homePlate w h : PathSeg list =
    let pt = w * 0.66
    poly [ 0.0, 0.0; pt, 0.0; w, h / 2.0; pt, h; 0.0, h ]
let private plusPath w h : PathSeg list =
    let ax = w * 0.25
    let ay = h * 0.25
    poly [ ax, 0.0; w - ax, 0.0; w - ax, ay; w, ay; w, h - ay; w - ax, h - ay
           w - ax, h; ax, h; ax, h - ay; 0.0, h - ay; 0.0, ay; ax, ay ]
let private parallelogram w h : PathSeg list =
    let off = w * 0.25
    poly [ off, 0.0; w, 0.0; w - off, h; 0.0, h ]
let private trapezoid w h : PathSeg list =
    let off = w * 0.25
    poly [ off, 0.0; w - off, 0.0; w, h; 0.0, h ]
let private bevelPath w h : PathSeg list =
    // outer rectangle (the inner highlight is a renderer detail we skip)
    rectPath w h

type private PresetPathCacheEntry = {
    Preset: string
    Width: float
    Height: float
    Path: PathSeg list
}

[<Struct>]
type private PresetPathCacheKey = {
    Preset: string
    Width: float
    Height: float
}

let private presetPathCacheKeyHash (key: PresetPathCacheKey) =
    let mutable hashCode =
        if isNull key.Preset then 0 else key.Preset.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ key.Width.GetHashCode()
    (hashCode * 397) ^^^ key.Height.GetHashCode()

let private presetPathCacheKeyEquals (left: PresetPathCacheKey) (right: PresetPathCacheKey) =
    left.Preset = right.Preset
    && left.Width = right.Width
    && left.Height = right.Height

let private presetPathCacheKeyComparer =
    { new IEqualityComparer<PresetPathCacheKey> with
        member _.Equals(left, right) = presetPathCacheKeyEquals left right
        member _.GetHashCode(value) = presetPathCacheKeyHash value }

let mutable private presetPathCache: PresetPathCacheEntry list = []
let private presetPathCacheByKey =
    Dictionary<PresetPathCacheKey, PresetPathCacheEntry>(presetPathCacheKeyComparer)
let private presetPathFastSlots = Array.zeroCreate<PresetPathCacheEntry> 256
let private presetPathCacheMaxEntries = 256
let private presetPathCacheMaxTotalPresetChars = 1000000

let private presetPathCacheMatches preset w h (cached: PresetPathCacheEntry) =
    cached.Preset = preset && cached.Width = w && cached.Height = h

let private presetPathCacheKey preset w h : PresetPathCacheKey = {
    Preset = preset
    Width = w
    Height = h
}

let private presetPathCacheKeyMatches (key: PresetPathCacheKey) (cached: PresetPathCacheEntry) =
    cached.Preset = key.Preset && cached.Width = key.Width && cached.Height = key.Height

let private presetPathCacheSlot (key: PresetPathCacheKey) slotCount =
    presetPathCacheKeyHash key &&& (slotCount - 1)

let private setPresetPathCache (entries: PresetPathCacheEntry list) =
    presetPathCache <- entries
    presetPathCacheByKey.Clear()
    for i = 0 to presetPathFastSlots.Length - 1 do
        presetPathFastSlots.[i] <- Unchecked.defaultof<PresetPathCacheEntry>
    for entry in entries do
        let key = presetPathCacheKey entry.Preset entry.Width entry.Height
        presetPathCacheByKey.[key] <- entry
        presetPathFastSlots.[presetPathCacheSlot key presetPathFastSlots.Length] <- entry

let private tryGetPresetPathCache (key: PresetPathCacheKey) =
    let slot = presetPathCacheSlot key presetPathFastSlots.Length
    let fast = presetPathFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && presetPathCacheKeyMatches key fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<PresetPathCacheEntry>
        if presetPathCacheByKey.TryGetValue(key, &cached) then
            presetPathFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private rememberPresetPath (entry: PresetPathCacheEntry) =
    if entry.Preset.Length <= presetPathCacheMaxTotalPresetChars then
        let trim (entries: PresetPathCacheEntry list) =
            let rec loop (remaining: PresetPathCacheEntry list) (kept: PresetPathCacheEntry list) count totalPresetChars =
                match remaining with
                | [] -> List.rev kept
                | candidate :: rest ->
                    let presetChars = max 1 candidate.Preset.Length
                    let nextTotalPresetChars = totalPresetChars + presetChars
                    if count < presetPathCacheMaxEntries
                       && nextTotalPresetChars <= presetPathCacheMaxTotalPresetChars then
                        loop rest (candidate :: kept) (count + 1) nextTotalPresetChars
                    else
                        loop rest kept count totalPresetChars
            loop entries [] 0 0
        entry
        :: (presetPathCache
            |> List.filter (presetPathCacheMatches entry.Preset entry.Width entry.Height >> not))
        |> trim
        |> setPresetPathCache

/// Presets we render as their true geometry (everything else uses the rect
/// fallback). Used only for diagnostics/coverage reporting.
let knownPresets =
    set [ "rect"; "rectangle"; "flowChartProcess"; "flowChartPredefinedProcess"; "flowChartInternalStorage"
          "roundRect"; "flowChartAlternateProcess"; "round1Rect"; "round2SameRect"; "round2DiagRect"; "snipRoundRect"
          "ellipse"; "circle"; "flowChartConnector"; "flowChartSummingJunction"; "flowChartOr"
          "triangle"; "isoscelesTriangle"; "flowChartExtract"; "flowChartMerge"; "rtTriangle"
          "diamond"; "flowChartDecision"; "flowChartSort"; "parallelogram"; "flowChartInputOutput"
          "trapezoid"; "flowChartManualOperation"; "pentagon"; "hexagon"; "flowChartPreparation"
          "heptagon"; "octagon"; "decagon"; "dodecagon"
          "star4"; "star5"; "star6"; "star7"; "star8"; "star10"; "star12"; "star16"; "star24"; "star32"
          "rightArrow"; "flowChartArrow"; "leftArrow"; "upArrow"; "downArrow"; "leftRightArrow"; "upDownArrow"
          "chevron"; "homePlate"; "plus"; "mathPlus"; "bevel"; "frame"; "flowChartTerminator"; "round"; "pie" ]

let isKnownPreset (preset: string) : bool = knownPresets.Contains preset

/// Path for a named OOXML preset in the box (w,h). Falls back to a rectangle.
let private presetPathUncached (preset: string) (w: float) (h: float) : PathSeg list =
    match preset with
    | "rect" | "rectangle" | "flowChartProcess" | "flowChartPredefinedProcess"
    | "flowChartInternalStorage" -> rectPath w h
    | "roundRect" | "flowChartAlternateProcess" -> roundRectPath w h (min w h * 0.16667)
    | "round1Rect" | "round2SameRect" | "round2DiagRect" | "snipRoundRect" ->
        roundRectPath w h (min w h * 0.16667)
    | "ellipse" | "circle" | "flowChartConnector" | "flowChartSummingJunction"
    | "flowChartOr" -> ellipsePath w h
    | "triangle" | "isoscelesTriangle" | "flowChartExtract" | "flowChartMerge" ->
        poly [ w / 2.0, 0.0; w, h; 0.0, h ]
    | "rtTriangle" -> poly [ 0.0, 0.0; 0.0, h; w, h ]
    | "diamond" | "flowChartDecision" | "flowChartSort" ->
        poly [ w / 2.0, 0.0; w, h / 2.0; w / 2.0, h; 0.0, h / 2.0 ]
    | "parallelogram" | "flowChartInputOutput" -> parallelogram w h
    | "trapezoid" | "flowChartManualOperation" -> trapezoid w h
    | "pentagon" -> regularPolygon 5 w h (-Math.PI / 2.0)
    | "hexagon" | "flowChartPreparation" -> regularPolygon 6 w h 0.0
    | "heptagon" -> regularPolygon 7 w h (-Math.PI / 2.0)
    | "octagon" -> regularPolygon 8 w h (Math.PI / 8.0)
    | "decagon" -> regularPolygon 10 w h 0.0
    | "dodecagon" -> regularPolygon 12 w h 0.0
    | "star4" -> starPath 4 w h 0.4
    | "star5" -> starPath 5 w h 0.382
    | "star6" -> starPath 6 w h 0.577
    | "star7" -> starPath 7 w h 0.5
    | "star8" -> starPath 8 w h 0.55
    | "star10" -> starPath 10 w h 0.6
    | "star12" -> starPath 12 w h 0.6
    | "star16" -> starPath 16 w h 0.65
    | "star24" -> starPath 24 w h 0.7
    | "star32" -> starPath 32 w h 0.75
    | "rightArrow" | "flowChartArrow" -> rightArrow w h
    | "leftArrow" -> leftArrow w h
    | "upArrow" -> upArrow w h
    | "downArrow" -> downArrow w h
    | "leftRightArrow" -> leftRightArrow w h
    | "upDownArrow" -> upDownArrow w h
    | "chevron" -> chevron w h
    | "homePlate" -> homePlate w h
    | "plus" | "mathPlus" -> plusPath w h
    | "bevel" | "frame" -> bevelPath w h
    | "flowChartTerminator" -> roundRectPath w h (h / 2.0)            // stadium
    | "flowChartDocument" -> rectPath w h
    | "round" | "pie" -> ellipsePath w h
    | _ -> rectPath w h

let presetPath (preset: string) (w: float) (h: float) : PathSeg list =
    let key = presetPathCacheKey preset w h
    match tryGetPresetPathCache key with
    | ValueSome cached -> cached.Path
    | ValueNone ->
        let path = presetPathUncached preset w h
        rememberPresetPath { Preset = preset; Width = w; Height = h; Path = path }
        path

/// Scale a custom path from its own coordinate space into the box (w,h).
let private scaleCustom (segs: PathSeg list) coordW coordH w h : PathSeg list =
    let sx = if coordW <= 0.0 then 1.0 else w / coordW
    let sy = if coordH <= 0.0 then 1.0 else h / coordH
    segs
    |> List.map (function
        | PMove(x, y) -> PMove(x * sx, y * sy)
        | PLine(x, y) -> PLine(x * sx, y * sy)
        | PQuad(cx, cy, x, y) -> PQuad(cx * sx, cy * sy, x * sx, y * sy)
        | PCubic(c1x, c1y, c2x, c2y, x, y) ->
            PCubic(c1x * sx, c1y * sy, c2x * sx, c2y * sy, x * sx, y * sy)
        | PClose -> PClose)

/// Resolve any geometry to an absolute path in the box (w,h), pre-translation.
let geometryPath (geom: ShapeGeometry) (w: float) (h: float) : PathSeg list =
    match geom with
    | GeoRect -> rectPath w h
    | GeoPreset preset -> presetPath preset w h
    | GeoCustom(segs, coordW, coordH) -> scaleCustom segs coordW coordH w h

/// Rotate a path by `deg` degrees around (cx,cy). The shape outline rotates;
/// box-flowed text is left unrotated (a known approximation).
let rotatePath cx cy deg (segs: PathSeg list) : PathSeg list =
    let r = deg * Math.PI / 180.0
    let ca, sa = cos r, sin r
    let rot x y = (cx + (x - cx) * ca - (y - cy) * sa, cy + (x - cx) * sa + (y - cy) * ca)
    segs
    |> List.map (function
        | PMove(x, y) -> let (nx, ny) = rot x y in PMove(nx, ny)
        | PLine(x, y) -> let (nx, ny) = rot x y in PLine(nx, ny)
        | PQuad(qx, qy, x, y) ->
            let (c1x, c1y) = rot qx qy
            let (nx, ny) = rot x y
            PQuad(c1x, c1y, nx, ny)
        | PCubic(a, b, c, d, x, y) ->
            let (c1x, c1y) = rot a b
            let (c2x, c2y) = rot c d
            let (nx, ny) = rot x y
            PCubic(c1x, c1y, c2x, c2y, nx, ny)
        | PClose -> PClose)

/// Translate a path by (dx,dy) — used to place the box on the page.
let translatePath dx dy (segs: PathSeg list) : PathSeg list =
    segs
    |> List.map (function
        | PMove(x, y) -> PMove(x + dx, y + dy)
        | PLine(x, y) -> PLine(x + dx, y + dy)
        | PQuad(cx, cy, x, y) -> PQuad(cx + dx, cy + dy, x + dx, y + dy)
        | PCubic(c1x, c1y, c2x, c2y, x, y) -> PCubic(c1x + dx, c1y + dy, c2x + dx, c2y + dy, x + dx, y + dy)
        | PClose -> PClose)
