module RedBat.Spreadsheet.Layout

open RedBat.Spreadsheet.Model

let private resizeArrayToList (items: System.Collections.Generic.List<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private clearArray (items: 'T array) =
    for i = 0 to items.Length - 1 do
        items.[i] <- Unchecked.defaultof<'T>

[<Struct>]
type Rect = {
    X: float
    Y: float
    Width: float
    Height: float
}

/// Measure a piece of text in pixels at a given font size/face. Optional;
/// when None, the cell painter approximates width as charCount × mdwPx.
/// Set by the native runner (real Renderer.Text.measureWidth) and by the
/// PDF emitter; left None by the Fable web canvas which doesn't have a
/// portable measurement API at this layer.
type TextMeasurer = string -> float -> string option -> bool -> bool -> float

type Viewport = {
    ScrollX: float
    ScrollY: float
    Width: float
    Height: float
    HeaderHeight: float
    RowHeaderWidth: float
    /// Optional real-font measurer used by buildCellPaintItems to decide
    /// wrap-text break points. When None, falls back to mdwPx × charCount.
    Measurer: TextMeasurer option
}

type private ViewportGeometryKey = {
    ScrollX: float
    ScrollY: float
    Width: float
    Height: float
    HeaderHeight: float
    RowHeaderWidth: float
}

type private ViewportGeometryFastEntry<'T> = {
    Key: ViewportGeometryKey
    Value: 'T
}

let private viewportGeometryKey (viewport: Viewport) = {
    ScrollX = viewport.ScrollX
    ScrollY = viewport.ScrollY
    Width = viewport.Width
    Height = viewport.Height
    HeaderHeight = viewport.HeaderHeight
    RowHeaderWidth = viewport.RowHeaderWidth
}

let private viewportGeometrySlot
    (slotCount: int)
    scrollX
    scrollY
    width
    height
    headerHeight
    rowHeaderWidth =
    let mutable hashCode = scrollX.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ scrollY.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ width.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ height.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ headerHeight.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ rowHeaderWidth.GetHashCode()
    hashCode &&& (slotCount - 1)

let private viewportGeometryKeyMatchesViewport key (viewport: Viewport) =
    key.ScrollX = viewport.ScrollX
    && key.ScrollY = viewport.ScrollY
    && key.Width = viewport.Width
    && key.Height = viewport.Height
    && key.HeaderHeight = viewport.HeaderHeight
    && key.RowHeaderWidth = viewport.RowHeaderWidth

let private tryGetViewportGeometryFastSlot
    (slots: ViewportGeometryFastEntry<'T> array)
    (viewport: Viewport) =
    let slot =
        viewportGeometrySlot
            slots.Length
            viewport.ScrollX
            viewport.ScrollY
            viewport.Width
            viewport.Height
            viewport.HeaderHeight
            viewport.RowHeaderWidth
    let fast = slots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && viewportGeometryKeyMatchesViewport fast.Key viewport then
        ValueSome fast.Value
    else
        ValueNone

let private setViewportGeometryFastSlot
    (slots: ViewportGeometryFastEntry<'T> array)
    key
    value =
    let slot =
        viewportGeometrySlot
            slots.Length
            key.ScrollX
            key.ScrollY
            key.Width
            key.Height
            key.HeaderHeight
            key.RowHeaderWidth
    slots.[slot] <- { Key = key; Value = value }

type VisibleRange = {
    StartRow: int
    EndRowExclusive: int
    StartCol: int
    EndColExclusive: int
}

[<Struct>]
type CellGeometry = {
    Address: CellAddress
    Rect: Rect
    IsMergedPrimary: bool
    Merge: MergedCell option
}

type ViewportZone =
    | BodyZone
    | ColumnHeaderZone
    | RowHeaderZone
    | CornerZone
    /// Cells in the top frozen-rows band (above the freeze split, to the
    /// right of any frozen columns). Scrolls horizontally with the body
    /// but stays pinned vertically.
    | FrozenRowsZone
    /// Cells in the left frozen-cols band. Scrolls vertically with the
    /// body but stays pinned horizontally.
    | FrozenColsZone
    /// Top-left corner: the intersection of frozen rows and frozen cols.
    /// Pinned in both axes (no scroll applied).
    | FrozenCornerZone

type HeaderGeometry =
    | ColumnHeader of col: int * rect: Rect
    | RowHeader of row: int * rect: Rect
    | CornerHeader of rect: Rect

type HitTarget =
    | Cell of CellAddress
    | HitColumnHeader of int
    | HitRowHeader of int
    | Corner
    | Empty

let defaultViewport width height = {
    ScrollX = 0.0
    ScrollY = 0.0
    Width = width
    Height = height
    HeaderHeight = 30.0
    RowHeaderWidth = 50.0
    Measurer = None
}

let private defaultRowHeight = 25.0
let private defaultColumnWidth = 100.0

let private lowerBound (values: int array) target =
    let mutable lo = 0
    let mutable hi = values.Length
    while lo < hi do
        let mid = lo + (hi - lo) / 2
        if values.[mid] < target then
            lo <- mid + 1
        else
            hi <- mid
    lo

let private containsSorted (values: int array) target =
    let index = lowerBound values target
    index < values.Length && values.[index] = target

type private PrefixSizeCache = {
    Keys: int array
    PrefixDeltas: float array
    LimitExclusive: int
}

type private PrefixSizeCacheEntry = {
    Source: Map<int, float>
    Cache: PrefixSizeCache
}

type private PrefixSizeCacheSlot = {
    mutable Entries: PrefixSizeCacheEntry list
    mutable LastEntrySet: bool
    mutable LastEntry: PrefixSizeCacheEntry
}

let private rowPrefixSizeCache = {
    Entries = []
    LastEntrySet = false
    LastEntry = Unchecked.defaultof<PrefixSizeCacheEntry>
}
let private columnPrefixSizeCache = {
    Entries = []
    LastEntrySet = false
    LastEntry = Unchecked.defaultof<PrefixSizeCacheEntry>
}
let private prefixSizeCacheMaxSources = 256
let private prefixSizeCacheMaxTotalKeys = 1000000
let private prefixSizeTinyUncachedThreshold = 32
let private prefixSizeDirectIndexThreshold = 4096
let private prefixSizeShallowCacheLimit = prefixSizeDirectIndexThreshold + 1
let private prefixSizeFullCacheLimit = System.Int32.MaxValue

type private HiddenPrefixSizeCacheEntry = {
    Source: Map<int, float>
    HiddenRows: int array
    CustomRows: int array
    PrefixDeltas: float array
    LimitExclusive: int
}

type private HiddenPrefixSizeCacheSlot = {
    mutable Entries: HiddenPrefixSizeCacheEntry list
    mutable LastEntrySet: bool
    mutable LastEntry: HiddenPrefixSizeCacheEntry
}

let private hiddenPrefixSizeCache = {
    Entries = []
    LastEntrySet = false
    LastEntry = Unchecked.defaultof<HiddenPrefixSizeCacheEntry>
}
let private hiddenPrefixSizeCacheMaxSources = 256
let private hiddenPrefixSizeCacheMaxTotalKeys = 1000000

let private prefixSizeCacheEntryMatches source (entry: PrefixSizeCacheEntry) =
    System.Object.ReferenceEquals(entry.Source, source)

let private prefixSizeCacheEntryKeyCount (entry: PrefixSizeCacheEntry) =
    max (Map.count entry.Source) entry.Cache.Keys.Length

let private prefixSizeCacheEntryRetainable (entry: PrefixSizeCacheEntry) =
    prefixSizeCacheEntryKeyCount entry <= prefixSizeCacheMaxTotalKeys

let private setPrefixSizeCacheEntries (slot: PrefixSizeCacheSlot) (entries: PrefixSizeCacheEntry list) =
    slot.Entries <- entries
    match entries with
    | entry :: _ ->
        slot.LastEntrySet <- true
        slot.LastEntry <- entry
    | [] ->
        slot.LastEntrySet <- false
        slot.LastEntry <- Unchecked.defaultof<PrefixSizeCacheEntry>

let private tryGetPrefixSizeCacheEntry (slot: PrefixSizeCacheSlot) overrides =
    if slot.LastEntrySet && prefixSizeCacheEntryMatches overrides slot.LastEntry then
        ValueSome slot.LastEntry
    else
        match slot.Entries |> List.tryFind (prefixSizeCacheEntryMatches overrides) with
        | Some entry ->
            slot.LastEntrySet <- true
            slot.LastEntry <- entry
            ValueSome entry
        | None -> ValueNone

let private prefixSizeCacheEntriesWithoutSource (slot: PrefixSizeCacheSlot) source =
    let rec loop acc remaining =
        match remaining with
        | [] -> slot.Entries
        | candidate :: tail when prefixSizeCacheEntryMatches source candidate ->
            List.rev acc @ (tail |> List.filter (prefixSizeCacheEntryMatches source >> not))
        | candidate :: tail ->
            loop (candidate :: acc) tail
    loop [] slot.Entries

let private rememberPrefixSizeCacheEntry (slot: PrefixSizeCacheSlot) (entry: PrefixSizeCacheEntry) =
    if not (prefixSizeCacheEntryRetainable entry) then
        let entries = prefixSizeCacheEntriesWithoutSource slot entry.Source
        if not (obj.ReferenceEquals(entries, slot.Entries)) then
            setPrefixSizeCacheEntries slot entries
    else
        let mutable keptCount = 1
        let mutable keptKeys = prefixSizeCacheEntryKeyCount entry
        let retained =
            prefixSizeCacheEntriesWithoutSource slot entry.Source
            |> List.choose (fun candidate ->
                let candidateKeys = prefixSizeCacheEntryKeyCount candidate
                if prefixSizeCacheEntryRetainable candidate
                   && keptCount < prefixSizeCacheMaxSources
                   && keptKeys + candidateKeys <= prefixSizeCacheMaxTotalKeys then
                    keptCount <- keptCount + 1
                    keptKeys <- keptKeys + candidateKeys
                    Some candidate
                else
                    None)
        entry :: retained
        |> setPrefixSizeCacheEntries slot

let private hiddenPrefixSizeCacheEntryMatches source hiddenRows (entry: HiddenPrefixSizeCacheEntry) =
    System.Object.ReferenceEquals(entry.Source, source)
    && System.Object.ReferenceEquals(entry.HiddenRows, hiddenRows)

let private hiddenPrefixSizeCacheEntryKeyCount (entry: HiddenPrefixSizeCacheEntry) =
    Map.count entry.Source + entry.CustomRows.Length + entry.HiddenRows.Length

let private hiddenPrefixSizeCacheEntryRetainable (entry: HiddenPrefixSizeCacheEntry) =
    hiddenPrefixSizeCacheEntryKeyCount entry <= hiddenPrefixSizeCacheMaxTotalKeys

let private setHiddenPrefixSizeCacheEntries (entries: HiddenPrefixSizeCacheEntry list) =
    hiddenPrefixSizeCache.Entries <- entries
    match entries with
    | entry :: _ ->
        hiddenPrefixSizeCache.LastEntrySet <- true
        hiddenPrefixSizeCache.LastEntry <- entry
    | [] ->
        hiddenPrefixSizeCache.LastEntrySet <- false
        hiddenPrefixSizeCache.LastEntry <- Unchecked.defaultof<HiddenPrefixSizeCacheEntry>

let private tryGetHiddenPrefixSizeCacheEntry overrides hiddenRows =
    if hiddenPrefixSizeCache.LastEntrySet
       && hiddenPrefixSizeCacheEntryMatches overrides hiddenRows hiddenPrefixSizeCache.LastEntry then
        ValueSome hiddenPrefixSizeCache.LastEntry
    else
        match hiddenPrefixSizeCache.Entries |> List.tryFind (hiddenPrefixSizeCacheEntryMatches overrides hiddenRows) with
        | Some entry ->
            hiddenPrefixSizeCache.LastEntrySet <- true
            hiddenPrefixSizeCache.LastEntry <- entry
            ValueSome entry
        | None -> ValueNone

let private hiddenPrefixSizeCacheEntriesWithoutSource source hiddenRows =
    let rec loop acc remaining =
        match remaining with
        | [] -> hiddenPrefixSizeCache.Entries
        | candidate :: tail when hiddenPrefixSizeCacheEntryMatches source hiddenRows candidate ->
            List.rev acc @ (tail |> List.filter (hiddenPrefixSizeCacheEntryMatches source hiddenRows >> not))
        | candidate :: tail ->
            loop (candidate :: acc) tail
    loop [] hiddenPrefixSizeCache.Entries

let private rememberHiddenPrefixSizeCacheEntry (entry: HiddenPrefixSizeCacheEntry) =
    if not (hiddenPrefixSizeCacheEntryRetainable entry) then
        let entries = hiddenPrefixSizeCacheEntriesWithoutSource entry.Source entry.HiddenRows
        if not (obj.ReferenceEquals(entries, hiddenPrefixSizeCache.Entries)) then
            setHiddenPrefixSizeCacheEntries entries
    else
        let mutable keptCount = 1
        let mutable keptKeys = hiddenPrefixSizeCacheEntryKeyCount entry
        let retained =
            hiddenPrefixSizeCacheEntriesWithoutSource entry.Source entry.HiddenRows
            |> List.choose (fun candidate ->
                let candidateKeys = hiddenPrefixSizeCacheEntryKeyCount candidate
                if hiddenPrefixSizeCacheEntryRetainable candidate
                   && keptCount < hiddenPrefixSizeCacheMaxSources
                   && keptKeys + candidateKeys <= hiddenPrefixSizeCacheMaxTotalKeys then
                    keptCount <- keptCount + 1
                    keptKeys <- keptKeys + candidateKeys
                    Some candidate
                else
                    None)
        entry :: retained
        |> setHiddenPrefixSizeCacheEntries

let private buildPrefixSizeCache defaultSize limitExclusive (overrides: Map<int, float>) =
    let capacity =
        if limitExclusive = prefixSizeFullCacheLimit then overrides.Count
        else min overrides.Count (max 0 limitExclusive)
    let keys = System.Collections.Generic.List<int>(capacity)
    let prefixDeltas = System.Collections.Generic.List<float>(capacity + 1)
    let mutable total = 0.0
    prefixDeltas.Add total
    use entries = (overrides :> seq<_>).GetEnumerator()
    let mutable keepGoing = true
    while keepGoing && entries.MoveNext() do
        let (KeyValue(i, size)) = entries.Current
        if i >= limitExclusive then
            keepGoing <- false
        elif i >= 0 then
            keys.Add i
            total <- total + (size - defaultSize)
            prefixDeltas.Add total
    {
        Keys = keys.ToArray()
        PrefixDeltas = prefixDeltas.ToArray()
        LimitExclusive = limitExclusive
    }

let private prefixSizeUncached defaultSize (overrides: Map<int, float>) index =
    let clipped = max 0 index
    let mutable total = 0.0
    use entries = (overrides :> seq<_>).GetEnumerator()
    let mutable keepGoing = true
    while keepGoing && entries.MoveNext() do
        let (KeyValue(i, size)) = entries.Current
        if i >= clipped then
            keepGoing <- false
        elif i >= 0 then
            total <- total + (size - defaultSize)
    defaultSize * float clipped + total

let private prefixSize defaultSize (overrides: Map<int, float>) (slot: PrefixSizeCacheSlot) (index: int) =
    let clipped = max 0 index
    if Map.isEmpty overrides then
        defaultSize * float clipped
    else
        match tryGetPrefixSizeCacheEntry slot overrides with
        | ValueSome entry when clipped < entry.Cache.LimitExclusive ->
            defaultSize * float clipped + entry.Cache.PrefixDeltas.[lowerBound entry.Cache.Keys clipped]
        | _ ->
            if clipped <= prefixSizeTinyUncachedThreshold then
                prefixSizeUncached defaultSize overrides index
            elif Map.count overrides > prefixSizeCacheMaxTotalKeys then
                prefixSizeUncached defaultSize overrides index
            else
                let limitExclusive =
                    if clipped <= prefixSizeDirectIndexThreshold then prefixSizeShallowCacheLimit
                    else prefixSizeFullCacheLimit
                let cache = buildPrefixSizeCache defaultSize limitExclusive overrides
                rememberPrefixSizeCacheEntry slot { Source = overrides; Cache = cache }
                defaultSize * float clipped + cache.PrefixDeltas.[lowerBound cache.Keys clipped]

let private buildHiddenPrefixSizeCache defaultSize (overrides: Map<int, float>) (hiddenRows: int array) limitExclusive =
    let customRows = System.Collections.Generic.List<int>()
    let prefixDeltas = System.Collections.Generic.List<float>()
    let mutable total = 0.0
    prefixDeltas.Add total
    use entries = (overrides :> seq<_>).GetEnumerator()
    let mutable keepGoing = true
    while keepGoing && entries.MoveNext() do
        let (KeyValue(row, height)) = entries.Current
        if row >= limitExclusive then
            keepGoing <- false
        elif row >= 0 && containsSorted hiddenRows row then
            customRows.Add row
            total <- total + (height - defaultSize)
            prefixDeltas.Add total
    {
        Source = overrides
        HiddenRows = hiddenRows
        CustomRows = customRows.ToArray()
        PrefixDeltas = prefixDeltas.ToArray()
        LimitExclusive = limitExclusive
    }

let private hiddenPrefixSizeUncached defaultSize (overrides: Map<int, float>) (hiddenRows: int array) clipped hiddenCount =
    let mutable total = 0.0
    use entries = (overrides :> seq<_>).GetEnumerator()
    let mutable keepGoing = true
    while keepGoing && entries.MoveNext() do
        let (KeyValue(row, height)) = entries.Current
        if row >= clipped then
            keepGoing <- false
        elif row >= 0 && containsSorted hiddenRows row then
            total <- total + (height - defaultSize)
    defaultSize * float hiddenCount + total

let private hiddenPrefixSize defaultSize (overrides: Map<int, float>) (hiddenRows: int array) (index: int) =
    if hiddenRows.Length = 0 || index <= 0 then 0.0
    else
        let clipped = max 0 index
        let hiddenCount = lowerBound hiddenRows clipped
        if hiddenCount = 0 then 0.0
        elif Map.isEmpty overrides then
            defaultSize * float hiddenCount
        else
            match tryGetHiddenPrefixSizeCacheEntry overrides hiddenRows with
            | ValueSome cache when clipped < cache.LimitExclusive ->
                defaultSize * float hiddenCount
                + cache.PrefixDeltas.[lowerBound cache.CustomRows clipped]
            | _ ->
                if clipped <= prefixSizeTinyUncachedThreshold then
                    hiddenPrefixSizeUncached defaultSize overrides hiddenRows clipped hiddenCount
                elif Map.count overrides + hiddenRows.Length > hiddenPrefixSizeCacheMaxTotalKeys then
                    hiddenPrefixSizeUncached defaultSize overrides hiddenRows clipped hiddenCount
                else
                    let limitExclusive =
                        if clipped <= prefixSizeDirectIndexThreshold then prefixSizeShallowCacheLimit
                        else prefixSizeFullCacheLimit
                    let cache = buildHiddenPrefixSizeCache defaultSize overrides hiddenRows limitExclusive
                    rememberHiddenPrefixSizeCacheEntry cache
                    defaultSize * float hiddenCount
                    + cache.PrefixDeltas.[lowerBound cache.CustomRows clipped]

[<Struct>]
type private RowTopCoordinateCacheEntry = {
    Set: bool
    RowHeights: Map<int, float>
    HasFilters: bool
    FilterHiddenRows: int array
    Index: int
    Value: float
}

[<Struct>]
type private ColumnLeftCoordinateCacheEntry = {
    Set: bool
    ColumnWidths: Map<int, float>
    Index: int
    Value: float
}

let private coordinateCacheSlotCount = 2048
let private rowTopFastSlots = Array.zeroCreate<RowTopCoordinateCacheEntry> coordinateCacheSlotCount
let private columnLeftFastSlots = Array.zeroCreate<ColumnLeftCoordinateCacheEntry> coordinateCacheSlotCount

let private rowTopHiddenRowsHash hasFilters (hiddenRows: int array) =
    if hasFilters && not (isNull hiddenRows) then
        System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(hiddenRows)
    else
        0

let private rowTopFastSlot rowHeights hasFilters hiddenRows index slotCount =
    let mutable hashCode = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(rowHeights)
    hashCode <- (hashCode * 397) ^^^ (if hasFilters then 1 else 0)
    hashCode <- (hashCode * 397) ^^^ rowTopHiddenRowsHash hasFilters hiddenRows
    hashCode <- (hashCode * 397) ^^^ index
    hashCode <- hashCode ^^^ (hashCode >>> 16)
    hashCode &&& (slotCount - 1)

let private tryGetRowTopFastSlot (slots: RowTopCoordinateCacheEntry array) rowHeights hasFilters hiddenRows index =
    let slot = rowTopFastSlot rowHeights hasFilters hiddenRows index slots.Length
    let entry = slots.[slot]
    if entry.Set
       && entry.Index = index
       && entry.HasFilters = hasFilters
       && obj.ReferenceEquals(entry.RowHeights, rowHeights)
       && (not hasFilters || obj.ReferenceEquals(entry.FilterHiddenRows, hiddenRows)) then
        ValueSome entry.Value
    else
        ValueNone

let private rowTopCoordinateSourceRetainable rowHeights hasFilters (hiddenRows: int array) =
    if Map.count rowHeights > prefixSizeCacheMaxTotalKeys then
        false
    elif hasFilters then
        Map.count rowHeights + hiddenRows.Length <= hiddenPrefixSizeCacheMaxTotalKeys
    else
        true

let private setRowTopFastSlot (slots: RowTopCoordinateCacheEntry array) rowHeights hasFilters hiddenRows index value =
    if rowTopCoordinateSourceRetainable rowHeights hasFilters hiddenRows then
        let slot = rowTopFastSlot rowHeights hasFilters hiddenRows index slots.Length
        slots.[slot] <- {
            Set = true
            RowHeights = rowHeights
            HasFilters = hasFilters
            FilterHiddenRows = if hasFilters then hiddenRows else Unchecked.defaultof<int array>
            Index = index
            Value = value
        }

let private columnLeftFastSlot columnWidths index slotCount =
    let mutable hashCode = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(columnWidths)
    hashCode <- (hashCode * 397) ^^^ index
    hashCode <- hashCode ^^^ (hashCode >>> 16)
    hashCode &&& (slotCount - 1)

let private tryGetColumnLeftFastSlot (slots: ColumnLeftCoordinateCacheEntry array) columnWidths index =
    let slot = columnLeftFastSlot columnWidths index slots.Length
    let entry = slots.[slot]
    if entry.Set
       && entry.Index = index
       && obj.ReferenceEquals(entry.ColumnWidths, columnWidths) then
        ValueSome entry.Value
    else
        ValueNone

let private setColumnLeftFastSlot (slots: ColumnLeftCoordinateCacheEntry array) columnWidths index value =
    if Map.count columnWidths <= prefixSizeCacheMaxTotalKeys then
        let slot = columnLeftFastSlot columnWidths index slots.Length
        slots.[slot] <- {
            Set = true
            ColumnWidths = columnWidths
            Index = index
            Value = value
        }

let rowTop (model: SpreadsheetModel) (row: int): float =
    let hasFilters = not (Map.isEmpty model.Filters)
    let hiddenRows = if hasFilters then model.FilterHiddenRows else Unchecked.defaultof<int array>
    match tryGetRowTopFastSlot rowTopFastSlots model.RowHeights hasFilters hiddenRows row with
    | ValueSome value -> value
    | ValueNone ->
        let total = prefixSize defaultRowHeight model.RowHeights rowPrefixSizeCache row
        let value =
            if not hasFilters then total
            else total - hiddenPrefixSize defaultRowHeight model.RowHeights model.FilterHiddenRows row
        setRowTopFastSlot rowTopFastSlots model.RowHeights hasFilters hiddenRows row value
        value

let columnLeft (model: SpreadsheetModel) (col: int): float =
    match tryGetColumnLeftFastSlot columnLeftFastSlots model.ColumnWidths col with
    | ValueSome value -> value
    | ValueNone ->
        let value = prefixSize defaultColumnWidth model.ColumnWidths columnPrefixSizeCache col
        setColumnLeftFastSlot columnLeftFastSlots model.ColumnWidths col value
        value

let private nextVisibleRowAtOrAfter (model: SpreadsheetModel) row =
    let row = max 0 row
    if row >= model.RowCount then model.RowCount
    elif Map.isEmpty model.Filters || model.FilterHiddenRows.Length = 0 then row
    else
        let hiddenRows = model.FilterHiddenRows
        let start = lowerBound hiddenRows row
        if start >= hiddenRows.Length || hiddenRows.[start] <> row then
            row
        else
            let baseDelta = row - start
            let mutable lo = start
            let mutable hi = hiddenRows.Length
            while lo < hi do
                let mid = lo + (hi - lo) / 2
                if hiddenRows.[mid] - mid = baseDelta then
                    lo <- mid + 1
                else
                    hi <- mid
            min model.RowCount (row + (lo - start))

let private indexAtOffset defaultSize (prefixAt: int -> float) (count: int) (sizeAt: int -> float) (contentOffset: float) =
    if count <= 0 || contentOffset <= 0.0 then 0 else
    let mutable index = min (count - 1) (max 0 (int (floor (contentOffset / defaultSize))))
    let mutable start = prefixAt index

    while index > 0 && start > contentOffset do
        index <- index - 1
        start <- start - sizeAt index

    while index < count - 1 && start + sizeAt index <= contentOffset do
        start <- start + sizeAt index
        index <- index + 1

    index

let private rowIndexAtOffset (model: SpreadsheetModel) (contentOffset: float) =
    if Map.isEmpty model.Filters then
        indexAtOffset defaultRowHeight (rowTop model) model.RowCount (fun r -> getRowHeight r model) contentOffset
    elif model.RowCount <= 0 || contentOffset <= 0.0 then
        0
    else
        let mutable lo = 0
        let mutable hi = model.RowCount
        while lo < hi do
            let mid = lo + (hi - lo) / 2
            let bottom = rowTop model mid + getRowHeight mid model
            if bottom <= contentOffset then
                lo <- mid + 1
            else
                hi <- mid
        min (model.RowCount - 1) lo

let columnX (model: SpreadsheetModel) (viewport: Viewport) (col: int): float =
    viewport.RowHeaderWidth + columnLeft model col

let rowY (model: SpreadsheetModel) (viewport: Viewport) (row: int): float =
    viewport.HeaderHeight + rowTop model row

let visibleRange (model: SpreadsheetModel) (viewport: Viewport): VisibleRange =
    // startCol/startRow walks forward while the next cell would still sit
    // entirely above the scroll boundary; y / x land on the *top* of the
    // first visible cell. endCol/endRow then add successive cells without
    // re-adding the start cell — earlier versions reused `y` (= bottom of
    // start cell) and the end loop's first iteration double-counted the
    // start row, dropping the last row when row heights weren't uniform.
    let colOffset = max 0.0 (viewport.ScrollX - viewport.RowHeaderWidth)
    let rowOffset = max 0.0 (viewport.ScrollY - viewport.HeaderHeight)

    let mutable startCol = indexAtOffset defaultColumnWidth (columnLeft model) model.ColCount (fun c -> getColumnWidth c model) colOffset
    let mutable x = columnX model viewport startCol
    while startCol < model.ColCount
          && x + getColumnWidth startCol model <= viewport.ScrollX do
        x <- x + getColumnWidth startCol model
        startCol <- startCol + 1

    let mutable endCol = startCol
    let mutable xEnd = x
    while endCol < model.ColCount && xEnd < viewport.ScrollX + viewport.Width do
        xEnd <- xEnd + getColumnWidth endCol model
        endCol <- endCol + 1

    let mutable startRow = nextVisibleRowAtOrAfter model (rowIndexAtOffset model rowOffset)
    let mutable y = rowY model viewport startRow
    while startRow < model.RowCount
          && y + getRowHeight startRow model <= viewport.ScrollY do
        y <- y + getRowHeight startRow model
        startRow <- nextVisibleRowAtOrAfter model (startRow + 1)

    let mutable endRow = startRow
    let mutable yEnd = y
    while endRow < model.RowCount && yEnd < viewport.ScrollY + viewport.Height do
        yEnd <- yEnd + getRowHeight endRow model
        endRow <- nextVisibleRowAtOrAfter model (endRow + 1)

    {
        StartRow = startRow
        EndRowExclusive = endRow
        StartCol = startCol
        EndColExclusive = endCol
    }

let cellRect (model: SpreadsheetModel) (viewport: Viewport) (row: int) (col: int): Rect =
    {
        X = columnX model viewport col
        Y = rowY model viewport row
        Width = getColumnWidth col model
        Height = getRowHeight row model
    }

let mergedRect (model: SpreadsheetModel) (viewport: Viewport) (merge: MergedCell): Rect =
    {
        X = columnX model viewport merge.StartCol
        Y = rowY model viewport merge.StartRow
        Width = columnLeft model (merge.EndCol + 1) - columnLeft model merge.StartCol
        Height = rowTop model (merge.EndRow + 1) - rowTop model merge.StartRow
    }

let private cellKey row col =
    ((int64 row) <<< 32) ||| int64 col

type private VisibleCellsCache = {
    Model: SpreadsheetModel
    SourceItemCount: int
    SourceTextCharCount: int
    Entries: System.Collections.Generic.Dictionary<ViewportGeometryKey, CellGeometry list>
    Order: System.Collections.Generic.Queue<ViewportGeometryKey>
    EntryFastSlots: ViewportGeometryFastEntry<CellGeometry list> array
    mutable TotalItemCount: int
}

type private HeadersCache = {
    Model: SpreadsheetModel
    SourceItemCount: int
    SourceTextCharCount: int
    Entries: System.Collections.Generic.Dictionary<ViewportGeometryKey, HeaderGeometry list>
    Order: System.Collections.Generic.Queue<ViewportGeometryKey>
    EntryFastSlots: ViewportGeometryFastEntry<HeaderGeometry list> array
    mutable TotalItemCount: int
}

let mutable private visibleCellsCache: VisibleCellsCache list = []
let mutable private headersCache: HeadersCache list = []
let private spreadsheetModelReferenceComparer =
    { new System.Collections.Generic.IEqualityComparer<SpreadsheetModel> with
        member _.Equals(left, right) = obj.ReferenceEquals(left, right)
        member _.GetHashCode(value) = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value) }
let private visibleCellsCacheByModel =
    System.Collections.Generic.Dictionary<SpreadsheetModel, VisibleCellsCache>(spreadsheetModelReferenceComparer)
let private headersCacheByModel =
    System.Collections.Generic.Dictionary<SpreadsheetModel, HeadersCache>(spreadsheetModelReferenceComparer)
let private visibleCellsFastSlots = Array.zeroCreate<VisibleCellsCache> 256
let private headersFastSlots = Array.zeroCreate<HeadersCache> 256
let private viewportGeometryCacheMaxEntries = 256
let private viewportGeometryCacheMaxModels = 256
let private viewportGeometryCacheMaxTotalItems = 200000
let private viewportGeometryCacheMaxTotalSourceItems = 200000
let private viewportGeometryCacheMaxTotalSourceTextChars = 5000000

let private viewportGeometryModelSlot (model: SpreadsheetModel) slotCount =
    System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(model) &&& (slotCount - 1)

let private viewportGeometryAddCapped limit value total =
    if total > limit then total
    else min (limit + 1) (total + max 0 value)

let private viewportGeometryListCount limit (items: 'T list) =
    let mutable count = 0
    let mutable remaining = items
    while count <= limit && not remaining.IsEmpty do
        count <- count + 1
        remaining <- remaining.Tail
    count

let private viewportGeometryOptionItemCount = function
    | Some _ -> 1
    | None -> 0

let private viewportGeometryFilterItemCount (filters: Map<int, ColumnFilter>) =
    if filters.Count = 0 then 0
    else
        let mutable total = 0
        use entries = (filters :> seq<_>).GetEnumerator()
        while total <= viewportGeometryCacheMaxTotalSourceItems && entries.MoveNext() do
            let filter = entries.Current.Value
            total <- viewportGeometryAddCapped viewportGeometryCacheMaxTotalSourceItems 1 total
            total <- viewportGeometryAddCapped viewportGeometryCacheMaxTotalSourceItems filter.VisibleValues.Count total
        total

let private viewportGeometryConditionalFormatItemCount (formats: Map<CellAddress, ConditionalFormat list>) =
    if formats.Count = 0 then 0
    else
        let mutable total = 0
        use entries = (formats :> seq<_>).GetEnumerator()
        while total <= viewportGeometryCacheMaxTotalSourceItems && entries.MoveNext() do
            let items = entries.Current.Value
            total <- viewportGeometryAddCapped viewportGeometryCacheMaxTotalSourceItems 1 total
            total <- viewportGeometryAddCapped viewportGeometryCacheMaxTotalSourceItems (viewportGeometryListCount viewportGeometryCacheMaxTotalSourceItems items) total
        total

let private viewportGeometryRichRunItemCount (runs: Map<CellAddress, (string * RunOverride) list>) =
    if runs.Count = 0 then 0
    else
        let mutable total = 0
        use entries = (runs :> seq<_>).GetEnumerator()
        while total <= viewportGeometryCacheMaxTotalSourceItems && entries.MoveNext() do
            let items = entries.Current.Value
            total <- viewportGeometryAddCapped viewportGeometryCacheMaxTotalSourceItems 1 total
            total <- viewportGeometryAddCapped viewportGeometryCacheMaxTotalSourceItems (viewportGeometryListCount viewportGeometryCacheMaxTotalSourceItems items) total
        total

let private viewportGeometryValidationRuleItemCount rule =
    match rule with
    | ValidationRule.ListValues values ->
        let mutable total = 0
        let mutable remaining = values
        while total <= viewportGeometryCacheMaxTotalSourceItems && not remaining.IsEmpty do
            total <- viewportGeometryAddCapped viewportGeometryCacheMaxTotalSourceItems 1 total
            remaining <- remaining.Tail
        total
    | ValidationRule.WholeNumber _
    | ValidationRule.DecimalRange _
    | ValidationRule.TextLength _
    | ValidationRule.DateRange _ -> 0

let private viewportGeometryValidationItemCount (validations: Map<CellAddress, DataValidation>) =
    if validations.Count = 0 then 0
    else
        let mutable total = 0
        use entries = (validations :> seq<_>).GetEnumerator()
        while total <= viewportGeometryCacheMaxTotalSourceItems && entries.MoveNext() do
            let validation = entries.Current.Value
            total <- viewportGeometryAddCapped viewportGeometryCacheMaxTotalSourceItems 1 total
            total <- viewportGeometryAddCapped viewportGeometryCacheMaxTotalSourceItems (viewportGeometryValidationRuleItemCount validation.Rule) total
        total

let private viewportGeometryPageSetupItemCount (pageSetup: PageSetup) =
    let limit = viewportGeometryCacheMaxTotalSourceItems
    let mutable total = viewportGeometryListCount limit pageSetup.ManualRowBreaks
    if total <= limit then
        total <- viewportGeometryAddCapped limit (viewportGeometryListCount limit pageSetup.ManualColBreaks) total
    if total <= limit then
        total <- viewportGeometryAddCapped limit (viewportGeometryOptionItemCount pageSetup.PrintArea) total
    if total <= limit then
        total <- viewportGeometryAddCapped limit (viewportGeometryOptionItemCount pageSetup.RepeatRows) total
    if total <= limit then
        total <- viewportGeometryAddCapped limit (viewportGeometryOptionItemCount pageSetup.RepeatCols) total
    total

let private viewportGeometryClipboardItemCount = function
    | Some clipboard -> commandRetainedClipboardItemCount clipboard
    | None -> 0

let private viewportGeometryHistoryItemCount (history: History) =
    let mutable total = 0
    let mutable past = history.Past
    while total <= viewportGeometryCacheMaxTotalSourceItems && not past.IsEmpty do
        let command = past.Head
        total <- viewportGeometryAddCapped viewportGeometryCacheMaxTotalSourceItems command.RetainedItemCount total
        past <- past.Tail
    let mutable future = history.Future
    while total <= viewportGeometryCacheMaxTotalSourceItems && not future.IsEmpty do
        let command = future.Head
        total <- viewportGeometryAddCapped viewportGeometryCacheMaxTotalSourceItems command.RetainedItemCount total
        future <- future.Tail
    total

let private viewportGeometrySourceItemCount (model: SpreadsheetModel) =
    let limit = viewportGeometryCacheMaxTotalSourceItems
    let mutable total = 0
    if total <= limit then total <- viewportGeometryAddCapped limit model.Cells.Count total
    if total <= limit then total <- viewportGeometryAddCapped limit (viewportGeometryListCount limit model.Tables) total
    if total <= limit then total <- viewportGeometryAddCapped limit (viewportGeometryConditionalFormatItemCount model.ConditionalFormats) total
    if total <= limit then total <- viewportGeometryAddCapped limit (viewportGeometryListCount limit model.MergedCells) total
    if total <= limit then total <- viewportGeometryAddCapped limit (viewportGeometryRichRunItemCount model.RichRuns) total
    if total <= limit then total <- viewportGeometryAddCapped limit model.RowHeights.Count total
    if total <= limit then total <- viewportGeometryAddCapped limit model.ColumnWidths.Count total
    if total <= limit then total <- viewportGeometryAddCapped limit (viewportGeometryFilterItemCount model.Filters) total
    if total <= limit then total <- viewportGeometryAddCapped limit model.FilterHiddenRows.Length total
    if total <= limit then total <- viewportGeometryAddCapped limit model.Spill.Count total
    if total <= limit then total <- viewportGeometryAddCapped limit (viewportGeometryValidationItemCount model.Validations) total
    if total <= limit then total <- viewportGeometryAddCapped limit model.SearchResults.Length total
    if total <= limit then total <- viewportGeometryAddCapped limit model.CfBoldCells.Count total
    if total <= limit then total <- viewportGeometryAddCapped limit (viewportGeometryPageSetupItemCount model.PageSetup) total
    if total <= limit then total <- viewportGeometryAddCapped limit (viewportGeometryClipboardItemCount model.Clipboard) total
    if total <= limit then total <- viewportGeometryAddCapped limit (viewportGeometryHistoryItemCount model.History) total
    max 1 total

let private viewportGeometryOptionStringCharCount (value: string option) =
    match value with
    | Some value -> value.Length
    | None -> 0

let private viewportGeometryBorderStyleStringCharCount (border: BorderStyle option) =
    match border with
    | Some border -> border.Style.Length + viewportGeometryOptionStringCharCount border.Color
    | None -> 0

let private viewportGeometryCellValueStringCharCount value =
    let mutable total = 0
    let mutable current = value
    let mutable doneCounting = false
    while not doneCounting do
        match current with
        | CellValue.String value
        | CellValue.Error value ->
            total <- total + value.Length
            doneCounting <- true
        | CellValue.Formula (formula, cachedValue) ->
            total <- total + formula.Length
            current <- cachedValue
        | CellValue.Empty
        | CellValue.Number _
        | CellValue.Boolean _ ->
            doneCounting <- true
    total

let private viewportGeometryCellFormatStringCharCount (format: CellFormat) =
    viewportGeometryOptionStringCharCount format.FontName
    + viewportGeometryOptionStringCharCount format.TextColor
    + viewportGeometryOptionStringCharCount format.BgColor
    + viewportGeometryOptionStringCharCount format.HorizontalAlign
    + viewportGeometryOptionStringCharCount format.VerticalAlign
    + viewportGeometryOptionStringCharCount format.NumberFormat
    + viewportGeometryBorderStyleStringCharCount format.BorderTop
    + viewportGeometryBorderStyleStringCharCount format.BorderBottom
    + viewportGeometryBorderStyleStringCharCount format.BorderLeft
    + viewportGeometryBorderStyleStringCharCount format.BorderRight
    + viewportGeometryBorderStyleStringCharCount format.BorderDiagonal
    + (match format.DataBar with Some (_, color) -> color.Length | None -> 0)

let private viewportGeometryCellCommentStringCharCount (comment: CellComment option) =
    match comment with
    | Some comment -> comment.Text.Length + viewportGeometryOptionStringCharCount comment.Author
    | None -> 0

let private viewportGeometryRunOverrideStringCharCount (run: RunOverride) =
    viewportGeometryOptionStringCharCount run.Color
    + viewportGeometryOptionStringCharCount run.FontName

let private viewportGeometryTableStyleStringCharCount (style: TableStyle option) =
    match style with
    | Some style ->
        style.Name.Length
        + viewportGeometryOptionStringCharCount style.HeaderFill
        + viewportGeometryOptionStringCharCount style.StripeFill
    | None -> 0

let private viewportGeometryConditionalRuleStringCharCount rule =
    match rule with
    | ConditionalFormatRule.ContainsText text -> text.Length
    | ConditionalFormatRule.GreaterThan _
    | ConditionalFormatRule.LessThan _
    | ConditionalFormatRule.Between _
    | ConditionalFormatRule.Equal _ -> 0

let private viewportGeometryValidationRuleStringCharCount rule =
    match rule with
    | ValidationRule.ListValues values ->
        let mutable total = 0
        let mutable seen = 0
        let mutable remaining = values
        while total <= viewportGeometryCacheMaxTotalSourceTextChars
              && seen <= viewportGeometryCacheMaxTotalSourceItems
              && not remaining.IsEmpty do
            total <- viewportGeometryAddCapped viewportGeometryCacheMaxTotalSourceTextChars remaining.Head.Length total
            seen <- seen + 1
            remaining <- remaining.Tail
        total
    | ValidationRule.WholeNumber _
    | ValidationRule.DecimalRange _
    | ValidationRule.TextLength _
    | ValidationRule.DateRange _ -> 0

let private viewportGeometryHeaderFooterStringCharCount (headerFooter: HeaderFooter) =
    viewportGeometryOptionStringCharCount headerFooter.OddHeader
    + viewportGeometryOptionStringCharCount headerFooter.OddFooter
    + viewportGeometryOptionStringCharCount headerFooter.EvenHeader
    + viewportGeometryOptionStringCharCount headerFooter.EvenFooter
    + viewportGeometryOptionStringCharCount headerFooter.FirstHeader
    + viewportGeometryOptionStringCharCount headerFooter.FirstFooter

let private viewportGeometryClipboardTextCharCount = function
    | Some clipboard -> commandRetainedClipboardTextCharCount clipboard
    | None -> 0

let private viewportGeometryHistoryTextCharCount (history: History) =
    let mutable total = 0
    let mutable past = history.Past
    while total <= viewportGeometryCacheMaxTotalSourceTextChars && not past.IsEmpty do
        let command = past.Head
        total <-
            viewportGeometryAddCapped
                viewportGeometryCacheMaxTotalSourceTextChars
                (command.Description.Length + command.RetainedTextCharCount)
                total
        past <- past.Tail
    let mutable future = history.Future
    while total <= viewportGeometryCacheMaxTotalSourceTextChars && not future.IsEmpty do
        let command = future.Head
        total <-
            viewportGeometryAddCapped
                viewportGeometryCacheMaxTotalSourceTextChars
                (command.Description.Length + command.RetainedTextCharCount)
                total
        future <- future.Tail
    total

let private viewportGeometrySourceTextCharCount (model: SpreadsheetModel) =
    let limit = viewportGeometryCacheMaxTotalSourceTextChars
    let cappedLimit = limit + 1
    let mutable total = viewportGeometryOptionStringCharCount model.SearchQuery + model.EditBuffer.Length + model.Macros.Length
    if total <= limit then
        total <- viewportGeometryAddCapped limit (viewportGeometryHeaderFooterStringCharCount model.PageSetup.HeaderFooter) total
    if total <= limit then
        total <- viewportGeometryAddCapped limit (viewportGeometryClipboardTextCharCount model.Clipboard) total
    if total <= limit then
        total <- viewportGeometryAddCapped limit (viewportGeometryHistoryTextCharCount model.History) total
    if total <= limit && model.Filters.Count > 0 then
        use filters = (model.Filters :> seq<_>).GetEnumerator()
        while total <= limit && filters.MoveNext() do
            let filter = filters.Current.Value
            use values = (filter.VisibleValues :> seq<string>).GetEnumerator()
            while total <= limit && values.MoveNext() do
                total <- viewportGeometryAddCapped limit values.Current.Length total
    if total <= limit && model.ConditionalFormats.Count > 0 then
        use conditionalFormats = (model.ConditionalFormats :> seq<_>).GetEnumerator()
        while total <= limit && conditionalFormats.MoveNext() do
            let mutable formats = conditionalFormats.Current.Value
            while total <= limit && not formats.IsEmpty do
                let format = formats.Head
                total <- viewportGeometryAddCapped limit (viewportGeometryConditionalRuleStringCharCount format.Rule) total
                total <- viewportGeometryAddCapped limit (cellFormatTextCharCountFor limit format.Style) total
                formats <- formats.Tail
    if total <= limit && model.RichRuns.Count > 0 then
        use richRuns = (model.RichRuns :> seq<_>).GetEnumerator()
        while total <= limit && richRuns.MoveNext() do
            let mutable runs = richRuns.Current.Value
            while total <= limit && not runs.IsEmpty do
                let text, run = runs.Head
                total <- viewportGeometryAddCapped limit (text.Length + viewportGeometryRunOverrideStringCharCount run) total
                runs <- runs.Tail
    let mutable tables = model.Tables
    let mutable seenTables = 0
    while total <= limit
          && seenTables <= viewportGeometryCacheMaxTotalSourceItems
          && not tables.IsEmpty do
        let table = tables.Head
        total <- viewportGeometryAddCapped limit table.Name.Length total
        if total <= limit then
            total <- viewportGeometryAddCapped limit (viewportGeometryTableStyleStringCharCount table.Style) total
        seenTables <- seenTables + 1
        tables <- tables.Tail
    if total <= limit && model.Validations.Count > 0 then
        use validations = (model.Validations :> seq<_>).GetEnumerator()
        while total <= limit && validations.MoveNext() do
            let validation = validations.Current.Value
            total <- viewportGeometryAddCapped limit (viewportGeometryValidationRuleStringCharCount validation.Rule) total
            total <- viewportGeometryAddCapped limit (viewportGeometryOptionStringCharCount validation.ErrorMessage) total
            total <- viewportGeometryAddCapped limit (viewportGeometryOptionStringCharCount validation.InputMessage) total
    if total <= limit && model.Cells.Count > 0 then
        model.Cells
        |> Map.exists (fun _ cell ->
            total <- viewportGeometryAddCapped limit (cellDataTextCharCountFor limit cell) total
            total > limit)
        |> ignore
    min total cappedLimit

let private viewportGeometrySourceTextCharCountFor sourceItemCount model =
    if sourceItemCount > viewportGeometryCacheMaxTotalSourceItems then
        viewportGeometryCacheMaxTotalSourceTextChars + 1
    else
        viewportGeometrySourceTextCharCount model

type private ViewportGeometrySourceMetrics = {
    Cells: Map<CellAddress, CellData>
    Tables: Table list
    ConditionalFormats: Map<CellAddress, ConditionalFormat list>
    MergedCells: MergedCell list
    RichRuns: Map<CellAddress, (string * RunOverride) list>
    RowHeights: Map<int, float>
    ColumnWidths: Map<int, float>
    Filters: Map<int, ColumnFilter>
    FilterHiddenRows: int array
    Spill: Map<CellAddress, CellAddress>
    Validations: Map<CellAddress, DataValidation>
    SearchResults: CellAddress array
    CfBoldCells: Set<CellAddress>
    PageSetup: PageSetup
    Clipboard: ClipboardData option
    History: History
    SearchQuery: string option
    EditBuffer: string
    Macros: string
    SourceItemCount: int
    SourceTextCharCount: int
}

let mutable private viewportGeometrySourceMetricsCache: ViewportGeometrySourceMetrics option = None

let private viewportGeometrySourceMetricsMatch (model: SpreadsheetModel) (cached: ViewportGeometrySourceMetrics) =
    obj.ReferenceEquals(cached.Cells, model.Cells)
    && obj.ReferenceEquals(cached.Tables, model.Tables)
    && obj.ReferenceEquals(cached.ConditionalFormats, model.ConditionalFormats)
    && obj.ReferenceEquals(cached.MergedCells, model.MergedCells)
    && obj.ReferenceEquals(cached.RichRuns, model.RichRuns)
    && obj.ReferenceEquals(cached.RowHeights, model.RowHeights)
    && obj.ReferenceEquals(cached.ColumnWidths, model.ColumnWidths)
    && obj.ReferenceEquals(cached.Filters, model.Filters)
    && obj.ReferenceEquals(cached.FilterHiddenRows, model.FilterHiddenRows)
    && obj.ReferenceEquals(cached.Spill, model.Spill)
    && obj.ReferenceEquals(cached.Validations, model.Validations)
    && obj.ReferenceEquals(cached.SearchResults, model.SearchResults)
    && obj.ReferenceEquals(cached.CfBoldCells, model.CfBoldCells)
    && obj.ReferenceEquals(cached.PageSetup, model.PageSetup)
    && obj.ReferenceEquals(cached.Clipboard, model.Clipboard)
    && obj.ReferenceEquals(cached.History, model.History)
    && obj.ReferenceEquals(cached.SearchQuery, model.SearchQuery)
    && obj.ReferenceEquals(cached.EditBuffer, model.EditBuffer)
    && obj.ReferenceEquals(cached.Macros, model.Macros)

let private viewportGeometrySourceMetrics (model: SpreadsheetModel) =
    match viewportGeometrySourceMetricsCache with
    | Some cached when viewportGeometrySourceMetricsMatch model cached ->
        cached.SourceItemCount, cached.SourceTextCharCount
    | _ ->
        let sourceItemCount = viewportGeometrySourceItemCount model
        let sourceTextCharCount = viewportGeometrySourceTextCharCountFor sourceItemCount model
        if sourceItemCount <= viewportGeometryCacheMaxTotalSourceItems
           && sourceTextCharCount <= viewportGeometryCacheMaxTotalSourceTextChars then
            viewportGeometrySourceMetricsCache <- Some {
                Cells = model.Cells
                Tables = model.Tables
                ConditionalFormats = model.ConditionalFormats
                MergedCells = model.MergedCells
                RichRuns = model.RichRuns
                RowHeights = model.RowHeights
                ColumnWidths = model.ColumnWidths
                Filters = model.Filters
                FilterHiddenRows = model.FilterHiddenRows
                Spill = model.Spill
                Validations = model.Validations
                SearchResults = model.SearchResults
                CfBoldCells = model.CfBoldCells
                PageSetup = model.PageSetup
                Clipboard = model.Clipboard
                History = model.History
                SearchQuery = model.SearchQuery
                EditBuffer = model.EditBuffer
                Macros = model.Macros
                SourceItemCount = sourceItemCount
                SourceTextCharCount = sourceTextCharCount
            }
        else
            viewportGeometrySourceMetricsCache <- None
        sourceItemCount, sourceTextCharCount

let private visibleCellsCacheRetainable (cache: VisibleCellsCache) =
    cache.Entries.Count > 0
    && cache.TotalItemCount <= viewportGeometryCacheMaxTotalItems
    && cache.SourceItemCount <= viewportGeometryCacheMaxTotalSourceItems
    && cache.SourceTextCharCount <= viewportGeometryCacheMaxTotalSourceTextChars

let private headersCacheRetainable (cache: HeadersCache) =
    cache.Entries.Count > 0
    && cache.TotalItemCount <= viewportGeometryCacheMaxTotalItems
    && cache.SourceItemCount <= viewportGeometryCacheMaxTotalSourceItems
    && cache.SourceTextCharCount <= viewportGeometryCacheMaxTotalSourceTextChars

let private setVisibleCellsCache entries =
    visibleCellsCache <- entries
    visibleCellsCacheByModel.Clear()
    clearArray visibleCellsFastSlots
    for entry in entries do
        visibleCellsCacheByModel.[entry.Model] <- entry
        visibleCellsFastSlots.[viewportGeometryModelSlot entry.Model visibleCellsFastSlots.Length] <- entry

let private setHeadersCache entries =
    headersCache <- entries
    headersCacheByModel.Clear()
    clearArray headersFastSlots
    for entry in entries do
        headersCacheByModel.[entry.Model] <- entry
        headersFastSlots.[viewportGeometryModelSlot entry.Model headersFastSlots.Length] <- entry

let private tryGetVisibleCellsCache model =
    let slot = viewportGeometryModelSlot model visibleCellsFastSlots.Length
    let fast = visibleCellsFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null)) && obj.ReferenceEquals(fast.Model, model) then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<VisibleCellsCache>
        if visibleCellsCacheByModel.TryGetValue(model, &cached)
           && obj.ReferenceEquals(cached.Model, model) then
            visibleCellsFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private tryGetHeadersCache model =
    let slot = viewportGeometryModelSlot model headersFastSlots.Length
    let fast = headersFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null)) && obj.ReferenceEquals(fast.Model, model) then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<HeadersCache>
        if headersCacheByModel.TryGetValue(model, &cached)
           && obj.ReferenceEquals(cached.Model, model) then
            headersFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private visibleCellsCacheWithoutModel model =
    let mutable cached = Unchecked.defaultof<VisibleCellsCache>
    if visibleCellsCacheByModel.TryGetValue(model, &cached) then
        visibleCellsCache
        |> List.filter (fun cached -> not (obj.ReferenceEquals(cached.Model, model)))
    else
        visibleCellsCache

let private removeVisibleCellsCacheForModel model =
    let entries = visibleCellsCacheWithoutModel model
    if not (obj.ReferenceEquals(entries, visibleCellsCache)) then
        setVisibleCellsCache entries

let private headersCacheWithoutModel model =
    let mutable cached = Unchecked.defaultof<HeadersCache>
    if headersCacheByModel.TryGetValue(model, &cached) then
        headersCache
        |> List.filter (fun cached -> not (obj.ReferenceEquals(cached.Model, model)))
    else
        headersCache

let private removeHeadersCacheForModel model =
    let entries = headersCacheWithoutModel model
    if not (obj.ReferenceEquals(entries, headersCache)) then
        setHeadersCache entries

let private rememberVisibleCellsCache (cache: VisibleCellsCache) =
    if not (visibleCellsCacheRetainable cache) then
        removeVisibleCellsCacheForModel cache.Model
    else
        let mutable keptCount = 1
        let mutable keptItems = cache.TotalItemCount
        let mutable keptSourceItems = cache.SourceItemCount
        let mutable keptSourceTextChars = cache.SourceTextCharCount
        let retained =
            visibleCellsCacheWithoutModel cache.Model
            |> List.choose (fun candidate ->
                let candidateItems = candidate.TotalItemCount
                let candidateSourceItems = candidate.SourceItemCount
                let candidateSourceTextChars = candidate.SourceTextCharCount
                if visibleCellsCacheRetainable candidate
                   && keptCount < viewportGeometryCacheMaxModels
                   && keptItems + candidateItems <= viewportGeometryCacheMaxTotalItems
                   && keptSourceItems + candidateSourceItems <= viewportGeometryCacheMaxTotalSourceItems
                   && keptSourceTextChars + candidateSourceTextChars <= viewportGeometryCacheMaxTotalSourceTextChars then
                    keptCount <- keptCount + 1
                    keptItems <- keptItems + candidateItems
                    keptSourceItems <- keptSourceItems + candidateSourceItems
                    keptSourceTextChars <- keptSourceTextChars + candidateSourceTextChars
                    Some candidate
                else
                    None)
        setVisibleCellsCache (cache :: retained)

let private rememberHeadersCache (cache: HeadersCache) =
    if not (headersCacheRetainable cache) then
        removeHeadersCacheForModel cache.Model
    else
        let mutable keptCount = 1
        let mutable keptItems = cache.TotalItemCount
        let mutable keptSourceItems = cache.SourceItemCount
        let mutable keptSourceTextChars = cache.SourceTextCharCount
        let retained =
            headersCacheWithoutModel cache.Model
            |> List.choose (fun candidate ->
                let candidateItems = candidate.TotalItemCount
                let candidateSourceItems = candidate.SourceItemCount
                let candidateSourceTextChars = candidate.SourceTextCharCount
                if headersCacheRetainable candidate
                   && keptCount < viewportGeometryCacheMaxModels
                   && keptItems + candidateItems <= viewportGeometryCacheMaxTotalItems
                   && keptSourceItems + candidateSourceItems <= viewportGeometryCacheMaxTotalSourceItems
                   && keptSourceTextChars + candidateSourceTextChars <= viewportGeometryCacheMaxTotalSourceTextChars then
                    keptCount <- keptCount + 1
                    keptItems <- keptItems + candidateItems
                    keptSourceItems <- keptSourceItems + candidateSourceItems
                    keptSourceTextChars <- keptSourceTextChars + candidateSourceTextChars
                    Some candidate
                else
                    None)
        setHeadersCache (cache :: retained)

let private removeVisibleCellsCacheEntry key (cache: VisibleCellsCache) =
    let mutable existing = Unchecked.defaultof<CellGeometry list>
    if cache.Entries.TryGetValue(key, &existing) then
        clearArray cache.EntryFastSlots
        cache.TotalItemCount <- max 0 (cache.TotalItemCount - existing.Length)
        cache.Entries.Remove key |> ignore

let private addVisibleCellsCacheEntryWithCount key (cells: CellGeometry list) cellCount (cache: VisibleCellsCache) =
    let mutable existing = Unchecked.defaultof<CellGeometry list>
    if cellCount > viewportGeometryCacheMaxTotalItems then
        removeVisibleCellsCacheEntry key cache
    else
        if cache.Entries.TryGetValue(key, &existing) then
            cache.TotalItemCount <- max 0 (cache.TotalItemCount - existing.Length)
        else
            cache.Order.Enqueue key
        cache.Entries.[key] <- cells
        cache.TotalItemCount <- cache.TotalItemCount + cellCount
        setViewportGeometryFastSlot cache.EntryFastSlots key cells
        let mutable removedEntry = false
        while (cache.Entries.Count > viewportGeometryCacheMaxEntries
               || cache.TotalItemCount > viewportGeometryCacheMaxTotalItems)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if oldest <> key then
                removedEntry <- true
                removeVisibleCellsCacheEntry oldest cache
        if cache.Entries.Count > viewportGeometryCacheMaxEntries
           || cache.TotalItemCount > viewportGeometryCacheMaxTotalItems then
            let current = cache.Entries.[key]
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue key
            cache.Entries.[key] <- current
            cache.TotalItemCount <- current.Length
            clearArray cache.EntryFastSlots
            setViewportGeometryFastSlot cache.EntryFastSlots key current
        elif removedEntry then
            clearArray cache.EntryFastSlots
            let mutable current = Unchecked.defaultof<CellGeometry list>
            if cache.Entries.TryGetValue(key, &current) then
                setViewportGeometryFastSlot cache.EntryFastSlots key current

let private addVisibleCellsCacheEntry key (cells: CellGeometry list) (cache: VisibleCellsCache) =
    addVisibleCellsCacheEntryWithCount key cells cells.Length cache

let private removeHeadersCacheEntry key (cache: HeadersCache) =
    let mutable existing = Unchecked.defaultof<HeaderGeometry list>
    if cache.Entries.TryGetValue(key, &existing) then
        clearArray cache.EntryFastSlots
        cache.TotalItemCount <- max 0 (cache.TotalItemCount - existing.Length)
        cache.Entries.Remove key |> ignore

let private addHeadersCacheEntryWithCount key (headers: HeaderGeometry list) headerCount (cache: HeadersCache) =
    let mutable existing = Unchecked.defaultof<HeaderGeometry list>
    if headerCount > viewportGeometryCacheMaxTotalItems then
        removeHeadersCacheEntry key cache
    else
        if cache.Entries.TryGetValue(key, &existing) then
            cache.TotalItemCount <- max 0 (cache.TotalItemCount - existing.Length)
        else
            cache.Order.Enqueue key
        cache.Entries.[key] <- headers
        cache.TotalItemCount <- cache.TotalItemCount + headerCount
        setViewportGeometryFastSlot cache.EntryFastSlots key headers
        let mutable removedEntry = false
        while (cache.Entries.Count > viewportGeometryCacheMaxEntries
               || cache.TotalItemCount > viewportGeometryCacheMaxTotalItems)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if oldest <> key then
                removedEntry <- true
                removeHeadersCacheEntry oldest cache
        if cache.Entries.Count > viewportGeometryCacheMaxEntries
           || cache.TotalItemCount > viewportGeometryCacheMaxTotalItems then
            let current = cache.Entries.[key]
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue key
            cache.Entries.[key] <- current
            cache.TotalItemCount <- current.Length
            clearArray cache.EntryFastSlots
            setViewportGeometryFastSlot cache.EntryFastSlots key current
        elif removedEntry then
            clearArray cache.EntryFastSlots
            let mutable current = Unchecked.defaultof<HeaderGeometry list>
            if cache.Entries.TryGetValue(key, &current) then
                setViewportGeometryFastSlot cache.EntryFastSlots key current

let private addHeadersCacheEntry key (headers: HeaderGeometry list) (cache: HeadersCache) =
    addHeadersCacheEntryWithCount key headers headers.Length cache

let iterVisibleCells (model: SpreadsheetModel) (viewport: Viewport) (emit: CellGeometry -> unit) =
    let range = visibleRange model viewport
    let visibleCols, visibleColCount =
        let count = max 0 (range.EndColExclusive - range.StartCol)
        let cols = Array.zeroCreate<int * float * float> count
        let mutable visibleCount = 0
        let mutable x = columnX model viewport range.StartCol
        for i in 0 .. count - 1 do
            let col = range.StartCol + i
            let width = getColumnWidth col model
            if width > 0.0 then
                cols.[visibleCount] <- (col, x, width)
                visibleCount <- visibleCount + 1
            x <- x + width
        cols, visibleCount
    let visibleRows, visibleRowCount =
        let rows = System.Collections.Generic.List<int * float * float>()
        let mutable row = nextVisibleRowAtOrAfter model range.StartRow
        let mutable y = rowY model viewport row
        while row < range.EndRowExclusive do
            let height = getRowHeight row model
            if height > 0.0 then
                rows.Add(row, y, height)
            y <- y + height
            row <- nextVisibleRowAtOrAfter model (row + 1)
        rows, rows.Count
    let visibleMerges =
        mergedCellsIntersectingRange range.StartRow range.EndRowExclusive range.StartCol range.EndColExclusive model
    if visibleMerges.IsEmpty then
        for rowIndex in 0 .. visibleRowCount - 1 do
            let row, y, height = visibleRows.[rowIndex]
            for colIndex in 0 .. visibleColCount - 1 do
                let col, x, width = visibleCols.[colIndex]
                emit {
                    Address = { Row = row; Col = col }
                    Rect = {
                        X = x
                        Y = y
                        Width = width
                        Height = height
                    }
                    IsMergedPrimary = false
                    Merge = None
                }
    else
        let mergedCoveredCells = System.Collections.Generic.HashSet<int64>()
        let columnLefts = System.Collections.Generic.Dictionary<int, float>()
        for colIndex in 0 .. visibleColCount - 1 do
            let col, x, width = visibleCols.[colIndex]
            columnLefts.[col] <- x
            columnLefts.[col + 1] <- x + width
        let rowTops = System.Collections.Generic.Dictionary<int, float>()
        for rowIndex in 0 .. visibleRowCount - 1 do
            let row, y, height = visibleRows.[rowIndex]
            rowTops.[row] <- y
            rowTops.[row + 1] <- y + height

        for merge in visibleMerges do
            let startRow = max merge.StartRow range.StartRow
            let endRow = min merge.EndRow (range.EndRowExclusive - 1)
            let startCol = max merge.StartCol range.StartCol
            let endCol = min merge.EndCol (range.EndColExclusive - 1)
            for row in startRow .. endRow do
                for col in startCol .. endCol do
                    mergedCoveredCells.Add(cellKey row col) |> ignore

        let mergeColumnX col =
            match columnLefts.TryGetValue col with
            | true, x -> x
            | false, _ -> columnX model viewport col

        let mergeRowY row =
            match rowTops.TryGetValue row with
            | true, y -> y
            | false, _ -> rowY model viewport row

        for merge in visibleMerges do
            let x = mergeColumnX merge.StartCol
            let y = mergeRowY merge.StartRow
            let x2 = mergeColumnX (merge.EndCol + 1)
            let y2 = mergeRowY (merge.EndRow + 1)
            emit {
                Address = { Row = merge.StartRow; Col = merge.StartCol }
                Rect = { X = x; Y = y; Width = x2 - x; Height = y2 - y }
                IsMergedPrimary = true
                Merge = Some merge
            }

        for rowIndex in 0 .. visibleRowCount - 1 do
            let row, y, height = visibleRows.[rowIndex]
            for colIndex in 0 .. visibleColCount - 1 do
                let col, x, width = visibleCols.[colIndex]
                if not (mergedCoveredCells.Contains(cellKey row col)) then
                    let addr = { Row = row; Col = col }
                    emit {
                        Address = addr
                        Rect = {
                            X = x
                            Y = y
                            Width = width
                            Height = height
                        }
                        IsMergedPrimary = false
                        Merge = None
                    }

let visibleCells (model: SpreadsheetModel) (viewport: Viewport): CellGeometry list =
    match tryGetVisibleCellsCache model with
    | ValueSome cached when visibleCellsCacheRetainable cached ->
        match tryGetViewportGeometryFastSlot cached.EntryFastSlots viewport with
        | ValueSome cells -> cells
        | ValueNone ->
            let key = viewportGeometryKey viewport
            let mutable cells = Unchecked.defaultof<CellGeometry list>
            if cached.Entries.TryGetValue(key, &cells) then
                setViewportGeometryFastSlot cached.EntryFastSlots key cells
                cells
            else
                let values = System.Collections.Generic.List<CellGeometry>()
                iterVisibleCells model viewport values.Add
                let cellCount = values.Count
                let cells = resizeArrayToList values
                addVisibleCellsCacheEntryWithCount key cells cellCount cached
                rememberVisibleCellsCache cached
                cells
    | ValueSome cached ->
        rememberVisibleCellsCache cached
        let values = System.Collections.Generic.List<CellGeometry>()
        iterVisibleCells model viewport values.Add
        resizeArrayToList values
    | ValueNone ->
        let values = System.Collections.Generic.List<CellGeometry>()
        iterVisibleCells model viewport values.Add
        let cellCount = values.Count
        let cells = resizeArrayToList values
        if cellCount <= viewportGeometryCacheMaxTotalItems then
            let key = viewportGeometryKey viewport
            let entries = System.Collections.Generic.Dictionary<ViewportGeometryKey, CellGeometry list>()
            let order = System.Collections.Generic.Queue<ViewportGeometryKey>()
            let sourceItemCount, sourceTextCharCount = viewportGeometrySourceMetrics model
            let cache: VisibleCellsCache = {
                Model = model
                SourceItemCount = sourceItemCount
                SourceTextCharCount = sourceTextCharCount
                Entries = entries
                Order = order
                EntryFastSlots = Array.zeroCreate<ViewportGeometryFastEntry<CellGeometry list>> 128
                TotalItemCount = 0
            }
            addVisibleCellsCacheEntryWithCount key cells cellCount cache
            rememberVisibleCellsCache cache
        cells

let iterHeaders (model: SpreadsheetModel) (viewport: Viewport) (emit: HeaderGeometry -> unit) =
    let range = visibleRange model viewport
    emit (CornerHeader {
            X = 0.0
            Y = 0.0
            Width = viewport.RowHeaderWidth
            Height = viewport.HeaderHeight
        })
    let mutable x = columnX model viewport range.StartCol
    for col in range.StartCol .. range.EndColExclusive - 1 do
        let width = getColumnWidth col model
        if width > 0.0 then
            emit (ColumnHeader (col, {
                X = x
                Y = 0.0
                Width = width
                Height = viewport.HeaderHeight
            }))
        x <- x + width

    let mutable row = nextVisibleRowAtOrAfter model range.StartRow
    let mutable y = rowY model viewport row
    while row < range.EndRowExclusive do
        let height = getRowHeight row model
        if height > 0.0 then
            emit (RowHeader (row, {
                X = 0.0
                Y = y
                Width = viewport.RowHeaderWidth
                Height = height
            }))
        y <- y + height
        row <- nextVisibleRowAtOrAfter model (row + 1)

let headers (model: SpreadsheetModel) (viewport: Viewport): HeaderGeometry list =
    match tryGetHeadersCache model with
    | ValueSome cached when headersCacheRetainable cached ->
        match tryGetViewportGeometryFastSlot cached.EntryFastSlots viewport with
        | ValueSome headers -> headers
        | ValueNone ->
            let key = viewportGeometryKey viewport
            let mutable headers = Unchecked.defaultof<HeaderGeometry list>
            if cached.Entries.TryGetValue(key, &headers) then
                setViewportGeometryFastSlot cached.EntryFastSlots key headers
                headers
            else
                let values = System.Collections.Generic.List<HeaderGeometry>()
                iterHeaders model viewport values.Add
                let headerCount = values.Count
                let headers = resizeArrayToList values
                addHeadersCacheEntryWithCount key headers headerCount cached
                rememberHeadersCache cached
                headers
    | ValueSome cached ->
        rememberHeadersCache cached
        let values = System.Collections.Generic.List<HeaderGeometry>()
        iterHeaders model viewport values.Add
        resizeArrayToList values
    | ValueNone ->
        let values = System.Collections.Generic.List<HeaderGeometry>()
        iterHeaders model viewport values.Add
        let headerCount = values.Count
        let headers = resizeArrayToList values
        if headerCount <= viewportGeometryCacheMaxTotalItems then
            let key = viewportGeometryKey viewport
            let entries = System.Collections.Generic.Dictionary<ViewportGeometryKey, HeaderGeometry list>()
            let order = System.Collections.Generic.Queue<ViewportGeometryKey>()
            let sourceItemCount, sourceTextCharCount = viewportGeometrySourceMetrics model
            let cache: HeadersCache = {
                Model = model
                SourceItemCount = sourceItemCount
                SourceTextCharCount = sourceTextCharCount
                Entries = entries
                Order = order
                EntryFastSlots = Array.zeroCreate<ViewportGeometryFastEntry<HeaderGeometry list>> 128
                TotalItemCount = 0
            }
            addHeadersCacheEntryWithCount key headers headerCount cache
            rememberHeadersCache cache
        headers

let contains (rect: Rect) x y =
    x >= rect.X && x < rect.X + rect.Width &&
    y >= rect.Y && y < rect.Y + rect.Height

let translateRect dx dy (rect: Rect): Rect =
    { rect with X = rect.X + dx; Y = rect.Y + dy }

let toViewportRect (viewport: Viewport) zone (rect: Rect): Rect =
    match zone with
    | BodyZone -> rect |> translateRect -viewport.ScrollX -viewport.ScrollY
    | ColumnHeaderZone -> rect |> translateRect -viewport.ScrollX 0.0
    | RowHeaderZone -> rect |> translateRect 0.0 -viewport.ScrollY
    | CornerZone -> rect
    // Frozen zones: pinned axes don't translate, free axis tracks scroll.
    | FrozenRowsZone -> rect |> translateRect -viewport.ScrollX 0.0
    | FrozenColsZone -> rect |> translateRect 0.0 -viewport.ScrollY
    | FrozenCornerZone -> rect

/// Width in CSS px of the frozen-cols band (zero when FrozenCols = 0).
let frozenColsWidth (model: SpreadsheetModel) : float =
    let mutable width = 0.0
    for c in 0 .. model.FrozenCols - 1 do
        width <- width + getColumnWidth c model
    width

/// Height in CSS px of the frozen-rows band (zero when FrozenRows = 0).
let frozenRowsHeight (model: SpreadsheetModel) : float =
    let mutable height = 0.0
    for r in 0 .. model.FrozenRows - 1 do
        height <- height + getRowHeight r model
    height

/// X coordinate (in viewport space) where the freeze split sits — that is,
/// the right edge of the frozen-cols band / start of horizontally-scrolling
/// body. Returns None when there are no frozen cols.
let frozenSplitX (model: SpreadsheetModel) (viewport: Viewport) : float option =
    if model.FrozenCols <= 0 then None
    else Some (viewport.RowHeaderWidth + frozenColsWidth model)

/// Y coordinate where the freeze split sits — bottom edge of frozen-rows.
let frozenSplitY (model: SpreadsheetModel) (viewport: Viewport) : float option =
    if model.FrozenRows <= 0 then None
    else Some (viewport.HeaderHeight + frozenRowsHeight model)

/// Decide which zone a body cell belongs to based on its row/col vs the
/// freeze split. Frozen-cell zones always paint (regardless of scroll
/// position) because they sit on top of the scrolling body.
let zoneForCell (model: SpreadsheetModel) (row: int) (col: int) : ViewportZone =
    let inFrozenRows = model.FrozenRows > 0 && row < model.FrozenRows
    let inFrozenCols = model.FrozenCols > 0 && col < model.FrozenCols
    match inFrozenRows, inFrozenCols with
    | true, true -> FrozenCornerZone
    | true, false -> FrozenRowsZone
    | false, true -> FrozenColsZone
    | false, false -> BodyZone

/// Clip rect for a zone. Callers pass the model in so frozen-pane zones
/// can derive their split positions from FrozenRows/FrozenCols.
let viewportClipFor (model: SpreadsheetModel) (viewport: Viewport) zone =
    let fx = frozenColsWidth model
    let fy = frozenRowsHeight model
    let bodyX = viewport.RowHeaderWidth + fx
    let bodyY = viewport.HeaderHeight + fy
    match zone with
    | BodyZone ->
        { X = bodyX
          Y = bodyY
          Width = max 0.0 (viewport.Width - bodyX)
          Height = max 0.0 (viewport.Height - bodyY) }
    | ColumnHeaderZone ->
        { X = viewport.RowHeaderWidth
          Y = 0.0
          Width = max 0.0 (viewport.Width - viewport.RowHeaderWidth)
          Height = viewport.HeaderHeight }
    | RowHeaderZone ->
        { X = 0.0
          Y = viewport.HeaderHeight
          Width = viewport.RowHeaderWidth
          Height = max 0.0 (viewport.Height - viewport.HeaderHeight) }
    | CornerZone ->
        { X = 0.0
          Y = 0.0
          Width = viewport.RowHeaderWidth
          Height = viewport.HeaderHeight }
    | FrozenRowsZone ->
        { X = bodyX
          Y = viewport.HeaderHeight
          Width = max 0.0 (viewport.Width - bodyX)
          Height = fy }
    | FrozenColsZone ->
        { X = viewport.RowHeaderWidth
          Y = bodyY
          Width = fx
          Height = max 0.0 (viewport.Height - bodyY) }
    | FrozenCornerZone ->
        { X = viewport.RowHeaderWidth
          Y = viewport.HeaderHeight
          Width = fx
          Height = fy }

/// Backward-compat shim: callers that don't have a model still get the
/// pre-freeze clip rectangles. Internal callers should prefer
/// `viewportClipFor` so frozen panes clip correctly.
let viewportClip (viewport: Viewport) zone =
    match zone with
    | BodyZone | FrozenRowsZone | FrozenColsZone | FrozenCornerZone ->
        { X = viewport.RowHeaderWidth
          Y = viewport.HeaderHeight
          Width = max 0.0 (viewport.Width - viewport.RowHeaderWidth)
          Height = max 0.0 (viewport.Height - viewport.HeaderHeight) }
    | ColumnHeaderZone ->
        { X = viewport.RowHeaderWidth
          Y = 0.0
          Width = max 0.0 (viewport.Width - viewport.RowHeaderWidth)
          Height = viewport.HeaderHeight }
    | RowHeaderZone ->
        { X = 0.0
          Y = viewport.HeaderHeight
          Width = viewport.RowHeaderWidth
          Height = max 0.0 (viewport.Height - viewport.HeaderHeight) }
    | CornerZone ->
        { X = 0.0
          Y = 0.0
          Width = viewport.RowHeaderWidth
          Height = viewport.HeaderHeight }

let intersectRect (a: Rect) (b: Rect): Rect option =
    let x1 = max a.X b.X
    let y1 = max a.Y b.Y
    let x2 = min (a.X + a.Width) (b.X + b.Width)
    let y2 = min (a.Y + a.Height) (b.Y + b.Height)
    if x2 <= x1 || y2 <= y1 then
        None
    else
        Some { X = x1; Y = y1; Width = x2 - x1; Height = y2 - y1 }

let private columnAtOffset (model: SpreadsheetModel) offset =
    let col = indexAtOffset defaultColumnWidth (columnLeft model) model.ColCount (fun c -> getColumnWidth c model) offset
    if col < 0 || col >= model.ColCount then None else
    let left = columnLeft model col
    if offset >= left && offset < left + getColumnWidth col model then Some col else None

let private rowAtOffset (model: SpreadsheetModel) offset =
    let row = indexAtOffset defaultRowHeight (rowTop model) model.RowCount (fun r -> getRowHeight r model) offset
    if row < 0 || row >= model.RowCount then None else
    let top = rowTop model row
    if offset >= top && offset < top + getRowHeight row model then Some row else None

let private cellHit model row col =
    let addr = { Row = row; Col = col }
    match findMergedCell addr model with
    | Some merge -> Cell { Row = merge.StartRow; Col = merge.StartCol }
    | None -> Cell addr

let hitTest (model: SpreadsheetModel) (viewport: Viewport) x y =
    if x < 0.0 || y < 0.0 || x >= viewport.Width || y >= viewport.Height then
        Empty
    elif contains { X = 0.0; Y = 0.0; Width = viewport.RowHeaderWidth; Height = viewport.HeaderHeight } x y then
        Corner
    elif y < viewport.HeaderHeight then
        columnAtOffset model (x + viewport.ScrollX - viewport.RowHeaderWidth)
        |> Option.map HitColumnHeader
        |> Option.defaultValue Empty
    elif x < viewport.RowHeaderWidth then
        rowAtOffset model (y + viewport.ScrollY - viewport.HeaderHeight)
        |> Option.map HitRowHeader
        |> Option.defaultValue Empty
    else
        let frozenW = frozenColsWidth model
        let frozenH = frozenRowsHeight model
        let bodyX = viewport.RowHeaderWidth + frozenW
        let bodyY = viewport.HeaderHeight + frozenH

        let offsets =
            if frozenW > 0.0 && frozenH > 0.0 && x < bodyX && y < bodyY then
                Some (x - viewport.RowHeaderWidth, y - viewport.HeaderHeight)
            elif frozenH > 0.0 && y < bodyY && x >= bodyX then
                Some (x + viewport.ScrollX - viewport.RowHeaderWidth, y - viewport.HeaderHeight)
            elif frozenW > 0.0 && x < bodyX && y >= bodyY then
                Some (x - viewport.RowHeaderWidth, y + viewport.ScrollY - viewport.HeaderHeight)
            elif x >= bodyX && y >= bodyY then
                Some (x + viewport.ScrollX - viewport.RowHeaderWidth, y + viewport.ScrollY - viewport.HeaderHeight)
            else
                None

        match offsets with
        | Some (contentX, contentY) ->
            match rowAtOffset model contentY, columnAtOffset model contentX with
            | Some row, Some col -> cellHit model row col
            | _ -> Empty
        | None -> Empty
