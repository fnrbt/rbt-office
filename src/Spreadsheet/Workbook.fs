module RedBat.Spreadsheet.Workbook

open System.Collections.Generic
open RedBat.Spreadsheet.Model

/// Sheet ID (unique identifier)
type SheetId = string

/// A cell addressed absolutely within a workbook — sheet + cell. The unit
/// the cross-sheet dependency graph is keyed on.
type SheetCell = {
    Sheet: SheetId
    Cell: CellAddress
}

/// Sheet in a workbook
type Sheet = {
    Id: SheetId
    Name: string
    Model: SpreadsheetModel
    Visible: bool
}

/// Table definition (structured reference). Coordinates are 0-indexed model
/// rows/cols (the form `getTableColumn`/`expandTableSel` read directly).
/// `HasTotals` marks a totals row occupying `EndRow` (Excel's `totalsRowCount`).
type Table = {
    Name: string
    SheetId: SheetId
    StartRow: int
    StartCol: int
    EndRow: int
    EndCol: int
    Columns: string list  // Column names (header order, left → right)
    HasHeaders: bool
    HasTotals: bool
}

/// Workbook containing multiple sheets
type Workbook = {
    Sheets: Map<SheetId, Sheet>
    SheetOrder: SheetId list  // Tab order
    ActiveSheet: SheetId
    Tables: Map<string, Table>  // Named tables
    /// Defined names → their refers-to formula text (e.g. "Sheet1!$A$1:$B$10").
    /// Keyed by (scope, UPPER-cased name): scope is the owning SheetId for
    /// sheet-local names, None for workbook-global names. A sheet-local name
    /// shadows a global one of the same name when evaluating on that sheet.
    Names: Map<SheetId option * string, string>
}

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private clearArray (items: 'T array) =
    for i = 0 to items.Length - 1 do
        items.[i] <- Unchecked.defaultof<'T>

let private referenceHash (value: obj) =
    System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value)

let private appendItemWithCapacity capacity item items =
    let values = ResizeArray(max 1 capacity)
    for value in items do
        values.Add value
    values.Add item
    resizeArrayToList values

let private filterOutWithCapacity capacity item items =
    let values = ResizeArray(max 0 capacity)
    for value in items do
        if value <> item then
            values.Add value
    resizeArrayToList values

type private SheetNamesCacheEntry = {
    Sheets: Map<SheetId, Sheet>
    SheetOrder: SheetId list
    NameCount: int
    NameCharCount: int
    SourceItemCount: int
    SourceTextCharCount: int
    Names: (SheetId * string) list
}

type private TableNamesCacheEntry = {
    Tables: Map<string, Table>
    NameCount: int
    NameCharCount: int
    SourceItemCount: int
    Names: string list
}

[<Struct>]
type private SheetNamesReferenceKey = {
    Sheets: Map<SheetId, Sheet>
    SheetOrder: SheetId list
}

let mutable private sheetNamesCache: SheetNamesCacheEntry list = []
let mutable private tableNamesCache: TableNamesCacheEntry list = []
let private sheetNamesCacheMaxEntries = 64
let private tableNamesCacheMaxEntries = 64
let private sheetNamesCacheMaxTotalNames = 200000
let private tableNamesCacheMaxTotalNames = 200000
let private sheetNamesCacheMaxTotalNameChars = 5000000
let private tableNamesCacheMaxTotalNameChars = 5000000
let private sheetNamesCacheMaxTotalSourceItems = 500000
let private sheetNamesCacheMaxTotalSourceTextChars = 5000000
let private tableNamesCacheMaxTotalSourceItems = 500000

let private addCapped limit value total =
    if total > limit then
        total
    else
        let value = max 0 value
        let remaining = limit - total
        if value > remaining then limit + 1
        else total + value

let private boundedListCount limit (items: 'T list) =
    let mutable count = 0
    let mutable remaining = items
    while count <= limit && not remaining.IsEmpty do
        count <- count + 1
        remaining <- remaining.Tail
    count

let private optionItemCount = function
    | Some _ -> 1
    | None -> 0

let private sheetNamesCacheEntryMatches sheets sheetOrder (entry: SheetNamesCacheEntry) =
    obj.ReferenceEquals(entry.Sheets, sheets)
    && obj.ReferenceEquals(entry.SheetOrder, sheetOrder)

let private tableNamesCacheEntryMatches tables (entry: TableNamesCacheEntry) =
    obj.ReferenceEquals(entry.Tables, tables)

let private sheetNamesReferenceHash sheets sheetOrder =
    (referenceHash sheets * 397) ^^^ referenceHash sheetOrder

let private sheetNamesReferenceKey sheets sheetOrder : SheetNamesReferenceKey =
    { Sheets = sheets
      SheetOrder = sheetOrder }

let private sheetNamesReferenceComparer =
    { new IEqualityComparer<SheetNamesReferenceKey> with
        member _.Equals(left, right) =
            obj.ReferenceEquals(left.Sheets, right.Sheets)
            && obj.ReferenceEquals(left.SheetOrder, right.SheetOrder)

        member _.GetHashCode(key) =
            sheetNamesReferenceHash key.Sheets key.SheetOrder }

let private tableNamesReferenceComparer =
    { new IEqualityComparer<Map<string, Table>> with
        member _.Equals(left, right) =
            obj.ReferenceEquals(left, right)

        member _.GetHashCode(value) =
            referenceHash value }

let private sheetNamesCacheBySource =
    Dictionary<SheetNamesReferenceKey, SheetNamesCacheEntry>(sheetNamesReferenceComparer)
let private sheetNamesFastSlots = Array.zeroCreate<SheetNamesCacheEntry> 256
let private tableNamesCacheBySource =
    Dictionary<Map<string, Table>, TableNamesCacheEntry>(tableNamesReferenceComparer)
let private tableNamesFastSlots = Array.zeroCreate<TableNamesCacheEntry> 256

let private setSheetNamesCache entries =
    sheetNamesCache <- entries
    sheetNamesCacheBySource.Clear()
    clearArray sheetNamesFastSlots
    for entry in entries do
        let key = sheetNamesReferenceKey entry.Sheets entry.SheetOrder
        let mutable existing = Unchecked.defaultof<SheetNamesCacheEntry>
        if not (sheetNamesCacheBySource.TryGetValue(key, &existing)) then
            sheetNamesCacheBySource.[key] <- entry
            sheetNamesFastSlots.[sheetNamesReferenceHash entry.Sheets entry.SheetOrder &&& (sheetNamesFastSlots.Length - 1)] <- entry

let private setTableNamesCache entries =
    tableNamesCache <- entries
    tableNamesCacheBySource.Clear()
    clearArray tableNamesFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<TableNamesCacheEntry>
        if not (tableNamesCacheBySource.TryGetValue(entry.Tables, &existing)) then
            tableNamesCacheBySource.[entry.Tables] <- entry
            tableNamesFastSlots.[referenceHash entry.Tables &&& (tableNamesFastSlots.Length - 1)] <- entry

let private tryGetSheetNamesCache sheets sheetOrder =
    let slot = sheetNamesReferenceHash sheets sheetOrder &&& (sheetNamesFastSlots.Length - 1)
    let fast = sheetNamesFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && sheetNamesCacheEntryMatches sheets sheetOrder fast then
        ValueSome fast
    else
        let key = sheetNamesReferenceKey sheets sheetOrder
        let mutable cached = Unchecked.defaultof<SheetNamesCacheEntry>
        if sheetNamesCacheBySource.TryGetValue(key, &cached) then
            sheetNamesFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private tryGetTableNamesCache tables =
    let slot = referenceHash tables &&& (tableNamesFastSlots.Length - 1)
    let fast = tableNamesFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && tableNamesCacheEntryMatches tables fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<TableNamesCacheEntry>
        if tableNamesCacheBySource.TryGetValue(tables, &cached) then
            tableNamesFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private sheetNamesCacheWithoutMatching sheets sheetOrder =
    let key = sheetNamesReferenceKey sheets sheetOrder
    let mutable cached = Unchecked.defaultof<SheetNamesCacheEntry>
    if sheetNamesCacheBySource.TryGetValue(key, &cached) then
        sheetNamesCache
        |> List.filter (sheetNamesCacheEntryMatches sheets sheetOrder >> not)
    else
        sheetNamesCache

let private tableNamesCacheWithoutMatching tables =
    let mutable cached = Unchecked.defaultof<TableNamesCacheEntry>
    if tableNamesCacheBySource.TryGetValue(tables, &cached) then
        tableNamesCache
        |> List.filter (tableNamesCacheEntryMatches tables >> not)
    else
        tableNamesCache

let private tableNamesCacheEntryRetainable (entry: TableNamesCacheEntry) =
    entry.NameCount <= tableNamesCacheMaxTotalNames
    && entry.NameCharCount <= tableNamesCacheMaxTotalNameChars
    && entry.SourceItemCount <= tableNamesCacheMaxTotalSourceItems

let private trimTableNamesCache (entries: TableNamesCacheEntry list) =
    let rec loop keptCount keptNames keptNameChars keptSourceItems acc remaining =
        match remaining with
        | [] -> List.rev acc
        | candidate :: tail ->
            let candidateNames = max 1 candidate.NameCount
            let candidateNameChars = max 1 candidate.NameCharCount
            let candidateSourceItems = max 1 candidate.SourceItemCount
            if tableNamesCacheEntryRetainable candidate
               && keptCount < tableNamesCacheMaxEntries
               && keptNames + candidateNames <= tableNamesCacheMaxTotalNames
               && keptNameChars + candidateNameChars <= tableNamesCacheMaxTotalNameChars
               && keptSourceItems + candidateSourceItems <= tableNamesCacheMaxTotalSourceItems then
                loop
                    (keptCount + 1)
                    (keptNames + candidateNames)
                    (keptNameChars + candidateNameChars)
                    (keptSourceItems + candidateSourceItems)
                    (candidate :: acc)
                    tail
            else
                loop keptCount keptNames keptNameChars keptSourceItems acc tail
    loop 0 0 0 0 [] entries

let private optionStringCharCount (value: string option) =
    match value with
    | Some value -> value.Length
    | None -> 0

let private runOverrideStringCharCount (run: RunOverride) =
    optionStringCharCount run.Color
    + optionStringCharCount run.FontName

let private tableStyleStringCharCount (style: TableStyle option) =
    match style with
    | Some style ->
        style.Name.Length
        + optionStringCharCount style.HeaderFill
        + optionStringCharCount style.StripeFill
    | None -> 0

let private conditionalRuleStringCharCount = function
    | ConditionalFormatRule.ContainsText text -> text.Length
    | ConditionalFormatRule.GreaterThan _
    | ConditionalFormatRule.LessThan _
    | ConditionalFormatRule.Between _
    | ConditionalFormatRule.Equal _ -> 0

let private validationRuleStringCharCount = function
    | ValidationRule.ListValues values ->
        let mutable total = 0
        let mutable seen = 0
        let mutable remaining = values
        while total <= sheetNamesCacheMaxTotalSourceTextChars
              && seen <= sheetNamesCacheMaxTotalSourceItems
              && not remaining.IsEmpty do
            total <- addCapped sheetNamesCacheMaxTotalSourceTextChars remaining.Head.Length total
            seen <- seen + 1
            remaining <- remaining.Tail
        total
    | ValidationRule.WholeNumber _
    | ValidationRule.DecimalRange _
    | ValidationRule.TextLength _
    | ValidationRule.DateRange _ -> 0

let private headerFooterStringCharCount (headerFooter: HeaderFooter) =
    optionStringCharCount headerFooter.OddHeader
    + optionStringCharCount headerFooter.OddFooter
    + optionStringCharCount headerFooter.EvenHeader
    + optionStringCharCount headerFooter.EvenFooter
    + optionStringCharCount headerFooter.FirstHeader
    + optionStringCharCount headerFooter.FirstFooter

let private validationRuleItemCount = function
    | ValidationRule.ListValues values ->
        let mutable total = 0
        let mutable remaining = values
        while total <= sheetNamesCacheMaxTotalSourceItems && not remaining.IsEmpty do
            total <- addCapped sheetNamesCacheMaxTotalSourceItems 1 total
            remaining <- remaining.Tail
        total
    | ValidationRule.WholeNumber _
    | ValidationRule.DecimalRange _
    | ValidationRule.TextLength _
    | ValidationRule.DateRange _ -> 0

let private validationItemCount (validations: Map<CellAddress, DataValidation>) =
    if validations.Count = 0 then 0
    else
        let mutable total = 0
        use entries = (validations :> seq<_>).GetEnumerator()
        while total <= sheetNamesCacheMaxTotalSourceItems && entries.MoveNext() do
            let validation = entries.Current.Value
            total <- addCapped sheetNamesCacheMaxTotalSourceItems 1 total
            total <- addCapped sheetNamesCacheMaxTotalSourceItems (validationRuleItemCount validation.Rule) total
        total

let private pageSetupItemCount (pageSetup: PageSetup) =
    let limit = sheetNamesCacheMaxTotalSourceItems
    let mutable total = boundedListCount limit pageSetup.ManualRowBreaks
    if total <= limit then
        total <- addCapped limit (boundedListCount limit pageSetup.ManualColBreaks) total
    if total <= limit then
        total <- addCapped limit (optionItemCount pageSetup.PrintArea) total
    if total <= limit then
        total <- addCapped limit (optionItemCount pageSetup.RepeatRows) total
    if total <= limit then
        total <- addCapped limit (optionItemCount pageSetup.RepeatCols) total
    total

let private filterItemCount (filters: Map<int, ColumnFilter>) =
    if filters.Count = 0 then 0
    else
        let mutable total = 0
        use entries = (filters :> seq<_>).GetEnumerator()
        while total <= sheetNamesCacheMaxTotalSourceItems && entries.MoveNext() do
            let filter = entries.Current.Value
            total <- addCapped sheetNamesCacheMaxTotalSourceItems 1 total
            total <- addCapped sheetNamesCacheMaxTotalSourceItems filter.VisibleValues.Count total
        total

let private conditionalFormatItemCount (formats: Map<CellAddress, ConditionalFormat list>) =
    if formats.Count = 0 then 0
    else
        let mutable total = 0
        use entries = (formats :> seq<_>).GetEnumerator()
        while total <= sheetNamesCacheMaxTotalSourceItems && entries.MoveNext() do
            let items = entries.Current.Value
            total <- addCapped sheetNamesCacheMaxTotalSourceItems 1 total
            total <- addCapped sheetNamesCacheMaxTotalSourceItems (boundedListCount sheetNamesCacheMaxTotalSourceItems items) total
        total

let private richRunItemCount (runs: Map<CellAddress, (string * RunOverride) list>) =
    if runs.Count = 0 then 0
    else
        let mutable total = 0
        use entries = (runs :> seq<_>).GetEnumerator()
        while total <= sheetNamesCacheMaxTotalSourceItems && entries.MoveNext() do
            let items = entries.Current.Value
            total <- addCapped sheetNamesCacheMaxTotalSourceItems 1 total
            total <- addCapped sheetNamesCacheMaxTotalSourceItems (boundedListCount sheetNamesCacheMaxTotalSourceItems items) total
        total

let private modelHistoryItemCount (history: History) =
    let mutable total = 0
    let mutable past = history.Past
    while total <= sheetNamesCacheMaxTotalSourceItems && not past.IsEmpty do
        let command = past.Head
        total <- addCapped sheetNamesCacheMaxTotalSourceItems (max 1 command.RetainedItemCount) total
        past <- past.Tail
    let mutable future = history.Future
    while total <= sheetNamesCacheMaxTotalSourceItems && not future.IsEmpty do
        let command = future.Head
        total <- addCapped sheetNamesCacheMaxTotalSourceItems (max 1 command.RetainedItemCount) total
        future <- future.Tail
    total

let private modelHistoryTextCharCount (history: History) =
    let mutable total = 0
    let mutable past = history.Past
    while total <= sheetNamesCacheMaxTotalSourceTextChars && not past.IsEmpty do
        let command = past.Head
        total <- addCapped sheetNamesCacheMaxTotalSourceTextChars (max 1 command.RetainedTextCharCount) total
        past <- past.Tail
    let mutable future = history.Future
    while total <= sheetNamesCacheMaxTotalSourceTextChars && not future.IsEmpty do
        let command = future.Head
        total <- addCapped sheetNamesCacheMaxTotalSourceTextChars (max 1 command.RetainedTextCharCount) total
        future <- future.Tail
    total

let private clipboardItemCount (clipboard: ClipboardData option) =
    match clipboard with
    | Some clipboard ->
        let mutable total = 0
        let mutable remaining = clipboard.Cells
        while total <= sheetNamesCacheMaxTotalSourceItems && not remaining.IsEmpty do
            total <- addCapped sheetNamesCacheMaxTotalSourceItems 1 total
            remaining <- remaining.Tail
        total
    | None -> 0

let private clipboardTextCharCount (clipboard: ClipboardData option) =
    match clipboard with
    | Some clipboard ->
        let mutable total = 0
        let mutable remaining = clipboard.Cells
        while total <= sheetNamesCacheMaxTotalSourceTextChars && not remaining.IsEmpty do
            let _, _, cell = remaining.Head
            total <- addCapped sheetNamesCacheMaxTotalSourceTextChars (cellDataTextCharCountFor sheetNamesCacheMaxTotalSourceTextChars cell) total
            remaining <- remaining.Tail
        total
    | None -> 0

let private modelCellMapTextCharCount (cells: Map<CellAddress, CellData>) =
    let mutable total = 0
    cells
    |> Map.exists (fun _ cell ->
        total <- addCapped sheetNamesCacheMaxTotalSourceTextChars (cellDataTextCharCountFor sheetNamesCacheMaxTotalSourceTextChars cell) total
        total > sheetNamesCacheMaxTotalSourceTextChars)
    |> ignore
    total

let private modelSourceItemCount (model: SpreadsheetModel) =
    let limit = sheetNamesCacheMaxTotalSourceItems
    let mutable total = 0
    if total <= limit then total <- addCapped limit model.Cells.Count total
    if total <= limit then total <- addCapped limit (clipboardItemCount model.Clipboard) total
    if total <= limit then total <- addCapped limit (modelHistoryItemCount model.History) total
    if total <= limit then total <- addCapped limit (filterItemCount model.Filters) total
    if total <= limit then total <- addCapped limit model.FilterHiddenRows.Length total
    if total <= limit then total <- addCapped limit model.RowHeights.Count total
    if total <= limit then total <- addCapped limit model.ColumnWidths.Count total
    if total <= limit then total <- addCapped limit model.SearchResults.Length total
    if total <= limit then total <- addCapped limit (conditionalFormatItemCount model.ConditionalFormats) total
    if total <= limit then total <- addCapped limit (boundedListCount limit model.MergedCells) total
    if total <= limit then total <- addCapped limit (richRunItemCount model.RichRuns) total
    if total <= limit then total <- addCapped limit (boundedListCount limit model.Tables) total
    if total <= limit then total <- addCapped limit model.CfBoldCells.Count total
    if total <= limit then total <- addCapped limit model.Spill.Count total
    if total <= limit then total <- addCapped limit (validationItemCount model.Validations) total
    if total <= limit then total <- addCapped limit (pageSetupItemCount model.PageSetup) total
    total

let private cellFormatTextCharCount (format: CellFormat) =
    cellDataTextCharCountFor sheetNamesCacheMaxTotalSourceTextChars { Value = Empty; Format = format; Comment = None }

let private modelSourceTextCharCount (model: SpreadsheetModel) =
    let limit = sheetNamesCacheMaxTotalSourceTextChars
    let mutable total =
        optionStringCharCount model.SearchQuery
        + model.EditBuffer.Length
        + model.Macros.Length
        + headerFooterStringCharCount model.PageSetup.HeaderFooter
    if total <= limit then
        total <- addCapped limit (clipboardTextCharCount model.Clipboard) total
    if total <= limit then
        total <- addCapped limit (modelHistoryTextCharCount model.History) total
    if total <= limit && model.Filters.Count > 0 then
        use filters = (model.Filters :> seq<_>).GetEnumerator()
        while total <= limit && filters.MoveNext() do
            let filter = filters.Current.Value
            use values = (filter.VisibleValues :> seq<string>).GetEnumerator()
            while total <= limit && values.MoveNext() do
                total <- addCapped limit values.Current.Length total
    if total <= limit && model.ConditionalFormats.Count > 0 then
        use conditionalFormats = (model.ConditionalFormats :> seq<_>).GetEnumerator()
        while total <= limit && conditionalFormats.MoveNext() do
            let mutable formats = conditionalFormats.Current.Value
            while total <= limit && not formats.IsEmpty do
                let format = formats.Head
                total <- addCapped limit (conditionalRuleStringCharCount format.Rule) total
                total <- addCapped limit (cellFormatTextCharCount format.Style) total
                formats <- formats.Tail
    if total <= limit && model.RichRuns.Count > 0 then
        use richRuns = (model.RichRuns :> seq<_>).GetEnumerator()
        while total <= limit && richRuns.MoveNext() do
            let mutable runs = richRuns.Current.Value
            while total <= limit && not runs.IsEmpty do
                let text, run = runs.Head
                total <- addCapped limit text.Length total
                total <- addCapped limit (runOverrideStringCharCount run) total
                runs <- runs.Tail
    if total <= limit && not model.Tables.IsEmpty then
        let mutable tables = model.Tables
        let mutable seenTables = 0
        while total <= limit
              && seenTables <= sheetNamesCacheMaxTotalSourceItems
              && not tables.IsEmpty do
            let table = tables.Head
            total <- addCapped limit table.Name.Length total
            if total <= limit then
                total <- addCapped limit (tableStyleStringCharCount table.Style) total
            seenTables <- seenTables + 1
            tables <- tables.Tail
    if total <= limit && model.Validations.Count > 0 then
        use validations = (model.Validations :> seq<_>).GetEnumerator()
        while total <= limit && validations.MoveNext() do
            let validation = validations.Current.Value
            total <- addCapped limit (validationRuleStringCharCount validation.Rule) total
            total <- addCapped limit (optionStringCharCount validation.ErrorMessage) total
            total <- addCapped limit (optionStringCharCount validation.InputMessage) total
    if total <= limit && model.Cells.Count > 0 then
        total <- addCapped limit (modelCellMapTextCharCount model.Cells) total
    total

let private modelSourceTextCharCountFor sourceItemCount (model: SpreadsheetModel) =
    if sourceItemCount > sheetNamesCacheMaxTotalSourceItems then
        sheetNamesCacheMaxTotalSourceTextChars + 1
    else
        modelSourceTextCharCount model

let private sheetNamesCacheEntryRetainable (entry: SheetNamesCacheEntry) =
    entry.NameCount <= sheetNamesCacheMaxTotalNames
    && entry.NameCharCount <= sheetNamesCacheMaxTotalNameChars
    && entry.SourceItemCount <= sheetNamesCacheMaxTotalSourceItems
    && entry.SourceTextCharCount <= sheetNamesCacheMaxTotalSourceTextChars

let private trimSheetNamesCache (entries: SheetNamesCacheEntry list) =
    let rec loop
        keptCount
        keptNames
        keptNameChars
        keptSourceItems
        keptSourceTextChars
        (acc: SheetNamesCacheEntry list)
        (remaining: SheetNamesCacheEntry list)
        =
        match remaining with
        | [] -> List.rev acc
        | candidate :: tail ->
            let candidateNames = max 1 candidate.NameCount
            let candidateNameChars = max 1 candidate.NameCharCount
            let candidateSourceItems = max 1 candidate.SourceItemCount
            let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
            if sheetNamesCacheEntryRetainable candidate
               && keptCount < sheetNamesCacheMaxEntries
               && keptNames + candidateNames <= sheetNamesCacheMaxTotalNames
               && keptNameChars + candidateNameChars <= sheetNamesCacheMaxTotalNameChars
               && keptSourceItems + candidateSourceItems <= sheetNamesCacheMaxTotalSourceItems
               && keptSourceTextChars + candidateSourceTextChars <= sheetNamesCacheMaxTotalSourceTextChars then
                loop
                    (keptCount + 1)
                    (keptNames + candidateNames)
                    (keptNameChars + candidateNameChars)
                    (keptSourceItems + candidateSourceItems)
                    (keptSourceTextChars + candidateSourceTextChars)
                    (candidate :: acc)
                    tail
            else
                loop keptCount keptNames keptNameChars keptSourceItems keptSourceTextChars acc tail
    loop 0 0 0 0 0 [] entries

let private sheetOrderTextCharCount (sheetOrder: SheetId list) =
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = sheetOrder
    while total <= sheetNamesCacheMaxTotalSourceTextChars
          && seen <= sheetNamesCacheMaxTotalSourceItems
          && not remaining.IsEmpty do
        total <- addCapped sheetNamesCacheMaxTotalSourceTextChars remaining.Head.Length total
        seen <- seen + 1
        remaining <- remaining.Tail
    total

let private tableDefinitionStringCharCount (table: Table) =
    let mutable total = table.Name.Length + table.SheetId.Length
    let mutable columns = table.Columns
    while total <= tableNamesCacheMaxTotalNameChars && not columns.IsEmpty do
        total <- addCapped tableNamesCacheMaxTotalNameChars columns.Head.Length total
        columns <- columns.Tail
    total

let private tableDefinitionStringCharCountFor sourceItemCount (table: Table) =
    if sourceItemCount > tableNamesCacheMaxTotalSourceItems then
        tableNamesCacheMaxTotalNameChars + 1
    else
        tableDefinitionStringCharCount table

let private tableDefinitionItemCount (table: Table) =
    let mutable total = 1
    let mutable columns = table.Columns
    while total <= tableNamesCacheMaxTotalSourceItems && not columns.IsEmpty do
        total <- addCapped tableNamesCacheMaxTotalSourceItems 1 total
        columns <- columns.Tail
    total

let private rememberSheetNamesCacheEntryWithMetrics sheets sheetOrder sourceItemCount sourceTextCharCount nameCount nameCharCount (names: (SheetId * string) list) =
    let entry: SheetNamesCacheEntry = {
        Sheets = sheets
        SheetOrder = sheetOrder
        NameCount = nameCount
        NameCharCount =
            if nameCount > sheetNamesCacheMaxTotalNames then sheetNamesCacheMaxTotalNameChars + 1
            else nameCharCount
        SourceItemCount = sourceItemCount
        SourceTextCharCount = sourceTextCharCount
        Names = names
    }
    if not (sheetNamesCacheEntryRetainable entry) then
        ()
    else
        (entry :: sheetNamesCacheWithoutMatching sheets sheetOrder)
        |> trimSheetNamesCache
        |> setSheetNamesCache
    names

let private rememberTableNamesCacheEntryWithMetrics tables sourceItemCount nameCount nameCharCount (names: string list) =
    let entry = {
        Tables = tables
        NameCount = nameCount
        NameCharCount = nameCharCount
        SourceItemCount = sourceItemCount
        Names = names
    }
    if not (tableNamesCacheEntryRetainable entry) then
        ()
    else
        (entry :: tableNamesCacheWithoutMatching tables)
        |> trimTableNamesCache
        |> setTableNamesCache
    names

let private validSheetOrder (sheets: Map<SheetId, Sheet>) (order: SheetId list) =
    let expectedCount = Map.count sheets
    let seen = HashSet<SheetId>()
    let mutable count = 0
    let mutable valid = true
    let mutable remaining = order
    while valid && not remaining.IsEmpty do
        match remaining with
        | sheetId :: tail ->
            count <- count + 1
            if count > expectedCount || not (Map.containsKey sheetId sheets) || not (seen.Add sheetId) then
                valid <- false
            remaining <- tail
        | [] -> ()
    valid && count = expectedCount

/// Create empty workbook with one sheet
let createWorkbook (): Workbook =
    let sheetId = "sheet1"
    let sheet = {
        Id = sheetId
        Name = "Sheet1"
        Model = createEmpty 1000 26
        Visible = true
    }
    {
        Sheets = Map.ofList [(sheetId, sheet)]
        SheetOrder = [sheetId]
        ActiveSheet = sheetId
        Tables = Map.empty
        Names = Map.empty
    }

/// Look up a defined name's refers-to text, scoped to the current sheet: a
/// sheet-local name shadows a workbook-global one of the same name.
/// Case-insensitive.
let private nameKey (name: string) =
    let mutable index = 0
    let mutable needsUpper = false
    while not needsUpper && index < name.Length do
        if System.Char.ToUpperInvariant name.[index] <> name.[index] then
            needsUpper <- true
        index <- index + 1
    if needsUpper then name.ToUpperInvariant() else name

let getName (currentSheet: SheetId option) (name: string) (workbook: Workbook): string option =
    let key = nameKey name
    let local = currentSheet |> Option.bind (fun s -> Map.tryFind (Some s, key) workbook.Names)
    match local with
    | Some _ -> local
    | None -> Map.tryFind (None, key) workbook.Names

/// Build the Names map from raw (scope, name, refersTo) triples. Keys are
/// upper-cased for case-insensitive lookup; later entries win on collision.
let namesOfTriples (triples: (SheetId option * string * string) seq): Map<SheetId option * string, string> =
    triples
    |> Seq.fold (fun acc (scope, n, refersTo) -> Map.add (scope, nameKey n) refersTo acc) Map.empty

/// Convenience for workbook-global names only (scope = None).
let namesOfPairs (pairs: (string * string) seq): Map<SheetId option * string, string> =
    pairs |> Seq.map (fun (n, refersTo) -> None, n, refersTo) |> namesOfTriples

/// Get active sheet
let getActiveSheet (workbook: Workbook): Sheet option =
    Map.tryFind workbook.ActiveSheet workbook.Sheets

/// Add a new sheet
let addSheet (name: string) (workbook: Workbook): Workbook =
    let sheetId = sprintf "sheet%d" (Map.count workbook.Sheets + 1)
    let sheet = {
        Id = sheetId
        Name = name
        Model = createEmpty 1000 26
        Visible = true
    }
    {
        workbook with
            Sheets = Map.add sheetId sheet workbook.Sheets
            SheetOrder = appendItemWithCapacity (Map.count workbook.Sheets + 1) sheetId workbook.SheetOrder
    }

/// Delete a sheet
let deleteSheet (sheetId: SheetId) (workbook: Workbook): Workbook option =
    if Map.count workbook.Sheets <= 1 then
        // Can't delete last sheet
        None
    elif not (Map.containsKey sheetId workbook.Sheets) then
        Some workbook
    else
        let newSheets = Map.remove sheetId workbook.Sheets
        let newOrder = filterOutWithCapacity newSheets.Count sheetId workbook.SheetOrder
        let newActive =
            if workbook.ActiveSheet = sheetId then
                List.head newOrder
            else
                workbook.ActiveSheet
        Some {
            workbook with
                Sheets = newSheets
                SheetOrder = newOrder
                ActiveSheet = newActive
        }

/// Rename a sheet
let renameSheet (sheetId: SheetId) (newName: string) (workbook: Workbook): Workbook =
    let mutable sheet = Unchecked.defaultof<Sheet>
    if not (workbook.Sheets.TryGetValue(sheetId, &sheet)) then
        workbook
    elif sheet.Name = newName then
        workbook
    else
        let updatedSheet = { sheet with Name = newName }
        { workbook with Sheets = Map.add sheetId updatedSheet workbook.Sheets }

/// Switch active sheet
let setActiveSheet (sheetId: SheetId) (workbook: Workbook): Workbook =
    if workbook.ActiveSheet = sheetId then
        workbook
    elif Map.containsKey sheetId workbook.Sheets then
        { workbook with ActiveSheet = sheetId }
    else
        workbook

/// Reorder sheets
let reorderSheets (newOrder: SheetId list) (workbook: Workbook): Workbook =
    // Validate every sheet appears exactly once.
    if newOrder = workbook.SheetOrder then
        workbook
    elif validSheetOrder workbook.Sheets newOrder then
        { workbook with SheetOrder = newOrder }
    else
        workbook

/// Parse cross-sheet reference: Sheet2!A1 or 'Sheet Name'!B2
let parseSheetReference (ref: string): (string option * string) =
    let mutable bang = -1
    let mutable multiple = false
    let mutable i = 0
    while i < ref.Length && not multiple do
        if ref.[i] = '!' then
            if bang >= 0 then
                multiple <- true
            else
                bang <- i
        i <- i + 1
    if bang < 0 || multiple then
        (None, ref)
    else
        let sheetName =
            if bang >= 2 && ref.[0] = '\'' && ref.[bang - 1] = '\'' then
                ref.Substring(1, bang - 2)
            else
                ref.Substring(0, bang)
        (Some sheetName, ref.Substring(bang + 1))

/// Resolve sheet name to sheet ID
let resolveSheetName (name: string) (workbook: Workbook): SheetId option =
    workbook.Sheets
    |> Map.tryPick (fun id sheet ->
        if sheet.Name.Equals(name, System.StringComparison.OrdinalIgnoreCase) then
            Some id
        else
            None
    )

/// Define a table (structured reference)
let defineTable (table: Table) (workbook: Workbook): Workbook =
    let mutable existing = Unchecked.defaultof<Table>
    if workbook.Tables.TryGetValue(table.Name, &existing)
        && (obj.ReferenceEquals(existing, table) || existing = table) then
        workbook
    else
        { workbook with Tables = Map.add table.Name table workbook.Tables }

/// Delete a table
let deleteTable (name: string) (workbook: Workbook): Workbook =
    if Map.containsKey name workbook.Tables then
        { workbook with Tables = Map.remove name workbook.Tables }
    else
        workbook

/// Get table by name (case-insensitive — Excel table names are).
let getTable (name: string) (workbook: Workbook): Table option =
    match Map.tryFind name workbook.Tables with
    | Some t -> Some t
    | None ->
        workbook.Tables
        |> Map.tryPick (fun k t ->
            if System.String.Equals(k, name, System.StringComparison.OrdinalIgnoreCase) then Some t
            else None)

/// Index of a column within a table by name (case-insensitive). None if absent.
let tableColumnIndex (table: Table) (columnName: string): int option =
    let mutable found = None
    let mutable index = 0
    let mutable remaining = table.Columns
    while found.IsNone && not remaining.IsEmpty do
        match remaining with
        | column :: tail ->
            if System.String.Equals(column, columnName, System.StringComparison.OrdinalIgnoreCase) then
                found <- Some index
            index <- index + 1
            remaining <- tail
        | [] -> ()
    found

/// The table on `sheet` whose body region contains `cell`, if any — used to
/// resolve unqualified structured refs (`[@Col]`) from the evaluating cell.
let findTableContaining (sheet: SheetId option) (cell: CellAddress) (workbook: Workbook): Table option =
    workbook.Tables
    |> Map.tryPick (fun _ t ->
        let sheetOk = match sheet with Some s -> s = t.SheetId | None -> true
        if sheetOk && cell.Row >= t.StartRow && cell.Row <= t.EndRow
           && cell.Col >= t.StartCol && cell.Col <= t.EndCol then
            Some t
        else
            None)

/// Build a structured-reference `Table` from a 1-indexed XLSX table region
/// (rows and cols 1-based, A=1 — the form `Xlsx.TableDefinition`/`Model.Table`
/// use) by shifting to the 0-indexed model coordinates this module stores.
/// `columns` are the declared `<tableColumn>` names; when empty they're read
/// left→right from the header row of `model`.
let makeTableFromRegion (sheetId: SheetId) (name: string)
                        (startRow1: int) (startCol1: int) (endRow1: int) (endCol1: int)
                        (headerRowCount: int) (totalsRowCount: int)
                        (columns: string list) (model: SpreadsheetModel): Table =
    let sr = startRow1 - 1
    let sc = startCol1 - 1
    let er = endRow1 - 1
    let ec = endCol1 - 1
    let cols =
        if not (List.isEmpty columns) then columns
        else
            let values = ResizeArray<string>(max 0 (ec - sc + 1))
            for c in sc .. ec do
                values.Add(displayValue (getCellValue { Row = sr; Col = c } model))
            resizeArrayToList values
    { Name = name
      SheetId = sheetId
      StartRow = sr
      StartCol = sc
      EndRow = er
      EndCol = ec
      Columns = cols
      HasHeaders = headerRowCount > 0
      HasTotals = totalsRowCount > 0 }

/// Resolve table column reference: Table1[Amount]
let parseTableReference (ref: string): (string * string) option =
    if ref.Contains("[") && ref.Contains("]") then
        let tablePart = ref.Substring(0, ref.IndexOf('['))
        let columnPart = ref.Substring(ref.IndexOf('[') + 1)
        let columnName = columnPart.Replace("]", "")
        Some (tablePart, columnName)
    else
        None

/// Get table column range
let getTableColumn (tableName: string) (columnName: string) (workbook: Workbook): (SheetId * CellAddress * CellAddress) option =
    match Map.tryFind tableName workbook.Tables with
    | Some table ->
        // Find column index
        let mutable foundIndex = None
        let mutable index = 0
        let mutable remaining = table.Columns
        while foundIndex.IsNone && not remaining.IsEmpty do
            match remaining with
            | column :: tail ->
                if column = columnName then
                    foundIndex <- Some index
                index <- index + 1
                remaining <- tail
            | [] -> ()
        match foundIndex with
        | Some colIndex ->
            let actualCol = table.StartCol + colIndex
            let startAddr = { Row = table.StartRow + (if table.HasHeaders then 1 else 0); Col = actualCol }
            let endAddr = { Row = table.EndRow; Col = actualCol }
            Some (table.SheetId, startAddr, endAddr)
        | None -> None
    | None -> None

/// Update sheet model
let updateSheetModel (sheetId: SheetId) (model: SpreadsheetModel) (workbook: Workbook): Workbook =
    let mutable sheet = Unchecked.defaultof<Sheet>
    if not (workbook.Sheets.TryGetValue(sheetId, &sheet)) then
        workbook
    elif obj.ReferenceEquals(sheet.Model, model) then
        workbook
    else
        let updatedSheet = { sheet with Model = model }
        { workbook with Sheets = Map.add sheetId updatedSheet workbook.Sheets }

/// Get all sheet names
let getSheetNames (workbook: Workbook): (SheetId * string) list =
    match tryGetSheetNamesCache workbook.Sheets workbook.SheetOrder with
    | ValueSome cached ->
        cached.Names
    | ValueNone ->
        let names = ResizeArray<SheetId * string>(workbook.SheetOrder.Length)
        let mutable nameCharCount = 0
        let nameCountOverLimit () = names.Count > sheetNamesCacheMaxTotalNames
        let initialSourceItemCount =
            addCapped
                sheetNamesCacheMaxTotalSourceItems
                (boundedListCount sheetNamesCacheMaxTotalSourceItems workbook.SheetOrder)
                workbook.Sheets.Count
        let initialSourceTextCharCount = sheetOrderTextCharCount workbook.SheetOrder
        let mutable sourceItemCount = max 1 initialSourceItemCount
        let mutable sourceTextCharCount = initialSourceTextCharCount
        for id in workbook.SheetOrder do
            match Map.tryFind id workbook.Sheets with
            | Some sheet ->
                names.Add(id, sheet.Name)
                if nameCharCount <= sheetNamesCacheMaxTotalNameChars then
                    nameCharCount <- addCapped sheetNamesCacheMaxTotalNameChars id.Length nameCharCount
                    if nameCharCount <= sheetNamesCacheMaxTotalNameChars then
                        nameCharCount <- addCapped sheetNamesCacheMaxTotalNameChars sheet.Name.Length nameCharCount
                if nameCountOverLimit () then
                    sourceItemCount <- sheetNamesCacheMaxTotalSourceItems + 1
                    sourceTextCharCount <- sheetNamesCacheMaxTotalSourceTextChars + 1
                if sourceTextCharCount <= sheetNamesCacheMaxTotalSourceTextChars then
                    sourceTextCharCount <- addCapped sheetNamesCacheMaxTotalSourceTextChars id.Length sourceTextCharCount
                    if sourceTextCharCount <= sheetNamesCacheMaxTotalSourceTextChars then
                        sourceTextCharCount <- addCapped sheetNamesCacheMaxTotalSourceTextChars sheet.Id.Length sourceTextCharCount
                    if sourceTextCharCount <= sheetNamesCacheMaxTotalSourceTextChars then
                        sourceTextCharCount <- addCapped sheetNamesCacheMaxTotalSourceTextChars sheet.Name.Length sourceTextCharCount
                    if sourceTextCharCount <= sheetNamesCacheMaxTotalSourceTextChars
                       && sourceItemCount <= sheetNamesCacheMaxTotalSourceItems then
                        sourceItemCount <- addCapped sheetNamesCacheMaxTotalSourceItems (modelSourceItemCount sheet.Model) sourceItemCount
                    if sourceTextCharCount <= sheetNamesCacheMaxTotalSourceTextChars then
                        sourceTextCharCount <-
                            addCapped
                                sheetNamesCacheMaxTotalSourceTextChars
                                (modelSourceTextCharCountFor sourceItemCount sheet.Model)
                                sourceTextCharCount
            | None -> ()
        if names.Count <> workbook.Sheets.Count then
            sourceItemCount <- max 1 initialSourceItemCount
            sourceTextCharCount <- initialSourceTextCharCount
            use sheets = (workbook.Sheets :> seq<_>).GetEnumerator()
            while sourceItemCount <= sheetNamesCacheMaxTotalSourceItems
                  && sourceTextCharCount <= sheetNamesCacheMaxTotalSourceTextChars
                  && sheets.MoveNext() do
                let id = sheets.Current.Key
                let sheet = sheets.Current.Value
                if sourceTextCharCount <= sheetNamesCacheMaxTotalSourceTextChars then
                    sourceTextCharCount <- addCapped sheetNamesCacheMaxTotalSourceTextChars id.Length sourceTextCharCount
                    if sourceTextCharCount <= sheetNamesCacheMaxTotalSourceTextChars then
                        sourceTextCharCount <- addCapped sheetNamesCacheMaxTotalSourceTextChars sheet.Id.Length sourceTextCharCount
                    if sourceTextCharCount <= sheetNamesCacheMaxTotalSourceTextChars then
                        sourceTextCharCount <- addCapped sheetNamesCacheMaxTotalSourceTextChars sheet.Name.Length sourceTextCharCount
                    if sourceTextCharCount <= sheetNamesCacheMaxTotalSourceTextChars
                       && sourceItemCount <= sheetNamesCacheMaxTotalSourceItems then
                        sourceItemCount <- addCapped sheetNamesCacheMaxTotalSourceItems (modelSourceItemCount sheet.Model) sourceItemCount
                    if sourceTextCharCount <= sheetNamesCacheMaxTotalSourceTextChars then
                        sourceTextCharCount <-
                            addCapped
                                sheetNamesCacheMaxTotalSourceTextChars
                                (modelSourceTextCharCountFor sourceItemCount sheet.Model)
                                sourceTextCharCount
        let nameCount = min names.Count (sheetNamesCacheMaxTotalNames + 1)
        rememberSheetNamesCacheEntryWithMetrics workbook.Sheets workbook.SheetOrder sourceItemCount sourceTextCharCount nameCount nameCharCount (resizeArrayToList names)

/// Get all table names
let getTableNames (workbook: Workbook): string list =
    match tryGetTableNamesCache workbook.Tables with
    | ValueSome cached ->
        cached.Names
    | ValueNone ->
        let names = ResizeArray<string>(workbook.Tables.Count)
        let mutable nameCharCount = 0
        let mutable sourceItemCount = 0
        let nameCountOverLimit = workbook.Tables.Count > tableNamesCacheMaxTotalNames
        if nameCountOverLimit then
            nameCharCount <- tableNamesCacheMaxTotalNameChars + 1
            sourceItemCount <- tableNamesCacheMaxTotalSourceItems + 1
        workbook.Tables
        |> Map.iter (fun name table ->
            names.Add name
            if nameCharCount <= tableNamesCacheMaxTotalNameChars then
                nameCharCount <- addCapped tableNamesCacheMaxTotalNameChars name.Length nameCharCount
            if nameCharCount <= tableNamesCacheMaxTotalNameChars
               && sourceItemCount <= tableNamesCacheMaxTotalSourceItems then
                sourceItemCount <- addCapped tableNamesCacheMaxTotalSourceItems (tableDefinitionItemCount table) sourceItemCount
            if nameCharCount <= tableNamesCacheMaxTotalNameChars then
                nameCharCount <-
                    addCapped
                        tableNamesCacheMaxTotalNameChars
                        (tableDefinitionStringCharCountFor sourceItemCount table)
                        nameCharCount)
        let nameCount = min names.Count (tableNamesCacheMaxTotalNames + 1)
        rememberTableNamesCacheEntryWithMetrics workbook.Tables sourceItemCount nameCount nameCharCount (resizeArrayToList names)

/// Workbook statistics
type WorkbookStats = {
    SheetCount: int
    TableCount: int
    TotalCells: int
    TotalFormulas: int
}

let getWorkbookStats (workbook: Workbook): WorkbookStats =
    let mutable totalCells = 0
    let mutable totalFormulas = 0
    workbook.Sheets
    |> Map.iter (fun _ sheet ->
        totalCells <- totalCells + Map.count sheet.Model.Cells
        sheet.Model.Cells
        |> Map.iter (fun _ cell ->
            match cell.Value with
            | CellValue.Formula _ -> totalFormulas <- totalFormulas + 1
            | _ -> ()))

    {
        SheetCount = Map.count workbook.Sheets
        TableCount = Map.count workbook.Tables
        TotalCells = totalCells
        TotalFormulas = totalFormulas
    }
