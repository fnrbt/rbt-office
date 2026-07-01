namespace RedBat.Word

module Vba =
    open Values
    open Ast

    type private CompiledSourceCacheKey = {
        Dialect: Dialect
        Tolerant: bool
        Source: string
        HostGlobalNames: string list
    }

    type private CompiledSourceCacheEntry = {
        Key: CompiledSourceCacheKey
        RetainedUnits: int
        Program: Ast.Program
        Bytecode: Bytecode.BytecodeProgram
    }

    let mutable private compiledSourceCache: CompiledSourceCacheEntry list = []
    let private compiledSourceFastSlots = Array.zeroCreate<CompiledSourceCacheEntry> 256
    let private compiledSourceCacheMaxEntries = 64
    let private compiledSourceCacheMaxTotalRetainedUnits = 1000000

    let private clearCompiledSourceFastSlots () =
        for index = 0 to compiledSourceFastSlots.Length - 1 do
            compiledSourceFastSlots.[index] <- Unchecked.defaultof<CompiledSourceCacheEntry>

    let private compiledSourceCacheSlot key =
        let mutable hashCode =
            if isNull key.Source then 0
            else System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.Source)
        hashCode <- (hashCode * 397) ^^^ System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.HostGlobalNames)
        hashCode <- (hashCode * 397) ^^^ hash key.Dialect
        hashCode <- (hashCode * 397) ^^^ if key.Tolerant then 1 else 0
        hashCode &&& (compiledSourceFastSlots.Length - 1)

    let private compiledSourceFastCacheMatches key (entry: CompiledSourceCacheEntry) =
        entry.Key.Dialect = key.Dialect
        && entry.Key.Tolerant = key.Tolerant
        && obj.ReferenceEquals(entry.Key.Source, key.Source)
        && entry.Key.HostGlobalNames = key.HostGlobalNames

    let private addRetainedUnits left right =
        if left > compiledSourceCacheMaxTotalRetainedUnits then
            left
        else
            let right = max 0 right
            if right > compiledSourceCacheMaxTotalRetainedUnits - left then compiledSourceCacheMaxTotalRetainedUnits + 1
            else left + right

    let private retainedUnitsOverBudget total =
        total > compiledSourceCacheMaxTotalRetainedUnits

    let private stringRetainedUnits (value: string) =
        if isNull value then 0 else value.Length

    let private optionRetainedUnits itemUnits value =
        match value with
        | Some item -> addRetainedUnits 1 (itemUnits item)
        | None -> 1

    let private stringListRetainedUnits values =
        let mutable total = 0
        let mutable remaining = values
        let mutable hasMore = true
        while hasMore && not (retainedUnitsOverBudget total) do
            match remaining with
            | [] ->
                hasMore <- false
            | value :: rest ->
                total <- addRetainedUnits total 1
                if not (retainedUnitsOverBudget total) then
                    total <- addRetainedUnits total (stringRetainedUnits value)
                remaining <- rest
        total

    let private stringArrayRetainedUnits (values: string array) =
        let mutable total = addRetainedUnits 0 values.Length
        let mutable index = 0
        while not (retainedUnitsOverBudget total) && index < values.Length do
            total <- addRetainedUnits total (stringRetainedUnits values.[index])
            index <- index + 1
        total

    let private listRetainedUnits itemUnits values =
        let mutable total = 0
        let mutable remaining = values
        let mutable hasMore = true
        while hasMore && not (retainedUnitsOverBudget total) do
            match remaining with
            | [] ->
                hasMore <- false
            | value :: rest ->
                total <- addRetainedUnits total 1
                if not (retainedUnitsOverBudget total) then
                    total <- addRetainedUnits total (itemUnits value)
                remaining <- rest
        total

    let rec private valueRetainedUnits value =
        match value with
        | VString text -> addRetainedUnits 1 (stringRetainedUnits text)
        | VArray values ->
            let mutable total = addRetainedUnits 1 values.Length
            let mutable index = 0
            while not (retainedUnitsOverBudget total) && index < values.Length do
                total <- addRetainedUnits total (valueRetainedUnits values.[index])
                index <- index + 1
            total
        | VObject obj ->
            addRetainedUnits 1 (stringRetainedUnits obj.ClassName)
        | VHostObject host ->
            addRetainedUnits 1 (stringRetainedUnits host.TypeName)
        | _ -> 1

    let private opcodeRetainedUnits opcode =
        match opcode with
        | Bytecode.GetMember name
        | Bytecode.SetMember name
        | Bytecode.TypeCheck name
        | Bytecode.CallBuiltin (name, _)
        | Bytecode.CallMethod (name, _) ->
            addRetainedUnits 1 (stringRetainedUnits name)
        | _ -> 1

    let private instructionRetainedUnits (instruction: Bytecode.Instruction) =
        addRetainedUnits 16 (opcodeRetainedUnits instruction.Opcode)

    let private instructionArrayRetainedUnits (instructions: Bytecode.Instruction array) =
        let mutable total = addRetainedUnits 0 instructions.Length
        let mutable index = 0
        while not (retainedUnitsOverBudget total) && index < instructions.Length do
            total <- addRetainedUnits total (instructionRetainedUnits instructions.[index])
            index <- index + 1
        total

    let private stringIntMapRetainedUnits (values: Map<string, int>) =
        let mutable total = addRetainedUnits 0 (Map.count values)
        use entries = (values :> seq<_>).GetEnumerator()
        while not (retainedUnitsOverBudget total) && entries.MoveNext() do
            let (KeyValue(name, _)) = entries.Current
            total <- addRetainedUnits total (stringRetainedUnits name)
        total

    let private bytecodeFunctionRetainedUnits (func: Bytecode.FunctionDef) =
        let mutable total = addRetainedUnits 1 (stringRetainedUnits func.Name)
        if not (retainedUnitsOverBudget total) then
            total <- addRetainedUnits total func.ByRefParams.Length
        if not (retainedUnitsOverBudget total) then
            total <- addRetainedUnits total (instructionArrayRetainedUnits func.Code)
        total

    let private classRetainedUnits (classDef: Bytecode.ClassDef) =
        let mutable total = addRetainedUnits 1 (stringRetainedUnits classDef.Name)
        if not (retainedUnitsOverBudget total) then
            total <- addRetainedUnits total (stringArrayRetainedUnits classDef.Fields)
        if not (retainedUnitsOverBudget total) then
            total <- addRetainedUnits total (stringIntMapRetainedUnits classDef.Methods)
        if not (retainedUnitsOverBudget total) then
            total <- addRetainedUnits total (stringIntMapRetainedUnits classDef.PropertyGetters)
        if not (retainedUnitsOverBudget total) then
            total <- addRetainedUnits total (stringIntMapRetainedUnits classDef.PropertyLetters)
        if not (retainedUnitsOverBudget total) then
            total <- addRetainedUnits total (stringIntMapRetainedUnits classDef.PropertySetters)
        total

    let private bytecodeProgramRetainedUnits (bytecode: Bytecode.BytecodeProgram) =
        let mutable total = instructionArrayRetainedUnits bytecode.Code
        if not (retainedUnitsOverBudget total) then
            total <- addRetainedUnits total bytecode.Constants.Length
        let mutable constantIndex = 0
        while not (retainedUnitsOverBudget total) && constantIndex < bytecode.Constants.Length do
            total <- addRetainedUnits total (valueRetainedUnits bytecode.Constants.[constantIndex])
            constantIndex <- constantIndex + 1
        if not (retainedUnitsOverBudget total) then
            total <- addRetainedUnits total bytecode.Globals
        if not (retainedUnitsOverBudget total) then
            total <- addRetainedUnits total (stringArrayRetainedUnits bytecode.GlobalNames)
        if not (retainedUnitsOverBudget total) then
            total <- addRetainedUnits total bytecode.Functions.Length
        let mutable functionIndex = 0
        while not (retainedUnitsOverBudget total) && functionIndex < bytecode.Functions.Length do
            total <- addRetainedUnits total (bytecodeFunctionRetainedUnits bytecode.Functions.[functionIndex])
            functionIndex <- functionIndex + 1
        if not (retainedUnitsOverBudget total) then
            total <- addRetainedUnits total bytecode.Classes.Length
        let mutable classIndex = 0
        while not (retainedUnitsOverBudget total) && classIndex < bytecode.Classes.Length do
            total <- addRetainedUnits total (classRetainedUnits bytecode.Classes.[classIndex])
            classIndex <- classIndex + 1
        total

    let private literalRetainedUnits literal =
        match literal with
        | Literal.String text -> addRetainedUnits 1 (stringRetainedUnits text)
        | _ -> 1

    let private typeRefRetainedUnits (typeRef: TypeRef) =
        addRetainedUnits 1 (stringRetainedUnits typeRef.Name)

    let rec private expressionRetainedUnits expression =
        match expression with
        | Literal literal -> addRetainedUnits 1 (literalRetainedUnits literal)
        | Identifier name -> addRetainedUnits 1 (stringRetainedUnits name)
        | Binary (_, left, right) ->
            let mutable total = addRetainedUnits 1 (expressionRetainedUnits left)
            total <- addRetainedUnits total (expressionRetainedUnits right)
            total
        | Unary (_, inner) ->
            addRetainedUnits 1 (expressionRetainedUnits inner)
        | Call (target, args) ->
            let mutable total = addRetainedUnits 1 (expressionRetainedUnits target)
            total <- addRetainedUnits total (listRetainedUnits expressionRetainedUnits args)
            total
        | Member (target, name) ->
            let mutable total = addRetainedUnits 1 (expressionRetainedUnits target)
            total <- addRetainedUnits total (stringRetainedUnits name)
            total
        | NewExpr name ->
            addRetainedUnits 1 (stringRetainedUnits name)
        | TypeOfIs (target, name) ->
            let mutable total = addRetainedUnits 1 (expressionRetainedUnits target)
            total <- addRetainedUnits total (stringRetainedUnits name)
            total
        | MeExpr
        | WithDotExpr ->
            1
    and private arraySpecRetainedUnits (arraySpec: ArraySpec) =
        addRetainedUnits 1 (listRetainedUnits expressionRetainedUnits arraySpec.Dimensions)
    and private varDeclaratorRetainedUnits (declarator: VarDeclarator) =
        let mutable total = addRetainedUnits 1 (stringRetainedUnits declarator.Name)
        total <- addRetainedUnits total (optionRetainedUnits arraySpecRetainedUnits declarator.ArraySpec)
        total <- addRetainedUnits total (optionRetainedUnits typeRefRetainedUnits declarator.Type)
        total
    and private declarationRetainedUnits declaration =
        match declaration with
        | Dim declarators ->
            addRetainedUnits 1 (listRetainedUnits varDeclaratorRetainedUnits declarators)
        | Const (name, typeRef, value) ->
            let mutable total = addRetainedUnits 1 (stringRetainedUnits name)
            total <- addRetainedUnits total (optionRetainedUnits typeRefRetainedUnits typeRef)
            total <- addRetainedUnits total (expressionRetainedUnits value)
            total
        | ReDim (name, size, _) ->
            let mutable total = addRetainedUnits 1 (stringRetainedUnits name)
            total <- addRetainedUnits total (expressionRetainedUnits size)
            total
    and private parameterRetainedUnits (parameter: Parameter) =
        let mutable total = addRetainedUnits 1 (stringRetainedUnits parameter.Name)
        total <- addRetainedUnits total (optionRetainedUnits typeRefRetainedUnits parameter.Type)
        total <- addRetainedUnits total (optionRetainedUnits expressionRetainedUnits parameter.DefaultValue)
        total
    and private onErrorActionRetainedUnits action =
        match action with
        | GoToLabel label -> addRetainedUnits 1 (stringRetainedUnits label)
        | _ -> 1
    and private caseTestRetainedUnits caseTest =
        match caseTest with
        | CaseValue value -> addRetainedUnits 1 (expressionRetainedUnits value)
        | CaseRange (first, last) ->
            let mutable total = addRetainedUnits 1 (expressionRetainedUnits first)
            total <- addRetainedUnits total (expressionRetainedUnits last)
            total
        | CaseComparison (_, value) ->
            addRetainedUnits 1 (expressionRetainedUnits value)
    and private doConditionRetainedUnits condition =
        let mutable total = addRetainedUnits 1 (expressionRetainedUnits condition.Condition)
        total <- addRetainedUnits total (listRetainedUnits statementRetainedUnits condition.Body)
        total
    and private statementRetainedUnits statement =
        match statement with
        | ExpressionStmt expression ->
            addRetainedUnits 1 (expressionRetainedUnits expression)
        | Let (name, value)
        | Set (name, value)
        | Assignment (name, value) ->
            let mutable total = addRetainedUnits 1 (stringRetainedUnits name)
            total <- addRetainedUnits total (expressionRetainedUnits value)
            total
        | IfStmt (condition, thenBody, elseIfs, elseBody) ->
            let mutable total = addRetainedUnits 1 (expressionRetainedUnits condition)
            total <- addRetainedUnits total (listRetainedUnits statementRetainedUnits thenBody)
            total <- addRetainedUnits total (
                listRetainedUnits
                    (fun (elseIfCondition, elseIfBody) ->
                        let mutable elseIfTotal = expressionRetainedUnits elseIfCondition
                        elseIfTotal <- addRetainedUnits elseIfTotal (listRetainedUnits statementRetainedUnits elseIfBody)
                        elseIfTotal)
                    elseIfs)
            total <- addRetainedUnits total (optionRetainedUnits (listRetainedUnits statementRetainedUnits) elseBody)
            total
        | SelectCase (target, cases) ->
            let mutable total = addRetainedUnits 1 (expressionRetainedUnits target)
            total <- addRetainedUnits total (
                listRetainedUnits
                    (fun (caseTests, body) ->
                        let mutable caseTotal = optionRetainedUnits (listRetainedUnits caseTestRetainedUnits) caseTests
                        caseTotal <- addRetainedUnits caseTotal (listRetainedUnits statementRetainedUnits body)
                        caseTotal)
                    cases)
            total
        | ForLoop (name, first, last, step, body) ->
            let mutable total = addRetainedUnits 1 (stringRetainedUnits name)
            total <- addRetainedUnits total (expressionRetainedUnits first)
            total <- addRetainedUnits total (expressionRetainedUnits last)
            total <- addRetainedUnits total (optionRetainedUnits expressionRetainedUnits step)
            total <- addRetainedUnits total (listRetainedUnits statementRetainedUnits body)
            total
        | ForEach (name, target, body) ->
            let mutable total = addRetainedUnits 1 (stringRetainedUnits name)
            total <- addRetainedUnits total (expressionRetainedUnits target)
            total <- addRetainedUnits total (listRetainedUnits statementRetainedUnits body)
            total
        | WhileLoop (condition, body) ->
            let mutable total = addRetainedUnits 1 (expressionRetainedUnits condition)
            total <- addRetainedUnits total (listRetainedUnits statementRetainedUnits body)
            total
        | DoLoop (topCondition, bottomCondition) ->
            let mutable total = 1
            total <- addRetainedUnits total (optionRetainedUnits doConditionRetainedUnits topCondition)
            total <- addRetainedUnits total (optionRetainedUnits doConditionRetainedUnits bottomCondition)
            total
        | CallStmt (target, args) ->
            let mutable total = addRetainedUnits 1 (expressionRetainedUnits target)
            total <- addRetainedUnits total (listRetainedUnits expressionRetainedUnits args)
            total
        | GoToStmt label
        | GoSubStmt label
        | LabelStmt label ->
            addRetainedUnits 1 (stringRetainedUnits label)
        | ReturnStmt ->
            1
        | Declaration (_, declaration) ->
            addRetainedUnits 1 (declarationRetainedUnits declaration)
        | IndexedAssignment (name, indexes, value) ->
            let mutable total = addRetainedUnits 1 (stringRetainedUnits name)
            total <- addRetainedUnits total (listRetainedUnits expressionRetainedUnits indexes)
            total <- addRetainedUnits total (expressionRetainedUnits value)
            total
        | MemberAssignment (target, name, value) ->
            let mutable total = addRetainedUnits 1 (expressionRetainedUnits target)
            total <- addRetainedUnits total (stringRetainedUnits name)
            total <- addRetainedUnits total (expressionRetainedUnits value)
            total
        | OnError action ->
            addRetainedUnits 1 (onErrorActionRetainedUnits action)
        | WithStmt (target, body) ->
            let mutable total = addRetainedUnits 1 (expressionRetainedUnits target)
            total <- addRetainedUnits total (listRetainedUnits statementRetainedUnits body)
            total
        | EraseStmt name ->
            addRetainedUnits 1 (stringRetainedUnits name)
        | ExitFor
        | ExitDo
        | ExitSub
        | ExitFunction
        | ExitProperty ->
            1

    let private functionRetainedUnits func =
        match func with
        | FunctionDecl (modifiers, name, parameters, typeRef, body)
        | PropertyGet (modifiers, name, parameters, typeRef, body) ->
            let mutable total = addRetainedUnits 1 modifiers.Length
            total <- addRetainedUnits total (stringRetainedUnits name)
            total <- addRetainedUnits total (listRetainedUnits parameterRetainedUnits parameters)
            total <- addRetainedUnits total (optionRetainedUnits typeRefRetainedUnits typeRef)
            total <- addRetainedUnits total (listRetainedUnits statementRetainedUnits body)
            total
        | SubDecl (modifiers, name, parameters, body)
        | PropertyLet (modifiers, name, parameters, body)
        | PropertySet (modifiers, name, parameters, body) ->
            let mutable total = addRetainedUnits 1 modifiers.Length
            total <- addRetainedUnits total (stringRetainedUnits name)
            total <- addRetainedUnits total (listRetainedUnits parameterRetainedUnits parameters)
            total <- addRetainedUnits total (listRetainedUnits statementRetainedUnits body)
            total

    let private enumMemberRetainedUnits (memberDef: EnumMember) =
        let mutable total = addRetainedUnits 1 (stringRetainedUnits memberDef.Name)
        total <- addRetainedUnits total (optionRetainedUnits expressionRetainedUnits memberDef.Value)
        total

    let private typeMemberRetainedUnits (memberDef: TypeMember) =
        let mutable total = addRetainedUnits 1 (stringRetainedUnits memberDef.Name)
        total <- addRetainedUnits total (typeRefRetainedUnits memberDef.Type)
        total <- addRetainedUnits total (optionRetainedUnits arraySpecRetainedUnits memberDef.ArraySpec)
        total

    let private declareInfoRetainedUnits (declareInfo: DeclareInfo) =
        let mutable total = addRetainedUnits 1 (stringRetainedUnits declareInfo.Name)
        total <- addRetainedUnits total (stringRetainedUnits declareInfo.LibName)
        total <- addRetainedUnits total (optionRetainedUnits stringRetainedUnits declareInfo.AliasName)
        total <- addRetainedUnits total (listRetainedUnits parameterRetainedUnits declareInfo.Parameters)
        total <- addRetainedUnits total (optionRetainedUnits typeRefRetainedUnits declareInfo.ReturnType)
        total

    let rec private topLevelRetainedUnits topLevel =
        match topLevel with
        | FunctionDef func ->
            addRetainedUnits 1 (functionRetainedUnits func)
        | ClassDecl (_, name, members) ->
            let mutable total = addRetainedUnits 1 (stringRetainedUnits name)
            total <- addRetainedUnits total (listRetainedUnits topLevelRetainedUnits members)
            total
        | EnumDecl (_, name, members) ->
            let mutable total = addRetainedUnits 1 (stringRetainedUnits name)
            total <- addRetainedUnits total (listRetainedUnits enumMemberRetainedUnits members)
            total
        | TypeDecl (_, name, members) ->
            let mutable total = addRetainedUnits 1 (stringRetainedUnits name)
            total <- addRetainedUnits total (listRetainedUnits typeMemberRetainedUnits members)
            total
        | DeclareDecl (_, declareInfo) ->
            addRetainedUnits 1 (declareInfoRetainedUnits declareInfo)
        | ImplementsDecl name
        | OptionStmt name ->
            addRetainedUnits 1 (stringRetainedUnits name)
        | EventDecl (_, name, parameters) ->
            let mutable total = addRetainedUnits 1 (stringRetainedUnits name)
            total <- addRetainedUnits total (listRetainedUnits parameterRetainedUnits parameters)
            total
        | WithEventsDecl (_, name, typeRef) ->
            let mutable total = addRetainedUnits 1 (stringRetainedUnits name)
            total <- addRetainedUnits total (typeRefRetainedUnits typeRef)
            total
        | TopLevelStatement statement ->
            addRetainedUnits 1 (statementRetainedUnits statement)

    let private programRetainedUnits (program: Ast.Program) =
        addRetainedUnits 1 (listRetainedUnits topLevelRetainedUnits program.TopLevels)

    let private compiledSourceRetainedUnits key program bytecode =
        let mutable total = max 1 key.Source.Length
        if not (retainedUnitsOverBudget total) then
            total <- addRetainedUnits total (stringListRetainedUnits key.HostGlobalNames)
        if not (retainedUnitsOverBudget total) then
            total <- addRetainedUnits total (programRetainedUnits program)
        if not (retainedUnitsOverBudget total) then
            total <- addRetainedUnits total (bytecodeProgramRetainedUnits bytecode)
        total

    let private setCompiledSourceCache entries =
        compiledSourceCache <- entries
        clearCompiledSourceFastSlots ()
        for entry in entries do
            let slot = compiledSourceCacheSlot entry.Key
            if obj.ReferenceEquals(compiledSourceFastSlots.[slot], null) then
                compiledSourceFastSlots.[slot] <- entry

    let private tryGetCompiledSourceCache key =
        let slot = compiledSourceCacheSlot key
        let fast = compiledSourceFastSlots.[slot]
        if not (obj.ReferenceEquals(fast, null))
           && compiledSourceFastCacheMatches key fast then
            Some fast
        else
            match compiledSourceCache |> List.tryFind (fun entry -> entry.Key = key) with
            | Some entry ->
                compiledSourceFastSlots.[slot] <- entry
                Some entry
            | None ->
                None

    let private rememberCompiledSource entry =
        let trim entries =
            let rec loop remaining kept count totalRetainedUnits =
                match remaining with
                | [] -> List.rev kept
                | candidate :: rest ->
                    let retainedUnits = max 1 candidate.RetainedUnits
                    let nextTotalRetainedUnits = addRetainedUnits totalRetainedUnits retainedUnits
                    if count < compiledSourceCacheMaxEntries
                       && nextTotalRetainedUnits <= compiledSourceCacheMaxTotalRetainedUnits then
                        loop rest (candidate :: kept) (count + 1) nextTotalRetainedUnits
                    else
                        loop rest kept count totalRetainedUnits
            loop entries [] 0 0
        entry
        :: (compiledSourceCache |> List.filter (fun existing -> existing.Key <> entry.Key))
        |> trim
        |> setCompiledSourceCache

    let private compileSourceCached (options: FsVb.ParseOptions) source =
        let key = {
            Dialect = options.Dialect
            Tolerant = options.Tolerant
            Source = source
            HostGlobalNames = options.HostGlobalNames
        }
        match tryGetCompiledSourceCache key with
        | Some entry ->
            Ok (entry.Program, entry.Bytecode)
        | None ->
            match FsVb.compileSource options source with
            | Error err -> Error err
            | Ok (program, bytecode) ->
                rememberCompiledSource {
                    Key = key
                    RetainedUnits = compiledSourceRetainedUnits key program bytecode
                    Program = program
                    Bytecode = bytecode
                }
                Ok (program, bytecode)

    let private runSourceCached (options: FsVb.RunOptions) source =
        if not options.Validate then
            FsVb.runSource options source
        else
            let hostGlobals = options.HostGlobals
            let parseOptions: FsVb.ParseOptions = {
                Dialect = options.Dialect
                Tolerant = options.Tolerant
                HostGlobalNames = hostGlobals |> Map.toSeq |> Seq.map fst |> Seq.toList
            }
            match compileSourceCached parseOptions source with
            | Error err -> Error err
            | Ok (program, bytecode) ->
                try
                    let state = FsVb.executeWithHostGlobals (hostGlobals |> Map.toSeq) bytecode
                    Ok {
                        Program = program
                        Bytecode = bytecode
                        State = state
                    }
                with ex ->
                    Error (FsVb.RuntimeError ex.Message)

#if FABLE_COMPILER
    open Fable.Core.JsInterop

    module Doc = Document
    module TextBuf = TextBuffer
    module WordApi = API

    let private eq (left: string) (right: string) =
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

    type private HostObject(typeName: string, getMember: string -> Value option, setMember: string -> Value -> bool, callMethod: string -> Value array -> Value option) =
        interface IHostObject with
            member _.TypeName = typeName
            member _.GetMember name = getMember name
            member _.SetMember name value = setMember name value
            member _.CallMethod name args = callMethod name args

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

    let private isWordDelimiter = function
        | ' ' | '\t' | '\r' | '\n' -> true
        | _ -> false

    let private countWords (text: string) =
        let mutable count = 0
        let mutable inWord = false
        for ch in text do
            if isWordDelimiter ch then
                inWord <- false
            elif not inWord then
                count <- count + 1
                inWord <- true
        count

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
        let paragraphs = ResizeArray<Doc.ParagraphBlock>()
        for block in doc.Blocks do
            match block with
            | Doc.Paragraph p
            | Doc.Heading (_, p) -> paragraphs.Add p
            | Doc.Table t ->
                for row in t.Rows do
                    for cell in row.Cells do
                        for para in cell.Content do
                            paragraphs.Add para
            | Doc.TextBox t ->
                for para in t.Content do
                    paragraphs.Add para
        List.ofSeq paragraphs

    let private iterParagraphs (doc: Doc.Document) f =
        for block in doc.Blocks do
            match block with
            | Doc.Paragraph p
            | Doc.Heading (_, p) -> f p
            | Doc.Table t ->
                for row in t.Rows do
                    for cell in row.Cells do
                        for para in cell.Content do
                            f para
            | Doc.TextBox t ->
                for para in t.Content do
                    f para

    let private paragraphCount (doc: Doc.Document) =
        let mutable count = 0
        iterParagraphs doc (fun _ -> count <- count + 1)
        count

    let private resolvePosition (doc: Doc.Document) (offset: int) =
        let mutable remaining = max 0 offset
        let mutable found = None
        let mutable lastParagraph = None
        iterParagraphs doc (fun para ->
            if found.IsNone then
                lastParagraph <- Some para
                let len = TextBuf.length para.Buffer
                if remaining <= len then
                    found <- Some (para.Id, remaining)
                else
                    // Match Document.getText's paragraph separator.
                    remaining <- remaining - len - 1)
        match found with
        | Some pos -> pos
        | None ->
            match lastParagraph with
            | Some last -> last.Id, TextBuf.length last.Buffer
            | None -> "", 0

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
                                countWords (Doc.getText editor.Doc)
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
                            if eq field "Count" then Some (VInteger (paragraphCount editor.Doc))
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
            FsVb.RunOptions.vba
            |> FsVb.RunOptions.withHostGlobals (hostGlobals editor)
        runSourceCached options source

    let runVbs (editor: WordApi.Editor) (source: string) =
        let options =
            FsVb.RunOptions.vbScript
            |> FsVb.RunOptions.withHostGlobals (hostGlobalsVbs editor)
        runSourceCached options source
#else
    open System.Collections.Generic

    module Doc = Document
    module TextBuf = TextBuffer
    module WordLayout = Layout

    let private clearArray (items: 'T array) =
        for i = 0 to items.Length - 1 do
            items.[i] <- Unchecked.defaultof<'T>

    let private documentBlocksReferenceComparer =
        { new IEqualityComparer<Doc.Block list> with
            member _.Equals(left, right) = obj.ReferenceEquals(left, right)
            member _.GetHashCode(value) = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value) }

    let private documentBlocksSlot blocks slotCount =
        System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(blocks)
        &&& (slotCount - 1)

    let private eq (left: string) (right: string) =
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

    type private HostObject(typeName: string, getMember: string -> Value option, setMember: string -> Value -> bool, callMethod: string -> Value array -> Value option) =
        interface IHostObject with
            member _.TypeName = typeName
            member _.GetMember name = getMember name
            member _.SetMember name value = setMember name value
            member _.CallMethod name args = callMethod name args

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
        let mutable count = 0
        let mutable inWord = false
        for ch in text do
            match ch with
            | ' ' | '\t' | '\r' | '\n' -> inWord <- false
            | _ ->
                if not inWord then
                    count <- count + 1
                    inWord <- true
        count

    let private countLines (text: string) =
        let mutable count = 1
        for ch in text do
            if ch = '\n' then
                count <- count + 1
        count

    let private paragraphsInOrder (doc: Doc.Document) =
        let paragraphs = ResizeArray<Doc.ParagraphBlock>()
        for block in doc.Blocks do
            match block with
            | Doc.Paragraph p
            | Doc.Heading (_, p) -> paragraphs.Add p
            | Doc.Table t ->
                for row in t.Rows do
                    for cell in row.Cells do
                        for para in cell.Content do
                            paragraphs.Add para
            | Doc.TextBox t ->
                for para in t.Content do
                    paragraphs.Add para
        List.ofSeq paragraphs

    let private iterParagraphs (doc: Doc.Document) f =
        for block in doc.Blocks do
            match block with
            | Doc.Paragraph p
            | Doc.Heading (_, p) -> f p
            | Doc.Table t ->
                for row in t.Rows do
                    for cell in row.Cells do
                        for para in cell.Content do
                            f para
            | Doc.TextBox t ->
                for para in t.Content do
                    f para

    let private paragraphCount (doc: Doc.Document) =
        let mutable count = 0
        iterParagraphs doc (fun _ -> count <- count + 1)
        count

    let private resolvePosition (doc: Doc.Document) (offset: int) =
        let mutable remaining = max 0 offset
        let mutable found = None
        let mutable lastParagraph = None
        iterParagraphs doc (fun para ->
            if found.IsNone then
                lastParagraph <- Some para
                let len = TextBuf.length para.Buffer
                if remaining <= len then
                    found <- Some (para.Id, remaining)
                else
                    remaining <- remaining - len - 1)
        match found with
        | Some pos -> pos
        | None ->
            match lastParagraph with
            | Some last -> last.Id, TextBuf.length last.Buffer
            | None -> "", 0

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
    }

    let private measureText (text: string) (style: TextBuf.TextStyle) =
        let size = if style.FontSize > 0.0 then style.FontSize else 11.0
        let weight = if style.Bold then 0.58 else 0.52
        float text.Length * size * weight

    let private measureRange = WordLayout.rangeMeasurerFromTextMeasurer measureText

    type private LayoutCacheValue = {
        SourceTextCharCount: int
        Layouts: WordLayout.BlockLayout list
        LayoutCount: int
        LayoutById: Dictionary<Doc.BlockId, WordLayout.BlockLayout>
        LayoutsByBlock: Dictionary<Doc.BlockId, ResizeArray<WordLayout.BlockLayout>>
        PageCount: int
    }

    type private LayoutCache = {
        CachedDoc: Doc.Document
        Value: LayoutCacheValue
    }

    let mutable private layoutCache: LayoutCache list = []
    let private layoutCacheByBlocks =
        Dictionary<Doc.Block list, ResizeArray<LayoutCache>>(documentBlocksReferenceComparer)
    let private layoutFastSlots = Array.zeroCreate<LayoutCache> 256
    let private layoutCacheMaxEntries = 64
    let private layoutCacheMaxTotalLayouts = 200000
    let private vbaSourceTextCacheMaxTotalChars = 5000000
    let private vbaSourceParagraphCountLimit = 100000

    let private vbaAddCapped total value =
        if total > vbaSourceTextCacheMaxTotalChars then
            total
        else
            let value = max 0 value
            let remaining = vbaSourceTextCacheMaxTotalChars - total
            if value > remaining then vbaSourceTextCacheMaxTotalChars + 1
            else total + value

    let private vbaOptionStringCharCount (value: string option) =
        match value with
        | Some value when not (isNull value) -> value.Length
        | _ -> 0

    let private vbaSectionColumnsStringCharCount (columns: Doc.SectionColumns option) =
        match columns with
        | Some columns ->
            let mutable total = 0
            let mutable seen = 0
            let mutable widths = columns.Widths
            while total <= vbaSourceTextCacheMaxTotalChars
                  && seen <= vbaSourceParagraphCountLimit
                  && not widths.IsEmpty do
                total <- vbaAddCapped total 1
                seen <- seen + 1
                widths <- widths.Tail
            if seen > vbaSourceParagraphCountLimit then
                vbaSourceTextCacheMaxTotalChars + 1
            else
                total
        | None -> 0

    let private vbaTextStyleStringCharCount (style: TextBuf.TextStyle) =
        let mutable total = 0
        let add value =
            total <- vbaAddCapped total (vbaOptionStringCharCount value)
        add style.Style
        add style.UnderlineStyle
        add style.UnderlineColor
        add style.EmphasisMark
        add style.FontFamily
        add style.Color
        add style.Highlight
        add style.Hyperlink
        add style.CommentId
        add style.CommentText
        add style.CommentAuthor
        add style.CommentDate
        add style.NoteId
        add style.NoteType
        add style.NoteText
        add style.BookmarkStartId
        add style.BookmarkName
        add style.BookmarkEndId
        add style.FieldCode
        add style.FieldResult
        add style.EquationXml
        add style.RevisionId
        add style.RevisionType
        add style.RevisionAuthor
        add style.RevisionDate
        add style.VertAlign
        total

    let private vbaImageStringCharCount (image: TextBuf.ImageRun) =
        let mutable total = 0
        total <- vbaAddCapped total image.Src.Length
        total <- vbaAddCapped total (vbaOptionStringCharCount image.Alt)
        total <- vbaAddCapped total (vbaOptionStringCharCount image.Name)
        total <- vbaAddCapped total (vbaOptionStringCharCount image.Title)
        total <- vbaAddCapped total (vbaOptionStringCharCount image.Wrap)
        total <- vbaAddCapped total (vbaOptionStringCharCount image.PositionHRelativeFrom)
        total <- vbaAddCapped total (vbaOptionStringCharCount image.PositionVRelativeFrom)
        total <- vbaAddCapped total (vbaOptionStringCharCount image.RelativeHeight)
        total <- vbaAddCapped total (vbaOptionStringCharCount image.CropLeft)
        total <- vbaAddCapped total (vbaOptionStringCharCount image.CropTop)
        total <- vbaAddCapped total (vbaOptionStringCharCount image.CropRight)
        total <- vbaAddCapped total (vbaOptionStringCharCount image.CropBottom)
        total

    let private vbaRunStringCharCount = function
        | TextBuf.Text textRun -> vbaTextStyleStringCharCount textRun.Style
        | TextBuf.Image image -> vbaImageStringCharCount image

    let private vbaBufferSourceTextCharCount (buffer: TextBuf.TextBuffer) =
        let mutable total = TextBuf.length buffer
        let mutable seen = 0
        let mutable runs = buffer.Runs
        while total <= vbaSourceTextCacheMaxTotalChars
              && seen <= vbaSourceParagraphCountLimit
              && not runs.IsEmpty do
            total <- vbaAddCapped total (vbaRunStringCharCount runs.Head)
            seen <- seen + 1
            runs <- runs.Tail
        total

    let private vbaParagraphBorderStringCharCount (border: Doc.ParagraphBorder option) =
        match border with
        | Some border ->
            let mutable total = 0
            total <- vbaAddCapped total (vbaOptionStringCharCount border.Style)
            total <- vbaAddCapped total (vbaOptionStringCharCount border.Color)
            total
        | None -> 0

    let private vbaParagraphBordersStringCharCount (borders: Doc.ParagraphBorders option) =
        match borders with
        | Some borders ->
            let mutable total = 0
            total <- vbaAddCapped total (vbaParagraphBorderStringCharCount borders.Top)
            total <- vbaAddCapped total (vbaParagraphBorderStringCharCount borders.Right)
            total <- vbaAddCapped total (vbaParagraphBorderStringCharCount borders.Bottom)
            total <- vbaAddCapped total (vbaParagraphBorderStringCharCount borders.Left)
            total
        | None -> 0

    let private vbaTabStopStringCharCount (tabStop: Doc.TabStop) =
        vbaAddCapped tabStop.Alignment.Length (vbaOptionStringCharCount tabStop.Leader)

    let private vbaParagraphMetadataStringCharCount (metadata: Doc.BlockMetadata) =
        let mutable total = 0
        total <- vbaAddCapped total (vbaOptionStringCharCount metadata.Style)
        total <- vbaAddCapped total (vbaOptionStringCharCount metadata.ContentControlId)
        total <- vbaAddCapped total (vbaOptionStringCharCount metadata.ContentControlTag)
        total <- vbaAddCapped total (vbaOptionStringCharCount metadata.ContentControlAlias)
        total <- vbaAddCapped total (vbaOptionStringCharCount metadata.ListId)
        total <- vbaAddCapped total (vbaOptionStringCharCount metadata.ListNumberFormat)
        total <- vbaAddCapped total (vbaOptionStringCharCount metadata.ListLevelText)
        total <- vbaAddCapped total (vbaOptionStringCharCount metadata.RightTabLeader)
        total <- vbaAddCapped total (vbaOptionStringCharCount metadata.LineSpacingRule)
        total <- vbaAddCapped total (vbaOptionStringCharCount metadata.BackgroundColor)
        total <- vbaAddCapped total (vbaParagraphBordersStringCharCount metadata.ParagraphBorders)
        total <- vbaAddCapped total (vbaOptionStringCharCount metadata.SectionBreak)
        total <- vbaAddCapped total (vbaSectionColumnsStringCharCount metadata.SectionColumns)
        let mutable seen = 0
        let mutable tabStops = metadata.TabStops
        while total <= vbaSourceTextCacheMaxTotalChars
              && seen <= vbaSourceParagraphCountLimit
              && not tabStops.IsEmpty do
            total <- vbaAddCapped total (vbaTabStopStringCharCount tabStops.Head)
            seen <- seen + 1
            tabStops <- tabStops.Tail
        let mutable headerReferences = metadata.HeaderReferences
        while total <= vbaSourceTextCacheMaxTotalChars
              && seen <= vbaSourceParagraphCountLimit
              && not headerReferences.IsEmpty do
            let kind, relationshipId = headerReferences.Head
            total <- vbaAddCapped total kind.Length
            total <- vbaAddCapped total relationshipId.Length
            seen <- seen + 1
            headerReferences <- headerReferences.Tail
        let mutable footerReferences = metadata.FooterReferences
        while total <= vbaSourceTextCacheMaxTotalChars
              && seen <= vbaSourceParagraphCountLimit
              && not footerReferences.IsEmpty do
            let kind, relationshipId = footerReferences.Head
            total <- vbaAddCapped total kind.Length
            total <- vbaAddCapped total relationshipId.Length
            seen <- seen + 1
            footerReferences <- footerReferences.Tail
        total

    let private vbaParagraphSourceTextCharCount (paragraph: Doc.ParagraphBlock) =
        let mutable total = paragraph.Id.Length
        total <- vbaAddCapped total (vbaBufferSourceTextCharCount paragraph.Buffer)
        total <- vbaAddCapped total (vbaParagraphMetadataStringCharCount paragraph.Metadata)
        total

    let private vbaTableBorderStringCharCount (border: Doc.TableBorder option) =
        match border with
        | Some border ->
            let mutable total = 0
            total <- vbaAddCapped total (vbaOptionStringCharCount border.Style)
            total <- vbaAddCapped total (vbaOptionStringCharCount border.Color)
            total
        | None -> 0

    let private vbaTableBordersStringCharCount (borders: Doc.TableBorders option) =
        match borders with
        | Some borders ->
            let mutable total = 0
            total <- vbaAddCapped total (vbaTableBorderStringCharCount borders.Top)
            total <- vbaAddCapped total (vbaTableBorderStringCharCount borders.Right)
            total <- vbaAddCapped total (vbaTableBorderStringCharCount borders.Bottom)
            total <- vbaAddCapped total (vbaTableBorderStringCharCount borders.Left)
            total <- vbaAddCapped total (vbaTableBorderStringCharCount borders.InsideH)
            total <- vbaAddCapped total (vbaTableBorderStringCharCount borders.InsideV)
            total
        | None -> 0

    let private vbaShadingStringCharCount (shading: Doc.Shading option) =
        match shading with
        | Some shading ->
            let mutable total = 0
            total <- vbaAddCapped total (vbaOptionStringCharCount shading.Pattern)
            total <- vbaAddCapped total (vbaOptionStringCharCount shading.ForegroundColor)
            total <- vbaAddCapped total (vbaOptionStringCharCount shading.BackgroundColor)
            total
        | None -> 0

    let private vbaTableStringCharCount (table: Doc.TableBlock) =
        let mutable total = table.Id.Length
        total <- vbaAddCapped total (vbaOptionStringCharCount table.Alignment)
        total <- vbaAddCapped total (vbaOptionStringCharCount table.BackgroundColor)
        total <- vbaAddCapped total (vbaShadingStringCharCount table.Shading)
        total <- vbaAddCapped total (vbaTableBordersStringCharCount table.Borders)
        total

    let private vbaTableRowStringCharCount (row: Doc.TableRow) =
        let mutable total = 0
        total <- vbaAddCapped total (vbaOptionStringCharCount row.HeightRule)
        total <- vbaAddCapped total (vbaTableBordersStringCharCount row.BordersOverride)
        total

    let private vbaTableCellStringCharCount (cell: Doc.TableCell) =
        let mutable total = 0
        total <- vbaAddCapped total (vbaOptionStringCharCount cell.BackgroundColor)
        total <- vbaAddCapped total (vbaShadingStringCharCount cell.Shading)
        total <- vbaAddCapped total (vbaOptionStringCharCount cell.BorderColor)
        total <- vbaAddCapped total (vbaTableBordersStringCharCount cell.Borders)
        total <- vbaAddCapped total (vbaOptionStringCharCount cell.VerticalAlignment)
        total

    let private vbaPathSegmentsStringCharCount (segments: Doc.PathSeg list) =
        let mutable total = 0
        let mutable seen = 0
        let mutable remaining = segments
        while total <= vbaSourceTextCacheMaxTotalChars
              && seen <= vbaSourceParagraphCountLimit
              && not remaining.IsEmpty do
            total <- vbaAddCapped total 1
            seen <- seen + 1
            remaining <- remaining.Tail
        if seen > vbaSourceParagraphCountLimit then
            vbaSourceTextCacheMaxTotalChars + 1
        else
            total

    let private vbaShapeGeometryStringCharCount = function
        | Doc.GeoRect -> 0
        | Doc.GeoPreset preset -> preset.Length
        | Doc.GeoCustom (segments, _, _) -> vbaPathSegmentsStringCharCount segments

    let private vbaGradientStringCharCount (gradient: Doc.GradientSpec option) =
        match gradient with
        | Some gradient ->
            let mutable total = 0
            let mutable seen = 0
            let mutable stops = gradient.Stops
            while total <= vbaSourceTextCacheMaxTotalChars
                  && seen <= vbaSourceParagraphCountLimit
                  && not stops.IsEmpty do
                total <- vbaAddCapped total stops.Head.Color.Length
                seen <- seen + 1
                stops <- stops.Tail
            if seen > vbaSourceParagraphCountLimit then
                vbaSourceTextCacheMaxTotalChars + 1
            else
                total
        | None -> 0

    let private vbaTextBoxStringCharCount (textBox: Doc.TextBoxBlock) =
        let mutable total = textBox.Id.Length
        total <- vbaAddCapped total (vbaOptionStringCharCount textBox.PositionHRelativeFrom)
        total <- vbaAddCapped total (vbaOptionStringCharCount textBox.PositionVRelativeFrom)
        total <- vbaAddCapped total (vbaOptionStringCharCount textBox.FillColor)
        total <- vbaAddCapped total (vbaOptionStringCharCount textBox.BorderColor)
        total <- vbaAddCapped total (vbaOptionStringCharCount textBox.RelativeHeight)
        total <- vbaAddCapped total (vbaShapeGeometryStringCharCount textBox.Geometry)
        total <- vbaAddCapped total (vbaGradientStringCharCount textBox.Gradient)
        total

    let private vbaNamedStyleStringCharCount (style: Doc.NamedStyle) =
        let mutable total = style.Id.Length
        total <- vbaAddCapped total (vbaOptionStringCharCount style.Name)
        total <- vbaAddCapped total style.StyleType.Length
        total <- vbaAddCapped total (vbaOptionStringCharCount style.BasedOn)
        total <- vbaAddCapped total (vbaOptionStringCharCount style.Next)
        total <- vbaAddCapped total (vbaOptionStringCharCount style.Linked)
        total <- vbaAddCapped total (vbaOptionStringCharCount style.Aliases)
        total <- vbaAddCapped total (vbaOptionStringCharCount style.LineSpacingRule)
        total <- vbaAddCapped total (vbaOptionStringCharCount style.BackgroundColor)
        match style.Text with
        | Some textStyle ->
            total <- vbaAddCapped total (vbaTextStyleStringCharCount textStyle)
        | None -> ()
        total

    let private vbaSelectionStringCharCount (selection: Doc.Selection option) =
        match selection with
        | Some selection ->
            let mutable total = selection.Anchor.BlockId.Length
            total <- vbaAddCapped total selection.Focus.BlockId.Length
            total
        | None -> 0

    let private vbaPageSetupStringCharCount (setup: Doc.PageSetup) =
        let mutable total = vbaOptionStringCharCount setup.SectionStart
        total <- vbaAddCapped total (vbaSectionColumnsStringCharCount setup.Columns)
        total

    let private vbaHeaderFooterPartStringCharCount (part: Doc.HeaderFooterPart) =
        let mutable total = part.Kind.Length
        total <- vbaAddCapped total (vbaOptionStringCharCount part.RelationshipId)
        total

    let private countDocumentBlocksSourceTextChars (blocks: Doc.Block list) =
        let limit = vbaSourceTextCacheMaxTotalChars + 1
        let mutable total = 0
        let mutable visitedItems = 0
        let canContinue () =
            total <= vbaSourceTextCacheMaxTotalChars
            && visitedItems <= vbaSourceParagraphCountLimit
        let visitItem () =
            visitedItems <- visitedItems + 1
        let addParagraph (paragraph: Doc.ParagraphBlock) =
            if canContinue () then
                visitItem ()
                total <- vbaAddCapped total (vbaParagraphSourceTextCharCount paragraph)
        let addTable (table: Doc.TableBlock) =
            let tables = Stack<Doc.TableBlock>()
            tables.Push table
            while tables.Count > 0 && canContinue () do
                let table = tables.Pop()
                visitItem ()
                total <- vbaAddCapped total (vbaTableStringCharCount table)
                let mutable rows = table.Rows
                while canContinue () && not rows.IsEmpty do
                    let row = rows.Head
                    visitItem ()
                    total <- vbaAddCapped total (vbaTableRowStringCharCount row)
                    let mutable cells = row.Cells
                    while canContinue () && not cells.IsEmpty do
                        let cell = cells.Head
                        visitItem ()
                        total <- vbaAddCapped total (vbaTableCellStringCharCount cell)
                        let mutable paragraphs = cell.Content
                        while canContinue () && not paragraphs.IsEmpty do
                            addParagraph paragraphs.Head
                            paragraphs <- paragraphs.Tail
                        let mutable nestedTables = cell.NestedTables
                        while canContinue () && not nestedTables.IsEmpty do
                            tables.Push nestedTables.Head
                            nestedTables <- nestedTables.Tail
                        cells <- cells.Tail
                    rows <- rows.Tail
        let mutable remaining = blocks
        while canContinue () && not remaining.IsEmpty do
            match remaining.Head with
            | Doc.Paragraph paragraph
            | Doc.Heading (_, paragraph) ->
                addParagraph paragraph
            | Doc.Table table ->
                addTable table
            | Doc.TextBox textBox ->
                visitItem ()
                total <- vbaAddCapped total (vbaTextBoxStringCharCount textBox)
                let mutable paragraphs = textBox.Content
                while canContinue () && not paragraphs.IsEmpty do
                    addParagraph paragraphs.Head
                    paragraphs <- paragraphs.Tail
            remaining <- remaining.Tail
        if visitedItems > vbaSourceParagraphCountLimit then limit
        else min total limit

    let private countDocumentBlocksSourceTextCharsFor sourceParagraphCount blocks =
        if sourceParagraphCount > vbaSourceParagraphCountLimit then
            vbaSourceTextCacheMaxTotalChars + 1
        else
            countDocumentBlocksSourceTextChars blocks

    let private countDocumentSourceTextChars (doc: Doc.Document) =
        let limit = vbaSourceTextCacheMaxTotalChars + 1
        let mutable total = 0
        let mutable visitedItems = 0
        let canContinue () =
            total <= vbaSourceTextCacheMaxTotalChars
            && visitedItems <= vbaSourceParagraphCountLimit
        let visitItem () =
            visitedItems <- visitedItems + 1
        let addText textCharCount =
            if canContinue () then
                total <- vbaAddCapped total textCharCount
        let addParagraph (paragraph: Doc.ParagraphBlock) =
            if canContinue () then
                visitItem ()
                addText (vbaParagraphSourceTextCharCount paragraph)
        addText (vbaSelectionStringCharCount doc.Selection)
        addText (vbaPageSetupStringCharCount doc.PageSetup)
        let mutable styles = doc.Styles
        while canContinue () && not styles.IsEmpty do
            visitItem ()
            addText (vbaNamedStyleStringCharCount styles.Head)
            styles <- styles.Tail
        if canContinue () then
            addText (countDocumentBlocksSourceTextChars doc.Blocks)
        let mutable headers = doc.Headers
        while canContinue () && not headers.IsEmpty do
            addParagraph headers.Head
            headers <- headers.Tail
        let mutable footers = doc.Footers
        while canContinue () && not footers.IsEmpty do
            addParagraph footers.Head
            footers <- footers.Tail
        let mutable headerParts = doc.HeaderParts
        while canContinue () && not headerParts.IsEmpty do
            let part = headerParts.Head
            visitItem ()
            addText (vbaHeaderFooterPartStringCharCount part)
            if canContinue () then
                addText (countDocumentBlocksSourceTextChars part.Blocks)
            headerParts <- headerParts.Tail
        let mutable footerParts = doc.FooterParts
        while canContinue () && not footerParts.IsEmpty do
            let part = footerParts.Head
            visitItem ()
            addText (vbaHeaderFooterPartStringCharCount part)
            if canContinue () then
                addText (countDocumentBlocksSourceTextChars part.Blocks)
            footerParts <- footerParts.Tail
        if visitedItems > vbaSourceParagraphCountLimit then limit
        else min total limit

    let private layoutInputsEqual (cached: Doc.Document) (doc: Doc.Document) =
        obj.ReferenceEquals(cached.Blocks, doc.Blocks)
        && cached.PageSetup = doc.PageSetup
        && obj.ReferenceEquals(cached.Styles, doc.Styles)
        && obj.ReferenceEquals(cached.Headers, doc.Headers)
        && obj.ReferenceEquals(cached.Footers, doc.Footers)
        && obj.ReferenceEquals(cached.HeaderParts, doc.HeaderParts)
        && obj.ReferenceEquals(cached.FooterParts, doc.FooterParts)

    let private setLayoutCache entries =
        layoutCache <- entries
        layoutCacheByBlocks.Clear()
        clearArray layoutFastSlots
        for entry in entries do
            let blocks = entry.CachedDoc.Blocks
            let mutable candidates = Unchecked.defaultof<ResizeArray<LayoutCache>>
            if not (layoutCacheByBlocks.TryGetValue(blocks, &candidates)) then
                candidates <- ResizeArray<LayoutCache>()
                layoutCacheByBlocks.[blocks] <- candidates
            candidates.Add entry
            let slot = documentBlocksSlot blocks layoutFastSlots.Length
            if obj.ReferenceEquals(layoutFastSlots.[slot], null) then
                layoutFastSlots.[slot] <- entry

    let private tryFindLayoutCacheCandidate (doc: Doc.Document) (candidates: ResizeArray<LayoutCache>) =
        let mutable index = 0
        let mutable found = false
        let mutable result = Unchecked.defaultof<LayoutCache>
        while not found && index < candidates.Count do
            let candidate = candidates.[index]
            if layoutInputsEqual candidate.CachedDoc doc then
                found <- true
                result <- candidate
            index <- index + 1
        if found then ValueSome result else ValueNone

    let private tryGetLayoutCache (doc: Doc.Document) =
        let slot = documentBlocksSlot doc.Blocks layoutFastSlots.Length
        let fast = layoutFastSlots.[slot]
        if not (obj.ReferenceEquals(fast, null))
           && layoutInputsEqual fast.CachedDoc doc then
            ValueSome fast
        else
            let mutable candidates = Unchecked.defaultof<ResizeArray<LayoutCache>>
            if layoutCacheByBlocks.TryGetValue(doc.Blocks, &candidates) then
                match tryFindLayoutCacheCandidate doc candidates with
                | ValueSome cached ->
                    layoutFastSlots.[slot] <- cached
                    ValueSome cached
                | ValueNone ->
                    ValueNone
            else
                ValueNone

    let private layoutCacheWithoutDoc (doc: Doc.Document) =
        match tryGetLayoutCache doc with
        | ValueSome _ ->
            layoutCache
            |> List.filter (fun existing -> not (layoutInputsEqual existing.CachedDoc doc))
        | ValueNone ->
            layoutCache

    let private rememberLayoutCache cached =
        let trim entries =
            let rec loop remaining kept count totalLayouts totalSourceTextChars =
                match remaining with
                | [] -> List.rev kept
                | entry :: rest ->
                    let layoutCount = max 1 entry.Value.LayoutCount
                    let sourceTextCharCount = max 1 entry.Value.SourceTextCharCount
                    let nextTotalLayouts = totalLayouts + layoutCount
                    let nextTotalSourceTextChars = totalSourceTextChars + sourceTextCharCount
                    if count < layoutCacheMaxEntries
                       && nextTotalLayouts <= layoutCacheMaxTotalLayouts
                       && nextTotalSourceTextChars <= vbaSourceTextCacheMaxTotalChars then
                        loop rest (entry :: kept) (count + 1) nextTotalLayouts nextTotalSourceTextChars
                    else
                        loop rest kept count totalLayouts totalSourceTextChars
            loop entries [] 0 0 0
        cached
        :: layoutCacheWithoutDoc cached.CachedDoc
        |> trim
        |> setLayoutCache

    let private layoutSegmentRetainedTextCharCount segment =
        match segment with
        | WordLayout.TextSegment (text, _, _, _, _) ->
            if isNull text then 0 else text.Length
        | WordLayout.ImageSegment (src, _, _, _, _, _, _, wrap, relativeHeight, cropLeft, cropTop, cropRight, cropBottom) ->
            let mutable total = if isNull src then 0 else src.Length
            total <- vbaAddCapped total (vbaOptionStringCharCount wrap)
            total <- vbaAddCapped total (vbaOptionStringCharCount relativeHeight)
            total <- vbaAddCapped total (vbaOptionStringCharCount cropLeft)
            total <- vbaAddCapped total (vbaOptionStringCharCount cropTop)
            total <- vbaAddCapped total (vbaOptionStringCharCount cropRight)
            total <- vbaAddCapped total (vbaOptionStringCharCount cropBottom)
            total
        | WordLayout.MathSegment _ ->
            0

    let private countLayoutRetainedTextChars (layouts: WordLayout.BlockLayout list) =
        let limit = vbaSourceTextCacheMaxTotalChars + 1
        let mutable total = 0
        let mutable remainingLayouts = layouts
        while total <= vbaSourceTextCacheMaxTotalChars
              && not remainingLayouts.IsEmpty do
            let layout = remainingLayouts.Head
            let mutable remainingLines = layout.Lines
            while total <= vbaSourceTextCacheMaxTotalChars
                  && not remainingLines.IsEmpty do
                let line = remainingLines.Head
                let mutable remainingSegments = line.Segments
                while total <= vbaSourceTextCacheMaxTotalChars
                      && not remainingSegments.IsEmpty do
                    total <- vbaAddCapped total (layoutSegmentRetainedTextCharCount remainingSegments.Head)
                    remainingSegments <- remainingSegments.Tail
                remainingLines <- remainingLines.Tail
            remainingLayouts <- remainingLayouts.Tail
        min total limit

    let private layoutRetainedTextCost textCharCount =
        vbaAddCapped textCharCount textCharCount

    let private buildLayoutCacheValue doc =
        let layouts = WordLayout.layoutDocumentWithRange (nativeSettings doc) measureText measureRange doc
        let layoutById = Dictionary<Doc.BlockId, WordLayout.BlockLayout>()
        let layoutsByBlock = Dictionary<Doc.BlockId, ResizeArray<WordLayout.BlockLayout>>()
        let pageOfY y =
            int (floor (max 0.0 y / (doc.PageSetup.PageHeight + WordLayout.pageGap))) + 1
        let mutable pageCount = 1
        for layout in layouts do
            if not (layoutById.ContainsKey layout.BlockId) then
                layoutById.[layout.BlockId] <- layout
            match layoutsByBlock.TryGetValue layout.BlockId with
            | true, blockLayouts -> blockLayouts.Add layout
            | false, _ ->
                let blockLayouts = ResizeArray<WordLayout.BlockLayout>()
                blockLayouts.Add layout
                layoutsByBlock.[layout.BlockId] <- blockLayouts
            for line in layout.Lines do
                pageCount <- max pageCount (pageOfY (line.Y + line.Height))
        let layoutCount = layouts.Length
        let sourceTextCharCount =
            if layoutCount > layoutCacheMaxTotalLayouts then
                vbaSourceTextCacheMaxTotalChars + 1
            else
                let sourceTextCharCount = countDocumentSourceTextChars doc
                let layoutTextCost = countLayoutRetainedTextChars layouts |> layoutRetainedTextCost
                vbaAddCapped sourceTextCharCount layoutTextCost
        { SourceTextCharCount = sourceTextCharCount
          Layouts = layouts
          LayoutCount = layoutCount
          LayoutById = layoutById
          LayoutsByBlock = layoutsByBlock
          PageCount = pageCount }

    let private layoutCacheFor doc =
        match tryGetLayoutCache doc with
        | ValueSome cached ->
            cached.Value
        | ValueNone ->
            let value = buildLayoutCacheValue doc
            if value.LayoutCount <= layoutCacheMaxTotalLayouts
               && value.SourceTextCharCount <= vbaSourceTextCacheMaxTotalChars then
                rememberLayoutCache { CachedDoc = doc; Value = value }
            value

    let private layoutsFor doc =
        (layoutCacheFor doc).Layouts

    let private tryLayoutFor doc blockId =
        let cache = layoutCacheFor doc
        match cache.LayoutById.TryGetValue blockId with
        | true, layout -> Some layout
        | false, _ -> None

    let private layoutsForBlock doc blockId =
        let cache = layoutCacheFor doc
        match cache.LayoutsByBlock.TryGetValue blockId with
        | true, layouts -> Some layouts
        | false, _ -> None

    let private pageCountFor doc =
        (layoutCacheFor doc).PageCount

    let private pageForY (doc: Doc.Document) y =
        int (floor (max 0.0 y / (doc.PageSetup.PageHeight + WordLayout.pageGap))) + 1

    let private pageTop (doc: Doc.Document) page =
        float (page - 1) * (doc.PageSetup.PageHeight + WordLayout.pageGap)

    let private countTableSourceParagraphs (table: Doc.TableBlock) =
        let mutable count = 0
        let mutable visitedItems = 0
        let canContinue () =
            count <= vbaSourceParagraphCountLimit
            && visitedItems <= vbaSourceParagraphCountLimit
        let visitItem () =
            visitedItems <- visitedItems + 1
        let addParagraphs (paragraphs: Doc.ParagraphBlock list) =
            let mutable remaining = paragraphs
            while canContinue () && not remaining.IsEmpty do
                count <- count + 1
                remaining <- remaining.Tail
        let tables = Stack<Doc.TableBlock>()
        tables.Push table
        while canContinue () && tables.Count > 0 do
            let table = tables.Pop()
            visitItem ()
            let mutable rows = table.Rows
            while canContinue () && not rows.IsEmpty do
                let row = rows.Head
                visitItem ()
                let mutable cells = row.Cells
                while canContinue () && not cells.IsEmpty do
                    let cell = cells.Head
                    visitItem ()
                    addParagraphs cell.Content
                    let mutable nestedTables = cell.NestedTables
                    while canContinue () && not nestedTables.IsEmpty do
                        tables.Push nestedTables.Head
                        nestedTables <- nestedTables.Tail
                    cells <- cells.Tail
                rows <- rows.Tail
        if visitedItems > vbaSourceParagraphCountLimit then
            vbaSourceParagraphCountLimit + 1
        else
            min count (vbaSourceParagraphCountLimit + 1)

    let private countDocumentSourceParagraphs (doc: Doc.Document) =
        let mutable count = 0
        let addParagraphs (paragraphs: Doc.ParagraphBlock list) =
            let mutable remaining = paragraphs
            while count <= vbaSourceParagraphCountLimit && not remaining.IsEmpty do
                count <- count + 1
                remaining <- remaining.Tail
        let mutable remaining = doc.Blocks
        while count <= vbaSourceParagraphCountLimit && not remaining.IsEmpty do
            match remaining.Head with
            | Doc.Paragraph _
            | Doc.Heading _ ->
                count <- count + 1
            | Doc.Table table ->
                count <- count + countTableSourceParagraphs table
            | Doc.TextBox textBox ->
                addParagraphs textBox.Content
            remaining <- remaining.Tail
        min count (vbaSourceParagraphCountLimit + 1)

    type private DocumentTextCache = {
        SourceBlocks: Doc.Block list
        SourceParagraphCount: int
        SourceTextCharCount: int
        TextLength: int
        Text: string
    }

    let mutable private documentTextCache: DocumentTextCache list = []
    let private documentTextCacheByBlocks =
        Dictionary<Doc.Block list, DocumentTextCache>(documentBlocksReferenceComparer)
    let private documentTextFastSlots = Array.zeroCreate<DocumentTextCache> 256
    let private documentTextCacheMaxEntries = 256
    let private documentTextCacheMaxTotalChars = 5000000
    let private documentTextCacheMaxTotalSourceParagraphs = vbaSourceParagraphCountLimit
    let private documentTextCacheMaxTotalSourceTextChars = vbaSourceTextCacheMaxTotalChars

    let private documentTextCacheMatches blocks (cached: DocumentTextCache) =
        obj.ReferenceEquals(cached.SourceBlocks, blocks)

    let private documentTextCacheWithoutBlocks blocks =
        let mutable cached = Unchecked.defaultof<DocumentTextCache>
        if documentTextCacheByBlocks.TryGetValue(blocks, &cached)
           && documentTextCacheMatches blocks cached then
            documentTextCache
            |> List.filter (documentTextCacheMatches blocks >> not)
        else
            documentTextCache

    let private setDocumentTextCache entries =
        documentTextCache <- entries
        documentTextCacheByBlocks.Clear()
        clearArray documentTextFastSlots
        for entry in entries do
            let mutable existing = Unchecked.defaultof<DocumentTextCache>
            if not (documentTextCacheByBlocks.TryGetValue(entry.SourceBlocks, &existing)) then
                documentTextCacheByBlocks.[entry.SourceBlocks] <- entry
                documentTextFastSlots.[documentBlocksSlot entry.SourceBlocks documentTextFastSlots.Length] <- entry

    let private tryGetDocumentTextCache blocks =
        let slot = documentBlocksSlot blocks documentTextFastSlots.Length
        let fast = documentTextFastSlots.[slot]
        if not (obj.ReferenceEquals(fast, null))
           && documentTextCacheMatches blocks fast then
            ValueSome fast
        else
            let mutable cached = Unchecked.defaultof<DocumentTextCache>
            if documentTextCacheByBlocks.TryGetValue(blocks, &cached)
               && documentTextCacheMatches blocks cached then
                documentTextFastSlots.[slot] <- cached
                ValueSome cached
            else
                ValueNone

    let private documentTextSourceParagraphCountFor textLength doc =
        if textLength > documentTextCacheMaxTotalChars then
            documentTextCacheMaxTotalSourceParagraphs + 1
        else
            countDocumentSourceParagraphs doc

    let private documentTextSourceTextCharCountFor textLength sourceParagraphCount blocks =
        if textLength > documentTextCacheMaxTotalChars
           || sourceParagraphCount > documentTextCacheMaxTotalSourceParagraphs then
            documentTextCacheMaxTotalSourceTextChars + 1
        else
            countDocumentBlocksSourceTextCharsFor sourceParagraphCount blocks

    let private rememberDocumentTextCache cached =
        let trim entries =
            let rec loop remaining kept count totalChars totalBlocks totalSourceTextChars =
                match remaining with
                | [] -> List.rev kept
                | entry :: rest ->
                    let textLength = max 1 entry.TextLength
                    let sourceParagraphCount = max 1 entry.SourceParagraphCount
                    let sourceTextCharCount = max 1 entry.SourceTextCharCount
                    let nextTotalChars = totalChars + textLength
                    let nextTotalBlocks = totalBlocks + sourceParagraphCount
                    let nextTotalSourceTextChars = totalSourceTextChars + sourceTextCharCount
                    if count < documentTextCacheMaxEntries
                       && nextTotalChars <= documentTextCacheMaxTotalChars
                       && nextTotalBlocks <= documentTextCacheMaxTotalSourceParagraphs
                       && nextTotalSourceTextChars <= documentTextCacheMaxTotalSourceTextChars then
                        loop rest (entry :: kept) (count + 1) nextTotalChars nextTotalBlocks nextTotalSourceTextChars
                    else
                        loop rest kept count totalChars totalBlocks totalSourceTextChars
            loop entries [] 0 0 0 0
        cached
        :: documentTextCacheWithoutBlocks cached.SourceBlocks
        |> trim
        |> setDocumentTextCache

    let private documentText (doc: Doc.Document) =
        match tryGetDocumentTextCache doc.Blocks with
        | ValueSome cached ->
            cached.Text
        | ValueNone ->
            let text = Doc.getText doc
            let sourceParagraphCount = documentTextSourceParagraphCountFor text.Length doc
            rememberDocumentTextCache {
                SourceBlocks = doc.Blocks
                SourceParagraphCount = sourceParagraphCount
                SourceTextCharCount = documentTextSourceTextCharCountFor text.Length sourceParagraphCount doc.Blocks
                TextLength = text.Length
                Text = text
            }
            text

    let private rangeText (doc: Doc.Document) startOffset endOffset =
        let text = documentText doc
        let start = max 0 (min startOffset text.Length)
        let finish = max start (min endOffset text.Length)
        if start = 0 && finish = text.Length then text
        else text.Substring(start, finish - start)

    type private TextStatisticsCacheKey = {
        StartOffset: int
        EndOffset: int
        Constant: int
    }

    type private TextStatisticsCache = {
        SourceBlocks: Doc.Block list
        SourceParagraphCount: int
        SourceTextCharCount: int
        Entries: Dictionary<TextStatisticsCacheKey, Value>
        Order: Queue<TextStatisticsCacheKey>
    }

    let mutable private textStatisticsCache: TextStatisticsCache list = []
    let private textStatisticsCacheMaxEntries = 64
    let private textStatisticsCacheMaxSources = 64
    let private textStatisticsCacheMaxTotalSourceParagraphs = vbaSourceParagraphCountLimit
    let private textStatisticsCacheMaxTotalSourceTextChars = vbaSourceTextCacheMaxTotalChars
    let private textStatisticsCacheByBlocks =
        Dictionary<Doc.Block list, TextStatisticsCache>(documentBlocksReferenceComparer)
    let private textStatisticsFastSlots = Array.zeroCreate<TextStatisticsCache> 256

    let private textStatisticsCacheMatches blocks (cached: TextStatisticsCache) =
        obj.ReferenceEquals(cached.SourceBlocks, blocks)

    let private textStatisticsCacheSlot blocks =
        documentBlocksSlot blocks textStatisticsFastSlots.Length

    let private setTextStatisticsCache entries =
        textStatisticsCache <- entries
        textStatisticsCacheByBlocks.Clear()
        clearArray textStatisticsFastSlots
        for entry in entries do
            let mutable existing = Unchecked.defaultof<TextStatisticsCache>
            if not (textStatisticsCacheByBlocks.TryGetValue(entry.SourceBlocks, &existing)) then
                textStatisticsCacheByBlocks.[entry.SourceBlocks] <- entry
                textStatisticsFastSlots.[textStatisticsCacheSlot entry.SourceBlocks] <- entry

    let private tryGetTextStatisticsCache blocks =
        let slot = textStatisticsCacheSlot blocks
        let fast = textStatisticsFastSlots.[slot]
        if not (obj.ReferenceEquals(fast, null))
           && textStatisticsCacheMatches blocks fast then
            ValueSome fast
        else
            let mutable cached = Unchecked.defaultof<TextStatisticsCache>
            if textStatisticsCacheByBlocks.TryGetValue(blocks, &cached)
               && textStatisticsCacheMatches blocks cached then
                textStatisticsFastSlots.[slot] <- cached
                ValueSome cached
            else
                ValueNone

    let private textStatisticsCacheWithoutBlocks blocks =
        let mutable cached = Unchecked.defaultof<TextStatisticsCache>
        if textStatisticsCacheByBlocks.TryGetValue(blocks, &cached) then
            textStatisticsCache
            |> List.filter (textStatisticsCacheMatches blocks >> not)
        else
            textStatisticsCache

    let private textStatisticsCacheRetainable (cached: TextStatisticsCache) =
        cached.Entries.Count > 0
        && cached.Entries.Count <= textStatisticsCacheMaxEntries
        && cached.SourceParagraphCount <= textStatisticsCacheMaxTotalSourceParagraphs
        && cached.SourceTextCharCount <= textStatisticsCacheMaxTotalSourceTextChars

    let private rememberTextStatisticsCache cached =
        if not (textStatisticsCacheRetainable cached) then
            let entries = textStatisticsCacheWithoutBlocks cached.SourceBlocks
            if not (obj.ReferenceEquals(entries, textStatisticsCache)) then
                setTextStatisticsCache entries
        else
            let mutable keptCount = 1
            let mutable keptSourceParagraphs = max 1 cached.SourceParagraphCount
            let mutable keptSourceTextChars = max 1 cached.SourceTextCharCount
            let retained =
                textStatisticsCacheWithoutBlocks cached.SourceBlocks
                |> List.choose (fun candidate ->
                    let candidateSourceParagraphs = max 1 candidate.SourceParagraphCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    if textStatisticsCacheRetainable candidate
                       && keptCount < textStatisticsCacheMaxSources
                       && keptSourceParagraphs + candidateSourceParagraphs <= textStatisticsCacheMaxTotalSourceParagraphs
                       && keptSourceTextChars + candidateSourceTextChars <= textStatisticsCacheMaxTotalSourceTextChars then
                        keptCount <- keptCount + 1
                        keptSourceParagraphs <- keptSourceParagraphs + candidateSourceParagraphs
                        keptSourceTextChars <- keptSourceTextChars + candidateSourceTextChars
                        Some candidate
                    else
                        None)
            setTextStatisticsCache (cached :: retained)

    let private textStatisticsCacheKey startOffset endOffset constant = {
        StartOffset = startOffset
        EndOffset = endOffset
        Constant = constant
    }

    let private addTextStatisticsCacheEntry key value (entries: Dictionary<TextStatisticsCacheKey, Value>) (order: Queue<TextStatisticsCacheKey>) =
        if not (entries.ContainsKey key) then
            order.Enqueue key
        entries.[key] <- value
        while entries.Count > textStatisticsCacheMaxEntries && order.Count > 0 do
            let oldest = order.Dequeue()
            if oldest <> key then
                entries.Remove oldest |> ignore
        if entries.Count > textStatisticsCacheMaxEntries then
            entries.Clear()
            order.Clear()
            order.Enqueue key
            entries.[key] <- value

    let private textStatisticsRangeTextLength startOffset endOffset =
        max 0 (endOffset - startOffset)

    let private textStatisticsSourceParagraphCountFor rangeTextLength doc =
        if rangeTextLength > textStatisticsCacheMaxTotalSourceTextChars then
            textStatisticsCacheMaxTotalSourceParagraphs + 1
        else
            countDocumentSourceParagraphs doc

    let private textStatisticsSourceTextCharCountFor rangeTextLength sourceParagraphCount blocks =
        if rangeTextLength > textStatisticsCacheMaxTotalSourceTextChars
           || sourceParagraphCount > textStatisticsCacheMaxTotalSourceParagraphs then
            textStatisticsCacheMaxTotalSourceTextChars + 1
        else
            countDocumentBlocksSourceTextCharsFor sourceParagraphCount blocks

    let private tryTextStatisticsCache (doc: Doc.Document) startOffset endOffset constant =
        let key = textStatisticsCacheKey startOffset endOffset constant
        match tryGetTextStatisticsCache doc.Blocks with
        | ValueSome cached when textStatisticsCacheRetainable cached ->
            let mutable value = VEmpty
            if cached.Entries.TryGetValue(key, &value) then Some value else None
        | ValueSome cached ->
            rememberTextStatisticsCache cached
            None
        | ValueNone -> None

    let private rememberTextStatistics (doc: Doc.Document) startOffset endOffset constant (value: Value) =
        let key = textStatisticsCacheKey startOffset endOffset constant
        match tryGetTextStatisticsCache doc.Blocks with
        | ValueSome cached ->
            addTextStatisticsCacheEntry key value cached.Entries cached.Order
            rememberTextStatisticsCache cached
        | ValueNone ->
            let entries = Dictionary<TextStatisticsCacheKey, Value>()
            let order = Queue<TextStatisticsCacheKey>()
            addTextStatisticsCacheEntry key value entries order
            let rangeTextLength = textStatisticsRangeTextLength startOffset endOffset
            let sourceParagraphCount = textStatisticsSourceParagraphCountFor rangeTextLength doc
            rememberTextStatisticsCache {
                SourceBlocks = doc.Blocks
                SourceParagraphCount = sourceParagraphCount
                SourceTextCharCount = textStatisticsSourceTextCharCountFor rangeTextLength sourceParagraphCount doc.Blocks
                Entries = entries
                Order = order
            }

    let private computeTextStatisticByConstant (doc: Doc.Document) startOffset endOffset constant =
        match constant with
        | 0 -> VInteger (countWords (rangeText doc startOffset endOffset))
        | 3 ->
            let text = rangeText doc startOffset endOffset
            VInteger text.Length
        | 4 ->
            let text = rangeText doc startOffset endOffset
            VInteger (countLines text)
        | 5 ->
            let text = rangeText doc startOffset endOffset
            VInteger text.Length
        | _ -> VEmpty

    let private textStatisticByConstant (doc: Doc.Document) startOffset endOffset constant =
        match tryTextStatisticsCache doc startOffset endOffset constant with
        | Some value -> value
        | None ->
            let value = computeTextStatisticByConstant doc startOffset endOffset constant
            rememberTextStatistics doc startOffset endOffset constant value
            value

    let private statisticsByConstant (doc: Doc.Document) startOffset endOffset constant =
        match constant with
        | 0
        | 3
        | 4
        | 5 -> textStatisticByConstant doc startOffset endOffset constant
        | 1
        | 2 ->
            let blockId, localStart = resolvePosition doc startOffset
            let _, localEnd = resolvePosition doc endOffset
            let pages = System.Collections.Generic.HashSet<int>()
            let mutable lineCount = 0
            match layoutsForBlock doc blockId with
            | None -> ()
            | Some layouts ->
                for layout in layouts do
                    for line in layout.Lines do
                        let intersects =
                            if localStart = localEnd then
                                localStart >= line.StartOffset && localStart <= line.EndOffset
                            else
                                localStart < line.EndOffset && localEnd > line.StartOffset
                        if intersects then
                            lineCount <- lineCount + 1
                            pages.Add(pageForY doc line.Y) |> ignore
            if constant = 1 then VInteger lineCount else VInteger pages.Count
        | _ -> VEmpty

    let private rangeInfoByConstant (doc: Doc.Document) startOffset endOffset constant =
        match constant with
        | 12 | 31 | 44 | 45 | 46 -> VBoolean false
        | 4 ->
            VInteger (pageCountFor doc)
        | 1 | 2 | 3 | 5 | 6 | 7 | 8 | 9 | 10 ->
            let blockId, localOffset = resolvePosition doc startOffset
            match tryLayoutFor doc blockId with
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
                | 5 -> VDouble x
                | 6 -> VDouble pageY
                | 7 -> VDouble (x - doc.PageSetup.MarginLeft)
                | 8 -> VDouble (max 0.0 (line.Y - doc.PageSetup.MarginTop))
                | 9 -> VInteger (max 1 (localOffset - line.StartOffset + 1))
                | 10 -> VInteger 1
                | _ -> VEmpty
        | _ -> VEmpty

    let private rangeBoundsObject (doc: Doc.Document) startOffset endOffset =
        let mutable cachedBounds: (float * float * float * float * float) option option = None

        let bounds () =
            match cachedBounds with
            | Some value -> value
            | None ->
                let blockId, localStart = resolvePosition doc startOffset
                let _, localEnd = resolvePosition doc endOffset
                let value =
                    match tryLayoutFor doc blockId with
                    | None -> None
                    | Some layout ->
                        let line = WordLayout.lineForOffset layout localStart
                        let alignOffset = WordLayout.alignmentOffset layout line
                        let x1 = doc.PageSetup.MarginLeft + alignOffset + WordLayout.xForOffsetWithRange measureRange line localStart
                        let x2 = doc.PageSetup.MarginLeft + alignOffset + WordLayout.xForOffsetWithRange measureRange line (max localStart localEnd)
                        let page = pageForY doc line.Y
                        Some (x1, line.Y, line.Y - pageTop doc page, max 0.0 (x2 - x1), line.Height)
                cachedBounds <- Some value
                value

        host "Word.RangeBounds"
            (fun (field: string) ->
                match bounds () with
                | None -> Some VEmpty
                | Some (x, y, pageY, width, height) ->
                    if eq field "X" then Some (VDouble x)
                    elif eq field "Y" then Some (VDouble y)
                    elif eq field "PageY" then Some (VDouble pageY)
                    elif eq field "Width" then Some (VDouble width)
                    elif eq field "Height" then Some (VDouble height)
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
                    let mutable found = false
                    iterParagraphs doc (fun para ->
                        if not found && para.Id = sel.Focus.BlockId then
                            result <- acc + sel.Focus.Offset
                            found <- true
                        if not found then
                            acc <- acc + TextBuf.length para.Buffer + 1)
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
                elif eq name "Content" then Some (rangeObject doc 0 (documentText doc).Length)
                elif eq name "Words" then
                    Some (host "Word.Words"
                        (fun field -> if eq field "Count" then Some (statisticsByConstant doc 0 (documentText doc).Length 0) else None)
                        (fun _ _ -> None))
                elif eq name "Characters" then
                    Some (host "Word.Characters"
                        (fun field -> if eq field "Count" then Some (statisticsByConstant doc 0 (documentText doc).Length 3) else None)
                        (fun _ _ -> None))
                elif eq name "Paragraphs" then
                    Some (host "Word.Paragraphs"
                        (fun field -> if eq field "Count" then Some (VInteger (paragraphCount doc)) else None)
                        (fun _ _ -> None))
                else None)
            (fun (name: string) (args: Value array) ->
                if eq name "Range" then
                    let startOffset = if args.Length >= 1 then valueToInt args.[0] else 0
                    let endOffset = if args.Length >= 2 then valueToInt args.[1] else startOffset
                    Some (rangeObject doc startOffset endOffset)
                elif eq name "ComputeStatistics" && args.Length >= 1 then
                    Some (statisticsByConstant doc 0 (documentText doc).Length (valueToInt args.[0]))
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
            FsVb.RunOptions.vba
            |> FsVb.RunOptions.withHostGlobals (hostGlobalsDocument doc)
        runSourceCached options source

    let runDocumentVbs (doc: Doc.Document) (source: string) =
        let options =
            FsVb.RunOptions.vbScript
            |> FsVb.RunOptions.withHostGlobals (hostGlobalsDocumentVbs doc)
        runSourceCached options source

    let hostGlobals (_: API.Editor) : Map<string, Value> = Map.empty
    let run (_: API.Editor) (_: string) =
        Error (FsVb.RuntimeError "Use Word.Vba.runDocument for native Word VBA host execution")
    let runVbs (_: API.Editor) (_: string) =
        Error (FsVb.RuntimeError "Use Word.Vba.runDocumentVbs for native Word VBS host execution")
#endif
