module RedBat.FableUI.Reactive

open Fable.Core
open System.Collections.Generic

[<Emit("queueMicrotask($0)")>]
let private queueMicrotask (fn: unit -> unit) : unit = jsNative

type Computation = {
    Execute: unit -> unit
    mutable Cleanups: ResizeArray<unit -> unit>
    Level: int
}

let mutable private currentComputation: Computation option = None
let private computationStack = Stack<Computation>()
let private pendingComputations = HashSet<Computation>()
let mutable private isScheduled = false
let mutable private isFlushing = false

let private flushPending () =
    isFlushing <- true
    while pendingComputations.Count > 0 do
        let computations = pendingComputations |> Seq.toArray |> Array.sortBy _.Level
        pendingComputations.Clear()
        for computation in computations do
            computation.Execute()
    isFlushing <- false
    isScheduled <- false

let private scheduleComputation computation =
    pendingComputations.Add(computation) |> ignore
    if not isScheduled then
        isScheduled <- true
        if not isFlushing then
            queueMicrotask flushPending

let batch (fn: unit -> unit) = fn ()

let untrack (fn: unit -> 'T) : 'T =
    let saved = currentComputation
    currentComputation <- None
    try fn ()
    finally currentComputation <- saved

[<AttachMembers>]
type Signal<'T when 'T: equality>(initial: 'T) =
    let mutable value = initial
    let subscribers = HashSet<Computation>()

    member _.Value
        with get () =
            match currentComputation with
            | Some computation ->
                if subscribers.Add(computation) then
                    computation.Cleanups.Add(fun () -> subscribers.Remove(computation) |> ignore)
            | None -> ()
            value
        and set newValue =
            if not (EqualityComparer<'T>.Default.Equals(value, newValue)) then
                value <- newValue
                let currentSubscribers = subscribers |> Seq.toArray
                for subscriber in currentSubscribers do
                    scheduleComputation subscriber

    member _.Peek() = value
    member this.Update(update: 'T -> 'T) = this.Value <- update this.Value

let signal<'T when 'T: equality> initial = Signal<'T>(initial)

[<AttachMembers>]
type EffectDisposable(disposeAction: unit -> unit) =
    member _.Dispose() = disposeAction()

let effect (fn: unit -> unit) : EffectDisposable =
    let mutable isDisposed = false
    let mutable cleanups = ResizeArray<unit -> unit>()

    let rec execute () =
        if not isDisposed then
            for cleanup in cleanups do cleanup()
            cleanups.Clear()

            match currentComputation with
            | Some outer -> computationStack.Push(outer)
            | None -> ()

            let computation = { Execute = execute; Cleanups = cleanups; Level = 1 }
            currentComputation <- Some computation
            try fn ()
            finally
                currentComputation <- if computationStack.Count > 0 then Some(computationStack.Pop()) else None
                cleanups <- computation.Cleanups

    execute ()

    EffectDisposable(fun () ->
        isDisposed <- true
        for cleanup in cleanups do cleanup()
        cleanups.Clear())

[<AttachMembers>]
type Computed<'T when 'T: equality>(compute: unit -> 'T) =
    let mutable cachedValue: 'T option = None
    let mutable isDirty = true
    let subscribers = HashSet<Computation>()
    let mutable cleanups = ResizeArray<unit -> unit>()

    let markDirty () =
        if not isDirty then
            isDirty <- true
            let currentSubscribers = subscribers |> Seq.toArray
            for subscriber in currentSubscribers do
                scheduleComputation subscriber

    member _.Value
        with get () =
            match currentComputation with
            | Some computation ->
                if subscribers.Add(computation) then
                    computation.Cleanups.Add(fun () -> subscribers.Remove(computation) |> ignore)
            | None -> ()

            if isDirty then
                for cleanup in cleanups do cleanup()
                cleanups.Clear()
                let saved = currentComputation
                let self = { Execute = markDirty; Cleanups = cleanups; Level = 0 }
                currentComputation <- Some self
                try
                    cachedValue <- Some (compute ())
                    isDirty <- false
                    cleanups <- self.Cleanups
                finally
                    currentComputation <- saved

            cachedValue.Value

    member _.Peek() =
        match cachedValue with
        | Some value -> value
        | None -> compute ()

let computed<'T when 'T: equality> compute = Computed<'T>(compute)

let on (deps: unit -> 'T) (fn: 'T -> unit) : EffectDisposable =
    effect (fun () ->
        let depValue = deps ()
        untrack (fun () -> fn depValue))
