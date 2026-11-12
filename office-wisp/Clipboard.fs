module Wisp.OfficeClipboard

open System
open System.Diagnostics
open System.Runtime.InteropServices
open System.Text
open System.Threading
open System.Threading.Tasks

/// Token is present only while this process can prove that its own write still owns the selection.
type Contents = { Text: string; Token: string option }

module private X11 =
    [<DllImport("libX11.so.6", EntryPoint = "XOpenDisplay")>]
    extern nativeint openDisplay(string name)
    [<DllImport("libX11.so.6", EntryPoint = "XCloseDisplay")>]
    extern int closeDisplay(nativeint display)
    [<DllImport("libX11.so.6", EntryPoint = "XInternAtom")>]
    extern nativeint internAtom(nativeint display, string name, bool onlyIfExists)
    [<DllImport("libX11.so.6", EntryPoint = "XGetSelectionOwner")>]
    extern nativeint selectionOwner(nativeint display, nativeint selection)
    [<DllImport("libXRes.so.1", EntryPoint = "XResQueryExtension")>]
    extern int queryExtension(nativeint display, int& eventBase, int& errorBase)
    [<DllImport("libXRes.so.1", EntryPoint = "XResQueryVersion")>]
    extern int queryVersion(nativeint display, int& major, int& minor)
    [<DllImport("libXRes.so.1", EntryPoint = "XResQueryClientIds")>]
    extern int queryClientIds(nativeint display, nativeint numSpecs, nativeint specs,
                              nativeint& numIds, nativeint& ids)
    [<DllImport("libXRes.so.1", EntryPoint = "XResGetClientPid")>]
    extern int getClientPid(nativeint value)
    [<DllImport("libXRes.so.1", EntryPoint = "XResClientIdsDestroy")>]
    extern void destroyClientIds(nativeint count, nativeint ids)

    let private withDisplay action =
        let display = openDisplay(null)
        if display = 0n then failwith "Cannot connect to X11 DISPLAY for clipboard"
        try action display
        finally closeDisplay display |> ignore

    let private clipboardOwner display =
        selectionOwner(display, internAtom(display, "CLIPBOARD", false))

    let owner () = withDisplay clipboardOwner

    /// The X-Resource 1.2 server reports the PID of the client that owns this
    /// exact XID. Window changes or equal text alone never prove provenance.
    let ownerAndPid () =
        withDisplay (fun display ->
            let mutable eventBase = 0
            let mutable errorBase = 0
            let mutable major = 0
            let mutable minor = 0
            try
                if queryExtension(display, &eventBase, &errorBase) = 0
                   || queryVersion(display, &major, &minor) = 0
                   || major < 1 || (major = 1 && minor < 2) then
                    failwith "X11 clipboard ownership proof requires X-Resource extension 1.2"
            with :? DllNotFoundException ->
                failwith "X11 clipboard ownership proof requires libXRes.so.1"
            let before = clipboardOwner display
            if before = 0n then 0n, None
            else
                // XResClientIdSpec: XID (native long), unsigned mask, native padding.
                let spec = Marshal.AllocHGlobal(IntPtr.Size * 2)
                try
                    Marshal.WriteIntPtr(spec, before)
                    Marshal.WriteInt32(spec, IntPtr.Size, 2) // XRES_CLIENT_ID_PID_MASK
                    let mutable count = 0n
                    let mutable ids = 0n
                    try
                        // Unlike QueryExtension/QueryVersion, this API returns X11 Success (0).
                        if queryClientIds(display, 1n, spec, &count, &ids) <> 0 then
                            failwith "X-Resource client PID query failed"
                        if count < 0n || count > 16n then failwith "Invalid X-Resource client PID response"
                        let mutable pid = None
                        for index in 0 .. int count - 1 do
                            let value = nativeint (int64 ids + int64 (index * IntPtr.Size * 4))
                            let reported = getClientPid value
                            if reported > 0 then pid <- Some reported
                        if before <> clipboardOwner display then failwith "Clipboard changed during ownership query"
                        before, pid
                    finally
                        if ids <> 0n then destroyClientIds(count, ids)
                finally Marshal.FreeHGlobal spec)

module private Win32 =
    [<DllImport("user32.dll", EntryPoint = "CreateWindowExW", CharSet = CharSet.Unicode, SetLastError = true)>]
    extern nativeint CreateWindowEx(uint32 extendedStyle, string className, string title, uint32 style,
                                    int x, int y, int width, int height, nativeint parent,
                                    nativeint menu, nativeint instance, nativeint parameter)
    [<DllImport("user32.dll", SetLastError = true)>]
    extern bool OpenClipboard(nativeint window)
    [<DllImport("user32.dll", SetLastError = true)>]
    extern bool CloseClipboard()
    [<DllImport("user32.dll", SetLastError = true)>]
    extern bool EmptyClipboard()
    [<DllImport("user32.dll", SetLastError = true)>]
    extern nativeint GetClipboardData(uint32 format)
    [<DllImport("user32.dll", SetLastError = true)>]
    extern nativeint SetClipboardData(uint32 format, nativeint handle)
    [<DllImport("user32.dll")>]
    extern uint32 GetClipboardSequenceNumber()
    [<DllImport("user32.dll")>]
    extern nativeint GetClipboardOwner()
    [<DllImport("user32.dll")>]
    extern bool IsClipboardFormatAvailable(uint32 format)
    [<DllImport("kernel32.dll", SetLastError = true)>]
    extern nativeint GlobalAlloc(uint32 flags, unativeint bytes)
    [<DllImport("kernel32.dll", SetLastError = true)>]
    extern nativeint GlobalLock(nativeint handle)
    [<DllImport("kernel32.dll", SetLastError = true)>]
    extern bool GlobalUnlock(nativeint handle)
    [<DllImport("kernel32.dll", SetLastError = true)>]
    extern nativeint GlobalFree(nativeint handle)

    let mutable private clipboardWindow = 0n
    let withClipboard action =
        let mutable attempts = 0
        while attempts < 20 && not (OpenClipboard clipboardWindow) do
            attempts <- attempts + 1
            Thread.Sleep 20
        if attempts = 20 then failwith "Cannot open Windows clipboard (busy)"
        try action ()
        finally CloseClipboard() |> ignore

    let ensureOwnerWindow () =
        if clipboardWindow = 0n then
            // Message-only window owned by this process, not a second UI toolkit.
            let window = CreateWindowEx(0u, "STATIC", "", 0u, 0, 0, 0, 0, -3n, 0n, 0n, 0n)
            if window = 0n then failwith "Cannot create Windows clipboard owner window"
            clipboardWindow <- window

    let ownsClipboard () = clipboardWindow <> 0n && GetClipboardOwner() = clipboardWindow

let private gate = obj ()
let mutable private ownedToken: string option = None
let mutable private ownedXclip: Process option = None
let mutable private ownedX11Window = 0n
let mutable private ownedWinSequence = 0u
let private timeLimit = TimeSpan.FromSeconds 2.0

let private stop (child: Process) =
    try
        if not child.HasExited then child.Kill()
        child.WaitForExit(500) |> ignore
    with _ -> ()
    child.Dispose()

let private startXclip (args: string) =
    let info = ProcessStartInfo("xclip", args)
    info.UseShellExecute <- false
    info.RedirectStandardInput <- true
    info.RedirectStandardOutput <- true
    info.RedirectStandardError <- true
    info.StandardInputEncoding <- UTF8Encoding(false)
    info.StandardOutputEncoding <- Encoding.UTF8
    try Process.Start info
    with :? System.ComponentModel.Win32Exception as ex ->
        failwithf "xclip is required for X11 clipboard: %s" ex.Message

let private writeX11 (text: string) token =
    // Fail before starting a writer when the display cannot prove ownership.
    X11.ownerAndPid () |> ignore
    let before = X11.owner ()
    let child = startXclip "-selection clipboard -quiet -in"
    // Foreground xclip logs each request; drain it so repeated pastes cannot fill stderr.
    try
        child.BeginErrorReadLine()
        let writer = Task.Run(fun () ->
            child.StandardInput.Write(text)
            child.StandardInput.Close())
        if not (writer.Wait(timeLimit)) then failwith "xclip clipboard write timed out"
        writer.GetAwaiter().GetResult()
        let deadline = Stopwatch.StartNew()
        let firstWindow, firstPid = X11.ownerAndPid ()
        let mutable window = firstWindow
        let mutable pid = firstPid
        while (window = 0n || window = before || pid <> Some child.Id)
              && not child.HasExited && deadline.Elapsed < timeLimit do
            Thread.Sleep 15
            let observedWindow, observedPid = X11.ownerAndPid ()
            window <- observedWindow
            pid <- observedPid
        if child.HasExited || window = 0n || window = before || pid <> Some child.Id then
            let detail = if child.HasExited then sprintf "xclip exited with code %d" child.ExitCode else "selection ownership timed out"
            failwithf "xclip clipboard write failed: %s" detail
        ownedXclip |> Option.iter stop
        ownedXclip <- Some child
        ownedX11Window <- window
        ownedToken <- Some token
        token
    with _ ->
        stop child
        reraise ()

let private readX11 () =
    let before = X11.owner ()
    if before = 0n then
        ownedToken <- None
        Ok { Text = ""; Token = None }
    else
        use child = startXclip "-selection clipboard -quiet -out"
        let output = child.StandardOutput.ReadToEndAsync()
        if not (child.WaitForExit(int timeLimit.TotalMilliseconds) && output.Wait(timeLimit)) then
            stop child
            failwith "xclip clipboard read timed out"
        if child.ExitCode <> 0 then failwithf "xclip clipboard read failed: %s" (child.StandardError.ReadToEnd())
        let after, pid = X11.ownerAndPid ()
        if after <> before then failwith "Clipboard changed during read; retry paste"
        let token =
            match ownedXclip, ownedToken with
            | Some writer, Some token when not writer.HasExited && after = ownedX11Window && pid = Some writer.Id ->
                Some token
            | _ -> None
        if token.IsNone then ownedToken <- None
        Ok { Text = output.GetAwaiter().GetResult(); Token = token }

let private writeWindows text token =
    let bytes = Encoding.Unicode.GetBytes(text + "\u0000")
    Win32.ensureOwnerWindow ()
    Win32.withClipboard (fun () ->
        if not (Win32.EmptyClipboard()) then failwith "Cannot clear Windows clipboard"
        let memory = Win32.GlobalAlloc(0x42u, unativeint bytes.Length)
        if memory = 0n then failwith "Cannot allocate Windows clipboard text"
        let mutable transferred = false
        try
            let data = Win32.GlobalLock memory
            if data = 0n then failwith "Cannot lock Windows clipboard text"
            try Marshal.Copy(bytes, 0, data, bytes.Length)
            finally Win32.GlobalUnlock memory |> ignore
            if Win32.SetClipboardData(13u, memory) = 0n then failwith "Cannot set Windows clipboard text"
            transferred <- true
        finally
            if not transferred then Win32.GlobalFree memory |> ignore)
    // CloseClipboard finalizes the sequence. Re-lock and prove ownership before
    // recording it, rather than accidentally blessing an intervening owner's write.
    let sequence =
        Win32.withClipboard (fun () ->
            if not (Win32.ownsClipboard ()) then failwith "Clipboard changed during write"
            Win32.GetClipboardSequenceNumber())
    ownedWinSequence <- sequence
    ownedToken <- Some token
    token

let private readWindows () =
    let text, sequence =
        Win32.withClipboard (fun () ->
            let text =
                if not (Win32.IsClipboardFormatAvailable 13u) then ""
                else
                    let handle = Win32.GetClipboardData 13u
                    if handle = 0n then failwith "Cannot read Windows clipboard text"
                    let data = Win32.GlobalLock handle
                    if data = 0n then failwith "Cannot lock Windows clipboard text"
                    try Marshal.PtrToStringUni data
                    finally Win32.GlobalUnlock handle |> ignore
            text, Win32.GetClipboardSequenceNumber())
    let token = if sequence <> 0u && sequence = ownedWinSequence then ownedToken else None
    if token.IsNone then ownedToken <- None
    { Text = (if isNull text then "" else text); Token = token }

/// Replace the system clipboard; failure leaves the caller's selection/document unchanged.
let writeText (text: string) : Result<string, string> =
    lock gate (fun () ->
        ownedToken <- None
        try
            let token = Guid.NewGuid().ToString("N")
            if OperatingSystem.IsWindows() then Ok(writeWindows text token)
            elif OperatingSystem.IsLinux() && not (String.IsNullOrEmpty(Environment.GetEnvironmentVariable "DISPLAY")) then
                Ok(writeX11 text token)
            else Error "OS clipboard requires Windows or Linux X11 (DISPLAY and xclip)"
        with ex -> Error ex.Message)

/// An empty selection is success with empty text; transport failures never fall back to cached text.
let readText () : Result<Contents, string> =
    lock gate (fun () ->
        try
            if OperatingSystem.IsWindows() then Ok(readWindows ())
            elif OperatingSystem.IsLinux() && not (String.IsNullOrEmpty(Environment.GetEnvironmentVariable "DISPLAY")) then readX11 ()
            else Error "OS clipboard requires Windows or Linux X11 (DISPLAY and xclip)"
        with ex ->
            ownedToken <- None
            Error ex.Message)
