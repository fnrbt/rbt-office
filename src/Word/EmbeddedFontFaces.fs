namespace Rbt.Word

#if FABLE_COMPILER

/// The fonts a document embeds (Office.Formats' EmbeddedFonts: word/fonts/*.odttf, de-obfuscated) as
/// the page's fonts: canvas text, measured and painted, is set in them. As Word, which sets text in an
/// embedded face only when the machine has no font of its family (docs/layout-parity-flow.md, fixtures
/// embedded-fonts), a family the page has a font of - installed, or a face the host added to
/// document.fonts - keeps it. A host that measures with its own measurers (ShaperText) adds the faces
/// to its font set itself: API.getEmbeddedFonts, ShaperText.addEmbeddedFaces.
module EmbeddedFontFaces =
    open Fable.Core
    open Fable.Core.JsInterop

    /// The faces a document embeds: [] for one made from scratch or still being read (its package's
    /// parts come with its end).
    let ofDocument (doc: Document.Document) : EmbeddedFonts.EmbeddedFace list =
        match doc.Package.Preserved with
        | Some preserved -> EmbeddedFonts.ofPreserved preserved
        | None -> []

    /// [{ family, bold, italic, subsetted, data }] (data: the font file, a Uint8Array).
    let toObjects (faces: EmbeddedFonts.EmbeddedFace list) : obj array =
        faces
        |> List.map (fun face ->
            createObj [ "family" ==> face.Family; "bold" ==> face.Bold; "italic" ==> face.Italic
                        "subsetted" ==> face.Subsetted; "data" ==> face.Data ])
        |> List.toArray

    // The faces registered here, by "family|bold|italic": a later document's face of a family takes
    // the place of an earlier one's (both are embedded: neither is the page's own font).
    let private registry : obj = JS.Constructors.Map.Create<string, obj>() |> box

    [<Emit("""(async (faces, registry) => {
    if (typeof document === 'undefined' || !document.fonts || typeof FontFace === 'undefined') return [];
    const norm = (s) => String(s).replace(/^["']|["']$/g, '').trim().toLowerCase();
    const ours = new Set(registry.values());
    const ownFamilies = new Set([...registry.values()].map((f) => norm(f.family)));
    let context = null;
    const probe = 'mmmmmmmmmmlli1WQ@#';
    // (Whether the page has a font of the family: a face in document.fonts not registered here, or an
    // installed font - text in it measures otherwise than in each generic family alone.)
    const pageHas = (family) => {
        const name = norm(family);
        if (ownFamilies.has(name)) return false;
        for (const face of document.fonts) if (!ours.has(face) && norm(face.family) === name) return true;
        context ??= document.createElement('canvas').getContext('2d');
        if (!context) return false;
        const quoted = '"' + family.replace(/["\\]/g, '') + '"';
        for (const generic of ['monospace', 'serif', 'sans-serif']) {
            context.font = '72px ' + generic;
            const alone = context.measureText(probe).width;
            context.font = '72px ' + quoted + ', ' + generic;
            if (context.measureText(probe).width !== alone) return true;
        }
        return false;
    };
    const families = new Map();
    for (const face of faces) {
        const name = norm(face.family);
        if (!families.has(name)) families.set(name, pageHas(face.family));
    }
    const added = new Set();
    await Promise.all(faces.map(async (face) => {
        if (families.get(norm(face.family))) return;
        const key = norm(face.family) + '|' + face.bold + '|' + face.italic;
        const known = registry.get(key);
        if (known && known.rbtData === face.data) return;
        try {
            const font = new FontFace(face.family, face.data, { weight: face.bold ? 'bold' : 'normal', style: face.italic ? 'italic' : 'normal' });
            font.rbtData = face.data;
            await font.load();
            if (known) document.fonts.delete(known);
            document.fonts.add(font);
            registry.set(key, font);
            added.add(face.family);
        } catch { }
    }));
    return [...added];
})($0, $1)""")>]
    let private registerFaces (_faces: obj array) (_registry: obj) : JS.Promise<string array> = jsNative

    [<Emit("new WeakMap()")>]
    let private newLoads () : obj = jsNative
    [<Emit("$0.get($1)")>]
    let private loadOf (_loads: obj) (_key: obj) : JS.Promise<string array> = jsNative
    [<Emit("$0.set($1, $2)")>]
    let private setLoad (_loads: obj) (_key: obj) (_load: JS.Promise<string array>) : unit = jsNative
    let private loads = newLoads ()

    [<Emit("$0.then($1)")>]
    let andThen (_promise: JS.Promise<'T>) (_next: 'T -> 'U) : JS.Promise<'U> = jsNative

    /// Register the faces `doc` embeds of the families the page has no font of (once per package):
    /// a promise of the families that were registered (loaded: canvas text is set in them now), which
    /// the layout lays out again.
    let load (doc: Document.Document) : JS.Promise<string array> =
        match doc.Package.Preserved with
        | None -> JS.Constructors.Promise.resolve [||]
        | Some preserved ->
            let known = loadOf loads (box preserved)
            if not (isNull (box known)) then known
            else
                let faces = ofDocument doc
                let pending =
                    if faces.IsEmpty then JS.Constructors.Promise.resolve [||]
                    else registerFaces (toObjects faces) registry
                setLoad loads (box preserved) pending
                pending

#endif
