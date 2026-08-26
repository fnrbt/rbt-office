# Damage-bounded paint for the spreadsheet

How a cursor move (arrow key) should rasterize ~150×25 px instead of the
whole viewport, and how the rest of the spreadsheet's interactions plug
into the same model.

This doc lives here because the spreadsheet is the first / hot consumer.
The cross-cutting infrastructure (renderer scissor, signal-level damage
accumulation, platform partial-commit) is to be added in `../widgets` /
`../renderer` — see notes in those repos for the platform side.

## The current bottleneck

Every signal change triggers full widget-tree rebuild → full `Cmd list`
→ full rasterization → full SHM upload → full surface commit. There's
no path to a sub-viewport bound today because:

- Raster backend `Init`/`InitInto` clears to black; `Submit` rasterizes
  every cmd. No scissor. (`renderer/Raster.fs:2740`)
- X11 window always `XShmPutImage`s the whole buffer
  (`widgets/platform/src/X11Window.fs:443`); Wayland always damages
  `(0,0,w,h)` (`widgets/platform/src/WaylandWindow.fs:236`).
- The previous-frame pixel buffer is overwritten each frame, so even
  with scissor there's no "preserved background" to scissor against.

My initial fix in `spreadsheet-native/SpreadsheetGrid.fs:560-644` caches
the `Cmd list` across cursor moves (the per-field ref-equality cache).
That cuts the F# paint-item construction but **not** rasterization —
the backend still rasterizes the cached cmd list every frame. The actual
bottleneck is glyph rasterization for hundreds of cells.

Damage bounding takes rasterization itself out of the per-frame budget.

## What other spreadsheets actually do

| System | Selection-move bound | Mechanism |
|---|---|---|
| Excel (Win32 / Direct2D) | **O(perimeter)** | `InvalidateRect(old) + InvalidateRect(new)` + clipped paint |
| LibreOffice Calc | **O(perimeter)** | VCL `VirtualDevice` backing-store + invalid-region union |
| Google Sheets / XL Online | **O(touched cells)** | Same canvas; redraw rows/cols touching new sel |
| Glide Data Grid | **O(visible cells)** | Full viewport redraw; canvas fill rate makes it work |
| AG-Grid | **O(flipped cells)** | DOM class toggle (DOM-bound at ~100k cells) |

Universal across native systems: scroll = blit-old-buffer-translated +
paint newly-exposed strip (`ScrollWindowEx` / `ctx.drawImage(self, dx,
dy)`). Nobody redraws the full viewport on scroll.

Sources: Kenny Kerr's Direct2D MSDN article (2013); LibreOffice
`vcl::Window::Invalidate`; Miklós Vajna's tiled editing posts;
`glideapps/glide-data-grid` `render/` directory.

## Design: damage rects sourced from signal writes

Our widget tree rebuilds every frame so it has no stable identity to
hang dirty markers on. Signals do. Damage originates at the signal-write
site, accumulates globally, drains at frame start.

```fsharp
// In widgets/Reactive.fs (to be added)
module Damage =
    let mutable private pending : Rect list = []
    let mark (r: Rect) = pending <- r :: pending
    let drain () : Rect list =
        let xs = pending in pending <- []; xs
```

The signal API stays unchanged. Handlers that know they caused a
visible change call `Damage.mark`. If a frame fires with no damage
marked, the platform falls back to full-surface commit — opt-in
performance, opt-out correctness.

## Spreadsheet-side integration

The damage rects in `model` space need translating to screen pixels.
The handler has the viewport in hand and is the right caller.

### `SpreadsheetEditor.fs`

- **`setFocus state row col`**, **`extendFocus state row col`** — compute
  `selectionScreenRect` for the old selection (before the mutation) and
  the new selection (after), mark `Rect.union` of the two. This is the
  hot path — arrow keys, click-to-select, drag-select.
- **`commitEdit state`** — mark the rect of the edited cell (or the
  affected merged range).
- **`applyChange state transform`** — generic path; without diffing
  cells we don't know the touched rect. Either mark full canvas
  (correct, slow), or pass damage info down with each `applyChange`
  caller (sort/filter would mark entire visible viewport; cell write
  would mark the cell rect). Start with full-canvas fallback; refine
  per caller as profiles demand.
- **Drag-resize column / row** — each `PointerMove` marks the column or
  row strip from the resize line to the right/bottom edge of the
  canvas.
- **`switchSheet state idx`** — full canvas. Correct: everything
  changed.
- **`startEditing` / `cancelEdit`** — the in-cell editor is a separate
  overlay widget; its appear/disappear marks the cell rect.

### `SpreadsheetGrid.fs`

Once end-to-end damage works, **delete** the per-field ref-equality
`Cmd list` cache (`spreadsheet-native/SpreadsheetGrid.fs:560-644`).
Damage IS the cache: the rasterizer only touches pixels inside damaged
rects, and those rects only contain things that actually changed.

### Scroll

Scroll is handled at the renderer / platform layer via the
`WindowBackend.Scroll(dx, dy)` API (see `../rbt-renderer/GL-DIRECT-
PRESENTATION.md` for the GL path; software path mirrors it with
`memmove`). The `ScrollView` widget wires its `ScrollX`/`ScrollY`
signal subscribers to call `Scroll`. The spreadsheet doesn't need to do
anything special for scroll — no `Damage.mark` calls, no special-cased
paint. The backend BitBlt's the existing pixels and rasterizes only the
newly-exposed strip; the cmd-list paint for that strip is the regular
canvas paint scissored to the strip rect.

Selection drawing returns to `buildCellPaintItems` — no need for the
separate `selectionPaintItems` overlay either. The rasterizer only
touches pixels inside the selection's damaged rects; drawing
selection-in-cell is the simplest model and the cost is the same.

### Computing `selectionScreenRect`

```fsharp
let selectionScreenRect (model: SpreadsheetModel) (viewport: Viewport)
                        (sel: Selection) : Rect =
    let r1, c1, r2, c2 =
        match sel with
        | SingleCell (r, c) -> r, c, r, c
        | Range (sr, sc, er, ec) ->
            min sr er, min sc ec, max sr er, max sc ec
    let x = columnX model viewport c1 - viewport.ScrollX
    let y = rowY model viewport r1 - viewport.ScrollY
    let w = [ c1 .. c2 ] |> List.sumBy (fun c -> getColumnWidth c model)
    let h = [ r1 .. r2 ] |> List.sumBy (fun r -> getRowHeight r model)
    // Pad by selection-stroke width so the green outline isn't clipped.
    { X = x - 2.0; Y = y - 2.0; Width = w + 4.0; Height = h + 4.0 }
```

For the native interactive path the viewport's ScrollX/Y are always 0
(scroll lives in the outer `ScrollView`), so the subtraction is a no-op
there — but writing it scroll-aware keeps the helper reusable for the
web path.

## Cost model

| Action | Today | After |
|---|---|---|
| Arrow key | rasterize ~5000 cmds, ~3M pixels | ~300 cmds, ~7500 pixels |
| Cell edit (commit) | full rasterize | one cell's rect (~2500 px) |
| Column resize (per drag frame) | full rasterize | column strip (~10k px) |
| Scroll | full rasterize | one `glBlitFramebuffer` + one strip rasterize |
| Sheet switch | full rasterize | full (correct — everything changed) |
| Sort / filter | full rasterize | full (visible area changed) |

## Phasing on the spreadsheet side

1. **`Damage.mark` on `setFocus` / `extendFocus`** — once the
   `widgets/renderer` plumbing lands, this is the smallest E2E slice.
   Validates the pipeline and is by far the most-common interaction.
2. **Cell-commit damage.** Mark the committed cell's rect from
   `commitEdit`.
3. **Drag-resize damage.** Per-move column/row strip mark.
4. **Sheet switch + full-canvas fallback.** Mark full canvas; verify
   `Damage.mark` of full-screen rect is harmless (equivalent to today).
5. **Delete the `Cmd list` cache in `SpreadsheetGrid.fs`** — once damage
   is end-to-end, the cache is redundant noise.
6. **Selection drawn in cells again.** Drop the `selectionPaintItems`
   overlay. Cleaner code, no perf cost given damage.

## Open spreadsheet-side questions

- **Frozen panes**: a selection extending into / out of frozen rows or
  columns has multiple visible rects (the body slice plus the frozen
  slice). `Damage.mark` accepts a `Rect list` (the union), so this is
  fine — but the helper that computes selection-screen-rects needs to
  return all visible pieces, not one bounding rect.
- **Hover highlights** (row/column header hovering): worth damage-marking
  too? Probably yes for consistency, but the hover signal lives outside
  the model — easy add via `Hoverable`'s callback.
- **Tooltips / popovers**: those are `Overlay`s above the canvas,
  separate paint path. They already invalidate cleanly on their own
  signals; damage rects from the canvas don't apply.
- **Drag-select crossing the freeze split**: the body's frozen-cols /
  frozen-rows / body / frozen-corner zones need to be considered
  individually. Solvable by walking zones in `selectionScreenRect`.
