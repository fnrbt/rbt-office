# Word editor: the host API

What a host calls on the Word editor of `@rbt/office`: every function the
module `@rbt/office/word/Word.js` exports, with its signature and a line on
what it does. The table is made from the source (`src/Word/Word.fs`: the
export list, the signatures and the comments of the functions); the comment
of a function there is the full description - what is refused, what is one
undo step, what a returned object holds.

```js
import Word from '@rbt/office/word/Word.js';
import officeFormats from '@rbt/office/office-formats/src/Main.js';   // reads and writes DOCX

const OF = typeof officeFormats === 'function' ? officeFormats() : officeFormats;
const editor = Word.initializeCanvas(canvas);
Word.setViewport(editor, { scale: 1, devicePixelRatio, scrollTop: 0, scrollLeft: 0, width, height });
Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));   // or OF.word.openProgressive(bytes)
// ... edit ...
const blob = await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor));
```

### Document save contract

Save imported documents with the complete package shown above, including its
`Preserved` state, headers, footers, notes and styles. Do not reconstruct a
package from `Paragraphs`: those are only the document's body.

Capture the revision **after** exporting (which completes any pending load),
then mark that revision saved only after the write and host storage succeed:

```js
const packageToSave = Word.exportDocumentPackage(editor);
const revisionToSave = Word.getRevision(editor);
const blobToSave = await OF.word.writeFromEditorPackage(packageToSave);
await storeDocument(blobToSave); // the host's storage operation
Word.markSaved(editor, revisionToSave);
```

Typing during the asynchronous save belongs to a new undo step and must not
be marked saved. A failed write or upload must not call `markSaved`.

The shell's Ctrl+S shortcut saves only when the key event belongs to the shell
or its editor input; a shortcut from another control on the page does not
start a shell save.

### Shared typed .NET boundary

The JavaScript API above is an adapter over the same typed package conversion
and source-preserving writer used by .NET. Native callers use:

```fsharp
let savedBytes =
    Rbt.Word.Docx.loadBytesNative bytes
    |> Option.map (Rbt.Word.Docx.exportTypedDocumentPackage >> DocumentWriter.writeTypedPackageBytes)
```

`loadBytesNative` returns `Document option`. Export the complete captured
document, including its package preservation state. `EditorPackage.Unchanged`
means equal to imported content, not equal to the last saved revision.
The native writer serializes its existing mutable package state internally.
Filesystem storage and UI belong to the host; mark the captured document saved
only after successful storage, and reject stale completion after replacement.
Wisp provides local Open/Save/Save As with these rules. Browser and Wisp share
the lifecycle session below, not full UI parity. The former Widgets host is retired.

### Shared editing-session lifecycle

`Rbt.Word.Session.EditingSession` owns document/history state, typing coalescing,
revision identity, pending character formatting and story transitions on both
runtimes. Native document signals are UI projections: make changes through the
host/session operations, not direct signal assignments. Native pending style
is read through `Session.PendingFormat`, not a separate host signal.

`CaptureSave()` ends the typing run and returns `Document`, `Revision`, `Epoch`
and `Unchanged`. Serialize that document; `Unchanged` refers to opened content,
not the last successful save. Call `MarkSaved(snapshot)` only after successful
storage. Later edits remain dirty, and captures from before `Replace` are
rejected. Replacement clears history and establishes fresh opened/saved identity.
The browser's numeric `getRevision` / `markSaved` contract remains supported;
old-document revisions cannot mark a replacement saved.

`Execute(Character command, tracking)` owns ordinary character formatting of
a selected range, the word surrounding an interior caret, or pending typing at
a word boundary. `CharacterCommand.NamedStyle` and `FontDialog` explicitly
choose future-typing formatting at a caret. Pending choices create no content
edit or history step; range edits commit once through the session. Insertion uses `WithPendingFormat`;
native adapters retain its typing continuation through `CommitPendingInsertion`. Moving the caret, changing
stories, replacement or a real undo/redo step drops the choice. Empty undo/redo
does not. Browser notifications and public JavaScript signatures are unchanged.

`EnterStory` / `ExitStory` operate on a shared `StorySession`. The active
`Document` is the editable story; `WholeDocument` and captured saves include the
body and all stories. Exit restores the body's selection; undo/redo restores
the story and selection associated with the edit. `headerFooterStory` prepares
an existing or lazy new part without modifying content merely by entering it.
Wisp exposes `enterStory` / `exitStory` adapters; native story UI
parity is not implied. Rendering and page-derived entry geometry remain host
responsibilities.

`Execute(command, tracking)` applies a shared `Session.WordCommand` and returns
`Result<CommandOutcome, string>`: `NoChange`, `SelectionChanged` or
`ContentChanged`. Character formatting, paragraph/list/table operations,
Enter/Tab and breaks use this boundary in browser and native hosts. Optional tracking supplies author/date
at the host boundary; tracked insertion and story restrictions remain shared.
Paragraph/table property transformations are atomic transactions, retaining the
existing document primitives' validation and preservation behavior.

No-op commands preserve the original document snapshot. Selection-only commands
change selection without changing content revision, saved identity or redo.
Content changes record one edit in the existing session history. Hosts publish
the resulting document and notify/render as appropriate; they must not record
the same command again. Existing browser JavaScript signatures are unchanged;
native command availability does not imply equivalent menus or dialogs.

`CaptureClipboard(htmlOf)` returns a rich `ClipboardCapture option` without editing.
Transfer its `Text`/`Html` first; `CommitCut(capture, tracking)` deletes only while
the captured document, selection, story and replacement epoch remain current.
Moving away and restoring the same selection does not revive a pending cut.
`PasteClipboard(text, html, decodeHtml, tracking)` owns rich/plain insertion,
bookmark identity handling, pending styles, tracked paste and story restrictions,
recording one transaction. HTML parsing and clipboard permissions remain host concerns.

`Navigate(Session.Navigation...)` owns character, word, paragraph and document-edge
movement, range extension and select-all. `VisualTarget` accepts a layout-derived
position from the host. Navigation ends typing and pending formatting without
creating a content revision or undo entry.

The browser package adapter exports header/footer parts with `Kind` (default,
first or even), matching the typed package records and existing host contract.


### Mounted Word shell

The Word shell installs its shared office and Word-specific styles once per
document; hosts do not need a separate stylesheet. Its `.rb-office-shell` /
`.rb-word-shell` root keeps the chrome scoped to the mounted editor. The
formatting controls wrap within the host, while the document workspace scrolls
horizontally on narrow hosts instead of scaling the canvas or changing its
hit-test coordinates. Allow the host element to shrink (`min-width: 0` in a
flex or grid layout).

The status bar reports body words and, for a non-collapsed selection, selected
words independently of action and error messages. `getWordCounts(editor)`
returns `{ BodyWords, SelectedWords, HasSelection }`; it does not edit the
document or force a progressive load to finish. During an incomplete or failed
load the shell labels the total “Body (loaded)”; it reflects the blocks read so
far and updates as more arrive.

The body total includes paragraphs in nested tables and body text boxes once,
not repeated rendered headers or footers, notes or comments. Selection counts
refer to the active editable story (including a header, footer, note, comment
or text box); a body selection does not include boxes it crosses. Each selected
paragraph fragment is counted separately. Like `computeRangeStatistics`, a
word is a nonempty run of UTF-16 text separated by Unicode whitespace; a
partial word counts as a word. This counts the model's text, including hidden
and tracked insertion/deletion text and object/reference placeholder
characters; it is not Microsoft Word's linguistic or revision-view count.

The document action buttons expose expanded state and open groups of ordinary
buttons. Tab navigates the controls; Escape closes an open group and restores
focus to its button, and moving focus or clicking outside closes it. A dialog
opened from a menu returns focus to that menu's visible trigger on dismissal. Toolbar
controls have descriptive names and formatting toggles expose their current
state, including mixed selections. Shell shortcuts remain confined to the
shell/editor and do not intercept text entry in dialog fields.

Right-clicking the shell's canvas opens a text/table context menu. A click
inside the selected range retains it; a click elsewhere targets the hit-tested
caret. Shift+F10 or the Context Menu key opens at the caret. Arrow keys,
Home/End and Enter navigate/invoke actions; Escape or Tab dismisses and returns
focus to the editor. Outside focus/pointer input, scrolling outside the menu,
resizing, document replacement and disposal dismiss it. Menus fit a narrow
viewport and scroll when necessary. External inputs and bare canvas widgets
retain their browser context menus; composition does not open the shell menu.

Context actions share the shell command registry: undo/redo, cut/copy/paste,
bold/italic/underline, font/paragraph dialogs, insert/edit/remove hyperlink,
and current-table row/column insertion/deletion and cell formatting/margins.
Read-only mode allows copying but disables editing. Dialogs opened here return
focus to the editor, not the removed context-menu button.

Clipboard commands use the browser Clipboard API and report permission errors;
unavailable capabilities disable their actions. Rich HTML and plain text are
written when `ClipboardItem` is available, otherwise copying is plain text.
Cut deletes only after a successful write. An asynchronous cut or paste is
cancelled if the document/selection changes, the editor becomes read-only,
or the shell is disposed. Native keyboard clipboard handling remains in the
editor; it is not replaced with shell shortcut interception.

The toolbar's Styles button opens a searchable catalogue of paragraph and
character styles. Hidden styles can be shown explicitly; a hidden current
style remains visible. The pane reports mixed selection state, uses resolved
style information where available, and applies styles to the retained editor
selection. It selects existing styles only; it does not author or remove
definitions. Escape or Close dismisses the pane and returns focus to Styles.
Style choices are disabled in read-only mode, and opening another document
refreshes the catalogue.

The shell's Insert > Hyperlink, Link toolbar button and Ctrl+K use the same
dialog. Ctrl+K is handled only from the editable editor input, not other
inputs or dialogs. The dialog edits or removes the whole link at the caret;
otherwise it links the selection, or inserts the entered text (the target if
text is empty) at a caret. Targets must be http, https, mailto, tel or a
`#bookmark`. Existing inactive targets can be removed or changed.

Insert > Table and the toolbar open the custom rows/columns dialog. The Table
menu also offers a default table and operations on the current table; those
operations are disabled outside a table. Table properties edits the caret's
innermost table and current row: preferred width (auto, twips, percent or
pixels), alignment, left/start indent (pixels), default cell margins and
spacing (pixels), table background and shading pattern, a replacement grid or no table borders (border widths in pixels),
row minimum height (twips), splitting and repeating header. Only consecutive
header rows at the start repeat. Row or cell overrides can mask table borders,
background and margins. Current cell style separately edits the caret's cell,
not a multi-cell selection; its width override is in twips.

The dialog keeps untouched imported properties, including layout mode,
per-side borders, shading patterns and exact row-height rules. The layout does
not implement autofit versus fixed layout, or exact row height; an imported
exact height is currently laid out as a minimum. Editing the height replaces
its rule with minimum height (or clears it when blank). Indent applies only to
left/start-aligned tables. Blank indent, margin, spacing and row-height fields
reset the corresponding unset/inherited value. Explicitly allowing row splitting
is saved as `w:cantSplit w:val="0"`, so it can override inherited restrictions.
Editor-package `PreferredWidth` values with `PreferredWidthType = "pct"` are
actual percentages for both tables and cells (1 means 1%, not 1/50%); the
DOCX reader/writer alone converts to/from fiftieths of a percent.
Cancel and unchanged Apply do not edit the document;
accepted table and row changes form one undo step. A changed document or
selection, a replaced editor, or a read-only input rejects the form.

The paragraph-only `exportToDocx` / `word.writeFromEditor` APIs and the
paragraph-only host/shell export wrappers have been removed. Use
`exportDocumentPackage` (or the mounted host's `exportMountedDocumentPackage`)
and `writeFromEditorPackage`; there is no compatibility alias.
For a **new** document assembled from paragraph objects, the writer accepts
`{ Paragraphs: paragraphs }`. That is creation, not a way to save an imported
document. The document builder also uses the package writer.


The first argument of nearly every function is the editor `initialize` /
`initializeCanvas` returned. A read-only viewer is a separate, smaller
module (`@rbt/office/word/WordViewer.js`, also `Word.createViewer`): see
[The viewer](#the-viewer).

## Units

- **Layout pixels**: 1/96 inch (twips / 15). Page sizes, margins, indents,
  what `getLayoutLines`, `getPageFrames`, `getShapeRects`, `getCaretRect`
  and the hit tests say and take.
- **CSS pixels**: layout pixels times the viewport's `scale`. `setViewport`
  (`scrollTop`, `scrollLeft`, `width`, `height`, `originX`, `originY`), the
  `Scaled*` fields of `getDocumentSize`, `hitTestViewport`,
  `viewportToDocument` / `documentToViewport`, the `ScrollTop` of a notice.
- **Points**: `setFontSizePt`, `setSpacingBefore`, `setSpacingAfter`
  (`setFontSize` takes layout pixels; `setLineSpacing` a multiple).
- A **position** is a block id and an offset into the text of that
  paragraph (`setSelection`, `getSelection`, the range queries). Block ids
  come from `getLayoutLines`, `getSelection`, `findText`, `getShapeRects`.

## What a host is told: `onChange`

`Word.onChange(editor, callback)` calls `callback(notice)` after every change
and returns the function that ends the subscription. A notice has:

| Field | |
| --- | --- |
| `DocumentChanged`, `SelectionChanged` | what changed (a caret move: the selection alone) |
| `FormatChanged` | a character format was chosen with nothing selected: `getSelectionFormat` says something else, the document does not |
| `CanUndo`, `CanRedo`, `Revision`, `Modified` | see `getRevision`, `markSaved`, `isModified` |
| `LayoutChanged`, `LayoutComplete` | more of a long document is laid out (more pages, a larger document), or all of it |
| `ScrollTop` | a number only in a notice that asks the host to scroll: see below |
| `ShapeSelectionChanged` | true only in a notice that another shape, or none, is selected (`getSelectedShape`) |
| `LoadChanged`, `LoadComplete`, `LoadFailed`, `LoadError` | a document that is still being read (`openProgressive`) |
| `FontsPending` | the document needs faces the host has not got (`setFontAvailability`) |

## Long documents: the layout comes in slices

Through a viewport a long document is not laid out as a whole before it is
shown. `isLayoutComplete` is false while the rest is laid out in slices
between the host's events; `getDocumentSize` and `getPageFrames` give the
pages there are so far, the listeners are told as they grow
(`LayoutChanged`), and `whenLayoutComplete` resolves at the end. What needs
the whole layout (`getPageCount`, the layout queries, find, `renderPage` of
a later page) lays the rest out at once.

The same happens after a change of most of the document at once - everything
selected and made bold, a font for all of it, replace all, their undo and
redo. With the view near the top the document is laid out again from its
start. With the view far down it (0.1.14) the blocks above the view are
kept **as they were laid out before the change**
(`getRenderStats(editor).StaleBlocks` of them) and the document is laid out
from the block at the top of the view; the layout from the start follows in
slices. Until it is complete the pages above the view, the page count and
the document's height are those from before the change.

When the heights above the view are known, the text in view stays where it
is on screen at another scroll position: the editor moves its own viewport
and tells the host with a notice `{ LayoutChanged: true, ScrollTop }`. **A
host scrolls its content to `ScrollTop`** (CSS pixels, as in `setViewport`;
`getViewport(editor).scrollTop` is the same number):

```js
Word.onChange(editor, (notice) => {
    if (typeof notice.ScrollTop === 'number') scroller.scrollTop = notice.ScrollTop;   // then setViewport as on any scroll
    if (notice.LayoutChanged) spacer.style.height = Word.getDocumentSize(editor).ScaledHeight + 'px';
});
```

A host that does not do this shows other text - the view jumps by what the
blocks above it grew or shrank - the next time it calls `setViewport` with
its own scroll position.

Text that is measured otherwise meanwhile leaves this as it is (0.1.17): a
face that arrives for the text that was just made bold
(`setTextMeasurersForFaces`, with or without the faces named), other
measurers (`setTextMeasurers`), a zoom or another width (`setViewport`). The
blocks above the view stay as they were laid out, the document never ends
above the view, what is in view is laid out again at once, and the layout
from the start begins again where it has to; the `ScrollTop` notice comes
when it has reached the view. (In 0.1.16 a face that arrived before the
layout from the start had reached the view made the document as short as
that layout was - three pages of six hundred - and no `ScrollTop` notice
followed.)

## Sections

A section's properties stand at its end in the file; the editor shows and
saves them so. For a host:

- `setSectionBreak(editor, kind)` puts a section break after the caret's
  paragraph. `kind` says how the section that **starts** at the break
  starts: `"continuous"`, `"nextPage"`, `"evenPage"`, `"oddPage"`; `"none"`
  takes the break away. (Up to 0.1.13 the kind was given to the section that
  ends at the caret, so a continuous break left the text after it on a new
  page.) The text before the break keeps the page, columns, headers, footers
  and page number start of the section it was in.
- A section that starts on an even or an odd page has a blank page before it
  where the next page is of the other kind. The blank page is one of the
  pages `getPageFrames` gives, is numbered, and shows the header and footer
  the section's rules give a page of its number.
- `w:vAlign` of a section (center, bottom, both) is read, laid out and
  saved. `setSectionVerticalAlignment(editor, align)` sets it for the section
  the caret is in (`"top"`, `"center"`, `"bottom"`, `"both"`; one undo
  step), `getSectionVerticalAlignment(editor)` says it.
- A page break inside a section with columns starts a new page; a column
  break (read from a file) goes on in the next column. `insertPageBreak`
  inserts a page break.

## Lists

A list item of a list that came from a file moved to another level -
`increaseIndent`, `decreaseIndent`, `setListLevel(editor, level)`, Tab and
Shift+Tab at the start of an item - has that level's number format, text
("%1.%2)"), start, indents and bullet at once, as after a reload. The number
of each level is written as its level writes it, a legal level (`w:isLgl`)
in decimal numbers; `w:lvlRestart`, `w:startOverride`, `w:suff` and
`w:lvlJc` (right-aligned numbers) are followed. A line of `getLayoutLines`
has the number or bullet as shown in `ListMarker` (on a paragraph's first
line). A number wider than its hanging indent is drawn over the start of
the text (the text is not moved to the next tab stop).

The browser shell's **Format > Bullets and numbering** dialog (command
`format.listsDialog`) changes the selected paragraphs, or the caret's
paragraph. It offers ordinary paragraphs, five bullet markers, decimal,
letter and Roman numbering, and levels 1–9. Unchanged or mixed fields are
left alone; changing a marker or kind splits the affected list instances so
unselected items and independently identified imported lists keep their
formatting. To retain the displayed counters and nested parent numbers, the
affected suffix is materialized as standard numbering instances with start
overrides; those instances remain distinct after reopening. A level change
uses the existing list's level definition. This is not an editor for
arbitrary multilevel templates.

For a numbered list, **Restart at** starts a new numbering instance at the
first selected item and includes the contiguous following items of the same
list, including nested items. **Continue preceding list** joins that run to
the nearest preceding numbered list at the same level in the same story,
even across ordinary paragraphs. It does not merge unrelated selected list
identities. Restart
requires a positive 32-bit integer. Apply without a change and continuing the
same instance are no-ops; each accepted edit is one undo step. The dialog
refuses a stale or read-only selection. Imported numbering definitions and
overrides not being changed are retained in the saved package. The native Wisp
host does not show this browser-only dialog.

## Text boxes

`insertTextBox(editor, text)` anchors a box in the caret's paragraph - in
the body, and in a header or footer that is being edited (saved in that
part); not while a text box is edited (a box does not hold another).
`getShapeRects` gives the boxes where they are painted - those of the body
(`Story: "body"`) and those of the headers and footers (`Story: "header"` /
`"footer"`, once each, on the first page that shows the header:
`PageIndex`) -, `enterTextBox(editor, id)` / `exitTextBox` edit the text of
any of them in place (a header's box on that page; what is typed is the
header's on every page), and
`setCurrentTextBoxStyle` sets its place, size, fill and border.

**Groups and drawing canvases.** A group (`wpg:wgp`, a VML `v:group`) or a
drawing canvas (`wpc:wpc`) is one shape of `getShapeRects`, with
`IsGroup: true` and its shapes in `Children` (Id, X, Y, Width, Height in the
same pixels, where the group's coordinates put them at the group's size
now, before its turn; Rotation, FlipH / FlipV, IsGroup, Text, Children of
their own). The group is selected, moved (`moveShape`) and resized
(`resizeShape`) as a shape is - its shapes scale with it, a canvas's stay
where they are - and saved as the element it was read from with its place
and size changed in it; whatever else of it changes is not saved, and its
shapes are not shapes of their own: `enterTextBox` refuses a group, and the
text of its shapes is drawn, not edited.

The pointer reaches them all (0.1.17). A click in a box of the body edits
its text. A double click on a page's header or footer edits it in place -
and where a text box of that header is under the pointer, the box's text,
with the caret where the click was (until 0.1.16 the caret went into the
header's own paragraph). While a header or footer is edited a click in one
of its boxes edits the box's text, on the page the click is on (it did
nothing); from the box a click elsewhere on a header or footer goes back to
that header, Escape or a double click on the body to the body. **A host
that handled the double click on a header's box itself - `getShapeRects`,
a hit test and `enterTextBox` - removes that**: the editor does it, and a
host that still calls `enterTextBox` from its own handler moves the caret
to the start of the box's text after the editor put it under the pointer.

**The selected shape (0.1.17).** The editor keeps a selected shape itself
and handles the pointer and the keys for it; a host need not draw anything.
`selectShape(editor, id)` selects a shape of the body, a header or a footer
(`null`: none), `getSelectedShape(editor)` says which, and a notice
`{ ShapeSelectionChanged: true }` tells of every change. A click into a
box's text selects that box. The editor paints the frame and the eight
handles of the selected shape on its own canvas, turned as the shape is.
With the pointer: a drag of a handle resizes the shape - the side or corner
across from the handle stays where it is, for a turned shape too -, a drag
of its frame moves it (of any point of a shape whose text a click does not
reach: behind the text, a watermark); the cursor says which. Each drag is
one undo step. With the keys, while no text is edited: Delete or Backspace
takes the shape out, the arrow keys move it by a pixel (Shift: ten), Escape
lets go of it. Escape in a box's text goes back to the body with the box
still selected. A click elsewhere lets go of it.

**A host that drew the frame and the handles itself and drove
`moveShape` / `resizeShape` from its own pointer handlers removes all of
that** - the overlay, the hit test against `getShapeHandles`, the drag, its
own Delete / arrow / Escape handling for a selected shape - or it moves the
shape twice. What stays the host's: showing its own controls for the shape
that `getSelectedShape` names (after a `ShapeSelectionChanged` notice), and
calling `selectShape` for a shape picked elsewhere (a list of shapes).
`resizeShapeBy(editor, id, handle, dx, dy)` is what a drag of a handle does
- one undo step also for a handle of the left or the top edge, which took a
`moveShape` and a `resizeShape` (two steps).

A host that still lets shapes be dragged itself: `getShapeHandles(editor, id)` gives the
eight handles of a shape of the body (`Name`: `"nw"`, `"n"`, `"ne"`, `"e"`,
`"se"`, `"s"`, `"sw"`, `"w"`; a square of eight pixels on screen around its
corner or the middle of its edge, in the pixels of `getShapeRects`; the CSS
`Cursor`), `hitTestShapeHandle(editor, id, x, y)` says what of the shape is
at a point - a handle's name, `"move"` inside it, `""` outside -, and
`moveShape(editor, id, dx, dy)`, `resizeShape(editor, id, width, height)`
and `deleteShape(editor, id)` do it, each one undo step (convert with
`documentToViewport` to draw over the canvas). A handle of the left or the
top edge is one `resizeShapeBy`. A shape that was aligned in the direction it is moved in
(centred, at the right margin) is by its offset from the page from then on.
They take the shapes of the body and (0.1.17) those of the headers and
footers (`Story: "header"` / `"footer"`: where the first page that shows the
header paints it, and the change is the header's on every page) - not while
the text of a box, a header or a footer is edited, and not for a shape that
stands in a line of text (`moveShape`).

A shape that was read and is only moved or resized is saved as the element
it was read from with the offsets (`wp:positionH` / `wp:positionV`) and the
extent (`wp:extent`, the shape's own `a:ext`) changed in it - its geometry
and whatever else the model does not hold stay the text they were. Its VML
fallback says the place and the size in its own way and is changed with
them (0.1.17: `margin-left` / `margin-top` / `width` / `height` of its
style, from the frame the anchor names; until 0.1.16 the fallback was left
out) - one that cannot be placed so (a line, a group) is left out. A group
or a drawing canvas, or a shape changed in more, is written from the model.

A save keeps the parts the file had and adds none for what was not changed
(0.1.17): a picture the file's relationships name and nothing in the part
uses - left behind by the tool that wrote the file - stays with its
relationship (only a picture that a change made unused goes), and a
document that was read without a styles part gets none, nor a relationships
part for its body, while no style is made or named.

What is done in the body does not reach into a box that is not entered: the
caret does not move into it, and typing, pasting or deleting over a
selection of the body never changes a box's text. A box goes with the start
of the paragraph it is anchored in: it is removed when that paragraph, or
its first character, is in the selection that is deleted, and stays when the
selection begins further on in the paragraph or ends before its first
character. `findText` and `replaceAllText` go by all the text of the story -
the text of its boxes too.

## Bookmarks

A run of text has every bookmark that starts before it and ends after it
(`getBookmarkInfos` lists them all; `removeBookmark` and `renameBookmark`
take one of several). Text typed with a format chosen at the caret is a run
of its own without the bookmarks of the text around it, and its paragraph is
still changed in the XML it was read from. A paragraph that is written from
the model writes each bookmark once, with the id and the name the file had,
and the document part is checked when it is put together: a range (a
bookmark, a comment's, a permission's, moved text's) that the model's part
would open twice, close twice, close without opening or leave open is put
right - see `balancedRanges` in the writer. The headers, the footers, the
notes and the comments are checked the same way when they are written from
the model.

In the editor package (`forEditorPackage`, `exportDocumentPackage`) a run's
bookmarks are typed: `BookmarkStarts: [{ Id, Name, EndsHere }]`, the bookmarks
that start before the run (`EndsHere`: it ends where it starts), and
`BookmarkEnds`, what ends after it - `{ Id }` a bookmark's end, `{ Id, Name,
EndsHere, Opens: true }` a bookmark opened after the run (after the last run
of a paragraph, as a file can have it). Packages of 0.1.28 and before had
strings instead (`BookmarkStartId`, `BookmarkName`, `BookmarkEndId`, their
tokens joined by U+001F); those fields are not read any more - a package is
made and used by the same release. A bookmark opened after the last run of a
paragraph is listed by `getBookmarkInfos` (its StartOffset and EndOffset at
that paragraph's end), and `removeBookmark` / `renameBookmark` take it.

What the model does not hold of ranges is written where the file had it when
its paragraph is written from the model: a bookmark in a paragraph without
text, the end of a bookmark before the first run of the next paragraph, a
bookmark's start or end between two blocks, and the ranges of comments and of
permissions (`w:commentRangeStart`, `w:permStart` and their ends). A block
that is deleted takes the bookmark that started in it: its end is not left
behind in the text after it.

## Comments

A comment that was read stands in the text as its mark, `[Comment 5]`: the
place of its reference. The mark is one thing. `setSelection` into it goes
to its nearer end (so does a click), an arrow key steps over it, a selection
has all of it or none of it, text is typed before or after it, and Backspace
after it (Delete before it) removes it whole. A host that kept a position
inside a mark gets the end of the mark back from `getSelection`.

## Footnotes and endnotes on the page

A footnote stands at the foot of the page its reference is on, an endnote
after the text (on further pages as it needs them), each laid out from its
body as the file has it - its paragraphs with their formatting, their
paragraph properties and style (FootnoteText / EndnoteText as the file
says), its tables - by the layout that lays out the body, with its mark (the
note's number, superscript) before its first paragraph's text. A page makes
room for its footnotes by their real height: a footnote too long for the
room below its reference goes on at the foot of the next page when the
layout has one (broken between the lines of a paragraph; a table goes on
whole), and the next page keeps room for it. A note whose reference's text
was set on its own (`setNote`) shows that text, and a note the package has
no body for shows the text its reference carries, in Calibri 10 pt.
`getLayoutLines` has a note's lines with block ids `__footnote__<page>_<n>#<note
id>#<paragraph id>` (`__endnote__<n>#...`), the paragraph being the note's own
(a cell's for a table), and the mark as the first line's `ListMarker`.

## Notes and comments edited in place

`enterNote(editor, noteId, noteType)` (`noteType` "footnote" | "endnote") and
`enterComment(editor, commentId)` put the caret into the body of that note or
comment, as `enterHeaderFooter` does into a header: from then on the document
the editor's commands and keys act on is that body - typing, Enter,
Backspace, selection, character and paragraph formatting, undo and redo work
there - until `exitNote`, Escape or a click off the note. A footnote or an
endnote is edited where the page shows it, as it is laid out there (see
above): the text typed is laid out in the note area at once, the page's text
making room for a note that grows, and a dashed frame around the note says
that it is the note that is edited; a comment is edited in a balloon below the
line its range starts on (the host's own comment pane, from `getCommentAnchors`, shows its
text as before, and the anchor's `Text` follows the edit). A double click on a
footnote or an endnote where the page shows it enters it with the caret where
the click was. `getCaretContext` says `Story` "footnote" | "endnote" |
"comment" and `NoteId` / `CommentId`; `getLayoutLines` and `getText` are those
of the body while it is edited, as in a header. Entering changes nothing, and
an undo step made there brings the caret back into the note.

The paragraphs edited are the note's paragraphs as they were read: the save
writes the part of notes (comments) as it was read with the edit made in the
note's paragraph - the other notes, the separators, the note's other
paragraphs and its mark run (`w:footnoteRef`, `w:annotationRef`) stay the
text they were, and the body's XML is not touched. (Before 0.1.27 a note's
body could only be replaced whole by `setNote` / `setComment`, as a new plain
paragraph.) Not yet: the note's mark itself is not part of what is edited; a
note or comment the text does not refer to cannot be entered.

## Fields

Text typed into the result of a field - also with a format chosen at the
caret - is of that result. A result of which parts are formatted
differently is as many runs in the model, each with the field's code
(`FieldCode`) and whole result (`FieldResult`) and its own part of the text;
it is read so, and saved in the paragraph's own XML with the result's run
split. (The number of a page is one run: the editor shows its own.) An entry
of a table of contents - one field over many paragraphs - is changed in its
own XML, and an entry that is written from the model leaves the other
entries the text they were.

## Keys

The editor handles these keys itself, on its canvas and its input element
(a host that also binds one of them does the thing twice):

| Key | |
| --- | --- |
| Enter | a new paragraph (in an empty list item: out of the list) |
| Shift+Enter | a line break in the paragraph (`insertLineBreak`) |
| Ctrl+Enter | a page break (`insertPageBreak`); nothing where a page cannot break - a text box, a header, a footer, a note |
| Ctrl+Shift+Enter | a column break (`insertColumnBreak`): the next column, the next page in a section of one column |
| Backspace, Delete | a character, or the selection; Backspace at the start of a list item takes it out of the list first, as in Word (0.1.26; it joined the item to the one before) |
| Ctrl+Backspace, Ctrl+Delete | the word before / after the caret, as far as Ctrl+Left / Ctrl+Right moves (Alt on a Mac) |
| Tab, Shift+Tab | in a table: the next / the previous cell (Tab in the last one adds a row); in a list: the item a level down / up; elsewhere Tab is a tab character and Shift+Tab nothing |
| Ctrl+Tab | a tab character in a table cell |
| Arrows, Home, End, PageUp, PageDown | the caret; with Shift the selection; Ctrl+Left / Right by words, Ctrl+Up / Down by paragraphs, Ctrl+Home / End the ends of the document |
| Ctrl+A, Ctrl+C, Ctrl+X, Ctrl+V | everything, copy, cut, paste - with formatting (see Pasting) |
| Ctrl+Z, Ctrl+Y, Ctrl+Shift+Z | undo, redo (keys typed one after the other are one step, as in Word - so are Backspaces, and Deletes; a caret move, another key or a format chosen starts the next) |
| Ctrl+B, Ctrl+I, Ctrl+U | bold, italic, underline |
| Ctrl+1, Ctrl+2, Ctrl+3 | heading 1 - 3 |
| Ctrl+E, Ctrl+L, Ctrl+R, Ctrl+J | centre, left, right, justify the paragraphs (0.1.26) |
| Ctrl+M, Ctrl+Shift+M | Increase / Decrease Indent: half an inch further in / out; a list item a level down / up (0.1.26) |
| Ctrl+Shift+>, Ctrl+Shift+< | Grow / Shrink Font: the next size of Word's size box (0.1.26) |
| Ctrl+], Ctrl+[ | a point larger / smaller (0.1.26) |
| Ctrl+=, Ctrl+Shift+= | subscript, superscript (on / off; 0.1.26) |
| Ctrl+Space | the character formatting cleared (0.1.26) |
| Escape | out of a header, a footer, a text box, a note or a comment that is edited |

The character keys work as Ctrl+B does: on the selection, or with nothing
selected for what is typed next at the caret. `getSelectionFormat` says
`VertAlign` ("superscript", "subscript", null). (On a Mac the Command key is Ctrl here.) Until 0.1.16 Ctrl+Enter and
Shift+Enter were Enter, Ctrl+Backspace and Ctrl+Delete took one character,
and Tab outside a list and a table did nothing. **A host that handles
Ctrl+Enter itself and calls `insertPageBreak` removes that handler** (two
breaks otherwise), unless it stops the event before it reaches the editor's
canvas and input element.

Virtual `beforeinput`/`input` paragraph insertion uses the same paragraph,
list-exit, table-cell and tracked-change edit as physical Enter, with one
undo step. Shift+Enter and `insertLineBreak` remain line breaks within a
paragraph. Paired physical and virtual events do not insert twice.

## Copying and pasting

Ctrl+C, Ctrl+X and Ctrl+V go through the browser's own copy, cut and paste
events on the editor's input element (0.1.26; before, the asynchronous
clipboard, which asks for a permission to read and never had HTML). Copy
writes the selection's text and its HTML (paragraphs, headings, lists,
alignment, bold, italic, underline, strike, sizes, faces, colours, links) and
keeps what it copied with explicit provenance in the returned HTML. Paste
uses that internal rich payload only when the supplied provenance and
content match; plain-text equality alone never identifies an internal copy.
Otherwise the supplied HTML determines formatting (from Word, a browser or
mail), or plain text is pasted when there is no HTML. Internal rich paste
keeps source formatting and the paragraph properties whose marks it copied;
the last paragraph joins the text after the caret and keeps that paragraph's
properties. With changes tracked it is the text. Each edit is one undo step.
What a copy leaves out: the marks of comments, bookmarks, fields (their result is text),
the objects of the file kept as they were read (charts). A host that handles
the events on the input element itself stops them first.

`pasteText(editor, text)` puts plain text in: its lines are paragraphs, the
whole is one undo step, and with nothing selected the text has the format
chosen at the caret (`toggleBold` and the like with nothing selected), every
line of it. A host that pastes from a menu calls `pasteClipboard(editor,
text, html)` with what it read (`navigator.clipboard.read()`), which pastes as
Ctrl+V does; `copySelection(editor)` / `cutSelection(editor)` give
`{ Text, Html }` to write (null with nothing selected) and keep the copy as
Ctrl+C / Ctrl+X do.

The browser's asynchronous cut fallback deletes only after a successful
clipboard write, and only while the captured document and selection remain
current. Rejected/unavailable writes leave the document untouched and report
a warning. The synchronous host `cutSelection` API instead transfers
ownership of `{ Text, Html }` to the caller and deletes immediately; hosts
requiring write confirmation should copy, await their transfer, then apply
their own selection-validated edit.

## The functions

### Creating an editor, loading and saving

| Function | What it does |
| --- | --- |
| `initialize(canvasId: string)` | Create an editor on the canvas element with this id. |
| `initializeCanvas(canvas: HTMLCanvasElement)` | Create an editor on a canvas element. |
| `initializeCanvasWithChange(canvas: HTMLCanvasElement, afterChange: () => void)` | As initializeCanvas, with a function called after every change of the document (before the onChange listeners). |
| `createViewer(canvas: HTMLCanvasElement): Viewer` | Create a read-only viewer on a canvas (see The viewer). |
| `getText(editor)` | The text of the document, a line per paragraph (the body's, its tables' and its text boxes'). |
| `dispose(editor)` | End the editor's input/layout lifecycle: detach its canvas, proxy and window listeners, remove the hidden input, and prevent pending input callbacks from changing the disposed editor. The canvas may be reused by another editor. |
| `getInputElement(editor): object` | The hidden text input that holds the keyboard focus while the editor is in use (IME, dead keys and virtual keyboards type into it); null once the editor is disposed. |
| `setAllowLinkedImages(editor, allow: boolean)` | Whether linked (not embedded) pictures may be loaded from the address the document gives. |
| `getAllowLinkedImages(editor): boolean` | Whether linked (external) pictures may be requested. |
| `setContent(editor, text: string)` | Set document content. |
| `setParagraphs(editor, texts: string[])` | Set document content as one paragraph per input string (unlike setContent, which puts the whole text in a single paragraph with hard breaks). |
| `testRope()` | Test rope operations. |
| `importFromDocx(editor, editorParagraphs: object[])` | Import document from EditorParagraphs (Office.Formats format) Use with: Office.Formats processors.forEditorPackage() Paragraphs output. |
| `importFromDocxWithHeaderFooters(editor, editorParagraphs: object[], headerParagraphs: object[], footerParagraphs: object[])` | Load the reader's paragraphs with one header and one footer (arrays of reader paragraphs). |
| `importFromDocxWithHeaderFooterParts(editor, editorParagraphs: object[], headerParts: object[], footerParts: object[])` | Load the reader's paragraphs with its header and footer parts (default / first / even). |
| `importFromDocxWithPageSetupAndHeaderFooterParts(editor, editorParagraphs: object[], pageSetup: object, headerParts: object[], footerParts: object[])` | As importFromDocxWithHeaderFooterParts, with the reader's page setup. |
| `importFromDocxPackage(editor, package: object)` | Show an editor package: Office.Formats' `forEditorPackage(bytes)`, or - for a document that is shown before it is read to its end - `word.openProgressive(bytes)`. |
| `setImageCache(editor, cache: object)` | Give the editor the decoded images of the document: a Map of image source (as in the editor package) -> HTMLImageElement. |
| `exportDocumentPackage(editor)` | Export the whole document for `OF.word.writeFromEditorPackage`. Completes pending reading and throws if loading failed. Ends the current typing run, making save an undo boundary. `Unchanged: true` means the document is back at its imported revision, so the writer can return the original bytes. |

### Browser shell files and PDF

`WordShell` File → Open and Ctrl/Cmd+O use a DOCX picker and the existing
Office.Formats package reader. The title shows the opened filename and a dirty
marker. Replacement asks before discarding unsaved edits; cancellation, read/
parse errors, newer opens, edits during reading and disposal cannot silently
replace the current document. Dirty sessions register a browser unload warning.
Browsers control whether that warning is shown.

File → Save and Ctrl/Cmd+S download the opened filename (initially
`Document1.docx`) through `exportDocumentPackage` → `OF.word.writeFromEditorPackage`.
The captured revision becomes the saved baseline only after browser download
handoff; later edits remain dirty. “Download requested” is **not confirmed disk
persistence** or acceptance of a save dialog. Keyboard handling belongs to this
shell and its editor input, not unrelated page inputs.

**Shell import signature change:** `WordShell.importDocumentPackage(package,
fileName)` now requires the filename and applies replacement protection. The
engine's `Word.importFromDocxPackage(editor, package)` remains unchanged.

After mounting, call `WordShell.setPdfFonts(fonts)` before File → Print PDF.
Fonts use the same shape as Spreadsheet: `{ family, faceName?, bold?, italic?,
data }`, where `data` is TrueType bytes (`Uint8Array` or `ArrayBuffer`); typed
Word.Pdf `FontFace` records are also accepted. Hosts supply licensed font files
and enough faces/glyph coverage for their documents. Missing fonts are an error,
not a textless PDF.

The shell lazily loads packaged `word-pdf/Browser.js` and renders a captured
whole-document snapshot through the existing Word PDF engine. The build stages
`word-pdf/ucd.dat` from the resolved Unicode package; custom Fable builds must run
`dotnet fsi scripts/stage-word-pdf-data.fsx <word-pdf-output>`. Keep that asset next
to `Browser.js`. The old `window.RedBatWordPdfExport` bridge is no longer used.
The user gesture reserves a PDF tab; blocked popups and rendering errors are
reported, and failed/stale exports close the reserved tab. PDF export does not
mark DOCX edits saved. Embedded images are decoded by the browser; JPEG bytes
are retained and other image formats become JPEG on white. Linked images remain
subject to the editor's allow-linked-images policy and browser CORS.

### A document that is still being read

| Function | What it does |
| --- | --- |
| `isLoadComplete(editor): boolean` | Whether the document on display is all there. |
| `whenLoadComplete(editor): Promise<boolean>` | Await the document load active when called: true on completion, false on failure, cancellation, replacement or disposal before completion. Each pending waiter settles once; replacing the document cannot turn the outgoing load's cancellation into success. |
| `completeLoad(editor): void` | Read the rest of the document now. |
| `isLoadFailed(editor): boolean` | Whether the document on display could not be read to its end (see isLoadComplete). |
| `cancelLoad(editor): boolean` | Stop reading the document on display: it stays what was read so far, as a document that could not be read to its end (isLoadFailed is true, getLoadError says that the reading was cancelled, whenLoadComplete resolves to false, the listeners are told, it is not saved). |
| `getLoadError(editor): string` | Why the document on display could not be read to its end; null when it was, or is being, read. |

### Changes, undo, revisions of the content

| Function | What it does |
| --- | --- |
| `undo(editor)` | Undo the last change, whether typed or made through the API. |
| `redo(editor)` | Redo the last undone change. |
| `canUndo(editor)` | Whether there is a change to undo. |
| `canRedo(editor)` | Whether there is an undone change to redo. |
| `onChange(editor, callback: (change) => void): () => void` | Call `callback` after every change of the document or the selection (typing, pointer, toolbar / API commands, undo, redo, loading a document). |
| `getRevision(editor): number` | The revision of the document's content: a number that changes with every change of the content (not with the caret or the selection), never repeats for different content, and comes back when undo or redo brings the content back. |
| `markSaved(editor, revision: object): void` | Remember the current content as the saved one: isModified is false until the content differs from it (and false again after undoing or redoing back to it). |
| `isModified(editor): boolean` | Whether the content differs from what was last loaded or marked saved. |

### Viewport, pages and painting

| Function | What it does |
| --- | --- |
| `isLayoutComplete(editor): boolean` | Whether the whole document is laid out. |
| `whenLayoutComplete(editor): Promise<object>` | A promise that resolves (with the document size, see getDocumentSize) once the whole document is laid out; at once when it is. |
| `completeLayout(editor): void` | Lay out the rest of the document now (what getPageCount does, without the count). |
| `getRenderStats(editor)` | Counters of the editor's renderer, for hosts and checks: { FrameBuilds (times a paint frame was built: the document was laid out and its paint commands collected - a caret move builds none), LaidOutBlocks / CollectedBlocks (of the last frame built: the top-level blocks that were laid out / whose paint commands were collected ... |
| `getShapeRects(editor): object[]` | The shapes (text boxes) where they are painted: Id, X (from the left edge of the shape's page), Y (down the document), Width, Height - layout pixels, as getLayoutLines -, Rotation, FlipH / FlipV (the shape mirrored across / upside down; its text is not), BehindText, RelativeHeight, Wrap, AnchorBlockId and Text; Story says where the shape is: "body", or "header" / "footer" for a text box of a header or ... A group or a drawing canvas: IsGroup true, its shapes in Children (see *Groups and drawing canvases*). |
| `getImageHitTest(editor, x: number, y: number)` | The picture under a point in layout coordinates: { BlockId, Offset, X, Y, Width, Height }, or null. |
| `getPageFrames(editor)` | The page frames as shown on the canvas, top to bottom: each page's position and size (sections can differ in page size and orientation) and margins. |
| `getPageCount(editor)` | The number of pages. Reads the rest of a document that is still being read and lays the whole document out. |
| `hitTestPosition(editor, x: number, y: number)` | Document position under a point in layout coordinates. |
| `hitTestViewport(editor, x: number, y: number)` | Document position under a point in layout coordinates (canvas pixels in whole-canvas mode; with a viewport, convert with viewportToDocument or use hitTestViewport), from the layout that is on screen. |
| `setViewport(editor, viewport: object)` | Show the document through a viewport: the canvas becomes exactly as large as the visible area (the engine sizes it: backing store width x height x devicePixelRatio, CSS size width x height), and only the pages and paint commands that intersect the visible band are drawn, at scale x devicePixelRatio. |
| `getViewport(editor)` | The viewport set with setViewport, or null in whole-canvas mode. |
| `getDocumentSize(editor)` | Total size of the document: { Width, Height } in layout pixels (widest page, bottom of the last page), { ScaledWidth, ScaledHeight } in CSS pixels at the viewport's scale (what a scroll spacer needs, plus the host's own origin / margins), and PageCount. |
| `viewportToDocument(editor, x: number, y: number)` | Layout (document) coordinates of a point given in CSS pixels relative to the canvas' top-left corner. |
| `documentToViewport(editor, x: number, y: number)` | CSS pixels relative to the canvas' top-left corner of a point in layout coordinates. |
| `getCaretRect(editor): object` | Where the caret is: { X, Y, Height } in layout coordinates and { ViewportX, ViewportY, ViewportHeight } in CSS pixels relative to the canvas. |
| `repaint(editor)` | Paint everything again (the host resized or cleared the canvas, or images finished loading). |
| `renderPage(editor, canvas: HTMLCanvasElement, pageIndex: number, scale: number)` | Paint one page of the editor's document onto another canvas (thumbnails, print preview). |

### Fonts, text measurers and painters

| Function | What it does |
| --- | --- |
| `onFontsDiscovered(package: object, callback: (faces: object[]) => void): () => void` | The faces a package that is still being read turns out to need as it is read on: see Viewer.onFontsDiscovered - the callback is called synchronously, before the blocks that need the faces are given to the editor, also when the rest is read at once by a call that needs the whole document. |
| `setFontAvailability(editor, isAvailable: object): void` | Say which faces the editor's measurers and painter can set text in: `isAvailable(family, bold, italic)` -> boolean, asked whenever the editor wants to know (so it is never out of date); null: the host does not say, and nothing is ever pending. |
| `setTextMeasurersForFaces(editor, measureText: (text, style) => number, measureRange: (text, style, start, length) => number, faces: object[]): number` | As setTextMeasurers, for a host whose font set gained faces: the new measurers measure as the ones installed, except text in the families of `faces` ([{ family, bold, italic }] - the faces the font set has now and had not; text that named them was set in a fallback). |
| `getPendingFonts(editor): object[]` | The faces the document needs that the host has not got: see setFontAvailability. |
| `setTextMeasurer(editor, measureText: (text, style) => number)` | Measure text with `measureText(text, style)` instead of canvas text. |
| `setTextMeasurers(editor, measureText: (text, style) => number, measureRange: (text, style, start, length) => number)` | As setTextMeasurer, with `measureRange(text, style, start, length)` for the width of a part of a shaped run. |
| `setTextRangeMeasurer(editor, measureRange: (text, style, start, length) => number)` | Install the measurer of a range of a run's text (see setTextMeasurers, which installs both). |
| `setTextPathPainter(editor, paintText: (text, style, x, baseline, width) => boolean)` | Paint text with `paintText(text, style, x, baseline, width)` - a plain JavaScript function returning true when it painted the text - instead of canvas text, on the editor's own canvas. |
| `clearTextPathPainter(editor)` | Remove the text path painter: text is painted as canvas text again. |
| `clearTextMeasurer(editor)` | Remove the text measurers: text is measured with the canvas again (the layout is made again). |
| `getRequiredFonts(target: object): object[]` | The font faces a document's text is set in: [{ family, bold, italic }], after style resolution, with the faces the engine itself introduces (a list marker in Symbol, Wingdings or Courier New; note text; headers and footers). |
| `setFontFamilies(editor, families: string[])` | Tell the editor which font families the host's text stack can draw (getFontFamilies then offers them). |
| `getFontFamilies(editor): string[]` | The font families to offer in a font box: the ones the text stack installed (setFontFamilies) and the ones the document uses (its text and its styles), sorted, each once. |
| `setTextPathPainterFactory(editor, painterFor: (context2d) => (text, style, x, baseline, width) => boolean)` | As setTextPathPainter, for every canvas the editor paints on: `painterFor(context2d)` returns the painter for that canvas context, so renderPage paints its pages with it too (a painter draws on one context). |

### Selection and what the caret is in

| Function | What it does |
| --- | --- |
| `getSelectionFormat(editor)` | The formatting of the current selection, for toolbar state. |
| `getCaretContext(editor): object` | What the caret is in, read off the document alone (no layout pass: a few microseconds whatever the document's length, where getSelectionInformation lays the caret's surroundings out): { InTable, Row, Column (0-based; -1 outside a table), RowCount, ColumnCount, CanMergeCells (the selection spans several cells), CanSplitCell, InHyperlink, InList, InHeaderFooter, Story ("body" \| "header" \| "footer" \| "textbox" \| "footnote" \| "endnote" \| "comment"), TextBoxId, NoteId / CommentId (the note or comment edited in place, else null), HeaderFooter, ImageSelected }. |
| `moveToAdjacentTableCell(editor, forward: boolean): boolean` | Tab (forward) / Shift+Tab in a table, for a host's own keys or buttons: true when the caret moved to another cell. |
| `getSelection(editor)` | The selection: { AnchorBlockId, AnchorOffset, FocusBlockId, FocusOffset, IsCollapsed }, or null when there is none. |
| `setSelection(editor, anchorBlockId: string, anchorOffset: number, focusBlockId: string, focusOffset: number): boolean` | Select from one document position to another (the same position twice places the caret). |
| `clearSelection(editor)` | Take the selection (and the caret) away. |
| `selectAll(editor)` | Select all text. |

Without Shift, Left/Right collapses a non-empty selection to its start/end
without stepping past that boundary, including at document edges and across
table cells. The browser and native host use the same navigation policy.

In the Word shell, the toolbar reflects the effective format of the current
selection. A formatting command changes that selection only; at a caret it
uses the editor's pending-format behavior (including formatting the word when
the caret is inside one). These edits participate in the editor's undo/redo
history and update the shell's modified indicator immediately. Selection or
caret formatting state is not a separate host-side copy.

`getSelectionFormat` reports effective values, `"mixed"` where selected
content differs, and `null` where no value is stated. Character values include
`AllCaps` and `SmallCaps`. `Paragraph.BackgroundColor` and
`Paragraph.ListNumberFormat` follow the same convention; a paragraph outside
a numbered list has no number format.

The shell uses `OfficeUi.Commands` for menu, toolbar and owned keyboard routes.
Ctrl/Cmd+B/I/U, undo/redo, Ctrl/Cmd+F/H and file shortcuts run once; dialog fields,
foreign inputs and composition retain their own input handling. Mixed toggles
use `aria-pressed="mixed"`; font and paragraph controls follow selection and
undo/import changes instead of storing independent formatting flags.

Format → Paragraph style reads the selection on every open. Mixed fields stay
unspecified; Apply patches only changed fields in one undo step. Opening,
cancelling or applying an unchanged dialog does not normalize mixed formatting.
Background accepts a color string; an empty edited background restores inheritance.

Format → Font edits only changed character properties of the selection in one
undo step; at a caret it sets the format for subsequently typed text. Mixed
values remain mixed until changed. Format → Borders and shading edits each
side independently and the background and pattern of selected paragraphs;
width and spacing are layout pixels. Available authored border styles are
single, dotted, dashed and dot-dash. Pattern choices are those painted by the
editor; other imported patterns remain available as their current value and
are preserved when unchanged. Both dialogs validate input and reject a stale
document or selection, retain errors, and leave Cancel and unchanged Apply
out of history.

Font fields cover family, point size, bold/italic, underline mode, strike,
caps/small caps, super/subscript, text/highlight colors and hidden text.
Unedited fields retain each run's formatting rather than flattening a mixed
selection. Border sides are independent; imported unsupported choices are
shown as current values instead of being replaced merely by opening the dialog.

`getSelectionFormat` also reports `Hidden` and `UnderlineMode` (`"none"`,
`"single"`, an underline style name, or `"mixed"`). `CharacterStyleId` is
the selected runs' named character style id, `null` when none is set, or
`"mixed"` when they differ; at a caret it includes a pending typing style.
Applying a character style preserves unrelated direct formatting. A selected
range is one undoable edit; a caret choice is pending until text is typed.
The dialog transactions are shell
implementation details, not additional JavaScript factory methods.


### Typing, pasting, find and replace

| Function | What it does |
| --- | --- |
| `findText(editor, query: string, matchCase: boolean)` | Find literal text matches across editable paragraphs. |
| `selectFindMatch(editor, matchObj: object)` | Select a previously returned find match. |
| `replaceAllText(editor, query: string, replacement: string, matchCase: boolean)` | Replace every occurrence of `query` with `replacement` (literal text; `matchCase`), as one undo step; returns how many were replaced. |

The shell's **Find and replace** dialog searches literal text. Find next
selects the next match, cycles back to the first, and honors the Match case
checkbox; its result reports the selected match number and total. Navigation
starts after the normalized selection end, including backwards selections.
Every invocation searches the current document rather than cached offsets.
Replace all replaces literal matches under the same case rule as one undoable
edit and reports the replacement count. An empty query asks for search text.
Both operations use the editor APIs above, so changes are reflected in the
document and modified indicator as they occur.

| `insertText(editor, text: string)` | Insert text at cursor. |
| `pasteText(editor, text: string): void` | Put plain text in at the caret, or over the selection, as Ctrl+V does with the clipboard's text (a host that reads the clipboard itself, or pastes from a menu, calls this): the lines of the text are paragraphs, the whole is one undo step, and with nothing selected the text has the format chosen at the caret for what is typed ... |
| `pasteClipboard(editor, text: string, html: string): void` | Paste supplied clipboard content: an explicitly identified, unchanged internal rich payload; otherwise HTML formatting (tables as tables between paragraphs; in a table cell or text box their rows as paragraphs, cells parted by tabs); otherwise plain text. Matching text alone never reuses an earlier copy's formatting. One undo step. |
| `copySelection(editor): { Text, Html } \| null` | Copy the selection as Ctrl+C does: the text and the HTML to write to the clipboard; kept for a paste. Null with nothing selected. |
| `cutSelection(editor): { Text, Html } \| null` | `copySelection`, and the selection deleted (one undo step). |

### Character formatting

| Function | What it does |
| --- | --- |
| `toggleBold(editor)` | Toggle bold on selection. |
| `toggleItalic(editor)` | Toggle italic on selection. |
| `toggleUnderline(editor)` | Toggle underline on selection. |
| `toggleAllCaps(editor)` | Toggle all caps on selection. |
| `toggleSmallCaps(editor)` | Toggle small caps on selection. |
| `toggleStrike(editor)` | Toggle strikethrough on selection. |
| `clearFormatting(editor)` | Clear direct character formatting on selection. |
| `setVertAlign(editor, align: string)` | Set vertical alignment on selection. |
| `setColor(editor, color: string)` | Set text color on selection. |
| `setHighlight(editor, color: string)` | Set highlight color on selection. |
| `setCharacterStyle(editor, styleId: string)` | Give selected text the named character style, or choose the style for text typed next at a caret (even inside a word). An empty id clears the named style. |
| `setFontSize(editor, size: number)` | Set font size on selection. |
| `setFontSizePt(editor, points: number)` | Set the font size of the selection in points (11, 12, 14 ...: the unit documents state sizes in and a size box shows). |
| `setFontFamily(editor, family: string)` | Set font family on selection. |

### Paragraph formatting, styles and lists

| Function | What it does |
| --- | --- |
| `setNumberFormat(editor, format: string)` | Set numbered-list marker format on current block. |
| `setParagraphStyle(editor, styleId: string)` | Give the selected paragraphs the named paragraph style (a style id of getStyleCatalog). |
| `getStyleCatalog(editor)` | The styles of the document, for a style gallery. |
| `getStyles(editor): object[]` | The document's styles, for a style box: [{ Id, Name, Type ("paragraph" \| "character" \| "table" \| "numbering"), IsDefault, IsQuickStyle, Hidden, HeadingLevel (1-9, or 0), BasedOn (id or null) }]. |
| `setLineSpacing(editor, multiple: number)` | Line spacing of the selected paragraphs as a multiple of single spacing (1, 1.15, 1.5, 2 ...); 0 or less goes back to the style's. |
| `setSpacingBefore(editor, points: number)` | Space before / after the selected paragraphs, in points; a negative value goes back to the style's. |
| `setSpacingAfter(editor, points: number)` | The space after the selected paragraphs, in points (a negative number: the style's again). |
| `applyParagraphStyleChanges(editor, changes: object)` | Patch only supplied fields in one undo step: `Alignment`, `IndentLeft`, `IndentRight`, `SpacingBefore`, `SpacingAfter`, `LineSpacing`, `FirstLineIndent`, `BackgroundColor`, `FontSizePt`. Indents are layout pixels, spacing and font size are points, line spacing is a multiple. Missing/null fields are unchanged; an empty background string restores inheritance. Font size affects selected text, not the whole document. Used by the browser paragraph dialog. |
| `applyParagraphStyle(editor, alignment: string, indentLeft: number, indentRight: number, spacingBefore: number, spacingAfter: number, lineSpacing: number, fontSize: number, firstLineIndent: number \| undefined, backgroundColor: string)` | Alignment, indents (px), spacing before / after (px), line spacing (multiple), font size, first-line indent and background of the selected paragraphs in one step. |
| `setHeadingLevel(editor, level: number)` | Apply Heading 1-3 to the current block. |
| `clearHeading(editor)` | Convert the current heading block back to a normal paragraph. |
| `alignLeft(editor)` | Align left. |
| `alignCenter(editor)` | Align center. |
| `alignRight(editor)` | Align right. |
| `alignJustify(editor)` | Align justify. |
| `toggleBulletList(editor)` | Make the selected paragraphs a bulleted list, or plain paragraphs when they all are one. |
| `toggleNumberedList(editor)` | Make the selected paragraphs a numbered list, or plain paragraphs when they all are one. |
| `increaseIndent(editor)` | The selected list items one level down (what Tab does at the start of a list item). |
| `decreaseIndent(editor)` | The selected list items one level up (Shift+Tab); from the first level out of the list. |
| `setListLevel(editor, level: number)` | The selected list items at a level of their list (1-9; see increaseIndent). |

### Pages, sections, headers and footers

| Function | What it does |
| --- | --- |
| `enterHeaderFooter(editor, pageIndex: number, kind: string): boolean` | Edit the header or the footer (`kind`: "header" \| "footer") of a page in place: the caret goes into it, the body is dimmed and a labelled line marks where it ends. |
| `exitHeaderFooter(editor): boolean` | Back to the body from a header or footer that is edited in place (Escape and a double click on the body do the same). |
| `enterNote(editor, noteId: string, noteType: string): boolean` | Edit the body of a footnote or endnote (`noteType` "footnote" \| "endnote") in place: the caret goes to the start of its text, where the page shows the note (see *Notes and comments edited in place*). False when there is no such note, or nothing in the text refers to it. |
| `enterComment(editor, commentId: string): boolean` | Edit the body of a comment in place, in a balloon below its anchor (see *Notes and comments edited in place*). False when there is no such comment, or nothing in the text refers to it. |
| `exitNote(editor): boolean` | Back to the body from a note or comment edited in place (Escape and a click off the note do the same); the body's selection is what it was. False when none was edited. |
| `getHeaderFooterLinkToPrevious(editor): boolean` | Whether the header or footer being edited is the one of the section before it ("Link to Previous"). |
| `setHeaderFooterLinkToPrevious(editor, linked: boolean): boolean` | "Link to Previous" for the header or footer being edited. |
| `insertPageNumberField(editor)` | Insert a page number field (PAGE) at the caret: every page shows its own number where the field stands in a header or footer - as its section counts and writes it (a section's first page number and number format, w:pgNumType). |
| `insertPageCountField(editor)` | Insert a page count field (NUMPAGES) at the caret: the number of pages of the document (what getPageCount says). |
| `insertSectionPageCountField(editor)` | Insert a field with the number of pages of the caret's section (SECTIONPAGES). |
| `insertPageBreak(editor): boolean` | Insert a page break at the caret (what Ctrl+Enter does). False - and nothing is done - in a text box, a header, a footer or a note. |
| `insertColumnBreak(editor): boolean` | (0.1.17) Insert a column break at the caret (Ctrl+Shift+Enter): what follows starts the next column of its section, the next page where the section has one column. Saved as `<w:br w:type="column"/>`. False where `insertPageBreak` is, and in a table cell. |
| `insertLineBreak(editor): void` | (0.1.17) Insert a line break at the caret (Shift+Enter): the paragraph goes on in the next line (`<w:br/>`). |
| `setSectionBreak(editor, kind: string)` | A section break after the caret's paragraph (a paragraph of the body). |
| `setSectionVerticalAlignment(editor, align: string): bool` | How the text of a page that is not full stands between the top and the bottom margin in the section the caret is in (w:vAlign): "top", "center", "bottom" or "both" (the paragraphs spread from the top to the bottom). |
| `getSectionVerticalAlignment(editor): string` | The vertical alignment of the section the caret is in: "top", "center", "bottom" or "both" (see setSectionVerticalAlignment). |
| `setPageSetup(editor, pageWidth: number, pageHeight: number, marginLeft: number, marginRight: number, marginTop: number, marginBottom: number, headerDistance: number, footerDistance: number)` | Page size, margins and header / footer distances (layout pixels) of the document (its last section). |
| `setPageSetupOptions(editor, pageWidth: number, pageHeight: number, marginLeft: number, marginRight: number, marginTop: number, marginBottom: number, headerDistance: number, footerDistance: number, differentFirstPage: boolean, evenAndOddHeaders: boolean)` | As setPageSetup, with whether the first page and the even pages have headers and footers of their own. |
| `getPageSetup(editor)` | The page setup of the document (its last section): PageWidth, PageHeight, margins, HeaderDistance, FooterDistance, DifferentFirstPage, EvenAndOddHeaders - layout pixels. |
| `setHeaderFooterText(editor, defaultHeader, defaultFooter, firstHeader, firstFooter, evenHeader, evenFooter)` | Plain-text default, first-page and even-page headers and footers (an empty string: none). |
| `getHeaderFooterText(editor)` | The default / first-page / even-page headers and footers as setHeaderFooterText, setHeaderFooterTableText and setHeaderFooterImageText take them (DefaultHeader, DefaultFooter, DefaultHeaderTableLeft, ...). |
| `setHeaderFooterTableText(editor, defaultHeaderLeft, defaultHeaderRight, defaultFooterLeft, defaultFooterRight)` | A default header and footer made of a one-row table with a left and a right cell. |
| `setHeaderFooterImageText(editor, defaultHeaderSrc, defaultHeaderAlt, defaultHeaderWidth, defaultHeaderHeight, defaultFooterSrc, defaultFooterAlt, defaultFooterWidth, defaultFooterHeight)` | A default header and footer made of a picture (source, alternative text, width, height). |
| `setCurrentSectionHeaderFooterText(editor, headerText, footerText)` | Plain-text header and footer of the section the caret is in. |

The browser shell's **File > Page setup** instead edits only the section containing the selection's focus. It offers Letter, Legal, A4, A5 and custom paper sizes, orientation, margins and header/footer distances in inches. It prefills from that section, preserves fields not edited, and applies the change as one undo step. It does not change every section; the API functions in the table above continue to address the final section. A changed document or selection, a replaced editor, or a read-only input requires reopening the dialog.

### Tables

| Function | What it does |
| --- | --- |
| `insertTableRowBefore(editor)` | Insert a table row above the row containing the current selection. |
| `insertTableColumnBefore(editor)` | Insert a table column to the left of the column containing the current selection. |
| `canMergeTableCells(editor): boolean` | Whether the selection spans several cells of one table (mergeTableCells would merge them). |
| `mergeTableCells(editor)` | Merge the selected table cells into one. |
| `splitTableCell(editor)` | Split the merged cell the caret is in back into single cells. |
| `insertTableRowAfter(editor)` | Insert a table row after the row containing the current selection. |
| `insertTableColumnAfter(editor)` | Insert a table column after the column containing the current selection. |
| `deleteTableRow(editor)` | Delete the table row containing the current selection. |
| `deleteTableColumn(editor)` | Delete the table column containing the current selection. |
| `insertTable(editor, rows: number, columns: number)` | Insert a table of `rows` x `columns` after the block the caret is in. |
| `setSelectedTableCellStyle(editor, backgroundColor: string, cellWidth: number, borderColor: string, borderWidth: number, verticalAlignment: string, paddingTop: number, paddingRight: number, paddingBottom: number, paddingLeft: number)` | Background, width, border colour and width, vertical alignment and padding (px) of the selected table cells. |
| `setSelectedTableCellStyleWithShading(editor, backgroundColor: string, shadingPattern: string, shadingForegroundColor: string, cellWidth: number, borderColor: string, borderWidth: number, verticalAlignment: string, paddingTop: number, paddingRight: number, paddingBottom: number, paddingLeft: number)` | As setSelectedTableCellStyle, with a shading pattern and its foreground colour. |
| `setSelectedTableCellShading(editor, backgroundColor: string, shadingPattern: string, shadingForegroundColor: string)` | The shading of the cell at the caret alone: a background colour (`shadingPattern` ""), or a pattern with its colours; "none" takes it away. The cell's width, borders, alignment and margins stay (the two above set all of them: an empty value clears one). |
| `setCurrentTableCellMargins(editor, paddingTop: number, paddingRight: number, paddingBottom: number, paddingLeft: number)` | The padding (px: top, right, bottom, left) of the cells of the caret's table. |
| `setCurrentTableCellMarginsAndSpacing(editor, paddingTop: number, paddingRight: number, paddingBottom: number, paddingLeft: number, cellSpacing: number)` | As setCurrentTableCellMargins, with the spacing between the cells. |

### Pictures, text boxes, links, notes, fields and other inserts

| Function | What it does |
| --- | --- |
| `setHyperlink(editor, target: string)` | Set hyperlink on selection. |
| `setComment(editor, commentId: string, commentText: string, author: string, date: string)` | Set comment metadata on selection. |
| `setNote(editor, noteId: string, noteType: string, noteText: string)` | Set footnote/endnote metadata on selection. |
| `refreshFields(editor)` | Bring the fields of the document up to date (page numbers, page counts, references, tables of contents) from the layout. |
| `insertHyperlink(editor, target: string, text: string)` | Insert a hyperlink to `target`: the selected text becomes the link; with a caret, `text` (or the target, when it is empty) is inserted as the link. |
| `getHyperlink(editor): object` | The hyperlink the caret is in: { Target, BlockId, StartOffset, EndOffset, Text }, or null. |
| `setHyperlinkTarget(editor, target: string)` | Change the target of the hyperlink the caret is in (the whole link); null or "" removes the link and keeps its text. |
| `insertImage(editor, src: string, width: number, height: number, alt: string)` | Insert image at cursor. |
| `resizeCurrentImage(editor, width: number, height: number)` | Give the selected picture a width and a height (layout pixels). |
| `insertTextBox(editor, text: string): boolean` | Insert a text box with `text`, anchored in the paragraph (or the table) the caret is in: 240 x 96 px, 24 px right of and below the start of that paragraph, in front of the text (no wrap). |
| `moveShape(editor, id: string, dx: number, dy: number): boolean` | Move a shape (a text box) of the body, or of a header or footer, by `dx` to the right and `dy` down, in layout pixels (`id`: as getShapeRects gives it). |
| `resizeShape(editor, id: string, width: number, height: number): boolean` | Make a shape of the body, a header or a footer `width` wide and `height` high (layout pixels, at least 1 each); its top left corner stays where it is (move it with moveShape to resize by another corner). |
| `selectShape(editor, id: string \| null): boolean` | (0.1.17) Select a shape of the body, a header or a footer (`null` or `""`: none): the editor paints its frame and handles and handles the pointer and the keys for it (see Text boxes). The text of another box, a header or a footer that is edited is left first. False when there is no such shape. |
| `getSelectedShape(editor): string \| null` | (0.1.17) The Id of the shape that is selected - by `selectShape` or a click into a box - or null. |
| `resizeShapeBy(editor, id: string, handle: string, dx: number, dy: number): boolean` | (0.1.17) Resize a shape by one of its handles ("nw", "n", "ne", "e", "se", "s", "sw", "w") dragged by `dx`, `dy` layout pixels: the side or corner across stays where it is, for a turned shape too (the drag in the shape's own directions). One undo step. At least four pixels wide and high. |
| `deleteShape(editor, id: string): boolean` | Take a shape of the body, a header or a footer out of the document, its text with it. |
| `getShapeHandles(editor, id: string): object[]` | The eight handles of a shape of the body, a header or a footer, for a host that shows it selected and lets it be dragged: Name ("nw", "n", "ne", "e", "se", "s", "sw", "w"), X, Y, Width, Height and Cursor. |
| `hitTestShapeHandle(editor, id: string, x: number, y: number): string` | What of a shape of the body is at a point (layout pixels, as getShapeHandles): the Name of the handle there, "move" inside the shape where no handle is, "" outside it. |
| `enterTextBox(editor, id: string): boolean` | Edit the text of a text box of the body in place (`id`: as getShapeRects gives it): the caret goes to the start of its text, and typing, selecting, formatting and undo work there as in the body. False for a group or a drawing canvas. |
| `exitTextBox(editor): boolean` | Back to the body from a text box that is edited in place. |
| `insertCaption(editor, label: string, captionText: string)` | Insert a caption paragraph (label, a SEQ number, text) after the caret's paragraph. |
| `insertCrossReferenceToLatestCaption(editor, label: string)` | Insert a reference field to the last caption with the given label. |
| `insertEquation(editor, equationText: string)` | Insert an equation (linear text) at the caret. |
| `setCurrentTextBoxStyle(editor, x: number, y: number, width: number, height: number, fillColor: string, borderColor: string, borderWidth: number)` | Position (x, y), size, fill, border colour and border width of the text box the caret is in. |

### Bookmarks and content controls

| Function | What it does |
| --- | --- |
| `setBookmark(editor, bookmarkName: string)` | Put a bookmark of the given name around the selection (or at the caret). |
| `setContentControl(editor, controlId: string, tag: string, alias: string)` | Wrap the selected paragraphs in a content control with the given id, tag and alias. |
| `getBookmarkInfos(editor)` | The bookmarks of the document: Id, Name, BlockId, StartOffset, EndOffset. |
| `selectBookmark(editor, bookmarkId: string)` | Select the text of the bookmark with the given name; false when there is none. |
| `removeBookmark(editor, bookmarkId: string)` | Remove the bookmark with the given name (its text stays). |
| `renameBookmark(editor, bookmarkId: string, bookmarkName: string)` | Give a bookmark another name. |
| `getContentControlInfos(editor)` | One entry per physical block-level control with an ID: Id, Tag, Alias, BlockId. Nested outer controls are not duplicated; table-only controls use their first cell paragraph as BlockId. |
| `selectContentControl(editor, controlId: string)` | Select the first physical control matching the ID or tag, including its nested paragraphs and table-cell text. Duplicate IDs cannot identify a later group through this API. |
| `removeContentControl(editor, controlId: string)` | Remove the content control with the given id (its paragraphs stay). |
| `updateContentControl(editor, controlId: string, newControlId: string, tag: string, alias: string)` | Change the id, tag and alias of a content control. |

Block-level bindings, unknown control properties, end properties and nested/adjacent
group boundaries survive supported edits and saves. ID-less bindings are preserved
but not exposed as editable IDs. Replacing the middle of a control creates separate
physical wrappers; discovery and selection keep those wrappers separate.
Imported inline (run-level) controls also retain their grouping and binding metadata
through supported edits and regenerated-model saves. The discovery/authoring APIs
above remain block-level; this is not a live custom-XML binding engine.
Row/cell-level controls and textbox control wrappers remain outside this model.

### Comments and tracked changes

| Function | What it does |
| --- | --- |
| `getTrackChanges(editor)` | Whether typing and deleting are recorded as tracked changes. |
| `setTrackChanges(editor, enabled)` | Switch change tracking on or off. |
| `setTrackChangesOptions(editor, author, date)` | The author and the date (ISO string) new tracked changes are recorded with. |
| `toggleTrackChanges(editor)` | Switch change tracking to the other state; returns the new state. |
| `getCommentAnchors(editor)` | The comments with where their anchors are painted: Id, Text, Author, Date, Excerpt, BlockId, StartOffset, EndOffset, X, Y, Width, Height, CalloutX, CalloutY, IsSelected. |
| `getRevisionInfos(editor)` | The tracked changes of the document, one object per run of changed text: BlockId, Start, End, RevisionType, Author, Date, Text. |
| `selectCommentAnchor(editor, anchor: object)` | Select the text a comment (an object of getCommentAnchors) is anchored to. |
| `selectRevisionInfo(editor, revisionObj: object)` | Select the text of a tracked change (an object of getRevisionInfos). |
| `setRevision(editor, revisionType: string, revisionId: string, author: string, date: string)` | Mark the selected text as a tracked change of the given type (revisionType, revisionId, author, date). |
| `acceptSelectedRevision(editor)` | Accept the tracked changes in the selection. |
| `rejectSelectedRevision(editor)` | Reject the tracked changes in the selection. |
| `acceptAllRevisions(editor)` | Accept every tracked change of the document. |
| `rejectAllRevisions(editor)` | Reject every tracked change of the document. |
| `selectNextRevision(editor)` | Select the next tracked change after the caret; false when there is none. |
| `selectPreviousRevision(editor)` | Select the tracked change before the caret; false when there is none. |

### Layout queries

| Function | What it does |
| --- | --- |
| `getDuplicateBlockIds(editor): string[]` | The block ids the document has more than once - among its paragraphs (those of table cells, of tables in cells, of text boxes, headers and footers too), its tables and its text boxes. |
| `getLayoutSegments(editor)` | Every laid-out segment (a run's part on a line): Kind, Text, BlockId, StartOffset, X, Y, Width and more - for hosts and tests that compare layouts. Lays the whole document out. |
| `getLayoutLines(editor)` | Every laid-out line: BlockId, Text, Page, X, Y, PageY, Width, Height, StartOffset, EndOffset, Indent, ListMarker, table and page position fields, Segments. Lays the whole document out. |
| `getLayoutFingerprint(editor)` | The laid-out document block by block, for comparing two layouts (see LayoutFingerprint): [{ Page, Y, Height, LineCount, LineHeights, Text, Paragraph, Runs }] with the resolved paragraph and run formatting as strings. |
| `getVbaLayoutLines(editor)` | The laid-out lines as Word's VBA object model reports them (points, positions relative to the page) - what the VBA layout comparison reads. |
| `getRangeInformation(editor, blockId: string, startOffset: number, endOffset: number)` | What Word's Range.Information reports for a range of a paragraph (page, line and column numbers, positions relative to the page and the text boundary, table position). |
| `getSelectionInformation(editor)` | getRangeInformation for the selection's focus. |
| `getRangeBounds(editor, blockId: string, startOffset: number, endOffset: number)` | Where a range of a paragraph's text is: BlockId, StartOffset, EndOffset, X, Y, PageX and the page it is on. |
| `getWordCounts(editor)` | Body and active-selection word counts; see the status bar semantics above. |
| `computeRangeStatistics(editor, blockId: string, startOffset: number, endOffset: number)` | Statistics of a range of a paragraph's text (what Word's Range.ComputeStatistics counts). |

## The viewer

`Word.createViewer(canvas)` (or the default export of
`@rbt/office/word/WordViewer.js`, which imports no editor code) returns an
object with these methods; `canvas` is used to measure text.

| Method | What it does |
| --- | --- |
| `load(editorPackage, options?)` | Show a package of `forEditorPackage` / `openProgressive`. `{ progressive: true }`: the start is laid out first, the rest in slices. |
| `isLayoutComplete()`, `whenLayoutComplete()`, `onLayoutProgress(callback)` | As the editor's; the callback gets `{ PageCount, LayoutComplete, LoadComplete, LoadFailed, LoadError, FontsPending }` and the call returns the function that ends the subscription. |
| `isLoadComplete()`, `whenLoadComplete()`, `completeLoad()`, `isLoadFailed()`, `getLoadError()`, `cancelLoad()` | A package that is still being read. |
| `getPageCount()`, `getPageFrames()`, `getDocumentSize()` | The pages (all of them: the rest is laid out at once), their frames, the document's size. |
| `renderPage(canvas, pageIndex, scale)` | Paint one page onto a canvas of the host's. |
| `setImageCache(map)` | Map of image source to decoded image. |
| `setTextMeasurers(measureText, measureRange)`, `setTextMeasurersForFaces(...)`, `setTextMeasurer`, `setTextRangeMeasurer`, `clearTextMeasurer()` | Measure text as an editor with the same measurers does. |
| `setTextPathPainter(painterFor)` | `painterFor(context2d)` gives the text painter for a canvas. |
| `setFontAvailability(isAvailable)`, `getPendingFonts()`, `getRequiredFonts()` | The faces the document is set in, and those the host has not got. |
| `getRenderStats()` | Counters for hosts and tests (frames built, blocks laid out). |
| `render(document)`, `dispose()` | Paint a document of the model on the viewer's canvas; end the viewer. |
