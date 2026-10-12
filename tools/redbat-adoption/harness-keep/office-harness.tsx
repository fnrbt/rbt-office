// TEMPORARY test harness - not committed.
import { render } from "solid-js/web";
import { Suspense, lazy, createSignal, Show } from "solid-js";
import { Id } from "@rbt/client";
import '@fontsource/inter/latin-400.css'
import '@fontsource/inter/latin-500.css'
import '@fontsource/inter/latin-600.css'
import '@fontsource/inter/latin-700.css'
import './app.css'
import './rb.css'
import './rb-chat.css'
import './rb-mega.css'
import { ThemeProvider } from '~/lib/theme'

const params = new URLSearchParams(location.search);
const component = params.get("c") ?? "";
const url = params.get("url") ?? "";
const name = params.get("name") ?? url.split("/").pop() ?? "file";
const w: any = window;
w.__saved = 0;
w.__grids = [];
w.__editors = [];
w.__closed = 0;
w.__created = [];
w.__formats = () => import("@rbt/office/office-formats/src/Main.js");

// The upload goes over the network (mocked by the test: recorded, delayed or failed there).
const saveFile = async (file: globalThis.File) => {
  const response = await fetch("/api/upload", { method: "POST", body: file });
  if (!response.ok) throw new Error(`The server did not accept the upload (${response.status})`);
  w.__saved++;
  return null;
};

async function hooks() {
  const isWord = /\.docx$/i.test(name);
  if (component.startsWith("sheet") || ((component === "chat-viewer" || component === "files-viewer") && !isWord)) {
    const mod: any = await import("@rbt/office/spreadsheet/Spreadsheet.js");
    w.__sheetModel = await import("@rbt/office/spreadsheet/Model.js");
    const mount = mod.Spreadsheet.prototype.Mount;
    mod.Spreadsheet.prototype.Mount = function () {
      w.__grids.push(this);
      return mount.call(this);
    };
  }
  if (component === "word-editor" || params.has("hookWord")) {
    const mod: any = await import("@rbt/office/word/Word.js");
    const init = mod.default.initializeCanvas;
    mod.default.initializeCanvas = (canvas: any) => {
      const editor = init(canvas);
      w.__editors.push(editor);
      return editor;
    };
    w.__word = mod.default;
    w.__wordModule = mod;
  }
}

if (params.has("flashcss")) void import("./apps/flashcards/Spreadsheet.css");
w.__viewerModule = () => import("@rbt/office/word/WordViewer.js");
w.__stack = () => import("~/lib/word-text-stack/browser.js");
const SpreadsheetEditor = lazy(() => import("~/components/office/SpreadsheetEditor"));
const WordEditor = lazy(() => import("~/components/office/WordEditor"));
const SpreadsheetContent = lazy(() => import("~/components/file/SpreadsheetContent"));
const WordDocumentContent = lazy(() => import("~/components/file/WordDocumentContent"));
const ChatFileViewer = lazy(() => import("~/components/file/FileViewer").then((m) => ({ default: m.FileViewer })));
const FilesViewer = lazy(() => import("~/components/files/FileViewer").then((m) => ({ default: m.FileViewer })));

function filesFile(): any {
  try {
    return {
      id: Id.fromRoute("file1"),
      name,
      mimeType: params.get("mime") ?? "",
      sizeBytes: 1234,
      storedFileId: Id.fromAttestedRoute("stored1:1:abc"),
      contentBlobId: Id.fromAttestedRoute("stored1:1:abc"),
      isStarred: false,
      isTrashed: false,
      createdAt: 1750000000,
      updatedAt: 1750000000,
    };
  } catch (err) {
    w.__idError = String(err);
    return null;
  }
}

function App() {
  const [open, setOpen] = createSignal(true);
  const [scale] = createSignal(Number(params.get("scale") ?? 1));
  w.__setOpen = setOpen;
  const close = () => { w.__closed++; setOpen(false); };
  const file = component === "files-viewer" ? filesFile() : null;
  return (
    <ThemeProvider>
      <Suspense fallback={<div>loading chunk</div>}>
        <Show when={open()}>
          {component === "sheet-editor" && <SpreadsheetEditor url={url} filename={name} kind={name.endsWith(".csv") ? "csv" : "xlsx"} saveFile={saveFile} onClose={close} />}
          {component === "word-editor" && <WordEditor url={url} filename={name} saveFile={saveFile} onClose={close} />}
          {component === "sheet-preview" && (
            <div class="rb-file-viewer-overlay"><div class="rb-file-viewer-shell"><div class="rb-file-viewer-content"><div class="rb-file-viewer-stage">
              <div class="rb-file-viewer-stage-scroll" data-testid="stage-scroll">
                <SpreadsheetContent url={url} filename={name} mimeType={params.get("mime") ?? undefined} />
              </div>
            </div></div></div></div>
          )}
          {component === "word-preview" && (
            <div class="rb-file-viewer-overlay"><div class="rb-file-viewer-shell"><div class="rb-file-viewer-content"><div class="rb-file-viewer-stage">
              <div class="rb-file-viewer-stage-scroll" data-testid="stage-scroll">
                <WordDocumentContent url={url} scale={scale()} onLoaded={(d: any) => { w.__wordLoaded = d; w.__wordFirst = d.pageCount; }} onPages={(n: number, complete: boolean) => { w.__wordSoFar = n; if (complete) w.__wordPages = n; }} />
              </div>
            </div></div></div></div>
          )}
          {component === "chat-viewer" && <ChatFileViewer file={{ name, url, mimeType: params.get("mime") ?? undefined }} onClose={close} />}
          {component === "files-viewer" && file && <FilesViewer file={file} onClose={close} onFileCreated={(f: any) => w.__created.push(f)} />}
        </Show>
      </Suspense>
    </ThemeProvider>
  );
}

hooks().then(() => {
  render(() => <App />, document.getElementById("root")!);
  w.__mounted = true;
});
