module Rbt.FableUI.Chrome

open Fable.Core

// A document owns its styles, not an individual editor instance. Both bundles
// may call this; disposing one shell must not unstyle another mounted shell.
[<Emit("(function(key, css) { const id = 'rbt-office-style-' + key; if (!document.getElementById(id)) { const style = document.createElement('style'); style.id = id; style.textContent = css; document.head.appendChild(style); } })($0, $1)")>]
let private install (_key: string) (_css: string) : unit = jsNative

let private stylesheet = """
.rb-office-shell, .rb-office-dialog {
  --office-surface: #ffffff;
  --office-desk: #f4f6f8;
  --office-text: #202124;
  --office-muted: #5f6368;
  --office-border: #dadce0;
  --office-hover: #f1f3f4;
  --office-selected: #e8f0fe;
  --office-accent: #1a73e8;
  --office-radius: 6px;
  color: var(--office-text);
  font: 14px/1.4 system-ui, -apple-system, sans-serif;
  box-sizing: border-box;
}
.rb-office-shell { background: var(--office-surface); min-width: 0; max-width: 100%; }
.rb-office-shell *, .rb-office-dialog * { box-sizing: border-box; }
.rb-office-shell button, .rb-office-shell input, .rb-office-shell select,
.rb-office-dialog button, .rb-office-dialog input, .rb-office-dialog select, .rb-office-dialog textarea {
  font: inherit; color: inherit; border-radius: var(--office-radius);
}
.rb-office-shell button, .rb-office-dialog button {
  min-height: 32px; padding: 5px 10px; border: 1px solid transparent;
  background: transparent; cursor: pointer; line-height: 1.3;
}
.rb-office-shell button:hover:not(:disabled), .rb-office-dialog button:hover:not(:disabled) { background: var(--office-hover); }
.rb-office-shell button[aria-pressed="true"], .rb-office-shell button[aria-checked="true"] {
  background: var(--office-selected); color: var(--office-accent); border-color: var(--office-accent);
}
.rb-office-shell button[aria-pressed="mixed"] { background: var(--office-selected); border-style: dashed; border-color: var(--office-accent); }
.rb-office-shell button:disabled, .rb-office-dialog button:disabled { opacity: .5; cursor: default; }
.rb-office-shell input:not([type="file"]):not([type="checkbox"]):not([type="radio"]), .rb-office-shell select,
.rb-office-dialog input:not([type="checkbox"]):not([type="radio"]), .rb-office-dialog select, .rb-office-dialog textarea {
  min-height: 32px; min-width: 0; padding: 5px 8px; background: var(--office-surface); border: 1px solid var(--office-border);
}
.rb-office-shell input[type="color"] { padding: 3px; width: 36px; }
.rb-office-shell :focus-visible, .rb-office-dialog :focus-visible {
  outline: 2px solid var(--office-accent); outline-offset: 2px;
}
.rb-office-shell [hidden], .rb-office-dialog [hidden] { display: none !important; }
.rb-office-dialog .rb-dialog-overlay {
  position: fixed; inset: 0; z-index: 1000; display: flex; align-items: center; justify-content: center;
  padding: 20px; background: rgb(0 0 0 / 28%);
}
.rb-office-dialog .rb-dialog-backdrop { display: none; }
.rb-office-dialog .rb-dialog-content {
  width: min(560px, 100%); max-height: calc(100dvh - 40px); overflow: auto;
  padding: 24px; background: var(--office-surface); border: 1px solid var(--office-border);
  border-radius: 10px; box-shadow: 0 12px 40px rgb(0 0 0 / 18%);
}
.rb-office-dialog .rb-dialog-title { margin: 0 0 18px; font-size: 18px; line-height: 1.3; }
.rb-office-dialog .rb-dialog-body { display: grid; gap: 14px; }
.rb-office-dialog .rb-dialog-body label { display: grid; gap: 5px; min-width: 0; }
.rb-office-dialog .rb-dialog-body label:has(input[type="checkbox"]) { display: flex; align-items: center; }
.rb-office-dialog .rb-dialog-body input:not([type="checkbox"]):not([type="radio"]),
.rb-office-dialog .rb-dialog-body select, .rb-office-dialog .rb-dialog-body textarea { width: 100%; }
.rb-office-dialog .rb-dialog-footer { display: flex; justify-content: flex-end; flex-wrap: wrap; gap: 8px; margin-top: 22px; }
.rb-office-dialog .rb-dialog-footer button { border-color: var(--office-border); }
.rb-office-dialog .rb-dialog-footer button:last-child { background: var(--office-accent); color: white; border-color: var(--office-accent); }
@media (max-width: 600px) {
  .rb-office-shell button { min-height: 36px; }
  .rb-office-dialog .rb-dialog-overlay { padding: 12px; }
  .rb-office-dialog .rb-dialog-content { padding: 18px; max-height: calc(100dvh - 24px); }
}
@media (forced-colors: active) {
  .rb-office-shell button, .rb-office-dialog button { border: 1px solid ButtonText; }
  .rb-office-shell button[aria-pressed="true"], .rb-office-shell button[aria-pressed="mixed"] { outline: 2px solid Highlight; }
}
"""

let ensureStyles () = install "chrome" stylesheet

/// Install editor-specific layout rules once, after the shared control rules.
let installStyles key css =
    ensureStyles ()
    install key css
