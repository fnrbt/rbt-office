# Writes the intermediate versions of the host files for one-topic commits.
# usage: python3 split-commits.py <stage>   (stage: v1 | v2 | final)  - reads final copies from ./final-host/
import subprocess, sys, os
W = "/home/dave/Dev/redbat-office-viewers"; F = W + "/mega/frontend/"
A = os.path.dirname(os.path.abspath(__file__)) + "/final-host/"
files = {"word": "src/components/office/WordEditor.tsx", "lib": "src/lib/word/editor.ts", "css": "src/components/office/office-editor.css"}
head = {k: subprocess.run(["git", "-C", W, "show", "HEAD0:mega/frontend/" + p], capture_output=True, text=True).stdout for k, p in files.items()} if False else None
def base(k): return subprocess.run(["git", "-C", W, "show", os.environ["BASE"] + ":mega/frontend/" + files[k]], capture_output=True, text=True, check=True).stdout
def final(k): return open(A + os.path.basename(files[k])).read()
def cut(s, start, end):
    i = s.index(start); j = s.index(end, i); return s[:i] + s[j:], s[i:j]
def no_ctrl_enter(fw):
    i = fw.index("    if (!editor || event.target !== wordApi.getInputElement(editor)) return;\n    // Ctrl+Enter starts a new page"); j = fw.index("    if (pickedShape === null) return;\n", i) + len("    if (pickedShape === null) return;\n")
    return fw[:i] + "    if (pickedShape === null || !editor || event.target !== wordApi.getInputElement(editor)) return;\n" + fw[j:]
stage = sys.argv[1]
if stage == "v1":
    w = base("word"); fw = final("word")
    a = '    const contentHeight = Math.ceil(size.Height * scale + 2 * pageMargin);\n'
    i = fw.index('    let contentHeight = Math.ceil('); j = fw.index('    spacerRef.style.width', i)
    assert w.count(a) == 1; w = w.replace(a, fw[i:j])
    a = '    if (change.DocumentChanged) followCaret = "edit";'
    i = fw.index('    // The engine moved its own viewport to stay on the text it shows'); j = fw.index('    // (A shape that is moved by its frame leaves the view where it is', i)
    assert w.count(a) == 1; w = w.replace(a, fw[i:j] + a)
    lib = base("lib"); fl = final("lib")
    i = fl.index('  /**\n   * A number only in a notice that asks the host to scroll'); j = fl.index('}\n', i)
    a = '  FontsPending?: boolean;\n}\n\n/** A formatting value'
    assert lib.count(a) == 1; lib = lib.replace(a, '  FontsPending?: boolean;\n' + fl[i:j] + '}\n\n/** A formatting value')
    out = {"word": w, "lib": lib, "css": base("css")}
elif stage == "v2":
    fw = final("word")
    fw, _ = cut(fw, '  /**\n   * A double click on a text box of a header or footer edits the box\'s text.', '  // ---- placeholder-never-found') if False else (fw, None)
    i = fw.index("  /**\n   * A double click on a text box of a header or footer edits the box's text."); j = fw.index("  };\n", fw.index("const onCanvasDoubleClick", i)) + len("  };\n")
    fw = fw[:i].rstrip("\n") + "\n" + fw[j:]
    a = '      canvasRef.addEventListener("dblclick", onCanvasDoubleClick);\n'
    assert fw.count(a) == 1; fw = fw.replace(a, "")
    fw = no_ctrl_enter(fw)
    out = {"word": fw, "lib": final("lib"), "css": final("css")}
elif stage == "v3":
    out = {"word": no_ctrl_enter(final("word")), "lib": final("lib"), "css": final("css")}
else:
    out = {k: final(k) for k in files}
for k, p in files.items(): open(F + p, "w").write(out[k])
print("wrote", stage)
