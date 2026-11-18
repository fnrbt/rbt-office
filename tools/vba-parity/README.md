# VBA parity with Microsoft Excel

The same VBA runs in Excel on the Windows VM (`winvm`, see `docs/excel-parity.md`) and in our engine
(Rbt.Vb with `src/Spreadsheet/MacroHost.fs`); what each leaves in the cells is compared. Excel is right.

```
dotnet build src/Spreadsheet
dotnet fsi tools/vba-parity/parity.fsx build tests/vba-parity/cases DOCS      # DOCS/<case>.xlsm and .bas
tools/vba-parity/run-on-vm.sh DOCS excel.tsv                                  # Excel's results
dotnet fsi tools/vba-parity/parity.fsx ours DOCS ours.tsv                     # ours
dotnet fsi tools/vba-parity/parity.fsx compare excel.tsv ours.tsv             # VBA_PARITY_ALL=1: every difference
```

A case (`tests/vba-parity/cases/*.bas`) is a standard module with `Sub Main()`. A probe line
`P "label", expression` writes on the next row the label (A), the value as a cell takes it (B),
`TypeName` (C) and `CStr` as text (D); an error in the expression writes `#ERR <number>` in B instead and
the case goes on. Anything else a case writes into cells is compared too (values, text, number format,
bold, fill, font colour).

**The VM:** the run uses Excel's VBA project object model to put each case into a blank workbook, which
needs "Trust access to the VBA project object model": `reg add
HKCU\Software\Microsoft\Office\16.0\Excel\Security /v AccessVBOM /t REG_DWORD /d 1 /f` for the VM's
user (set on `winvm` 2026-11-18; the same for Word). Excel runs visible so that a stuck dialog can be seen
(`tools/vba-parity/screen.ps1`, run like the other scripts with `WORD_VM_ENGINE=powershell`, captures the
desktop and lists the windows with their texts).

**What stops a run:** anything Excel answers with a dialog - a compile error (a declaration after the
first procedure; syntax Excel rejects) or a macro that never ends - holds the batch until the task's
timeout (`VBA_PARITY_TIMEOUT`, seconds). `run-on-vm.sh` copies back `OUT.tsv.progress.txt`, which names
the case it stopped in. Mind VBA's own rules when writing cases: `Resume` re-executes the statement that
raised, so a retry must change what that statement sees.

Engine differences found this way are kept as minimal repros in `tests/vba-parity/engine-gaps/`.

## Compile errors (early binding)

VBA checks members of declared types when it compiles. Our engine does the same against Excel's and Word's
type libraries (`src/Spreadsheet/ExcelTypeLibrary.fs`, `src/Word/WordTypeLibrary.fs`, made from the VM's
Office by `tools/vba-parity/typelib.sh`; the TSV is `typelib.ps1`'s). The cases are
`tests/vba-parity/compile-cases/*.bas` (and `<case>.<Class>.cls` class modules):

```
tools/vba-parity/compile-on-vm.sh tests/vba-parity/compile-cases excel-compile.tsv     # Excel's compile errors
dotnet fsi tools/vba-parity/compile-parity.fsx tests/vba-parity/compile-cases excel-compile.tsv
```

`compile-cases.ps1` puts each case into a blank workbook and runs VBE's "Compile VBAProject"; a compile
error is a dialog, which a watcher process reads and closes. What Excel showed: a member an *extensible*
interface does not list (Excel's Range, Font, Application) compiles and is found at run time (438); a
non-extensible one (Excel's Worksheet, nearly all of Word's, a VBA class) gives "Method or data member not
found". Known members are checked either way (arguments, read-only properties). The task runner takes an
output file that is already there as done: each run uses a fresh folder.
