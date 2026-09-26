# EdSharp Conversion Scripts -- Audit and Notes

This documents the conversion batch files, how EdSharp runs them, what was wrong
with the old ones, and the new `.cmd` replacements. New files in this drop:
`Convert\doc2txt.cmd`, `Convert\MinGW.cmd`, `Convert\pbw.cmd`.

## How EdSharp runs a conversion (from EdSharp.cs)

Import/Export and Compiler commands are defined in `EdSharp.ini` (overridable in
`EdSharp.inix`). When EdSharp runs one it:

1. Substitutes the placeholders via `Util.ExpandCommandLine`. Two forms exist for
   each path: the plain form (`%ProgDir%`, `%Source%`, `%SourceDir%`, `%Target%`)
   is the **8.3 short path** (no spaces), and the `...Long%` form
   (`%ProgDirLong%`, `%SourceLong%`, ...) is the full path (may contain spaces).
   The short forms were the old way of dodging spaces; quoting is the robust fix,
   because 8.3 names are absent on volumes where short-name creation is disabled.
2. For **Import/Export**: it creates a temp **target** file with the right
   extension and expects the command to WRITE that file. EdSharp then re-encodes
   the target to UTF-8-with-BOM and reads it back. So an Import/Export script's
   job is to produce `%Target%`; stdout is not used.
3. For **Compilers**: the `.ini` line ends with `2>&1`, EdSharp captures
   **stdout** and applies the line's error-matching regex. So a Compiler script's
   job is to emit the tool's messages to stdout.
4. The command is launched hidden; if the target file does not appear it is
   retried as `cmd.exe /c <command>` (which is also how `.cmd`/`.bat` files are
   actually executed). So `.cmd` works exactly as `.bat` did here.

## Audit of the old scripts

### doc2txt.bat -> doc2txt.cmd (Import: Word -> text)
Old: `cls` (clears your screen mid-convert), and every path (`%1 %2 %3`)
unquoted, so a space anywhere broke it. Logic was sound: try `WdVert.exe`
(Word COM), and if it produced no file, fall back to `GetText.exe`.
New: dropped `cls`, quoted every path, kept the same try/fallback.

### Mingw.bat -> MinGW.cmd (Compiler: C++)
Old: `cls`; unquoted `%1.cpp` / `%1.o`; hard-coded `C:\MinGW\bin`; and the
compile step `g++ -c %1.cpp` writes its `.o` to the current directory, not
necessarily next to the source, so the `if exist %1.o` test was fragile.
New: dropped `cls`, quoted paths, made the toolchain dir overridable with the
`MINGW_BIN` environment variable (defaulting to `C:\MinGW\bin`), and gave both
`g++` steps explicit `-o` targets so the `.o` and `.exe` land exactly where
expected.

### pbw.bat -> pbw.cmd (Compiler: PowerBASIC)
Old: `cls`; hard-coded `C:\PBWin10`; and crucially `copy %2 %3` where the `.ini`
passes only two arguments, so `%3` was empty and nothing useful was captured --
this is the "I didn't know how to get the output" workaround. PB/Win writes
errors to a `.log` file, not the console.
New: dropped `cls`, quoted paths, made `PBWIN_BIN`/`PBWIN_INC` overridable, and
after compiling it does `type "%~2"` to echo the `.log` to stdout, which is what
EdSharp captures (the `.ini` line appends `2>&1`). Now compiler errors actually
reach the Compiler output.

## EdSharp.ini reference lines to update

Renaming `.bat` to `.cmd` means the `.ini` (or `.inix`) lines that call them
must point to the new names; I also added quoting around `%Target%`. Replace the
old line with the new one in the matching section.

`[Import]` (also appears in `[Export]`):

    old:  doc2txt=%ProgDir%\Convert\doc2txt.bat "%ProgDir%" "%SourceLong%" %Target%
    new:  doc2txt=%ProgDir%\Convert\doc2txt.cmd "%ProgDir%" "%SourceLong%" "%Target%"

`[Compilers]` (keep each line's trailing `~...` regex part exactly as it is; only
the program path changes):

    old:  MinGW C++="%ProgDir%\Convert\MinGW.Bat %SourceDir%\%SourceRoot% 2>&1~...
    new:  MinGW C++="%ProgDir%\Convert\MinGW.cmd "%SourceDir%\%SourceRoot%" 2>&1~...

    old:  PowerBASIC="%ProgDir%\Convert\pbw.bat "%Source%" "%SourceDir%\%SourceRoot%.log" 2>&1~...
    new:  PowerBASIC="%ProgDir%\Convert\pbw.cmd "%Source%" "%SourceDir%\%SourceRoot%.log" 2>&1~...

If you would rather I make these edits automatically on build (with a backup,
the way the Pandoc flag fix works), say so and I will add it.

## Still needed to finish the audit

Three conversion scripts live in folders that were not in anything uploaded, so
I could not audit them precisely (guessing the liblouis table names or the
NFBTrans flags could break working conversions):

- `Convert\liblouis\brl2txt.bat`  (Import `brl2txt`)
- `Convert\NFBTrans\BackTran.bat` (Import `brf2txt`)
- `Convert\Tidy\tidy.bat`         (the commented-out `HTML Tidy` compiler)

Please upload those three files (or the `Convert\liblouis`, `Convert\NFBTrans`,
and `Convert\Tidy` folders) and I will give you audited `.cmd` versions to match.

## 2htm replaces GetText.exe and the Office/HTML text converters

The generic extractor `GetText.exe` (used by `hlp2txt`, `wpd2txt`, and as the
`doc2txt` fallback) hangs on modern Windows. It is now removed. In its place,
EdSharp bundles `Convert\2htm\2htm.exe` (by Jamal Mazrui, MIT) -- a single-file
converter that turns `.docx .doc .rtf .odt .pdf .xlsx .xls .pptx .ppt .csv
.html .htm .md .json .txt` into readable text with its `-p` (plain-text) switch.

A thin wrapper, `Convert\any2txt.cmd`, adapts 2htm to EdSharp's Import contract:
2htm names its output `<sourcebasename>.txt` in an output directory, so the
wrapper points 2htm at the target's directory and then renames the result to the
exact `%Target%` EdSharp expects. Invocation:

    <key>=%ProgDir%\Convert\any2txt.cmd "%ProgDirLong%" "%SourceLong%" "%TargetLong%"

The default `EdSharp.ini` now routes these `[Import]`/`[Export]` text
conversions through it: `doc2txt`, `html2txt`, `ppt2txt`, `pptx2txt`, `xls2txt`,
`xlsx2txt`, `hlp2txt`, `wpd2txt` (Import) and `html2txt` (Export). The Pandoc
conversions are unchanged, `pdf2txt` still uses Xpdf's `pdftotext` (no Office
needed), and `xls2csv` / `xlsx2csv` keep `XlVert` because 2htm cannot emit CSV.
This retires `GetText.exe`, `HTM2TXT\htm2txt.exe`, and the `WdVert`/`PpVert`
text paths; FileDir can drop its own `GetText.exe` dependency the same way.

### Applying this to an existing install

Active converter definitions live in your **data-folder** `EdSharp.ini` (the
per-user copy), which a zip cannot reach -- so rebuilding alone will not change
them on a machine that already ran EdSharp. Two ways to switch over:

- Drop an `EdSharp.inix` next to that `EdSharp.ini` with the override block from
  the shipped `EdSharp.inix` template (uncomment its `[Import]`/`[Export]`
  lines). EdSharp reads `.inix` entries in preference to `.ini`, so this is
  immediate and non-destructive -- delete the file to revert.
- Or edit the `[Import]`/`[Export]` lines in your data-folder `EdSharp.ini`
  directly to match the new default.

New installs get the 2htm-based converters automatically from the default
`EdSharp.ini`.
