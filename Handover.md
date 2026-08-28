# EdSharp Handover

Written 26 August 2026, at the end of the chat that took EdSharp 5.0 out of
beta. It exists so the next conversation starts where this one ended: it says
what was decided, what was tried and abandoned, and where the traps are. The
last section is addressed to FileDir development rather than to EdSharp.

## How to Work on EdSharp

**Run the audit before believing anything.** auditEdSharp.cmd checks what a
compiler cannot, and the build runs it before compiling. Every check exists
because something broke: duplicate keys, access key collisions, undescribed
commands, illegal patterns in the compiler table, missing conversion scripts,
unbalanced braces in the C# and PowerShell sources, spell check interfaces
that must match Windows' published layouts slot for slot, public methods
exposing private types, bare type names from unimported namespaces, installer
sources that do not exist, and a documentation set that must be complete.
Adding a check when something breaks is the habit that made this week
productive; keep it.

**When a build dies without writing a log**, run auditEdSharp.cmd by hand. A
PowerShell script that will not parse never reaches its own logging.

**Verify edits by reading the file back.** Twice this week an edit script
computed a change, hit an assertion, and exited without writing — leaving the
file untouched while the transcript said otherwise. Grep for the new text
afterwards.

**Two lists decide what the repository carries, and nothing else does.**
EdSharp_Setup.iss says what EdSharp installs; repoPolicy.py names the files
that build, audit and release it. A file belongs because it is named in one
of those two places. Everything else -- this brief, saved web pages, saved
mailing list messages, old drafts, logs -- stays on disk and out of the
repository. tidyRepo.cmd untracks whatever does not belong, without deleting
anything, and auditEdSharp checks the same rule on every build.

That rule replaces a pattern that cost weeks. tidyRepo's test used to admit
any file at the top of the folder ending in .md or .htm, meaning to spare the
documentation set. Every saved page and every old draft ends in .htm or .md
and sits at the top of the folder, so 38 of them were declared needed and the
survey reported the repository clean, run after run. Two lessons worth
keeping: a rule that admits files by the look of their name will admit
things nobody named, and .gitignore cannot repair it afterwards, because
.gitignore has no effect at all on a file that is already tracked.

**Beware commented examples when anchoring an edit.** EdSharp.inix contains a
"Personal overrides (examples)" block with commented section headers. An edit
anchored on `[Compiler C#]` matches the example first and damages the wrong
place. Anchor with a line-start match.

## Decisions That Should Not Be Relitigated

**Spell check runs on Hunspell, not the Windows API.** The Windows Spell
Checking API works, but on the developer's machine the object it returns
supports ISpellChecker2 and refuses ISpellChecker, which cost four rounds of
wrong guesses about vtable slots and interface identifiers. Hunspell is a
managed library with plain dictionary files, it is the default, and the COM
path remains as a fallback that now tries both interfaces and reports what the
object actually supports.

**Everything that touches Hunspell goes through reflection.** Naming its types
makes the compiler read member signatures declared with Span, which .NET
Framework 4.8 lacks, and the build fails demanding System.Memory. This is the
general trap with modern packages on this framework: PdfPig, Roslyn scripting
and Hunspell all hit it. Reflection or a different approach, not another
package.

**PDF conversion uses PyMuPDF4LLM through Python, not a .NET library.** The
requirement is rich output — headings, lists, tables — not text extraction.
PdfPig was tried and rejected for producing plain text; Word's PDF reflow was
retired because it required Office.

**The C# console compiles each line with csc and loads it in process.**
Roslyn's scripting package would be elegant and drags a dozen span-typed
assemblies. dotnet-script needs the SDK. Compiling costs about a second a
line, which is acceptable for a console used a line at a time, and in-process
loading is what lets the code touch the live editor — the entire point.

**Python means the python.org build.** Windows ships an app execution alias at
WindowsApps that answers `where python` and then advertises the Microsoft
Store. Every probe rejects any path containing WindowsApps and verifies the
candidate answers `--version`. A machine may also carry several Pythons;
installPdfTools records which interpreter it used so the summary asks that
one.

**Office is reached only by named fallbacks.** SpellCheckWord, ThesaurusWord,
MailWord, WordFile2String and WordSource2TargetFormat — the last two for
legacy .doc, .ppt and .xls, which the developer explicitly agreed may
subordinate to Office. The audit enforces that no other code path reaches it.

**Descriptions live in a table inside EdSharp.cs.** They used to live only in
Hotkeys.ini, which the installer ships with a flag that never overwrites an
existing copy — so any machine that had EdSharp before never received new
descriptions, and Key Describer answered "no description available" for every
new command. The file is still read first, as a user override.

**Speech supplements the screen reader; it never repeats it.** Window titles
and the name, role, state and value of the focused control are announced by
JAWS, NVDA and Narrator on their own. EdSharp speaks only what it alone knows.
This principle is also in the HomerApp documentation from 2010.

## Conventions

- **Camel Type** throughout: Hungarian prefixes, lower camel case, functions
  not subprocedures, one-line simple conditions, declarations grouped at the
  top.
- **Every command names its key at least once in context** when written about:
  "the Spell Check command, F7". In an announcement, drop keys unless the key
  itself carries meaning.
- **Lists rather than tables**, since a list reads better aloud. Tables only
  in reports aimed at sighted readers.
- **No bare URLs**; use link text with the address behind it.
- **Match the noun to the count**: "1 match", not "1 matches"; zero is a real
  answer, not an error.
- **No console pauses in the installer.** One Results box at the very end
  reports every checkbox by name.
- **Logs go in the EdSharp logs folder** under local application data, written
  line by line as they happen, never buffered.
- **Ship replacement files rather than patch scripts** when a file can express
  the change.

## The Homer Guidelines for Forms

These are the defaults for every dialog. A form may depart from them when it
is designed to, but not by accident.

- **Every control has its own trigger letter.** A control means a tab stop: a
  label and its text box count as one control, the label giving the box its
  name; a list is a control; each button is a control. The letter is marked
  with an ampersand in the control's name, so Alt and that letter reach it
  from anywhere in the form.
- **The letter must begin a word.** The first word's initial by preference,
  the second word's initial when that clashes. A letter from the middle of a
  word is almost never acceptable; the rare exception is a strong mnemonic,
  such as X for Export.
- **When two controls want one letter, rename a control.** That is the
  sanctioned answer, not a mid-word letter. In the spell dialog the
  suggestion box became Correction because Replace held R, and the context
  line became "In context" because Skip held S.
- **OK and Cancel carry no ampersand.** Their keys are Control+Enter and
  Escape. This also frees O and C for other controls.
- **Help is added by the toolkit**, sits after OK and Cancel, and holds H.
  Because buttons are added right to left, Help is created first and keeps
  its letter when a run-time label would have collided.
- Buttons are numbered in the order given, so one tab from the last field
  reaches the first button.
- Every field gets a label and a focus tip; F1 describes them all.
- One dialog rather than a chain of prompts.

The audit checks all of this: every control offers a letter, every letter
begins a word, and no two controls in a dialog share one.

## Where Things Stand

Version 5.0 is ready to release. The documentation set is complete: ReadMe,
EdSharp (user guide), Tutorials (twelve roles), Hotkeys (generated from the
program), FAQ, History, Development, Announce. The build regenerates each
.htm; commit new Markdown files before tagging or the HTML will not exist.

Known open items, none blocking: nine settings are not yet documented in the
guide (the audit lists them as a note); the JAWS scripts scrub cannot rewrite
key maps while JAWS holds them open, so the build scrubs the repository copy
instead; and the spell check COM path has never been seen to work on the
developer's machine, though it is now written to try both interfaces.

## For FileDir Development

FileDir is a file and directory manager in C#, version 5.0 beta, whose public
repository currently holds only documentation. Its guide describes an
installer, JAWS scripts, and a hotkey-driven interface much like EdSharp's.
Everything below is transferable, and several pieces should be shared rather
than copied.

### Shared Homer classes to take from EdSharp

- **Say.cs** — the speech dispatcher: JAWS through its automation interface,
  then NVDA through its controller library, then a native UIA notification
  that Narrator reads, stopping at the first that answers. Take it whole.
  With it, take the policy: never speak a window title or a focus change,
  because the screen reader already does, and never send one message by two
  mechanisms.
- **Lbc.cs** — the dialog toolkit. Fields added in order, so reading order,
  tab order and visual order are identical; automatic Help button; F1
  describing every field; Control+Enter and Escape; and the trigger-letter
  rules above enforced by the toolkit itself. FileDir's dialogs should be
  rebuilt on it rather than hand-laid, and its audit should carry the same
  check.
- **Inix.cs** — the settings reader, with the three-file arrangement: shipped
  defaults in the program folder, personal overrides in the data folder,
  changed settings in the ini. Upgrades replace only the first.
- **KeyMap.cs** — key to command mapping, and the in-code description table
  pattern that feeds Key Describer, the Alternate Menu and the generated
  hotkey reference. Do not put descriptions only in a shipped file.

### The build script

Copy BuildEdSharp.ps1's shape rather than its details:

1. Derive the next version from the release tags on the remote and write it
   into the installer script, so the tag and the built version cannot
   disagree.
2. Fetch pinned packages, verifying the version on disk. Note that some
   packages stamp their assembly with a rounder number than the package
   version — Markdig 1.3.2 stamps 1.3.0.0 — so record what was fetched in a
   small note file, or the build downloads the same file for ever.
3. **Run a source audit and stop on failure.** This is the single highest
   value piece. FileDir's version should check its own invariants: duplicate
   keys, described commands, balanced braces, installer sources present.
4. Compile, then regenerate the .htm for every tracked Markdown file.
5. Log everything to one file, appended line by line, and say plainly which
   file to send when a build fails.

### The installer

The EdSharp installer's arrangement is worth copying exactly:

- **No Tasks or Components page.** Optional pieces are finish-page checkboxes,
  each with a label computed at run time saying whether it will install,
  update or reinstall, with the version and the size.
- **Group by action** — install, then update, then reinstall — alphabetical
  within each group, reinstall entries never ticked.
- **Probe behind the progress bar.** Version queries take a second or two
  each; run them at the post-install step where a status line can say
  "Checking Python" rather than during the finish page, which otherwise takes
  a silent minute to appear.
- **Nothing pauses.** One Results box at the very end, started from
  DeinitializeSetup so it runs after every checkbox, reporting each item by
  name with its version or the exact command to add it later.
- Run helper scripts with runascurrentuser: setup is elevated, and winget
  installs into the profile of whoever is signed in.
- Two closing unchecked boxes: launch the program, open the guide, both
  runasoriginaluser so settings land in the right profile.

### The local AI

If FileDir uses the same Ollama models as EdSharp and HomerScribe, they share
one installation and one set of models — several gigabytes downloaded once,
not per program. Take these decisions with it:

- Offer Ollama as one checkbox with a named model, llama3.2, and the larger
  models as separate checkboxes naming them: qwen2.5:7b for translation,
  qwen2.5-coder:7b for code. Never say "a stronger model" without naming it.
- Choose the model in code by what is installed, asking `ollama list` once per
  session, rather than making the user configure anything.
- Speak progress during a long wait — a count every fifteen seconds — because
  silence is indistinguishable from a hang.
- Run every probe with a real no-window process. PowerShell's Start-Process
  silently ignores its hide-window request once output is redirected, which
  opened a console window during setup and alarmed the developer. Better
  still, ask Ollama over its web interface and start no process at all.
- For a file manager, the obvious uses are summarizing a selected document,
  describing what a folder contains, and answering questions about a file's
  text — all of which are single-document work that a small model does well.
  Cross-file reasoning is not.

### Traps FileDir will otherwise rediscover

- Modern packages on .NET Framework 4.8 that use Span will fail the build in a
  way the error message does not explain. Check a package's target framework
  before pinning it, and reach it by reflection if in doubt.
- PowerShell joins argument lists with spaces and quotes nothing, so
  `python -c "import x"` arrives as three words.
- The Windows command interpreter's quote stripping after /c defeated two
  attempts at running a quoted program with a quoted argument. Do not fight
  it: Inno's Exec and .NET's process object both take the program and its
  arguments separately, and then no quoting rule applies. This was the cause
  of the document tools appearing perpetually uninstalled.
- **Never run the ollama command to ask a question.** It starts the server in
  a console of its own when the server is not already running, and that
  window sits on screen looking like a fault. Ollama answers over a local web
  interface at port 11434 -- /api/tags lists the models -- which opens
  nothing. Only a download needs the command, and it should be launched
  hidden.
- Windows' Python alias under WindowsApps is not Python.
- Decide what the repository carries by naming files, never by matching
  names. A pattern such as "any markdown at the top of the folder" reads as
  a description of the documentation and behaves as an open door. And
  .gitignore is not a repair: it governs untracked files only, so a name
  added to it after a file is committed changes nothing until the file is
  untracked. FileDir should take repoPolicy.py, the tracked-files audit
  check, and this rule together.
- A file the installer ships with the never-overwrite flag will never reach a
  machine that already has the program.

This brief is a development aid, not part of EdSharp. It lives in C:\EdSharp,
it is named in .gitignore, and it is named in neither EdSharp_Setup.iss nor
repoPolicy.py, which is exactly how a file is kept out of the public
repository.

Jamal Mazrui's standing preferences — Camel Type, ninth grade reading level in
user documentation, lists over tables, no bare URLs, detailed logs beside the
script, replacement files over patch scripts — apply to FileDir exactly as
they do here.
