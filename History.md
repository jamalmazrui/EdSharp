# EdSharp History

A human-readable record of fixes and enhancements, newest first. Each entry
says what changed and why, so a future reader -- or a future maintainer --
can judge the decision, not just observe it.

## 26 August 2026 -- version 5.0, out of beta

The beta released earlier this year invited people to try EdSharp again after
a decade. What they reported, and what testing found, shaped everything below.
Each entry says what changed and why.

**Documents convert, and keep their shape.** Plain text, Markdown and HTML are
supported as both input and output, alongside Word, PowerPoint, spreadsheets,
rich text and web pages. PDF was rebuilt: it now goes through a free reader
that keeps headings, lists and tables, so a converted PDF can be navigated by
block rather than read as a wall. The route that depended on Microsoft Word's
PDF reflow is gone.

**Spell check and thesaurus without Microsoft Office.** F7 uses Hunspell, with
a dictionary that ships in the box, and walks the document one misspelling at
a time the way Compile walks errors -- word spoken and spelled, position in
the pass announced, suggestions in an editable box. Shift+F7 uses WordNet,
grouped by meaning. The Windows spell checking service is a fallback, and Word
remains an option; neither is needed. Along the way the Windows interfaces
were got right by asking the object which ones it supports and using whichever
it accepts, after three rounds of failures that no error message explained.

**Compilers arrive configured.** Picking a compiler now brings its compile
command, its error pattern, the output to abbreviate, its comment prefix, its
indentation and its interactive shell. Python, JavaScript through Node, C#
with the compiler inside Windows, PowerShell, JAWS script and JScript .NET
ship configured; nine unmaintained or unverifiable entries were retired, and
VBScript joined them when Microsoft began removing it from Windows. Compile
speech starts at the earliest error in the file rather than the first one
printed, with the caret placed by reading the tool's own marker.

**A console for writing snippets.** Control+Shift+G opens a prompt with the
editor window and its text box already in scope. The JScript one is the
Interactive JScript program of 2010, recovered from a damaged archive,
rewritten in Camel Type and given the editor it never had. The C# one compiles
each line with the compiler Compile uses and loads it into the running editor,
which is what lets it touch the live document.

**AI on the machine, not on a server.** F12 asks a question, sending the
document when the wording refers to it and not when it does not; Shift+F12
always sends it. A source file goes to a coding model when one is installed.
Alt+Shift+F7 translates between eighteen languages, using a better model when
one is installed. The old translation command, which called a web service
withdrawn years ago, is rebuilt on this.

**Gathering and batch work.** Append from Clipboard turns a window into a
collector for research from several sources. File Find builds a list of paths
you can edit; Transform Files applies search and replace tasks to every file
in it. Regular expressions run through find, replace and the two commands that
extract matching text into a new document.

**Tutorials by role**, opened with Control+Shift+F1: twelve of them, from
Python developer to web researcher, each naming the settings and keys that
matter for that work rather than describing the program in the abstract.

**Speech that does not repeat itself.** A screen reader already announces
window titles and the focused control; EdSharp now leaves those alone and
speaks only what it alone knows. Messages go to one voice, never two. Key
Describer was finished: every command has a description, checked automatically
against the real bindings, and Alt+F4 leaves the mode and closes the program
rather than describing itself.

**A source audit runs before every build**, checking what a compiler cannot:
duplicate keys, access key collisions, undescribed commands, illegal patterns
in the compiler table, missing conversion scripts, unbalanced braces in the C#
and PowerShell sources, interface layouts that must match what Windows
publishes, public methods exposing private types, and bare type names from
namespaces the file does not import. Each check exists because something once
broke, and each now fails in a second rather than in a tester's hands.

**The installer explains itself.** Optional pieces are checkboxes that say
what they will do and how large they are, grouped by whether they install,
update or reinstall, alphabetical within each group, with the versions probed
behind the progress bar so the page appears at once. Nothing pauses for a
keypress, and one Results box at the very end reports every item by name.

## 23 August 2026 -- audit follow-up, applied on standing authority

Jamal authorized best-judgment implementation of the remaining items from
the 22 August documentation-and-code audit, with alternatives weighed and
the reasoning recorded. Before any of this is committed, the
prepareAuditFixes script marks the last released commit as the branch
snapshotBeforeAuditFixes_20260823, so one command reverts everything:
checking out that branch (or restoring single files from it) brings back
the exact pre-change code.

What changed, and the judgment behind each piece:

- Open Other Format's duplicate picker row is gone (EdSharp.cs,
  ConvertFile2String). When a real to-text converter exists for a format,
  the bare keep-the-extension entry -- which displayed as "txt" but opened
  the file raw -- no longer appears. The alternative considered was
  relabeling the bare row as its own extension; removal won because two
  rows meaning "convert to text" and "show raw source" under one label
  was the confusion, and raw opening already belongs to the ordinary Open
  command. When no converter exists, the bare row stays, and its raw-read
  fallback is exactly what its offer means.

- A blocked converter can no longer freeze EdSharp (EdSharp.cs,
  runShell). The hidden converter process now gets two minutes, and one
  still running after that is ended, letting the existing error dialog
  show the command line. The alternative -- moving conversion to a
  background thread -- was rejected for now as a larger rewrite with new
  failure modes; a bounded wait fixes the harm with eight lines.

- The conversion tables live in EdSharp.inix alone, under the policy:
  Pandoc directly wherever Pandoc reads the format; 2htm or the
  OfficeConvert utilities where it cannot (old Word, PDF through Word's
  PDF Reflow, PowerPoint, Excel); .cmd batch files only, never .bat.
  Entries whose tools were removed from Convert (braille through
  liblouis, HTML Tidy, WinHelp and WordPerfect through GetText) are
  tombstoned -- an empty value hides them everywhere, including stale
  copies in a user's old EdSharp.ini. Excel to Markdown goes through a
  CSV made by OfficeConvert, because Pandoc has no Excel reader. New
  routes: doc, pdf, ppt, and pptx to HTML and Markdown via 2htm, and rtf
  to HTML, Markdown, and text via Pandoc's RTF reader.

- The user guide matches the program again (EdSharp.md): two commands
  that no longer exist are out of the catalog, three malformed catalog
  lines are repaired, ten commands are newly documented (Preview
  Markdown and its browser twin, the three menu-only conversion
  commands, Say Braces, Hotkey Summary, Lookup Term, Translate Language,
  and Tutorial), and the braille back-translation promise is replaced by
  the truth: braille files always open raw, their converters having been
  retired with their tools.

- The build keeps the documentation pairs fresh (BuildEdSharp.ps1). Every
  git-tracked root Markdown file regenerates its .htm through Pandoc when
  the Markdown is newer or the .htm is missing. Tracked-only is the
  judgment: personal notes living in the same folder must never sprout
  .htm files. The alternative -- a separate regeneration script -- was
  rejected because a manual step is a step someone forgets, which is how
  the shipped EdSharp.htm drifted behind EdSharp.md in the first place.

- The JAWS scripts join the repository (prepareAuditFixes runs
  git add -f Scripts). A fresh clone previously built an installer with
  no JAWS scripts at all, because the folder existed only on the
  development machine and the repository's ignore rules -- which treat
  installer wildcards as permission, not desire -- would never add it.
  The forced add was chosen over editing .gitignore because it is one
  action, leaves the protective ignore rules intact for future strays,
  and is permanent: files already tracked stay tracked.

Reverting: git checkout snapshotBeforeAuditFixes_20260823 restores the
whole pre-change state; git checkout snapshotBeforeAuditFixes_20260823 --
EdSharp.cs (or any single path) restores one file.

## Earlier in August 2026

- v5.0.20 restored a working EdSharp after the assembly-name collision
  era: the C# sources compile into one EdSharp.exe, and EdSharp.dll is
  the JScript .NET evaluator built from EdSharp.js, loaded by reflection.
  Menu items may deliberately have no hotkey; a blank key means
  menu-only. The installer gained a single consolidated setup log, a
  script-based JAWS install running as the original user, and a Results
  box reporting observed facts. Releases v5.0.11 through v5.0.19 ship
  executables that cannot start and should not be installed.
