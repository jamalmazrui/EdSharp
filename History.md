# EdSharp History

A human-readable record of fixes and enhancements, newest first. Each entry
says what changed and why, so a future reader -- or a future maintainer --
can judge the decision, not just observe it.

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
