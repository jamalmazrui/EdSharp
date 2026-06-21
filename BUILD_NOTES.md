# EdSharp 5.0 baseline starter -- build notes (revision 52)

Revision 52: refine the rev 51 Append from Clipboard change per follow-up.

Two adjustments: (1) drop the spoken "Append from Clipboard is on" reminder;
(2) suppress only copies made in the SAME window that is collecting -- not copies
made in other EdSharp windows.  The rev 51 test
(GetWindowThreadProcessId(GetForegroundWindow(),0) == GetCurrentThreadId()) was
true for any EdSharp window, so a copy in a different document was wrongly
suppressed.  Replaced it with a single check, "if (this.ContainsFocus)
sClipboard = \"\";".  ContainsFocus is true only when the input focus is in this
collecting document, which excludes both other EdSharp windows (a different child
has focus) and other applications (EdSharp is not foreground).  So a copy made in
this document is skipped, while a copy made in another EdSharp document or any
other application is collected as before.  No status message is shown.  Brace
balance +11.

Docs updated: the Append From Clipboard section now says clips from another
application or from a different EdSharp document are collected, and only text
copied within the collecting document itself is skipped; no reminder is
mentioned.  Regenerated EdSharp.htm.

---

# EdSharp 5.0 baseline starter -- build notes (revision 51)

Revision 51: Append from Clipboard (Alt+7) no longer appends your own copies.

User request (Jim Homme): with Append from Clipboard on, copying within EdSharp
appended the copy back into the current document; he wanted that suppressed,
or at least a reminder that the mode is on (he keeps forgetting).

In the WM_DRAWCLIPBOARD handler (MdiChild.WndProc), when AppendFromClipboard==1
the clip is now appended only if the copy did NOT originate in EdSharp.  Detection
reuses the existing ForceWindow idiom:
GetWindowThreadProcessId(GetForegroundWindow(), 0) == GetCurrentThreadId() is true
when EdSharp's own (single-UI-thread) window is in the foreground, i.e. the copy
came from EdSharp.  In that case the append is skipped (sClipboard cleared) and
App.Frame.AddMessage("Append from Clipboard is on") reminds the user -- covering
both of Jim's suggestions at once.  Clips copied in other applications still
append and beep as before.  Brace balance +11.

Docs: the Append From Clipboard section now states that only clips from other
applications are collected and that copying inside EdSharp announces the mode is
on instead of appending; regenerated EdSharp.htm.

---

# EdSharp 5.0 baseline starter -- build notes (revision 50)

Revision 50 (JAWS keymap only): finish the sentence-key fix that rev 49 missed.

Rev 49 set the FIRST of two Alt+UpArrow/Alt+DownArrow bindings to SilentKey, on
the assumption that the first duplicate wins.  It did not: the key was still
spoken, which proves JAWS resolves a duplicated key in a .jkm section to the
LAST occurrence.  The later binding was Alt+DownArrow=OpenListBox /
Alt+UpArrow=CloseListBox (JAWS built-ins), so those were active -- they passed
the key through (sentence navigation worked) but announced the key name.

Fix: removed the OpenListBox/CloseListBox duplicate so a single
Alt+UpArrow=SilentKey / Alt+DownArrow=SilentKey binding remains for each key.
One binding means no precedence ambiguity, so the result no longer depends on
first-vs-last. Sentence navigation stays silent, and combo boxes still open
(see below).

Findings behind the fix:
- Sentence-navigation key convention: there is no OS- or screen-reader-standard
  keystroke for sentence-by-sentence caret movement.  Alt+Down/UpArrow for
  sentence is EdSharp's own choice and collides with the Windows meaning of
  Alt+DownArrow (open a combo box).  By contrast Control+Up/DownArrow for
  paragraph IS standard (Word, and the RichEdit control moves by paragraph
  natively), which is why those keys were easier.
- JAWS does NOT do sentence navigation itself in a RichTextBox.  It has reading
  commands and auto-reads after caret moves, but no sentence-unit caret nav on
  Alt+Down/Up.  Its DEFAULT for Alt+DownArrow is OpenListBox (combo open),
  because Alt+DownArrow is the Windows combo key -- that default is exactly what
  was speaking the key.  EdSharp performs the sentence navigation itself.
- SilentKey works for combo boxes as well as sentences.  Alt+DownArrow is the
  native Windows keystroke that drops down a combo box, so a silent pass-through
  opens the combo (Windows handles it) and JAWS reads the list, with no key
  announcement -- in the editor the same pass-through gives EdSharp the key for
  sentence navigation.  EdSharp's dialogs use standard combo boxes, which also
  respond to F4.

Not changed (still offered): read-by-word (Control+Left/RightArrow) and
read-by-chunk (Alt+Left/RightArrow) remain on TypeCurrentScriptKey and would
announce the key the same way; they can get the SilentKey treatment too.

Deploy unchanged: edsharp.jkm is read live -- drop into each JAWS version's
per-user Settings\<lang> and reload, or rerun --install-jaws-settings.
EdSharp.cs unchanged.

---

# EdSharp 5.0 baseline starter -- build notes (revision 49)

Revision 49 (JAWS keymap only -- no recompile): extend the SilentKey fix to
sentence navigation.

Reading by sentence uses Alt+UpArrow (prior sentence) and Alt+DownArrow (next
sentence).  In edsharp.jkm these keys were bound to TypeCurrentScriptKey (the
first binding in [Common Keys]); a later duplicate binds the same keys to
Open/CloseListBox, but JAWS resolves a duplicated key to the first occurrence, so
TypeCurrentScriptKey was active -- which is why the key name was announced (only
EdSharp's custom TypeCurrentScriptKey speaks it, via SayCurrentScriptKeyLabel()
when UIIsEditorWindow() is false).  Changed both to SilentKey, matching the
Control+Up/DownArrow paragraph fix from rev 48: silent pass-through, sentence
still moved to and read, no key announcement.  Left the later
Open/CloseListBox duplicates in place (they remain overridden; combo boxes in
dialogs still open via the native Alt+Down/Up the pass-through delivers).

Control+UpArrow was already SilentKey (rev 48), so no change was needed there.

Not changed (offered): the read-by-word keys (Control+Left/RightArrow) and
read-by-chunk keys (Alt+Left/RightArrow) are bound to TypeCurrentScriptKey too
and would announce the key name the same way; they can get the identical
SilentKey treatment if wanted.

Deploy is unchanged from rev 48: edsharp.jkm is a keymap (read live), so no
recompile -- drop it into each JAWS version's per-user settings
(%APPDATA%\Freedom Scientific\JAWS\<ver>\Settings\<lang>) and reload, or rerun
EdSharp.exe --install-jaws-settings.  EdSharp.cs unchanged.

---

# EdSharp 5.0 baseline starter -- build notes (revision 48)

Revision 48 (JAWS keymap only -- no recompile): stop JAWS announcing the key
name on paragraph navigation.

Control+UpArrow / Control+DownArrow were bound to TypeCurrentScriptKey. EdSharp
redefines that script (EdSharp.JSS ~577); its body calls
SayCurrentScriptKeyLabel() -- speaking the key name -- whenever UIIsEditorWindow()
is false, then passes the key through. In the real WinForms RICHEDIT edit window
UIIsEditorWindow() returns false (the window-class test is too strict, the same
root cause noted in rev 40), so the key passed through and the paragraph was read
(correct), but "Control+DownArrow" was spoken first (unwanted).

Fix: bind both keys to SilentKey instead. SilentKey (EdSharp.JSS ~1038) is just
TypeCurrentScriptKey() with no SayCurrentScriptKeyLabel(), so it is a silent
pass-through -- the paragraph is still moved to and read, with no key
announcement. This is a keymap (.jkm) change only; JAWS reads edsharp.jkm live,
so no .jss/.jsb recompile is required. (The deeper alternative -- correcting
UIIsEditorWindow() in homer.jss and recompiling EdSharp.jsb -- would also fix it,
and globally for every TypeCurrentScriptKey-bound key, but needs JAWS to compile;
the SilentKey keymap change is the distributable no-recompile fix.)

Deploy: ships in Scripts\edsharp.jkm; the installer copies Scripts\* to
{app}\Scripts and the Finish-page "EdSharp.exe --install-jaws-settings" step copies
the family into each JAWS version's per-user settings
(%APPDATA%\Freedom Scientific\JAWS\<ver>\Settings\<lang>). For an existing
install, dropping the updated edsharp.jkm into that Settings\<lang> folder (e.g.
enu) is enough -- no recompile, just reload (restart JAWS or its scripts).

EdSharp.cs unchanged this revision.

---

# EdSharp 5.0 baseline starter -- build notes (revision 47)

Revision 47: document the Transform Files job format, and fix a bug found while
documenting it (this is very likely the "Alt+= tended not to work" beta report).

Bug: in TransFormFiles(), a task wrote the file only when
"bHasReplace && iCount > 0", where bHasReplace required a non-empty Replace
value.  So a deletion task (Find with an empty Replace -- e.g. the common "trim
trailing whitespace") changed the text in memory but was never saved.  A job
appeared to work only when it also contained a non-empty-replacement task that
happened to match, making the feature behave unpredictably.  Reworked the apply
branch: extract tasks now use Regex.Matches (collect to clipboard, never touch
the file or later tasks' input), and replace/delete tasks use Regex.Replace and
mark the file changed on any match (iCount > 0), so deletions persist.  Removed
the now-unused bHasReplace.  Brace balance unchanged (+11).

Docs: the EdSharp.md Transform Files section already described the new
Regexer-style .inix job format in prose, but the embedded sample still showed
the old positional .job format.  Tightened the key descriptions (an empty
Replace deletes matches; Extract collects to the clipboard without changing the
file; Divider defaults to form feed + newline) and replaced the stale sample
with a real .inix sample (deletion, replacement, extraction) that matches the
shipped Transform_Example.inix.  Regenerated EdSharp.htm.

Format summary (for reference): a job is an .inix file (.ini or .inix
extension), one [Section] per task; the section name is a description shown in
Test/Verbose output.  Keys: Find (regex, required), Replace (replacement text;
\n \t etc. and $1/$# honored; empty deletes), Extract (true to collect matches
to the clipboard instead of editing), Divider (separator between collected
matches), Options (comma-separated .NET RegexOptions names).  Values may span
multiple lines.  The current document supplies the source file list, one path
per line.  Modes: Test (count only), Run (apply), Verbose (apply with per-task
detail).

---

# EdSharp 5.0 baseline starter -- build notes (revision 46)

Revision 46: make the installer (EdSharp_Setup.exe) run on ARM, completing the
ARM support begun in rev 43 (AnyCPU EdSharp.exe).

EdSharp_Setup.iss had ArchitecturesAllowed=x64, which makes Inno Setup refuse to
run on a Windows-on-ARM machine ("can only be run on x64"). Changed both
ArchitecturesAllowed and ArchitecturesInstallIn64BitMode from x64 to
x64compatible, which (Inno Setup 6.3+) matches x64 AND ARM64, so the installer
runs and installs in 64-bit mode on both, and the AnyCPU EdSharp.exe then runs
natively. This matches the already-modernized DbDo_setup.iss, so the installed
Inno Setup is new enough for the identifier. Added MinVersion=10.0 (matches the
.NET Framework 4.8 / Windows 10+ requirement and DbDo).

The ngen [Run] steps are unchanged and remain safe on ARM64: NgenExe resolves to
{win}\Microsoft.NET\Framework64\v4.0.30319\ngen.exe, which on an ARM64 system
is the ARM64 framework's ngen, and the HasNgen (FileExists) Check skips it
gracefully if it is absent, so installation never fails on that account. Updated
the script comments accordingly (header, ngen notes).

No other architecture-gated logic exists in the script (the only Check is the
file-existence HasNgen). EdSharp.cs is unchanged this revision.

---

# EdSharp 5.0 baseline starter -- build notes (revision 45)

Revision 45: open raw by default; auto-convert only binary/document formats.

Per request, no text format is auto-converted when a file is opened from outside
the editor (Explorer, "Open with", command line, Recent Files). Previously
GetViewLevel() returned a fallback of 1 (convert) for any extension not listed
in the ViewLevels option, so a text format that happened to have an Import
converter and was not listed (for example .rst) would back-translate/convert on
those paths. .json/.csv/.inix already opened raw only because they have no
converter -- a fragile guarantee that a future converter could break.

GetViewLevel() now decides in this order: (1) an explicit ViewLevels entry wins
(user can force any extension, e.g. "docx:0" or "rst:1"); (2) a built-in set of
binary/document formats converts -- doc docx xls xlsx ppt pptx pdf epub epub3
hlp wpd rtf; (3) everything else opens raw (fallback changed from 1 to 0). This
guarantees text/markup/data/source/unknown formats open raw now and as new
converters are added, while binary formats whose raw bytes are unreadable still
convert. The binary set lives in code (const sBinaryFormats), so existing
installs get correct binary conversion even though their saved ViewLevels
predates this change -- their old entries simply act as overrides and stay
consistent (htm:0 = raw, pdf:1 = convert, etc.); the only behavior change they
see is that a text-with-converter format like .rst now opens raw, which is the
intent. No existing-install action is required.

Control+O still always opens raw; Control+Shift+O still always converts. The
default ViewLevels in EdSharp.ini is now empty (the policy lives in code; the
option is purely for per-extension overrides). Docs (ViewLevels section) and the
EdSharp.inix override note were rewritten to the new model; regenerated
EdSharp.htm.

This supersedes the rev 44 braille-specific default (brl:0 brf:0), which is now
subsumed by the general "text opens raw" rule.

---

# EdSharp 5.0 baseline starter -- build notes (revision 44)

Revision 44: consistency fix for braille open behavior.

The ordinary Open command (Control+O) never converts -- it opens raw (only .rtf
prompts). The Explorer / "Open with" / command-line / Recent Files paths instead
use GetViewLevel(), which converts when an extension's ViewLevels entry is >= 1
(or is unlisted, since the fallback default is 1) AND an [Import] converter
exists. .brl/.brf were unlisted, so they back-translated on those paths only --
an asymmetry with Control+O. Added "brl:0 brf:0" to the default ViewLevels in
EdSharp.ini so braille opens raw on every path by default, matching Control+O.
Open Other Format (Control+Shift+O) still back-translates on demand; set
"brl:1 brf:1" to restore auto-back-translation on open. Docs and the EdSharp.inix
override note were updated (the override now serves existing installs, whose
data-folder ViewLevels predates this default).

Design note (not implemented): making EVERY format open raw on those paths
(removing auto-conversion) would regress binary/container formats -- PDF, .docx,
.doc, .xls/.xlsx, .ppt/.pptx, .odt, .epub -- which would open as raw binary
(unreadable, and especially noisy for a screen reader) instead of auto-extracted
text. Those formats are valuable to auto-convert; text-like formats (htm, html,
md, xml, tex, source code, and now braille) are the ones that should open raw.
The per-extension ViewLevels mechanism already expresses exactly this split, so
the targeted brl/brf entries are the right fix rather than a global change.

---

# EdSharp 5.0 baseline starter -- build notes (revision 43)

Revision 43 addresses beta feedback (Dean Martineau via BITS).

1. ARM compatibility. BuildEdSharp.cmd now compiles EdSharp.exe AnyCPU instead
   of x64. AnyCPU runs as native 64-bit on x64 AND as native ARM64 on Windows on
   ARM (an x64 build does not run natively there, and on Win10-on-ARM not at
   all). The only architecture-specific dependency is the native
   nvdaControllerClient.dll (x64); Say.cs already guards that P/Invoke
   (DllNotFoundException + catch-all), so on ARM64 it degrades gracefully to the
   other speech paths instead of crashing. No manifest change was needed (it
   carries no processorArchitecture).

2. Append From Clipboard (Alt+7) crash on Windows 11. EdSharp watches the
   clipboard as a viewer and read it with a bare Clipboard.GetText() inside
   WndProc; on Windows 11 the clipboard is frequently locked (cloud clipboard /
   history) and GetText throws ExternalException, which is fatal in a window
   procedure. Added Util.GetClipboardText() and Util.SetClipboardText() that
   retry on contention (10 x 40 ms) and return ""/clear instead of throwing, and
   routed all active clipboard reads (7) and writes (9) through them. This fixes
   the reported crash and hardens every copy/paste/append path.

3. Opening .brl / .brf raw. EdSharp already supports this through the ViewLevels
   option (GetViewLevel): a level of 0 opens a type without conversion. The
   default ViewLevels did not list brl/brf, so they fell through to the default
   level 1 and back-translated. Rather than change the default (some braille
   users want back-translation on open), the user guide now documents adding
   "brl:0 brf:0" to ViewLevels, and EdSharp.inix carries a ready commented
   override for existing installs. Open Other Format (Control+Shift+O) still
   forces back-translation on demand.

Not changed (need more information / out of scope as confident fixes):
- Alt+= is the Transform Files command; "tended not to work" is too vague to fix
  without a reproducible case. Logged for follow-up.
- A macro recorder is a new feature, out of scope for a feedback-fix pass.
- Defaulting braille files to raw vs back-translated is a product decision left
  to the author; a future toggle/UI could make it one keystroke.

Docs: requirements and rebuild notes now state AnyCPU / x64 + ARM64; regenerated
EdSharp.htm.

---

# EdSharp 5.0 baseline starter -- build notes (revision 42)

Revision 42: installer + documentation housekeeping.

- EdSharp_Setup.iss: EdSharp.exe.config was already shipped (ignoreversion);
  regrouped it directly beneath EdSharp.exe with a comment, since it carries the
  startup tuning and must stay in sync with the executable. ignoreversion means
  every install/update refreshes it.
- EdSharp.md: updated the title block from "Version 4.0 / May 29, 2017 /
  Copyright 2007 - 2017" to "Version 5.0 beta / June 2026 / Copyright
  2007 - 2026", and the requirements sentence from ".NET Framework 4.0 or above
  ... Windows 7 or later" to ".NET Framework 4.8, built into Windows 10 and 11"
  plus a note that 5.0 is a 64-bit application. Regenerated EdSharp.htm.
  (Tutorial.md/Announce.md already carried no stale version text; the PowerBASIC
  10.0 reference is a third-party tool version and is left as-is.)

---

# EdSharp 5.0 baseline starter -- build notes (revision 41)

Revision 41 clarifies and tunes EdSharp's startup performance.

Background: EdSharp.exe is a managed .NET Framework 4.8 assembly (MSIL), JIT-
compiled by the CLR at launch -- not a native (machine-code) binary. ngen (the
Native Image Generator) pre-compiles it to a cached native image so the JIT cost
is skipped at startup.

What changed:
- Removed the vestigial BUILD-TIME ngen step from BuildEdSharp.cmd. It ran in the
  build folder, usually without admin, so it printed "Access is denied" and did
  nothing for the installed program (native images are tied to the installed
  location). This was the confusing part.
- Native pre-JIT is, and remains, handled by the INSTALLER: EdSharp_Setup.iss
  already runs "ngen install" against {app}\EdSharp.exe during setup (elevated,
  silent, with /AppBase), and "ngen uninstall" on removal. So a copy installed
  via EdSharp_Setup.exe is backed by a native image and starts faster than a
  copy run straight from the build folder.
- EdSharp.exe.config now disables generatePublisherEvidence. Without this the CLR
  can stall for seconds at startup verifying Authenticode publisher evidence
  against certificate-revocation lists, especially offline or behind a slow
  proxy. gcConcurrent=true is stated explicitly (it is the workstation default)
  to keep the UI responsive. Both are read at runtime -- no recompile needed.

How to confirm the native image (elevated command prompt):
    ngen display EdSharp
A line listing EdSharp with a "Native image ..." entry means it is installed.

Optional further gains (not done here):
- Multicore background JIT: a few lines in startup
  (System.Runtime.ProfileOptimization.SetProfileRoot/StartProfile, pointed at the
  data folder) speed warm starts when no native image is present. Largely
  redundant once the installer's ngen image exists; needs a recompile.
- True native/single-file AOT would require migrating from .NET Framework 4.8 to
  .NET 8 with ReadyToRun or NativeAOT -- a large, separate effort, and WinForms
  NativeAOT support is limited.

---

# EdSharp 5.0 baseline starter -- build notes (revision 40)

Revision 40 fixes JAWS intercepting Ctrl+Up / Ctrl+Down in the editor instead of
letting EdSharp's native paragraph navigation run (reported by a user whose
prior/next paragraph worked everywhere except EdSharp).

Diagnosis: EdSharp already maps Ctrl+Up = Prior Paragraph and Ctrl+Down = Next
Paragraph natively (menu accelerators). The keymap edsharp.jkm bound those keys
to the ControlUpArrow / ControlDownArrow scripts, which pass the key through to
EdSharp only when UIIsEditorWindow() is true. But UIIsEditorWindow() tests the
focus window class with an exact match (sClass == "WindowsForms10.RichEdit"),
which does not match the actual WinForms RichEdit class
(WindowsForms10.RICHEDIT50W.app.0.<hash>). So the test failed in the editor and
the scripts fell to their Else branch -- PerformScript ControlDownArrow() --
which, per the JAWS context search, runs the Default.jss ControlDownArrow and
applies JAWS's own (inferior) Ctrl+arrow behavior, exactly the symptom.

Fix (edsharp.jkm only): bind Control+UpArrow and Control+DownArrow directly to
the built-in TypeCurrentScriptKey, so JAWS unconditionally passes those keys to
EdSharp whenever EdSharp is focused and never falls through to the Default.jss
override. This is the approach the keymap's own commented-out
;Control+...Arrow=TypeCurrentScriptKey lines originally used, and matches the
active Alt+UpArrow/Alt+DownArrow=TypeCurrentScriptKey pattern. The redundant
commented lines were removed.

Why keymap-only: the .jkm keymap is read live by JAWS, so this needs no
recompilation. Changing EdSharp.JSS (e.g. to repair UIIsEditorWindow) would
require recompiling EdSharp.jsb in JAWS on Windows; the compiled .jsb is what
runs, so a source-only edit would have no effect. The ControlUpArrow /
ControlDownArrow scripts remain defined in EdSharp.JSS but are now unreferenced
by the keymap -- harmless, and left untouched so the shipped .jss and .jsb stay
in sync. (UIIsEditorWindow's class check is worth repairing in a future pass
that recompiles the scripts, since other editor scripts depend on it too.)

Delivery: the corrected edsharp.jkm reaches users through EdSharp's JAWS-settings
install (the Finish-page option / --install-jaws-settings), which copies the
Scripts files into each JAWS version's settings folder; reload JAWS scripts (or
restart JAWS) afterward.

---

# EdSharp 5.0 baseline starter -- build notes (revision 39)

Revision 39 fixes Alt+Ctrl+E launching the OLD EdSharp after a 5.0 install, and
aligns the desktop hot-key shortcut with the DbDo installer model.

Cause: the old installer assigned Alt+Ctrl+E to a shortcut on the USER's
desktop ({userdesktop}\EdSharp.lnk) pointing at the old exe.  The first 5.0
installer created its desktop shortcut on the COMMON desktop ({autodesktop}
-> common desktop under an admin install) with no hot key, so the old
user-desktop shortcut kept owning Alt+Ctrl+E and the old exe.

Fix (EdSharp_Setup.iss), following DbDo_setup.iss:
- The single hot-key shortcut is created with {autodesktop} + HotKey:
  Alt+Ctrl+E, exactly as DbDo creates its Alt+Control+D desktop shortcut.
  {autodesktop} adapts to the install scope (user vs all-users); no Start Menu
  item carries a hot key, so Alt+Ctrl+E has exactly one owner.
- EdSharp differs from DbDo in one respect: it has a legacy installer that left
  a conflicting Alt+Ctrl+E shortcut behind.  So [InstallDelete] removes any
  pre-existing EdSharp.lnk from BOTH {userdesktop} and {commondesktop} before
  [Icons] recreates the single shortcut, guaranteeing the new one is the sole
  owner regardless of where the old shortcut sat.
- No -activate parameter is needed.  DbDo passes -activate because its shortcut
  must choose GUI-activate over CLI mode; EdSharp is GUI-only and its
  WindowsFormsApplicationBase OnStartupNextInstance already foregrounds the
  running instance, so a plain relaunch activates rather than duplicating.

After recompiling EdSharp_Setup.iss and running it, log off/on (or reboot) if
Explorer has not already re-registered the changed shortcut.  Uninstalling the
old EdSharp separately is optional cleanup; it is not required for the hot key.

Immediate manual alternative (no reinstall): open the desktop EdSharp shortcut's
Properties and change Target to "C:\Program Files\EdSharp\EdSharp.exe" and
Start in to "C:\Program Files\EdSharp"; the existing Alt+Ctrl+E then launches
the new version.

---

# EdSharp 5.0 baseline starter -- build notes (revision 38)

Revision 38 reworks document conversion around the bundled 2htm utility and
makes the HTML Format command a Markdown-to-HTML converter.

HTML Format (Control+H): the command now treats the current document as
Markdown source and renders it to a complete HTML page with the bundled Pandoc
(`Util.Markdown2Html`: writes the buffer to a temp `.md`, runs
`pandoc -f gfm -t html5 -s --metadata title=...`, opens the result in a new
window). The old literal HtmlEncode-plus-Contents-TOC builder is gone. Modern
Pandoc flags only (no removed `-S`).

Document-to-text conversion: EdSharp now bundles `Convert\2htm\2htm.exe` and a
`Convert\any2txt.cmd` wrapper that runs it with `-p` (plain text) and renames
its `<basename>.txt` output to the exact `%Target%`. The default `EdSharp.ini`
routes `doc2txt`, `html2txt`, `ppt2txt`, `pptx2txt`, `xls2txt`, `xlsx2txt`,
`hlp2txt`, `wpd2txt` (Import) and `html2txt` (Export) through it. This retires
`GetText.exe` (which hung on modern Windows) and the `WdVert`/`PpVert`/
`htm2txt.exe` text paths. Pandoc lines, `pdf2txt` (Xpdf), and `xls2csv`/
`xlsx2csv` (XlVert, since 2htm cannot output CSV) are unchanged.

Because active converter definitions live in the per-user data-folder
`EdSharp.ini`, the shipped `EdSharp.inix` template now carries a ready
`[Import]`/`[Export]` override block: copy it next to the data-folder
`EdSharp.ini` and uncomment to switch an existing install over immediately
(inix entries are read in preference to ini), or edit that ini directly. New
installs get the 2htm converters from the updated default `EdSharp.ini`.

Files: added `Convert/2htm/2htm.exe` and `Convert/2htm/License.md` (committed
binary; `.gitattributes` marks `*.exe` binary, `.gitignore` does not exclude
`Convert/2htm`), `Convert/any2txt.cmd`, and an updated default `EdSharp.ini`.
The installer's recursive `Convert\*` rule already ships the new files.

---

# EdSharp 5.0 baseline starter -- build notes (revision 37)

Revision 37 rewrites the Elevate Version command (F11) to update from GitHub
Releases instead of the old PowerBASIC-era AppStamp.ini stamp file, following
the model of the sibling DbDo project.

How it now works:

- A new App.VersionString constant ("5.0.0") is the single source of truth for
  the version used in comparisons (the About dialog still shows "5.0 beta").
- Util.FetchLatestReleaseTag("JamalMazrui/EdSharp") reads the latest release
  tag.  It first calls the GitHub REST API
  (api.github.com/repos/JamalMazrui/EdSharp/releases/latest) and parses
  tag_name; if that fails it fetches github.com/.../releases/latest and takes
  the tag from the post-redirect address.  Both go through Homer.Web, which
  sends the required User-Agent header and uses TLS 1.2/1.3.
- Util.CompareVersions does a dotted-numeric comparison ("5.0" == "5.0.0").
- ElevateVersion() compares the tag with App.VersionString and shows a clear
  message: newer available (offer to install, default Yes), up to date (offer
  to reinstall), or running newer than public.  On confirmation it downloads
  github.com/JamalMazrui/EdSharp/releases/latest/download/EdSharp_Setup.exe to
  the temp folder via Homer.Web.download and starts it with
  ProcessStartInfo.UseShellExecute = true so the installer can request UAC.
  EdSharp is NOT closed; the Inno Setup installer detects the running EdSharp
  and offers to close it.  Every failure path gives a friendly message with a
  manual link to the releases page.

Removed from this path: the AppStamp.ini download, the plain-HTTP
EmpowermentZone URLs, Win32.Url2File (no callers remain; the helper itself is
left in place, unused), and Application.Exit.  The asset name EdSharp_Setup.exe
matches the installer's OutputBaseFilename (GitHub asset URLs are
case-sensitive).  The user guide and tutorial were updated to describe the
GitHub releases flow and the correct installer filename.

Caveat: F11 needs a published GitHub release containing EdSharp_Setup.exe.
Until the first release exists the command reports that it could not check for
updates and points to the releases page -- by design, not a failure.

---

# EdSharp 5.0 baseline starter -- build notes (revision 36)

Revision 36 is a pre-release QA pass (rev-35 built cleanly, confirming the
Transform Files rewrite, Web.cs, and the Help Tutorial item all compile).
Findings and fixes:

- User guide: a stray form feed before the Hotkey Summary heading and an
  unclosed code fence had trapped the Hotkey Summary list and the Development
  Notes, Contributors, and Third Party Utilities headings inside a code block,
  so they were not real headings. Fixed by removing the form feed and making
  the hotkey list one properly-closed code block; all four headings and all
  table-of-contents anchors now resolve. (Also fixed an EdSharp.htm/.md typo,
  "EdSharpapplication".)
- Installer: nvdaControllerClient.dll now uses skipifsourcedoesntexist, since
  the build treats it as optional; without the flag a developer lacking that
  DLL could not compile the installer.
- Verified: brace balances (+11 / 0), no c_ constants, no duplicate key
  bindings, version strings consistent (5.0 beta), removed symbols (VB.GetLinks,
  EasyEncode) absent, no bare URLs, Tutorial bindings match source, .gitignore
  ignores only reproducible artifacts, all text files CRLF.
- Known minor: some legacy List<T> locals in EdSharp.cs still use the l prefix
  rather than ls; left as-is to avoid churn (new code uses ls).

---

# EdSharp 5.0 baseline starter -- build notes (revision 35)

Revision 35 is the 5.0 beta release packaging.

## Tutorial in the Help menu
Added a Help > Tutorial command (Control+Shift+F1) that opens Tutorial.htm in
the default browser for the .htm extension, the same way Documentation (F1)
opens EdSharp.htm. The About box now reads "EdSharp 5.0 beta".

## HTML versions of the docs
EdSharp.htm, Tutorial.htm, and Announce.htm are generated from the matching
.md files with Pandoc (GitHub-flavored, standalone). EdSharp.htm is what F1
opens, so it must ship. Announce.md/.htm is the new release announcement.

## Program icon
EdSharp.ico (a multi-size E# badge) is added. BuildEdSharp.cmd already embeds
it via /win32icon when present, so EdSharp.exe carries the icon and the
shortcuts use it. The installer sets SetupIconFile and UninstallDisplayIcon.

## Installer (EdSharp_Setup.iss)
AppVerName is now "EdSharp 5.0 beta" (VersionInfoVersion 5.0). The icon is
set. The remaining C# sources (Lbc/Say/Inix/KeyMap/Web.cs) and the new docs
(EdSharp.htm, Tutorial.md/.htm, Announce.md/.htm, Transform_Example.inix) and
EdSharp.ico are shipped. Start-menu entries added for the Tutorial and the
announcement; the Manual entry now opens EdSharp.htm.

## Repository hygiene for GitHub (git add -A)
.gitignore excludes everything reproducible: EdSharp.exe, EdSharp.dll,
EdSharp_Setup.exe, BuildEdSharp.log, the fetched Ude.dll, and the downloaded
Convert tool folders (astyle, Pandoc, Tidy, liblouis, Xpdf) plus Tools.lock and
*.pandoc-bak. Committed non-built binaries (Tektosyne.dll,
nvdaControllerClient.dll, EdSharp.ico) are kept. .gitattributes keeps text in
CRLF and marks binaries. A developer section in the user guide explains
rebuilding with BuildEdSharp.cmd and building the installer with Inno Setup.
No EdSharp.cs logic change beyond the Help menu item; balance remains +11.

---

# EdSharp 5.0 baseline starter -- build notes (revision 34)

Your rev-33 log built cleanly (the rev-32 compiler C# default and .inix
compiler sections compile fine). Revision 34 reworks two features.

## Transform Files now uses the Regexer .inix job format

The Transform Files command (Alt+Equals) previously read a job as plain text,
four lines per task (comment, find, replace, blank). It now reads a Regexer-
style .inix job: one [Section] per task with Find, Replace, Options, Extract,
and Divider keys, and values may span multiple lines. The per-task engine is
ported from Regexer's Program.cs: Replace is run through Regex.Unescape (so \n,
\t, \" work) and $# inserts the running match count; Options is a comma list of
.NET RegexOptions names (multiline, ignorecase, singleline, ...); Extract=true
collects each match to the clipboard, joined by Divider (default form feed +
newline). The current document still supplies the file list (one path per
line) and the Test/Run/Verbose modes are kept. New Util helpers ToBool and
RegexOptionsFromString; parsing uses the existing Homer InixCodec (multi-line
values). Note: options now default to None per task (Regexer behavior) instead
of always-multiline -- specify Options=multiline where ^/$ should be per line.
Transform_Example.inix ships as a template; the user guide is updated.

## New Homer.Web module for web get/download (durl.py ideas in C#)

Web.cs (namespace Homer, BCL-only) brings the practical ideas of your durl.py
downloader to C#: modern TLS negotiation (the main fix for HTTPS failures on
older .NET defaults), a realistic desktop User-Agent, automatic redirect
following, filenames taken from Content-Disposition (including RFC 5987
filename*), an extension guessed from the MIME type when the URL has none,
sanitized and uniquified output names, and HTML link extraction (href/src plus
bare http and www URLs in text). The Web Download command now gathers links
with Homer.Web.getLinks -- replacing the old Internet-Explorer-based path,
which no longer works on current Windows -- and saves each file with
Homer.Web.download. Util.DownloadFile also gained the User-Agent and TLS setup.
Certificate validation is deliberately left on (durl's verify=False was not
carried over, to keep the update path safe); say the word to add a per-call
opt-out. Web.cs is added to the csc source list.

---

# EdSharp 5.0 baseline starter -- build notes (revision 33)

Documentation-only revision (no code changes from rev 32).

## Refreshed the dated MSDN links in EdSharp.md

All seven distinct MSDN library links (eight occurrences) were repointed from
the retired msdn2.microsoft.com VS.80 pages to their current Microsoft Learn
equivalents (regular-expression-language-quick-reference; the DateTimeFormatInfo,
RichTextBox, and Keys API pages under dotnet/api; custom-date-and-time-format-
strings; the JScript page under previous-versions; and the dotnet/csharp hub).
The descriptive link text is unchanged; only the destinations moved.

## New: Tutorial.md

Converted Jim Homme's TextPal tutorial (a predecessor of EdSharp) into an
EdSharp tutorial, adjusted to EdSharp terminology, concepts, and bindings:
sections/section breaks and Text Contents (not "topics"), Tab/Shift+Tab to
indent/outdent, Alt+Up/Down for sentences, Control+Up/Down for paragraphs,
Alt+F3 for find-at-cursor, F6/Shift+F6 for Go to Section/Contents, the C#
compile default, and EdSharp install specifics (EdSharp_setup.exe, the
Alt+Control+E desktop hot key, F11 to update). Pandoc Markdown, attribution
preserved. Ships in the zip; not yet wired into EdSharp_Setup.iss.

---

# EdSharp 5.0 baseline starter -- build notes (revision 32)

Revision 32 fine-tunes the Compile feature (Control+F5) and adds a built-in C#
default, plus a more solid way to define compilers via EdSharp.inix. It also
refreshes the user guide. No tool/encoding changes from rev 31.

## Built-in C# compiler default

With no compiler configured, pressing Control+F5 on a .cs file now compiles it
with the latest available .NET Framework C# compiler -- Roslyn csc.exe if VS
Build Tools are installed (newest C# language version), otherwise the csc.exe
that always ships with the .NET Framework (Util.FindCscPath locates it). The
default JumpPosition regex \(\d+,\d+\) matches csc's line,column, so EdSharp
jumps to the first error automatically. With no compiler and a non-.cs file,
Control+F5 now says "No compiler configured. Press Control+Shift+F5 to pick
one" instead of doing nothing.

## More solid compiler settings via .inix

Compilers are still stored as one tilde-delimited value in [Compilers], but
that packing breaks if a field contains a tilde and is hard to read. Pick
Compiler now also reads an optional section named "Compiler <name>" with one
key per setting (CompileCommand, JumpPosition, AbbreviateOutput, NavigatePart,
QuotePrefix, ExtensionDefault, GoToEnvironment); when present its keys override
the packed fields. In EdSharp.inix each value is kept verbatim, so regexes need
no escaping. Fully backward compatible: no such section means the old behavior.
EdSharp.inix now documents the format with a [Compiler C#] example.

## User guide (EdSharp.md)

All 18 URLs in the guide were folded into descriptive Markdown links
([text](<url>), angle-wrapped so the parenthesized MSDN URLs stay valid) so no
raw URL is shown to the reader. Content unchanged otherwise.

---

# EdSharp 5.0 baseline starter -- build notes (revision 31)

From your rev-30 log: clean build, and all five Convert tools now resolve
(AStyle 3.6 installs; Pandoc 3.10 / Tidy 5.8.0 / liblouis 3.38.0 / Xpdf 4.06
up to date). The tool subsystem is complete.

Revision 31 retires the external EasyEncode utf8b.exe from the import/convert
path. After a conversion wrote its target file, EdSharp used to shell out to
Convert\EasyEncode\utf8b.exe to re-encode that file, then read it. Since the
rev-27 encoding work, Util.File2String already detects the target's encoding
(byte-order mark, then content detection) and decodes it correctly, so the
re-encode pass was redundant. ConvertFile2String now just reads the target
with File2String. One fewer external tool in the conversion path; no behavior
change for correctly-converted files. Rename-safe, brace-balanced (+11).

---

# EdSharp 5.0 baseline starter -- build notes (revision 30)

Revision 30 audits the conversion batch files and ships improved .cmd versions,
and pins AStyle correctly. From your rev-29 log: Xpdf now installs (4.06) and
Pandoc/Tidy/liblouis are current. AStyle 404'd because best_release.json
reported a build number (3.6.16) with no matching download path; it is now
pinned to astyle-3.6-x64.zip on the byte-serving master.dl mirror.

New conversion scripts under Convert\ (consistent, quoted, reliable):
doc2txt.cmd, MinGW.cmd, pbw.cmd. Each drops the screen-clearing `cls`, quotes
every path, and fixes a real defect: MinGW.cmd gives g++ explicit -o targets,
and pbw.cmd now `type`s PB/Win's .log to stdout (the old `copy %2 %3` had an
empty %3, so compiler errors were never captured). Toolchain dirs are now
overridable via MINGW_BIN / PBWIN_BIN / PBWIN_INC. See Convert\CONVERT_NOTES.md
for the full audit, how EdSharp.cs invokes these, and the exact EdSharp.ini
line changes (the .bat->.cmd rename means the [Import]/[Compilers] references
must be updated; the notes give old/new lines to paste). Three scripts
(brl2txt.bat, BackTran.bat, tidy.bat) were not in any upload and still need to
be sent for audit. C# unchanged.

---

# EdSharp 5.0 baseline starter -- build notes (revision 29)

UPDATE from your rev-28 log: the build succeeded and Pandoc (3.10, flags
modernized automatically), Tidy (5.8.0), and liblouis (3.38.0) all updated.
Two download URLs were wrong and are now fixed:

* AStyle: the downloads.sourceforge.net link returned an HTML interstitial,
  not the zip. It now downloads from the byte-serving master.dl mirror and
  resolves its version from SourceForge best_release.json.
* Xpdf: the pinned 4.05 file 404'd because 4.06 superseded it. Xpdf now reads
  the current version from the xpdfreader download page and builds the URL.

Also: downloads now send a User-Agent and verify the file really is a ZIP (PK
signature) before extracting, so a bad URL fails with a clear message instead
of "End of Central Directory record could not be found". The manifest gained a
`page` resolver (url|version-regex|url-template) for hosts without a release
API. Rebuild; AStyle and Xpdf should now report ready. C# unchanged.

---

# EdSharp 5.0 baseline starter -- build notes (revision 28)

Revision 28 makes the build script keep the third-party Convert tools at their
LATEST versions, driven by a manifest, and auto-modernizes the Pandoc command
flags so the newer Pandoc keeps working. No C# changes from revision 27 (which
your log shows built cleanly).

## Tools.inix -- the tool manifest

The tool list moved out of FetchConvertTools.ps1 into Tools.inix (inix format,
one [Section] per tool: dir, file, flatten, src, optional sfrss/winget/choco).
This is build-time data and stays separate from the runtime EdSharp.inix on
purpose -- end users never need it. Add or adjust a tool by editing one block.

## Ensure-latest fetching

FetchConvertTools.ps1 was rewritten. For each tool it resolves the latest
version (GitHub "releases/latest" for Pandoc/Tidy/liblouis; the SourceForge RSS
or the pinned URL for AStyle; the pinned URL for Xpdf) and installs it only when
the version differs from what is recorded in Convert\Tools.lock, or when the
tool is missing. So normal rebuilds do nothing once you are current, and a new
upstream release is picked up automatically. As before it is best-effort and
never fails the build.

Current latest at this edit: AStyle 3.6, Pandoc 3.10, Xpdf 4.05; Tidy and
liblouis follow GitHub latest.

### winget / choco are a fallback only

EdSharp ships portable tools beside the program so end users need no package
manager, and winget/choco install system-wide rather than into Convert\<dir>.
So the direct downloads populate the distribution; winget/choco are tried only
if a direct download fails AND you have filled in a verified package id in
Tools.inix (the id fields are blank by default).

## Pandoc flags modernized automatically

Your EdSharp.ini Pandoc lines use 1.x syntax (-S, markdown_github,
--reference-docx) that Pandoc 2.0+ rejects. Whenever Pandoc is installed or
upgraded, the fetcher runs ModernizePandocConfig.ps1, which (idempotently, with
a one-time .pandoc-bak backup) rewrites only the pandoc.exe lines in EdSharp.ini
and EdSharp.inix: drops -S, changes markdown_github to gfm, and
--reference-docx to --reference-doc. If EdSharp keeps an active EdSharp.ini in a
separate user data folder, apply the same change there (or copy this folder's
EdSharp.ini over it). You can also run ModernizePandocConfig.ps1 by hand once.

## Verify

Rebuild. The [tools] lines now report "up to date (<version>)", "installing
<version>", or "upgrading to <version>", and Convert\Tools.lock records each
version. If Pandoc upgrades, a [pandoc] line reports the flag modernization.
Please send the new BuildEdSharp.log so we can confirm each tool resolved and
tune any download URL or asset pattern that needs it.

## On AStyle and a .NET replacement

There is no pure-.NET equivalent of AStyle. The .NET formatters are either
C#-only (CSharpier, dotnet format) or only syntax highlighters; AStyle uniquely
beautifies C, C++, C#, Objective-C, and Java in one small tool, so it stays an
external utility. AStyle does ship a shared DLL with an AStyleMain entry point
that could be called by P/Invoke, but that still bundles AStyle's native code,
so it is no simpler than the current astyle.exe.
