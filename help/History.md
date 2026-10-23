# EdSharp History



## 19 October 2026 -- The old converters retire

- **Your settings live in EdSharp.inix alone.** The first start copies your EdSharp.ini into your EdSharp.inix, in your EdSharp data folder, a value at a time and never over a value already there, and keeps the old file as EdSharp.ini.migrated; the log says how many values moved. From then on every setting is read from and written to the .inix, over the shipped defaults. Edit Settings File opens it, and an edit you save takes effect at once. Reset Configuration clears it, so the defaults show through. To go back, rename EdSharp.ini.migrated to EdSharp.ini. Favorites, Recent, the hotkey list and each compiler's own file are unchanged.
- **The window commands are the Homer suite's own:** Next Window, Previous Window, Say Windows Open and Close All but Current Window run the same code as every Homer app with windows inside. Close All but Current now says how many windows closed and how many were kept open for unsaved work, and Say Windows Open follows the suite's one, two or three presses.
- **PDFs are read with the Homer kit's own PDF reader,** as in every Homer app, keeping a tagged PDF's structure and reading an untagged one by its layout. The build fetches PdfPig by itself; where PdfPig and Markdig need the same library, the newer one is kept, and EdSharp.exe.config redirects every signed library to the version installed. It needs HomerDev 1.88.2 or later.
- **The guide describes opening other formats as EdSharp does it:** through the Homer kit first, with your own converter in the settings file used only for a format the kit does not cover. It no longer says a PDF converter is configured by default.
- **The guide's keys are the program's keys.** Its Hotkey Summary was a copy of an older key list, which still gave the bookmark commands Control+K, Control+Shift+K and Alt+K and gave the block commands keys they no longer have; it now points to the Hotkeys document, which is made from the program each build, and to the four ways to find a key inside EdSharp. Export Format is Alt+Shift+X, as the guide now says.
- **Recordings and speech, as in every Homer app:** a recording converts to a transcript, and a document to spoken audio, through the Homer kit. The build needs HomerDev 1.87.5 or later.
- **Web pages in the Homer page form,** as every Homer app makes them: one heading for the title, the page's parts marked for a screen reader, and pictures carried inside the page. The build needs HomerDev 1.87.1 or later.
- **Setup writes one log per run,** dated, as every Homer log is, instead of one ever-growing EdSharp_setup.log: EdSharp-setup-<date>.log for the installer's record, and one for each script it runs.
- **EdSharp's private converters are retired,** now that the Homer kit's conversions do their work: any2htm, any2md, any2txt, doc2txt, pdf2md, sheet2md, the Office converters and EdSharp's own copy of 2htm. The build moves them into notes\retired\convert, never deleting them, and the installer removes them from an earlier installation.
- **The Import and Export settings keep only what the kit does not do,** such as MediaWiki, LaTeX and MDX; 63 entries the kit serves were removed, since none could be reached.
- **The build needs HomerDev 1.79.1 or later,** for the kit's routes from old PowerPoint and Excel files.

## 19 October 2026 -- Key Describer, and command names said by the suite's rule

- **Control+F1 turns Key Describer off again,** a fault from the rename to Key Describer Toggle: the check that lets it through still looked for its old name. It now finds the command itself.
- **A command's name is said by the suite's rule,** as in DbDo and FileDir: not for a Say command, a dialog, whose title says it, or a move the screen reader announces. Commands whose handlers speak for themselves stay quiet, as before.

## 19 October 2026 -- The last key decisions

- **Hotkeys (Control+Shift+F1)** opens the Hotkeys page in your web browser, as in every Homer app. **Play Tutorials** moves to Alt+F12, beside the log and the AI chat in the F12 family.
- **Extra Speech Toggle is Alt+Shift+Z,** Z for toggles, as in DbDo and FileDir. **Reverse Items is Alt+Shift+R,** joining the other item commands on Alt+Shift with their first letter. **Rename is Shift+F2,** in the F2 family for renaming; plain F2 stays Special Character.
- **Go to Special Folder is Control+Shift+G,** as in FileDir. Control+Shift+0 is taken by Windows on many computers with more than one keyboard language. **Go to Environment** moves to Control+Shift+F12, beside Chat about Document.
- **Say Styles (Alt+Slash) also says the font and color,** so Say Font is no longer needed; its key, Alt+Dash, opens a document window's system menu in Windows.
- **Toggle Indentation is JAWS+Shift+GraveAccent** in EdSharp's JAWS scripts, beside Toggle Punctuation, instead of a Windows key.
- **Every choice can be taken back on its own.** The reasons and the old keys are in the HomerDev kit's KeyDecisions page.

## 19 October 2026 -- Suite names and keys

- **Commands now have the names they carry in every Homer app.** Prior becomes Previous throughout, such as Previous Paragraph and Previous Section; the regular expression commands say Regex, such as Forward Find with Regex; Forward and Reverse Find Again are Search Next and Search Previous. Configuration Options is Edit Settings, Manual Options is Edit Settings File, Set Bookmark is Save Bookmark, Go to Bookmark is List Bookmarks, Explorer Folder is Open in Explorer, Command Prompt is Open Command Prompt, Key Describer is Key Describer Toggle, and Tutorials is Play Tutorials.
- **Some keys moved, so the same key means the same thing in every Homer app:**
  - Edit Settings is Alt+Shift+E, as in DbDo, and Edit Settings File is Alt+Shift+F2.
  - Export Format is Alt+Shift+X, X for export.
  - Extra Speech Log is Alt+Shift+F12, beside Copy Log on Control+F12.
  - Search for Topic is Alt+F6.
  - Tile Vertical is Shift+F11, since Control+Shift+F11 belongs to JAWS.
  - Next Section and Previous Section also answer Control+F6 and Control+Shift+F6.
- **Search Next, F3, repeats Search for Topic** when that was your last search, so Search for Topic Again is gone.
- **New, as in the other apps:** Save Clipboard (Control+Apostrophe), Append Clipboard (Control+Shift+Apostrophe), Clear Clipboard (Alt+Shift+Apostrophe), and ReadMe (Alt+Shift+F1).
- **The key you hear is the key the command has.** A Hotkeys.ini kept from an earlier install could name an old key; the key now always comes from the command itself.

## 19 October 2026 -- Files convert through the Homer kit

- **Open Other Format, Text Convert and Export use the kit's shared conversions,** the same code as FileDir and the other Homer apps. Text comes straight out of Word, OpenDocument, Excel and PowerPoint files, Pandoc handles the markup formats, and a PDF keeps its headings, lists and tables through the kit's PDF route, as before. Microsoft Office is the last resort, with one Alert box when it is missing.
- **Your Import and Export commands still work** for targets the kit does not cover, such as MediaWiki, LaTeX and braille.
- **Export's .doc choice goes through the kit.** It still needs Microsoft Word, the only program that writes a 1997 Word file. A document of a type the kit does not know, such as a .cs file, is exported as plain text.
- **Export to Word 2003 XML is retired.** The format was replaced by .docx in 2007, which EdSharp exports through Pandoc.
- **The old Word text readers are gone,** with their code for PowerPoint and Excel, which nothing used.
- **The build needs HomerDev 1.76.0 or later.**

## 19 October 2026 -- The suite's shared commands and the kit's log

- **Four commands now run the same code as every Homer app,** the kit's SuiteCommands:
  - **Say Clipboard** (Alt+Apostrophe) says the text, or how many files are copied and their names. One press speaks, a quick second press spells, and a third shows the answer in a window you can review and copy.
  - **Say Time** (Alt+SemiColon) says the time, then the date.
  - **History of Changes** (Shift+F1) opens this history as a web page, as in the other apps, instead of a text file in an EdSharp window.
  - **Copy Log** (Control+F12) puts the log on the clipboard as a file and as its path, as before.
- **The session log is written by the kit.** It is now named EdSharp-yyyyMMdd-HHmmss.log, as every Homer app's log is. It opens with the full environment, including whether JAWS and NVDA are running. EdSharp's own lines about the outside commands it runs still go into it.
- **EdSharp needs HomerDev 1.74.0 or later to build.**

## 19 October 2026 -- EdSharp on the Homer framework's key router

- **Every key goes through the Homer suite's one router.** Each command EdSharp declares is now also a command in the kit's KeyMap, and EdSharp hands its keys to KeyMap.route, as every app on the Homer framework does. The document is marked as EdSharp's editor, so its own keys -- Shift+Space, Control+C and the rest -- still work while it has focus, and an unbound key still types. A quick second press of a key is counted the same way in every Homer app, and the kit's checks see EdSharp's keys like the others'.
- **Menus are quicker to open.** Each menu item's status-line handler used to scan every item of every menu each time it was drawn. An item now looks only at itself.
- **Bookmarks are on the B keys, as in DbDo:**
  - Control+B sets a bookmark;
  - Alt+B goes to it;
  - Control+Shift+B clears it.
  
  K was a letter from the middle of "bookmark", which the Homer key rule does not allow.
- **Block movement is gone.** Next Block, Prior Block and Say Block, which nothing used, left the menus and the tutorials. Next Indent and Prior Indent (Control+I and Control+Shift+I) move by indentation as before.

## 14 October 2026 -- a save that kept the file but not its backup says so

- **When Save succeeds but the .bak cannot be kept, EdSharp says so,** in its messages and its log, rather than nothing; and the staged copy is kept when it may be the only complete copy of what was typed. This follows a change to the kit's safe replacement (from a second audit by another AI).

## 14 October 2026 -- every write of your files accounted for

- **Writing a file is whole or not at all, wherever EdSharp writes your text.** Save was made safe earlier, but other paths still wrote straight over a file: the four text-writing utilities that many commands use, and export in another encoding or line-break style. They now write beside the file and swap it in, through the kit, with the same text and encoding as before, so the bytes written are unchanged (from an audit by another AI).
- **Export through a converter no longer loses the file it replaces.** The earlier file at the export path was deleted before the converter ran, so a conversion that failed lost it. It is now set aside, put back if no output comes, and removed only once the new file is there.
- **The developer guide lists every place EdSharp writes, moves or deletes a file,** sorted into those that replace your file through the kit, the one that sets the old file aside, your own deliberate actions, and EdSharp's own files that hold nothing you wrote.
## 14 October 2026 -- the user guide in plainer words

- **EdSharp.md reads at grade 9.0, down from 10.3,** within the ninth-grade target. About ninety of its densest paragraphs were rewritten, with every key, setting name, size and file type kept; reference blocks, code and lists of file types were left as they were. Two paragraphs that had run together on the page are separated, and two typos are fixed.
- **One out-of-date passage is corrected:** a paragraph still listed Boo, Java, Perl, PHP, PowerBASIC and Ruby among the compilers EdSharp has settings for; it now gives the current set, as the Compilers section already did.
- **Passages that may be out of date, kept as written for the author to decide:** a PowerBASIC batch file, though PowerBASIC left the compiler list; EdSharp.ini, Delphi.ini and a Roaming data folder, where other passages say EdSharp.inix; Keep Unique Items given as Alt+Shift+U in one place and Alt+Shift+K in the command list; Web Download taking its address from Internet Explorer; and the build fetching liblouis, though braille back-translation and its converter tools were retired in August.
## 14 October 2026 -- contents lists, and plainer guides

- **Every help page has a contents list.** The build ran Pandoc without --toc, so a long guide had no list of its sections. Pages are now built with one, and a page that lacks one is built again once to gain it.

## 14 October 2026 -- the announcement in plainer words

- **Announce.md now reads at grade 8.4, down from 9.8,** inside the ninth-grade target for documents shared publicly. Long sentences were split and dense phrases put plainly; every fact, key and link is kept, and first-person passages keep their author's own words. The section about the companion tools reads the same in each announcement that carries it, with the tools' hotkeys in alphabetical order.

## 14 October 2026 -- the newest Pandoc for the help pages

- **History.htm and the other pages are made by the newest Pandoc, not the first on the PATH,** which was Pandoc 2.19 in C:\bin while Pandoc 3.12.1 was installed in Program Files. The build asks the kit's newest.ps1, as the other apps' builds now do, running it as a process of its own so its answer can be read; the log names every copy found and its version (from an audit by another AI).

## 14 October 2026 -- the build checks its program where it is made

- **No more ABSENT for files that are there.** The build's closing report looked for EdSharp.exe and its libraries in the project folder, while they are built into exec, so every build said they were ABSENT. Each is now looked for where it is made, the installer in the project folder, and a program or library that is truly missing stops the build instead of being noted (from an audit by another AI). The NVDA controller DLL stays optional.

## 14 October 2026 -- saves that never truncate the file

- **A save writes a complete copy first, then swaps it in.** Text and RTF were written straight over the file, so a failed or interrupted save could leave it half written. The bytes saved are unchanged: the same encoding and line breaks (from an audit by another AI).
- **The old .bak is kept until the new one exists.** It was deleted before the new backup was made.

## 11 October 2026 -- Markdown to HTML works again

- **What Markdig needs is fetched, shipped and redirected.** Markdown to HTML stopped with "Could not load file or assembly 'System.Memory, Version=4.0.5.0'": Markdig 1.3.2 targets .NET Standard 2.0 and needs System.Memory, which needs System.Buffers, System.Runtime.CompilerServices.Unsafe and System.Numerics.Vectors, and the build fetched Markdig alone. The build now reads each package's own .nuspec for the lib target it chose, fetches what that names, and what those name in turn, records them in the package's .fetched note so a later build knows them without fetching, copies them into exec, and writes binding redirects for them into EdSharp.exe.config from the shipped DLLs; the installer ships exec\\System.*.dll. A .fetched note from before records no dependencies, so its package is fetched once more to learn them.
- **The error is logged.** It was shown in a box but not written to EdSharp's log, which therefore could not say what went wrong; the full exception is logged now.

## 11 October 2026 -- from an audit and a report

- **Launch after setup works on a first install.** The installer's Launch box writes a marker file into %LOCALAPPDATA%\\<App>\\logs, but on a first install that folder did not exist yet -- it was made later, when setup kept its log -- so the marker could not be written, and the program did not start after setup. The marker line now makes the folder first. Found by an audit of bookFido by another AI, and fixed in every Homer installer.

## 9 October 2026 -- from the build and release logs

- **The walks' reader speaks concrete words.** The Overview gently reminds of three reader keys, Insert plus Up Arrow, Insert plus Tab and Insert plus T, each with what EdSharp's editor would say, and that Caps Lock serves in place of Insert where set up; the Do Math walk's reader says the expressions typed ("1250 times 12 plus 300") in place of "(each character echoed as it is typed)".

- **A reinstall keeps its folder, and the screen reader boxes are the kit's.** EdSharp's installer did not remember an earlier install's folder, so every reinstall asked for one; it now reuses it and skips that page, as every Homer installer does. It also had its own copy of the kit's screen reader code; it now includes HomerComponents.iss, whose functions word and tick the JAWS and NVDA boxes by asking installJawsScripts.cmd the kit's question, "state jaws <file>", which that script now answers.

- **One level-one heading on every help page (9 October 2026).** build.ps1 converts the documents after the build's encoding pass, so their pages kept two level-one headings, the front matter's title and the document's own; the build now runs the pass again after build.ps1, and the kit's fixEncoding leaves each page with one.

- **Two fixes from an outside audit (9 October 2026).** A .cmd or .bat file now saves without a byte order mark, which cmd can misread; and Replace now treats a $ in the replacement as the character typed, where $1 or $0 had been taken as regular-expression substitution. The audit's other findings are triaged in notes.

- **The spelling dictionary ships again (9 October 2026).** The build fetched the dictionaries into a top-level Dictionaries folder, while EdSharp reads, and the installer ships, data\\dictionaries; so the installer found no dictionary to ship. The build now fetches them into data\\dictionaries and moves any already in the old folder there. The old top-level Snippets folder is retired into notes\\retired, like Convert and Samples; the installer ships templates\\snippets. Both were found by the kit's layout check once it stopped crashing.

- **Every help page has a title.** History.md has no front matter, so its page was built with no title and Pandoc warned on every build. A document with no title now takes one from EdSharp and its file name, such as EdSharp History.
- **The subtitle appears once.** Announce.md repeated its subtitle as a heading under the title, so the page said it twice, once as a heading that skipped a level.

## Unreleased -- 9 October 2026

- **Two old folders retired.** The top-level Convert and Samples folders were left from the layout before configs\convert and templates\samples; EdSharp reads neither, and all but one of their nine files were already kept, word for word, where it does read. The build now moves them into notes\retired, file by file, logging each, and deletes nothing.

## 8 October 2026 -- the pattern of ten, delivered and corrected

The conversion to ten spoken walks, made on 7 October, had not reached GitHub, where the twelve older walks remained; it is delivered again, and the build retires the twelve and their audio on its own. Checked against the current Homer practices, the walks were also corrected:

- **Walk 1, User Interface, rewritten to show, then tell.** One step had been a three-minute speech; now each idea is named and then shown, with its key and what the screen reader says: the window, Say Address and Say Status, the status line, dialogs, remembered answers, selection, help, the Key Describer and the Alternate Menu.
- **A command that does not exist, removed.** Walks 1 and 5 taught "Convert File Format, Control plus Shift plus V"; Control plus Shift plus V is Paste File, and EdSharp has no such command. Converting files on disk is Text Convert, Control plus T, which converts each file named in the document, one to a line, to a text file beside it; Text Combine, Control plus Shift plus T, puts them in one window.
- **Walk 8, Do Math, shows what it promised.** Powers, roots and rounding with Control plus Equals, and mathematics written the LaTeX way in Markdown, turned into a web page whose formula the screen reader reads as math.
- **Shorter lines, longer walks.** Walk 0's opening is three steps, not one long one, and walk 4 has more on moving through a long document.
- **The installer** ships the walk scripts and the audio's fingerprints as well.
- **Kit tools** updated from HomerDev 1.64.0.

## 8 October 2026 -- tutorials in the pattern of ten

The tutorials follow the Homer pattern of ten. The interface and key-pattern walks are now one, 1 User Interface, its narration folded into the reader's demonstrations; the five task walks are 3 to 7; 8 Do Math is new, Control plus Equals on the line or a selection, and math that reads aloud; and 9 Conclusion joins the conclusion, the glossary and more information. The installer now ships the tutorials' audio, playlist and feed, which it had left out.

A human-readable record of fixes and enhancements, newest first. Each entry
says what changed and why, so a future reader -- or a future maintainer --
can judge the decision, not just observe it.

## 6 October 2026 -- the walks run three to five minutes

**More substance in every part from 01 to 10.** The first spoken set measured
under two and a half minutes a walk; each was given the adjacent things its
want needs -- what each installer box is for, the result and preview windows,
a guessing exchange for the key rules, Undo and Save As and the window list,
Replace and regular expressions and File Find, batch conversion and
encodings, Say Compiler and the code-block runner, choosing the model and
asking well -- a planned misstep with its recovery where one fits, fourteen
more glossary terms, and a conclusion that keeps one thing from each walk in
two voices. Needs HomerDev 1.52.7.

## 6 October 2026 -- twelve spoken walks, in the Homer pattern

**Twelve walks replace the one.** 00 Overview and Table of Contents, 01
Install and Launch, 02 User Interface Concepts, 03 Key Patterns, five tasks --
04 Open, Edit and Save a File; 05 Find Your Way in a Long Document; 06 Convert
Between Formats; 07 Write Code; 08 Ask the AI on Your Own Computer -- then 09
Glossary, 10 Conclusion and 11 More Information. Each is built around a plain
want, each is under five minutes, and every recap and the glossary are two
voices: the host says the key or the term, the reader says the command or the
meaning. The earlier overview's real speech was kept where it fit.

**Pages made by pandoc get the Homer encoding.** The build regenerated
help\Tutorials.htm after fixEncoding had run, so it shipped without a byte
order mark and the release check refused it two days running; the page is now
given the mark and CRLF line endings as it is made. Needs HomerDev 1.52.6.

## 27 September 2026 -- setup fixes

- **Setup.** The Results box at the end of setup is titled "EdSharp Setup Results", and the finish page uses the Homer wording: the verb first, no "recommended", and "Launch EdSharp (desktop hotkey ...)". The JAWS scripts were never installed: the finish page ran scripts\installJawsScripts.cmd, which the installer had put at the top of the program folder, and "Scripts\*" copied the whole scripts folder into scripts\jaws. Both go where they belong now. The Results box said Pandoc was "not present" beside a working install, because it looked in configs\convert\Pandoc; it asks the same check as the finish page. Updating Python upgraded Python 3.13 whatever was installed, so with 3.14 it did nothing; it upgrades the release that is there. The boxes read "Install JAWS scripts" and "Install NVDA add-on (NVDA must be running)".

## 26 September 2026 -- built with HomerDev 1.43.19

- **The kit's tools carry the day's fixes.** `scripts\tidy` keeps a file where the project says it lives, so the two carried libraries stay at the top; `scripts\check` reads keys and access letters without false alarms and never waits for a key; and `scripts\release` publishes a draft -- 5.0.15 was left as one by an earlier run -- and confirms the release is GitHub's latest before calling it published.
- **The grave accent convention is now the kit's.** EdSharp's arrangement -- Alt for volume, Control for rate, Shift reversing, and the screen reader key for punctuation -- is how every Homer app with speech commands uses the key.

## 26 September 2026 -- brought up to HomerDev 1.43.10

- **The installer is written to the top of the project**, where the kit's release script looks for it. The release that day stopped with "EdSharp_Setup.exe not found" because the installer was in exec. The file keeps its name, EdSharp_Setup.exe: it is the asset every installed EdSharp downloads with F11.
- **The kit's scripts under their plain names.** The build refreshes check, push, release, tidy, unpushed, finish and installCommon into scripts and deletes the old checkHomerApp, gitPush, gitUnpushed, homerFinish, homerInstall, homerTidy and tagRelease copies. EdSharp's own install scripts and the installer use installCommon.cmd, the shared half of every install script.
- **The table conversions find inixVert.exe.** The build wrote it to Convert, the layout-before-the-kit's folder; the Import and Export tables run it from configs\convert, which is where it goes now.
- **The documents are converted again.** The build looked for pandoc in Convert\Pandoc, which the Homer layout no longer has, so no .htm had been regenerated since the move. It uses the pandoc on the PATH, as the conversions do, and now converts the documents in help as well as the ReadMe and License.
- **The two carried libraries come back.** A tidy on 26 September moved Tektosyne.dll and nvdaControllerClient.dll from the top of the project into notes, and git recorded them as deleted; kit 1.43.11 leaves them where they are. The build restores each from exec when the top copy is missing, so the next push carries them again.
- **Other people's files keep their encoding.** KeepEncoding.txt names configs\convert, data\dictionaries and templates\snippets\LaTeX, which fixEncoding and the kit's check now leave as they came -- a byte order mark on a Lua filter or a tool's config is read as part of its first line.
- **The fruit basket sample names its fields once.** It set an accessible name repeating each label, so a screen reader said "Fruit" twice; the labels now come just before their fields in tab order, which names them.
- **Tidier repository.** The copies of the kit's classes and the root tagRelease pair are removed from the top of the project; git records buildEdSharp.cmd, buildEdSharp.ps1 and EdSharp_setup.iss under their current capitals; and `.gitattributes` keeps the Homer CRLF line endings as they are, rather than converting them.

## 26 September 2026 -- moved to the Homer Development Kit

**Two reports from a beta tester, both taken.** A file opened by pressing
Enter on it in Explorer, or by Open With, did not move the remembered
folder: only the Open dialog did. The call that remembers the folder had
been commented out in the path every other kind of open goes through. It
runs there now, after the load, so the next Open starts where you are
working.

And EdSharp kept a private copy of Pandoc inside its own folder, and its
installer copied an existing Pandoc in if it found one, so a tester who had
it through Chocolatey ended up with two. That broke the Homer rule that a
shared component installs machine-wide to its own folder, where every Homer
program shares one copy and an app upgrade cannot destroy it. Every
conversion now asks for pandoc by name and takes the one on the PATH; the
installer offers Pandoc only when no pandoc answers anywhere; and the
install script installs it machine-wide with winget rather than copying it
anywhere.

**The NVDA controller client is carried by the repository.** NV Access
publishes it beside each NVDA release at an address that carries the
release number, so the address the build fetched from returned 404 for
three builds running once 2026.1.1 replaced 2026.1. The library has not
changed since NVDA 2024.1, and every word EdSharp speaks to an NVDA user
goes through it, so a fetch that fails on a fresh clone is the wrong side
of the trade. It joins Tektosyne as a named binary; the build still tries
to refresh it and says quietly when it cannot.

**The install scripts log through the kit.** Each install script carried
its own forty lines of log setup, all writing one shared EdSharp_setup.log
with no timestamp, so a second run wrote over the first. They now call the
kit's homerInstall.cmd first, which gives each run its own stamped log in
%LOCALAPPDATA%\\EdSharp\\logs and works out the quiet flag in one place.
The installer ships homerInstall.cmd beside them -- without it every script
would stop on its first line -- and passes noPause to each, since Inno runs
them with no console to read from. EdSharp's own installOllama.cmd is gone;
the kit's is refreshed into scripts by every build. Eighteen script
invocations inside the installer's Parameters strings were also still
naming the top of the folder rather than scripts, missed by the earlier
repoint; all moved.

**The finish page stays as it is, on purpose.** The kit's HomerComponents.iss
describes itself as a faithful port of what EdSharp and FileDir already had
right. EdSharp's installer is the original that file was ported from, so it
already behaves exactly as the include does. Replacing working Pascal with
its own port, with no Inno Setup on hand to compile the result, would be
risk without gain. If the kit's copy ever moves ahead, that is the time.

**The acceptance checks are rewritten** without parentheses, as one dir of
one file each. The first form failed nine out of nine on a quoting error.

**What the first build after the move found.** Six things, all of them
the migration only half done.

On the third attempt the build succeeded end to end: the program, the
JScript support assembly and the installer, all in exec. The release then
failed on its own rebuild, because homerTidy had carried Tektosyne.dll into
notes\other -- nothing named it, and the build does not fetch it, because
nothing can: it is a 2007-era library with no package and no download. It
is now named in RepoFiles.txt as the repository's one binary, the exception
the rule allows, and the build brings it home from notes if a tidy ever
carries it off again.


On the second attempt the program compiled against the kit's classes with
one warning and no errors -- the first proof that EdSharp and the kit's
newer Lbc, Say and Inix agree. The build then stopped in Inno Setup, which
was told to ship EdSharp.exe from the top of the folder when the build now
writes it to exec. The installer's binary lines were half repointed: the
ones with aligned columns had been missed, the same slip that caught the
documents a day earlier. And fixing that alone would have shipped a program
with no libraries beside it, since the fetched libraries land beside the
sources for the compiler and a program in exec cannot load from the folder
above. The build now copies every library it loads into exec after the
compile.


The build still ran EdSharp's own auditEdSharp.py, whose repoPolicy knows
nothing of the new folders, so it failed on "Tracked files are ones the
project needs" -- a check about a layout that no longer exists. The kit's
scripts\checkHomerApp does that job now, from tagRelease and by hand; the
build no longer calls the retired one. EdSharp's own invariants that the kit
does not cover, such as duplicate shortcuts and the compiler table, are
worth moving into the kit's check or into accept.inix rather than losing.

The build also still worked the version out from the release tags on the
remote and wrote it into the installer script. With version.txt saying 5.0.1
and this saying 5.0.43, two systems were minting numbers, and the second one
read the literal "{#MyAppVersion}" as a version because the installer now
computes it from version.txt. The version comes from version.txt alone.

And a failed build left no evidence: the PowerShell half logged to
BuildEdSharp.log at the top of the folder rather than into logs\, and the
build script did not capture its output. Both halves now write one stamped
log in logs\.

**EdSharp now shares one copy of the Homer classes with every other Homer
program.** It carried its own Inix, KeyMap, Lbc, Say, Web and inixVert, and
all six had drifted from the kit's. They are gone; the build compiles the
kit's, by path, out of C:\HomerDev, so a fix there reaches EdSharp on its
next build. Elevate came with them, which is what gives the Help box its
version section and F11 its update offer.

**The folder is laid out the way every Homer program is.** Sources, build
files, ReadMe and License at the top; configs, data, exec, help, logs,
scripts and templates below, each starting with a different letter so the
list can be walked by initial. Convert became configs\convert, Dictionaries
became data\dictionaries, Samples and Snippets became templates\samples and
templates\snippets, the documents moved into help, the JAWS scripts into
scripts\jaws, and the built program and installer into exec. The installed
tree mirrors it, so every path the program uses moved with it.

**The version lives in version.txt.** It was AppVersion in the installer
script, worked out from the release tags on the remote. Now the build steps
version.txt, writes Version.cs from it, the installer reads it, and
tagRelease reads it back out of what was actually built, so the program, the
installer and the tag cannot disagree.

**Nine of EdSharp's own tools were retired**, each an earlier edition of
something the kit now does: tidyRepo, repoPolicy, moveNotes, restoreMissing,
prepareAuditFixes, applyConvertPolicy, dropLatexJawsKeys, auditEdSharp and
summarizeSetup, along with the root tagRelease and the old BuildEdSharp
pair. The kit's homerTidy, checkHomerApp, gitPush, gitUnpushed, fixEncoding,
buildTutorials and tagRelease take their place, refreshed into scripts on
every build.

**Two documents were merged and one renamed.** Tutorial and Tutorials were
two documents covering the same ground and are now one; Development became
Developer, which is what the kit's check looks for. The CamelType documents
are the kit's now, so EdSharp's copies went.

**A first spoken walk** is in help\Tutorial_00_Overview.inix, which the
build speaks into an .mp3.

## 28 August 2026 -- version 5.0.40

**Three reports from a beta tester, addressed.**

*A document translated to English came back in Spanish.* The model was not
at fault. The language pair is remembered between uses and started at
English to Spanish, so the model was told to translate English into Spanish,
handed Spanish, and sensibly returned it much as it found it. The source
language now offers Detect, which is what a first use gets, and with Detect
chosen the prompt names no source at all -- naming the wrong one is worse
than naming none. A translation that comes back the same as the text that
went in now says so rather than opening a window that looks translated until
you read it.

*A replacement of a space and an x produced only the x.* A settings value on
one line is trimmed, so the space was gone before the job ran. Quoting the
value keeps it -- Replace=" x" -- and that already worked, but nothing said
so. Two things made it hard to find: \s, which is what a person reaches
for, means nothing in a replacement and is silently handed back as the
letter s by the .NET unescaper; and an older build did not yet strip the
quotation marks, so quoting appeared not to work either. No new escape was
invented for this -- a job file should stay readable as ordinary .NET
regular expressions -- and the quoting is now documented instead.

*Numbering passages that share an ending* turned out to need no new feature.
$# in a Replace value is the number of the match, counting from one, and has
been there all along without being written down anywhere.

A new Samples\\Jobs folder holds three worked jobs -- numbering passages,
turning a braille-ready file into readable text, and gathering headings to
the clipboard -- with a page explaining the format and the three things
easiest to miss.

**The language settings follow the file you are in.** The compiler was one
setting for the whole program, changed only by Pick Compiler. Everything
resting on it therefore described whatever file you last picked it for: with
C# chosen, a Python file compiled with csc, Alt+PageDown navigated by C#
declarations, comments were quoted with two slashes, indentation was four
spaces and Alt+V offered C# snippets. Five things wrong at once, and the
only cure was to remember.

A new CompilerExtensions section maps sixty file extensions to nineteen
languages, and EdSharp consults it when a document is opened or activated.
An extension nobody claims changes nothing, so a plain text file leaves the
last choice alone; Pick Compiler still overrides; and FollowDocument set to
0 turns it off.

**Twelve languages were added to the table**, chosen so that each gains
something real rather than an entry. Visual Basic .NET compiles with the
vbc.exe that sits beside csc.exe on every Windows machine. XML and JSON are
checked by the readers Windows already carries, through two small batch
files in Convert so the quoting is settled in one place rather than fought
with in a settings value. Perl, PHP and Ruby use their own syntax-only
switches. LaTeX returns to the table, having been retired six days earlier
as no longer relevant: it earns its place back because LaTeX is how
accessible mathematics is written, and Pandoc now turns that into MathML.
VBScript, HTML, Markdown, Inix, Batch and SQL have no checker worth having
and say so, but all six gained block navigation.

Fifteen of the nineteen can now report what is wrong before anything runs,
and all nineteen move by blocks that mean something in that language -- a
declaration in C#, a heading or landmark in HTML, a section in a settings
file, a label in a batch file, a sectioning command in LaTeX.

A new audit check holds the map to its sections, so an extension can never
point at a language that does not exist or be claimed by two.

**Mathematics now converts to MathML rather than MathJax.** The Markdown to
HTML conversions asked Pandoc for MathJax, which fetched a large script from
the internet every time a page opened, so the mathematics did not read at all
without a connection. They now ask for MathML, which is part of the page.

The reason to change is that the ground moved. Every major browser has drawn
MathML itself since MathML Core reached Chromium in 2023. NVDA reads it from
2025.1, and from 2026.1 it carries MathCAT, the math speech and braille
engine, built in with no add-on to fetch. JAWS reads it from its 2026
releases. Meanwhile MathJax version 4 turned off the hidden MathML that
screen readers had been relying on, preferring speech of its own. Native
MathML is smaller, works offline, and is read by the screen reader itself.

The guide's mathematics section was rewritten with it. It had been telling
readers to install MathPlayer and open the page in Internet Explorer, advice
that was sound around 2010 and has been wrong for years: MathPlayer has not
been developed since 2017 and Internet Explorer is retired.

**Lbc can build a dialog from a file, which is IniForm brought forward.**
IniForm, written in PowerBASIC between 2005 and 2015, turned a settings file
into a real Windows dialog and wrote the answers to a second file, so a
program in any language could put up an accessible form. The new
LbcInixForm class does the same from .inix, and being an ordinary Lbc dialog
it is 64-bit like the program hosting it, where the original built Windows
dialog templates by hand and could only ever be 32-bit.

The definition is .inix, so a list of items spans lines instead of being
crammed onto one separated by vertical bars, and the companion .txt file that
existed only to get round that limit is gone: an .inix value holds as many
lines as it needs. The bars still work, so every old definition still runs.
Four control types are added to the original nine -- combo, spin, heading and
separator -- and password, which was an attribute, is a type as well.

**The names come from IniForm, spelled in Camel Type.** The three passes
keep their own names -- dialogGroupControls, dialogSizeControls,
dialogPositionControls -- in their own order, the DialogBand record keeps its
fields, and the COM server's runForm, showResults and getResult are on
LbcInixForm under those names, so a caller written against IniForm needs
nothing new learnt. Only the spelling changes: lower camel case for methods
and properties, a Hungarian prefix on every variable, so i_borderPad becomes
iBorderPad and a_dlgBand becomes ldBands.

**A combo box now means what it has always meant.** The `combo` type is a box
you may type in, with a list beside it of the values already known, sorted
alphabetically with upper and lower case treated alike; the starting value
need not be one of them. The drop-down that allows only its listed values is
`droplist`. Lbc gained addComboEditBox for the first, beside the
addComboPickBox it already had for the second and addComboHistoryBox for a
list of what was typed before.

**Both layouts are kept, named in parallel.** LbcInixForm's Layout property
chooses between them, and the methods pair off: buildStack and buildBand,
runStack and runBand. Stack is Lbc's own, one control to a row, which is
what guarantees that reading order, tab order and visual order are one
order, and it stays the default. Band is IniForm's arithmetic, transcribed
whole into LbcBandLayout and checked against the layout IniForm itself
recorded for the Customer Information sample: all eighteen positions and
sizes, and the form size of 421 by 286, come out identical. Two details had
to be right for that -- the even-spacing division rounds half to even, as
PowerBASIC did when assigning to a Long, and dialog units become pixels
through base units measured from the form's own font rather than the 8 point
figures IniForm could assume. One slip is not carried forward: a band's
shared button width was written to the band but read from an array indexed
by control, so a band took its last control's width rather than its widest.
The documented intent was plain, and the intent is what is implemented.

A multi-line .inix value is read verbatim, from the start of its first line
to the end of its last, with nothing trimmed at either end at read or at
write. What to do with it afterwards belongs to the program reading it: a
range trims each item, because surrounding space there could only be an
accident of layout, while a memo trims nothing, because every space in it is
content. So a memo whose text should not be indented is written flush left,
or between fences, which also lets it hold lines that would otherwise look
like keys.

**A snippet can now describe its own dialog.** A block between [[form]] and
[[end]] is a form definition in that same format, and each answer replaces
the control's name between percent signs in the rest of the snippet. The
%Label=Value% syntax asks for text and nothing else; this gives a snippet
check boxes, lists, memos and spin controls, all obeying the same form rules
as every other EdSharp dialog. Sample definitions are in Samples\\Forms.

**The audit now checks every C# source for balanced braces.** It checked
EdSharp.cs and Inix.cs alone, which was enough until Lbc grew a class of its
own. Lbc.cs, Say.cs, KeyMap.cs and Web.cs are checked too.

**Snippet variables with more than one word in the label now work.** They
never had. The pattern that found them required the label to be a single run
of word characters, and a space is not one, so "%First participant=User%" was
invisible: no question was asked and the raw text was left in the finished
document. Of the eleven snippets that use variables, the three Mermaid ones
were affected, every variable in them but one. A label may now hold anything
except a percent sign, an equals sign or a line break.

**A value with vertical bars in it is now a list to choose from.** Writing
`flowchart %Direction=TD|LR|BT|RL%` gives a list box with TD selected instead
of a box to type into. This is the only addition to the syntax. Mustache and
Jinja braces were considered and set aside: they collide with the shell,
PowerShell and LaTeX snippets EdSharp already ships, and a template engine
would be a large dependency for a feature whose whole job is to ask a few
questions.

**Snippet dialogs follow the Homer form rules.** Every field now takes its
Alt letter from the start of a word, and no two fields in one dialog share a
letter; previously an ampersand was pushed in front of whatever the label
happened to be, so two fields starting with the same letter both answered the
same key. The rule moved into Lbc, where the button row already kept it, so
buttons and fields cannot come to different conclusions. Every field also
carries a focus tip, which is what F1 reads out.

**Snippets are saved to the Homer standard.** Save Snippet wrote plain ASCII
text as ANSI with no byte-order mark; it now writes UTF-8 with one and CRLF
endings, which is what the reader detects most reliably. Reading was already
correct, through the Ude detector. Two kinds of file take no mark: those the
command interpreter runs, and .gitignore and .gitattributes, where git reads
the mark as part of the first pattern so that line silently stops working.

**Mermaid diagrams are checked for accessibility.** Check Markdown, Alt+F9,
now looks inside a fenced mermaid block rather than skipping it, and applies
the checklist published by Princeton University Library: a diagram wants an
accessible name and a detailed description, should not set a theme in its
init block, should have its colours checked three ways if it sets any, and
should have a heading before it when a page holds several. The reason these
matter is that Mermaid does not tell assistive technology how its nodes
connect, so the description is not a caption for the picture -- for some
readers it is the picture.

The ten shipped snippets prompt for the name and the description first, so
the accessible version is the one that costs nothing extra, and seven new
ones cover class, state, entity relationship, Gantt, mindmap, timeline and XY
chart diagrams. A new Samples\\Mermaid folder holds one short example of each
alongside a page explaining the five rules. Every snippet and every example
was checked against the new rules and reports nothing.

**EdSharp is now under the MIT License.** It was released under the GNU
Lesser General Public License from 2007 through version 5.0. The MIT terms
ask less of anyone who wants to build on the code: use it, change it, ship it
in something of your own, including something you sell, as long as the
copyright notice goes with it. The full terms are in License.md, which also
names the parts of EdSharp that other people wrote and that keep their own
licenses -- the Ude encoding detector under the LGPL, the 2htm extractor in
the Convert folder, and the libraries and tools EdSharp fetches rather than
carries. Earlier releases stay under the terms they were published with.

**The working folder and the repository were separated properly.** What the
repository carries is now decided in one place, repoPolicy.py, read by both
the tidy script and the audit so they cannot disagree: a file belongs because
EdSharp_Setup.iss names it or because repoPolicy names it as a build file,
never because of what its name looks like. A pattern that admitted any
markdown or web page at the top of the folder had been quietly keeping saved
web pages, old mailing list messages and abandoned drafts in the repository
while every report said it was clean. Reference material now lives in a notes
folder that git ignores wholesale, the audit fails the build when anything
unnecessary is tracked, and a new restoreMissing script brings back any
needed file that has gone missing from the folder and the current commit
alike.

**The JAWS scripts are in the repository at last.** An ignore entry naming
the scripts folder, matched with regard to case where Windows does not, had
kept them out for years, so a fresh clone built an EdSharp that offered JAWS
users nothing. Two new audit checks guard the arrangement: the script sources
must be present to ship, and no compiled .jsb may be shipped, since a
compiled script belongs to the JAWS version that compiled it and each
installed version compiles its own.

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
