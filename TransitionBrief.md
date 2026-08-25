# EdSharp Transition Brief -- 20 August 2026

Start-of-chat context for continuing work on EdSharp. Everything below was
true when v5.0.20 was released and the repository was certified clean.

## Where things stand

- EdSharp v5.0.20 is released and working. It is the first good release
  since v5.0.10; every release from v5.0.11 through v5.0.19 ships a program
  that cannot start, for reasons fixed and recorded below.
- The repository at [github.com/JamalMazrui/EdSharp](https://github.com/JamalMazrui/EdSharp)
  is certified: 466 files tracked and none unnecessary, nothing missing,
  every remote tag matching its local copy, the pack at 83.15 MiB, no file
  of 5 MB or more left in history, and zero untracked files that are not
  gitignored. The saved emails, personal book drafts, help notes, and the
  copyrighted PDF are out of both tracking and history.
- The working folder is C:\EdSharp. The release routine is: buildEdSharp;
  launch C:\EdSharp\EdSharp.exe as a check; install; git add -A; git
  commit; git push; tagRelease. All of it now behaves: add -A is instant
  and stages only real changes.

## How EdSharp is built (do not rediscover this the hard way)

- The C# is ONE assembly. Every source file -- EdSharp.cs, Lbc.cs, Say.cs,
  Inix.cs, KeyMap.cs, Web.cs, plus the generated Version.cs -- compiles into
  EdSharp.exe. Never compile the support sources into a library called
  EdSharp.dll: two weak-named assemblies sharing the simple name "EdSharp"
  compile perfectly and can never run, because .NET answers requests for
  the library's types with the exe itself. That was the day-long silent
  launch failure.
- EdSharp.dll still exists and must keep existing: it is the JScript .NET
  EVALUATOR, compiled from EdSharp.js by jsc.exe, loaded by reflection at
  run time for evaluating expression strings (the FileDir and DbDo model).
  The build rebuilds it every run; the installer ships it beside EdSharp.js.
- Three libraries are pinned and version-verified at build time:
  ReverseMarkdown 4.7.1, HtmlAgilityPack 1.11.72, Markdig 0.42.0 (the last
  dependency-free Markdig). The build checks the version stamped inside any
  dll already on disk and refetches on mismatch. A stray Markdig 1.3 that
  passed a mere exists-check caused startup death by missing dependencies.
- The compiler is Roslyn csc found via vswhere, with the netstandard facade
  referenced for Markdig. buildEdSharp bumps the version by counting from
  the highest released tag; "buildEdSharp nobump" keeps the number;
  "buildEdSharp console" makes a temporary console-mode exe for debugging
  startup errors (runConsole.cmd captures its output), skips the installer,
  and never bumps. The build closes a running EdSharp politely (window
  close, never kill) before compiling, because a running EdSharp holds
  EdSharp.exe and EdSharp.dll open.
- Menu key facts: bindings are BUILT INTO the source as literals, by
  design; the legacy ini rebinding is gone; Hotkeys.ini supplies
  descriptions only. A blank key string means a menu-only command --
  String2Key returns Keys.None for it and CreateMenuItem skips shortcut
  bookkeeping. Four Markdown-era menu items are deliberately menu-only.

## How the installer works

- One consolidated log: %LOCALAPPDATA%\EdSharp\logs\EdSharp_setup.log.
  installPandoc, installJawsScripts, and Inno Setup's own record all append
  under dated banners (Inno's log is snapshot-copied first, because Inno
  holds it open for writing).
- JAWS scripts are the INSTALLER's job, not the editor's: installJawsScripts.ps1
  copies the Scripts folder into each JAWS version's settings and compiles
  with that version's scompile. It runs as the ORIGINAL user
  (runasoriginaluser), because JAWS settings live in the user's profile and
  the installer is elevated; the two sides cannot see each other's
  profiles, so the script reports through a small result file in C:\temp
  that the closing Results box reads. A silent-install twin entry exists
  because /SILENT skips postinstall tasks. The uninstall entry must NOT
  carry runasoriginaluser -- that flag is Run-only and breaks compilation.
- pandoc is fetched by installPandoc into Convert\Pandoc, checkbox-gated
  when absent. The Results box reports observed facts only.

## The repository lessons (apply to every project)

- Certify by PROOF, not assertion. "Untracked and left alone" must mean
  "and gitignored", verified with git check-ignore over every such path.
  Unignored leftovers are exactly what git add -A sweeps in.
- Always run git with core.quotepath=false in scripts. Otherwise non-ASCII
  paths flow through every list in an octal-quoted disguise, and ignore
  entries written for them can never match.
- Square brackets in .gitignore are glob character classes. Escape them (and
  * and ?) when writing literal entries, or names like
  "\[program-l\] Re ....txt" are never matched by their own entries.
- An installer pattern (root *.md, a wholesale folder) proves a tracked
  file is ALLOWED. It never proves an absent file is WANTED. Only exact
  names justify adding missing files, or personal documents get swept in.
- Fetch before counting ahead and behind. A push rejection should be a
  survey line, not a surprise.
- Two kinds of divergence, two cures: same lineage behind the remote means
  rebase then push; two different lineages mean ADOPT the good remote
  history -- bookmark the old lineage as a branch, mixed-reset to the
  remote, restore files the good tree has that the folder lacks, keep
  modified files as pending changes, and replace local tags with the
  remote's so nothing old force-pushes over good tags.
- History rewrites need proportionate verification: check reachability of
  the removed paths EXACTLY (not as substrings) over the branch, tags, and
  remotes -- not over bookmark branches, which hold old objects on purpose
  -- and let a residue smaller than the bulk threshold warn rather than
  abort.
- Every destructive step gets a dated backup first, and an aborted run must
  push NOTHING. That safety saved the good GitHub history three times in
  one day.
- tidyRepo.py (v10, on disk) embodies all of this. One future hardening is
  queued: skip nameless objects (unreachable debris) when listing rewrite
  candidates; garbage collection removes them anyway.

## The debugging lessons (apply to every program)

- A windowed .NET program that fails before Main's first line dies in
  perfect silence, and hooks inside Main never see it. The cheap decisive
  instruments: a console-mode build that prints the loader exception, and
  the Windows Application event log, which records the exception type for
  every silent .NET startup death.
- The error dialog can share the disease: if its own types depend on the
  broken library, the program that would explain the failure is the
  failure. Diagnose with instruments that do not depend on the patient.
- "It used to work" plus "it compiles" proves nothing about running: an
  assembly-name collision and a wrong-version library both compile clean
  and both kill at startup. Test a LAUNCH after build changes, not just
  the build.
- Evidence beats plausible stories. The menu crash was solved by scanning
  the real source for the actual inputs (four empty key strings), not by
  blaming the converter that had always worked.

## Editing scripts safely (the delivery lessons)

- Locate Inno Setup sections by line-anchored regex, never substring; a
  section token inside a comment once amputated half the file. Verify
  structurally after editing: the ten sections in order, Pascal code only
  after the Code header.
- Edit C# by brace-matching with a scanner that blanks strings, verbatim
  strings, char literals, and comments first; refuse to write unless the
  whole file still balances and no reference survives.
- Chain packaging on edit success (the shell "and" operator), or a failed
  edit ships a stale file.
- Match the noun to the count in every announcement, and deliver scripts,
  never manual steps, with debug-grade logs that make the next round
  unnecessary.

## Open items, none urgent

- Delete the nine broken releases v5.0.11 through v5.0.19 and their tags
  from GitHub so nobody installs one; a script for this is on offer.
- Reclaim local disk: three dated C:\EdSharp_backup_20260820 folders and
  the localBefore_20260820_095055 bookmark branch hold several gigabytes;
  deletable once v5.0.20 has settled.
- Optional source cleanups, scripts already on disk but never run:
  removeJawsFeature.py deletes the now-orphaned JAWS-install code from
  EdSharp.cs; addStartupLog.py gives every EdSharp run a timestamped log
  with crash capture. Either requires a rebuild afterward.
- JAWS content question: the old installer reported Notifications and
  VoiceProfiles buckets (72 files); the new script installs exactly what
  the Scripts folder contains (36). If that extra content matters, add
  Scripts\Notifications and Scripts\VoiceProfiles subfolders and the
  next install carries them automatically.
- The mystery of how C:\EdSharp's history got replaced by an old lineage
  mid-day was never solved; the bookmark branch preserves the evidence.

## Summary

EdSharp v5.0.20 is released and working, the repository is certified clean
in tracking and history, and the whole build, installer, and release
pipeline is scripted with its hard-won rules written into the scripts
themselves. The one-assembly rule, the evaluator's identity, version-
verified fetching, proof-based gitignore coverage, quotepath and glob
escaping, exact-and-proportionate rewrite verification, and launch-testing
after build changes are the lessons that made the nightmare survivable and
should keep it from recurring here or in any other project.
