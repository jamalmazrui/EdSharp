; EdSharp_Setup.iss -- Inno Setup script for the AnyCPU EdSharp baseline (x64 and ARM64).
;
; Compile with ISCC.exe (Inno Setup 6.3+ required for x64compatible and for the
; per-user fallback described below). Run BuildEdSharp.cmd first so EdSharp.exe,
; EdSharp.dll, and nvdaControllerClient.dll exist. Produces EdSharp_Setup.exe
; in C:\EdSharp.
;
; Robustness revisions in this version:
; 1. AppId added so upgrades and uninstall entries stay tied to one identity
;    even if AppName ever changes.
; 2. PrivilegesRequiredOverridesAllowed=commandline. The interactive install
;    is unchanged -- no extra screen, standard elevation, default destination
;    C:\Program Files\EdSharp with the usual directory page for the rare user
;    who wants another folder. The directive only enables the /ALLUSERS and
;    /CURRENTUSER command-line switches as a documented escape hatch for an
;    account that cannot elevate (for example a standard account on a machine
;    where UAC has been disabled, so no elevation prompt can ever appear).
;    With /CURRENTUSER the app lands in {localappdata}\Programs\EdSharp via
;    {autopf}, the App Paths key is written under HKCU via the HKA root, and
;    the admin-only ngen steps are skipped by the isAdminNgen check.
; 3. Full VersionInfo block so the file properties identify the publisher and
;    product. This does not bypass SmartScreen, but it removes the anonymous
;    look that makes antivirus heuristics and cautious users more suspicious.
; 4. SignTool hook (commented) ready for Authenticode signing, which is the
;    real cure for the SmartScreen "Windows protected your PC" block on
;    downloaded copies. See the notes at the end of this header.
; 5. pandoc is fetched, not packaged, following the HomerView pattern. It is
;    roughly 200 megabytes, GitHub warns about it on every push, and not every
;    user converts documents. installPandoc.cmd and installPandoc.ps1 are
;    installed into {app}; a Finish-page checkbox (hidden when pandoc is
;    already in place) runs them elevated so they can write pandoc.exe into
;    {app}\Convert; and the same scripts can be run by hand later. Remember to
;    also take pandoc.exe out of the Git repository:
;      git rm --cached Convert/pandoc.exe
;    and add it to .gitignore, or the push warning will continue regardless of
;    what this installer does.
;
; SmartScreen note: an unsigned installer downloaded from the web carries the
; Mark of the Web and will be interposed by Microsoft Defender SmartScreen on
; first launch. No .iss directive can suppress that. Remedies, best first:
; sign the installer (SignPath Foundation is free for open-source projects;
; Azure Artifact Signing, formerly Trusted Signing, is Microsoft's low-cost
; service); publish to winget; or tell users to unblock the file before
; running it (file Properties dialog, Unblock checkbox on the General page, or
; in PowerShell: unblock-file EdSharp_Setup.exe).

[Setup]
AppId={{9F4E2C7A-1B5D-4E8A-B6C3-2D7F0A9E5481}
AppName=EdSharp
AppVersion=5.0.11
AppVerName=EdSharp 5.0.11 beta
VersionInfoVersion=5.0.11
VersionInfoCompany=NonvisualDevelopment.org
VersionInfoProductName=EdSharp
VersionInfoDescription=EdSharp 5.0 Setup
VersionInfoCopyright=Copyright 2006-2026 by Jamal Mazrui
SetupIconFile=EdSharp.ico
UninstallDisplayIcon={app}\EdSharp.exe
AppPublisher=NonvisualDevelopment.org
AppPublisherURL=https://github.com/JamalMazrui/EdSharp
AppCopyright=Copyright 2006-2026 by Jamal Mazrui
DefaultDirName={autopf}\EdSharp
DefaultGroupName=EdSharp
; x64compatible matches both x64 and ARM64 (Inno Setup 6.3+), so the AnyCPU
; EdSharp.exe installs and runs natively on both.  MinVersion 10.0 matches the
; .NET Framework 4.8 / Windows 10+ requirement.
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
Compression=lzma2/max
SolidCompression=yes
OutputBaseFilename=EdSharp_Setup
OutputDir=C:\EdSharp
SourceDir=C:\EdSharp
; Standard all-users install: elevation via UAC, no extra screens, default
; destination C:\Program Files\EdSharp.  The commandline value adds no UI; it
; only enables the /ALLUSERS and /CURRENTUSER switches, so documentation can
; tell a user whose account cannot elevate to run:
;   EdSharp_Setup.exe /CURRENTUSER
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=commandline
; The [Dirs] section below intentionally creates per-user data folders under
; {userappdata} even in an elevated install; EdSharp also recreates them at
; run time for each user, so a mismatch is self-healing.  Silence the compiler
; warning about user areas in an admin install.
UsedUserAreasWarning=no
ChangesAssociations=yes
ChangesEnvironment=yes
DisableProgramGroupPage=yes
DisableStartupPrompt=yes
Uninstallable=yes
SetupLogging=yes
; Authenticode signing hook.  After configuring a signing tool in the Inno
; Setup IDE (Tools menu, Configure Sign Tools) or via ISCC /S, uncomment:
; SignTool=edsharpsign
; SignedUninstaller=yes

[Files]
; Built artifacts (present after BuildEdSharp.cmd).
Source: "EdSharp.exe";        DestDir: "{app}"; Flags: ignoreversion
; Runtime configuration for EdSharp.exe -- carries the startup tuning (disables
; Authenticode publisher-evidence/CRL checks, enables concurrent GC).  It must
; sit next to EdSharp.exe, and ignoreversion ensures it is always refreshed so
; it stays in sync with the executable.
Source: "EdSharp.exe.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "EdSharp.dll";        DestDir: "{app}"; Flags: ignoreversion
Source: "nvdaControllerClient.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
; Source and build inputs (shipped so users can recompile, EdSharp-style).
Source: "EdSharp.cs";         DestDir: "{app}"; Flags: ignoreversion
Source: "Lbc.cs";             DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "Say.cs";             DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "Inix.cs";            DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "KeyMap.cs";          DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "Web.cs";             DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "EdSharp.ico";        DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "EdSharp.js";         DestDir: "{app}"; Flags: ignoreversion
Source: "EdSharp.manifest";   DestDir: "{app}"; Flags: ignoreversion
Source: "BuildEdSharp.cmd";   DestDir: "{app}"; Flags: ignoreversion
Source: "FetchConvertTools.ps1";   DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "FetchUde.ps1";            DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "ModernizePandocConfig.ps1"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "Tools.inix";              DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
; pandoc fetch scripts, installed so the Finish-page checkbox below can run
; them and so a user can run installPandoc.cmd by hand at any later time.
Source: "installPandoc.cmd";  DestDir: "{app}"; Flags: ignoreversion
Source: "installPandoc.ps1";  DestDir: "{app}"; Flags: ignoreversion
Source: "EdSharp_Setup.iss";  DestDir: "{app}"; Flags: ignoreversion
Source: "Tektosyne.dll";      DestDir: "{app}"; Flags: ignoreversion
Source: "Ude.dll";            DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
; JAWS settings family (compiled into each installed JAWS version by [Code]).
Source: "Scripts\*";        DestDir: "{app}\Scripts"; Flags: ignoreversion recursesubdirs skipifsourcedoesntexist
; NVDA add-on (installed on the Finish page via [Run]).
Source: "EdSharp.nvda-addon"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
; Configuration: do not clobber a user's existing settings on upgrade.
Source: "EdSharp.ini";        DestDir: "{app}"; Flags: onlyifdoesntexist
Source: "Hotkeys.ini";        DestDir: "{app}"; Flags: onlyifdoesntexist
; Documentation.
Source: "EdSharp.md";         DestDir: "{app}"; Flags: ignoreversion
Source: "EdSharp.htm";        DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "Tutorial.md";        DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "Tutorial.htm";       DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "Announce.md";        DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "Announce.htm";       DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "Transform_Example.inix"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "CamelType_JAWSScript.md"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "CamelType_CSharp.md"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "EdSharp.inix"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "history.txt";        DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "lgpl.txt";           DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
; Data trees.  pandoc.exe is excluded from Convert even if a copy is sitting
; there at compile time: the Run section below fetches it on the user's
; machine instead, so the installer stays small enough for GitHub.
Source: "Snippets\*"; DestDir: "{app}\Snippets"; Flags: recursesubdirs ignoreversion skipifsourcedoesntexist
Source: "Convert\*";  DestDir: "{app}\Convert";  Excludes: "pandoc.exe,*.sln,*.vcproj,*.vcxproj,*.vcxproj.filters,*.suo,*.user,*.c,*.asm,*.cs,*.obj,*.zip,temp.htm,temp.txt"; Flags: recursesubdirs ignoreversion skipifsourcedoesntexist

[Dirs]
Name: "{userappdata}\EdSharp";
Name: "{userappdata}\EdSharp\Temp";

[InstallDelete]
; Clear out any pre-existing EdSharp desktop shortcut before the [Icons] section
; recreates the single hot-key shortcut below.  This matters because EdSharp --
; unlike DbDo, which is a brand-new app -- has a legacy installer that placed an
; Alt+Ctrl+E shortcut on the USER's desktop pointing at the old exe, and an
; earlier 5.0 install placed a hot-key-less shortcut on the COMMON desktop.
; Removing both leaves the {autodesktop} shortcut below as the sole owner of
; Alt+Ctrl+E.  (InstallDelete runs before [Icons], so the recreate still wins.
; In a per-user install the common-desktop deletion may fail for lack of
; rights; InstallDelete ignores such failures, so this is harmless.)
Type: files; Name: "{userdesktop}\EdSharp.lnk"
Type: files; Name: "{commondesktop}\EdSharp.lnk"

[Icons]
Name: "{group}\Launch EdSharp";   Filename: "{app}\EdSharp.exe"; WorkingDir: "{app}"
Name: "{group}\EdSharp Manual";   Filename: "{app}\EdSharp.htm"
Name: "{group}\EdSharp Tutorial"; Filename: "{app}\Tutorial.htm"
Name: "{group}\EdSharp 5.0 beta Announcement"; Filename: "{app}\Announce.htm"
Name: "{group}\Uninstall EdSharp"; Filename: "{uninstallexe}"
; Single hot-key shortcut, following the DbDo model: the one shortcut that owns
; Alt+Ctrl+E is created with {autodesktop} (the user desktop for a per-user
; install, the common desktop for an all-users install) and HotKey.  No Start
; Menu item carries a hot key, so Alt+Ctrl+E has exactly one owner.  EdSharp is
; single-instance: OnStartupNextInstance brings the running copy to the
; foreground, so a plain relaunch activates rather than starting a second copy
; (no -activate parameter is needed, unlike DbDo's dual GUI/CLI shortcut).
Name: "{autodesktop}\EdSharp"; Filename: "{app}\EdSharp.exe"; WorkingDir: "{app}"; IconFilename: "{app}\EdSharp.ico"; HotKey: Alt+Ctrl+E; Comment: "Launch or activate EdSharp 5.0 (Alt+Control+E)"

[Run]
; Install EdSharps JAWS scripts (Finish-page option, like DbDo). Delegates to
; EdSharp.exe --install-jaws-settings, whose C# implementation copies the
; settings family into every installed JAWS version and compiles them there.
Filename: "{app}\EdSharp.exe"; Parameters: "--install-jaws-settings"; WorkingDir: "{app}"; Description: "Install JAWS scripts for EdSharp (recommended if you use JAWS)"; Flags: postinstall skipifsilent
; Install the NVDA add-on by shell-executing the .nvda-addon file (NVDA
; registers itself as the handler). Unchecked by default; checking it opens
; NVDA's add-on install dialog. NVDA must be running, and be restarted after.
Filename: "{app}\EdSharp.nvda-addon"; Description: "Install NVDA add-on (NVDA must be running; restart NVDA afterward)"; Flags: postinstall shellexec skipifdoesntexist unchecked
; Fetch pandoc into {app}\Convert, following the HomerView pattern: checked by
; default because the user who needs it cannot tell in advance that they do,
; and hidden entirely (needPandoc) once a copy is in place, so a reinstall
; asks nothing.  Unlike HomerView this entry runs ELEVATED on purpose -- no
; runasoriginaluser -- because the destination is under Program Files and only
; an elevated process can write there.  The console window stays visible so
; the download progress can be read; the script also writes a detailed log to
; the user's local application data, EdSharp\logs.
Filename: "{app}\installPandoc.cmd"; WorkingDir: "{app}"; Description: "Download and install pandoc, used to convert document formats (about 200 MB)"; Flags: postinstall skipifsilent; Check: needPandoc
; Pre-generate native images for faster startup (64-bit ngen).  ngen writes to
; the machine-wide native image cache, so it needs an elevated install; the
; isAdminNgen check skips it gracefully in a per-user install, where EdSharp
; simply JIT-compiles on first launch.
Filename: "{code:ngenExe}"; Parameters: "uninstall EdSharp /nologo /silent"; Flags: runhidden; Check: isAdminNgen
Filename: "{code:ngenExe}"; Parameters: "install ""{app}\EdSharp.exe"" /AppBase:""{app}"" /nologo /silent"; Flags: runhidden; Check: isAdminNgen

[UninstallRun]
Filename: "{code:ngenExe}"; Parameters: "uninstall EdSharp /nologo /silent"; Flags: runhidden; Check: isAdminNgen

[UninstallDelete]
Type: files; Name: "{app}\EdSharp.exe"
Type: files; Name: "{app}\EdSharp.dll"
Type: files; Name: "{app}\BuildEdSharp.log"
; pandoc.exe was placed by installPandoc, not by this installer, so Inno does
; not know to remove it; named here so an uninstall leaves no 200 MB orphan.
Type: files; Name: "{app}\Convert\pandoc.exe"

[Registry]
; HKA maps to HKLM in an elevated (all-users) install and to HKCU in a
; per-user install.  App Paths is honored from either hive, so "edsharp" keeps
; working from the Run dialog in both modes.
Root: HKA; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EdSharp.exe"; ValueType: string; ValueName: ""; ValueData: "{app}\EdSharp.exe"; Flags: uninsdeletekey

[Code]
function ngenExe(sParam: string): string;
begin
  // ngen ships with the 64-bit .NET Framework runtime; on an ARM64 system the
  // Framework64 path is the ARM64 framework.  isAdminNgen guards a missing
  // file and a non-elevated install.
  result := ExpandConstant('{win}\Microsoft.NET\Framework64\v4.0.30319\ngen.exe');
end;

function isAdminNgen(): boolean;
begin
  // ngen updates the machine-wide native image cache, which requires an
  // elevated install; in per-user mode we skip it and let the JIT handle
  // compilation at first launch.
  result := IsAdminInstallMode and FileExists(ExpandConstant('{code:ngenExe}'));
end;

function needPandoc(): boolean;
begin
  // pandoc is fetched by installPandoc.cmd rather than packaged.  The offer
  // is hidden once a copy is already in place, so a reinstall over a working
  // installation never even shows the checkbox (the HomerView lesson: an
  // offer the program always declines is an offer that should not be made).
  result := not FileExists(ExpandConstant('{app}\Convert\pandoc.exe'));
end;
