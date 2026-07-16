; EdSharp_Setup.iss -- Inno Setup script for the AnyCPU EdSharp baseline (x64 and ARM64).
;
; Compile with ISCC.exe (Inno Setup 5.6+ or 6.x). Run BuildEdSharp.cmd
; first so EdSharp.exe, EdSharp.dll, and nvdaControllerClient.dll exist.
; Produces EdSharp_Setup.exe in C:\EdSharp.
;
; This is a slimmed, 64-bit replacement for the old edsharp_setup.iss.
; The legacy Java / JRE detection block and the obsolete 32-bit support
; assemblies (JsSupport, VbSupport, saapi32, nvdaControllerClient32) have
; been removed. Native code generation via ngen is kept; on ARM64 it targets
; the ARM64 framework, and HasNgen skips it gracefully if ngen is absent.

; ---- Version -----------------------------------------------------------------
; The version number is NOT stored in this script.  It lives in version.txt, one
; line, which Build<App>.cmd increments on every build.  Inno reads it here, and
; Build<App>.cmd also generates Version.cs from it, so the program, the installer,
; and the release tag always report the same number -- which is what Elevate
; Version (F11) compares.  Because no version literal appears in this file, a
; stale copy of it can never rewind the version.
#define VerFile FileOpen(AddBackslash(SourcePath) + "version.txt")
#define AppVersion Trim(FileRead(VerFile))
#expr FileClose(VerFile)
#undef VerFile

[Setup]
AppName=EdSharp
AppVersion={#AppVersion}
AppVerName=EdSharp {#AppVersion} beta
VersionInfoVersion={#AppVersion}
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
PrivilegesRequired=admin
ChangesAssociations=yes
ChangesEnvironment=yes
DisableProgramGroupPage=yes
DisableStartupPrompt=yes
Uninstallable=yes
SetupLogging=yes

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
; Data trees.
Source: "Snippets\*"; DestDir: "{app}\Snippets"; Flags: recursesubdirs ignoreversion skipifsourcedoesntexist
Source: "Convert\*";  DestDir: "{app}\Convert";  Excludes: "*.sln,*.vcproj,*.vcxproj,*.vcxproj.filters,*.suo,*.user,*.c,*.asm,*.cs,*.obj,*.zip,temp.htm,temp.txt"; Flags: recursesubdirs ignoreversion skipifsourcedoesntexist

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
; Alt+Ctrl+E.  (InstallDelete runs before [Icons], so the recreate still wins.)
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
; The four Finish-page checkboxes, in this order.  All are checked by default
; except the user guide.  The order here IS the order shown.
;
; 1. JAWS scripts.  "EdSharp.exe --install-jaws-settings" copies the script family into
;    every installed version of JAWS and compiles it there.  The implementation is the
;    shared Homer.JawsSettingsInstaller (in Say.cs), so EdSharp, FileDir, and DbDo all
;    install scripts by the same code, and the command can be re-run later.
FileName: "{app}\EdSharp.exe"; \
  Parameters: "--install-jaws-settings"; \
  WorkingDir: "{app}"; \
  Description: "Install scripts for improving use with the JAWS screen reader"; \
  Flags: postinstall waituntilterminated runhidden skipifsilent

; 2. NVDA add-on.  Shell-executing the .nvda-addon hands it to NVDA's own file
;    association, so NVDA shows its native add-on install dialog.  skipifdoesntexist
;    means the checkbox simply does not appear if the app ships no add-on yet.
FileName: "{app}\EdSharp.nvda-addon"; \
  WorkingDir: "{app}"; \
  Description: "Install add-on for improving use with the NVDA screen reader"; \
  Flags: postinstall shellexec waituntilterminated skipifsilent skipifdoesntexist

; 3. Launch the app.
FileName: "{app}\EdSharp.exe"; \
  WorkingDir: "{app}"; \
  Description: "Launch EdSharp (Alt+Control+E)"; \
  Flags: nowait postinstall skipifsilent

; 4. User guide -- the ONLY box not checked by default.
FileName: "{app}\EdSharp.htm"; \
  Description: "Open user guide for EdSharp"; \
  Flags: postinstall shellexec skipifsilent skipifdoesntexist unchecked

; Native image generation.  Not checkboxes: these run automatically and elevated, so
; the installed copy starts from a cached native image instead of JIT-compiling.
; Identical in all three apps.  HasNgen skips them if ngen.exe is absent.
FileName: "{code:NgenExe}"; Parameters: "uninstall EdSharp /nologo /silent"; Flags: runhidden; Check: HasNgen
FileName: "{code:NgenExe}"; Parameters: "install ""{app}\EdSharp.exe"" /AppBase:""{app}"" /nologo /silent"; Flags: runhidden; Check: HasNgen
; EdSharp.dll is loaded late-bound, so it is not in the exe's static dependency closure
; and ngen of the exe does not cover it.  Pre-compile it too.
FileName: "{code:NgenExe}"; Parameters: "install ""{app}\EdSharp.dll"" /AppBase:""{app}"" /nologo /silent"; Flags: runhidden; Check: HasNgen

[UninstallRun]
Filename: "{code:NgenExe}"; Parameters: "uninstall EdSharp /nologo /silent"; Flags: runhidden; Check: HasNgen

[UninstallDelete]
Type: files; Name: "{app}\EdSharp.exe"
Type: files; Name: "{app}\EdSharp.dll"
Type: files; Name: "{app}\BuildEdSharp.log"

[Registry]
Root: HKLM; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EdSharp.exe"; ValueType: string; ValueName: ""; ValueData: "{app}\EdSharp.exe"; Flags: uninsdeletekey

[Code]
function NgenExe(sParam: string): string;
begin
  // ngen ships with the 64-bit .NET Framework runtime; on an ARM64 system the
  // Framework64 path is the ARM64 framework.  HasNgen guards a missing file.
  result := ExpandConstant('{win}\Microsoft.NET\Framework64\v4.0.30319\ngen.exe');
end;

function HasNgen(): boolean;
begin
  result := FileExists(ExpandConstant('{code:NgenExe}'));
end;


