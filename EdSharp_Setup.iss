; EdSharp_Setup.iss -- Inno Setup script for the x64 EdSharp baseline.
;
; Compile with ISCC.exe (Inno Setup 5.6+ or 6.x). Run BuildEdSharp.cmd
; first so EdSharp.exe, EdSharp.dll, and nvdaControllerClient.dll exist.
; Produces EdSharp_Setup.exe in C:\EdSharp.
;
; This is a slimmed, 64-bit replacement for the old edsharp_setup.iss.
; The legacy Java / JRE detection block and the obsolete 32-bit support
; assemblies (JsSupport, VbSupport, saapi32, nvdaControllerClient32) have
; been removed. Native code generation via ngen is kept (64-bit).

[Setup]
AppName=EdSharp
AppVersion=5.0
AppVerName=EdSharp 5.0 beta
VersionInfoVersion=5.0
SetupIconFile=EdSharp.ico
UninstallDisplayIcon={app}\EdSharp.exe
AppPublisher=NonvisualDevelopment.org
AppPublisherURL=https://github.com/JamalMazrui/EdSharp
AppCopyright=Copyright 2006-2026 by Jamal Mazrui
DefaultDirName={autopf}\EdSharp
DefaultGroupName=EdSharp
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
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
Source: "EdSharp.exe.config"; DestDir: "{app}"; Flags: ignoreversion
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

[Icons]
Name: "{group}\Launch EdSharp";   Filename: "{app}\EdSharp.exe"; WorkingDir: "{app}"
Name: "{group}\EdSharp Manual";   Filename: "{app}\EdSharp.htm"
Name: "{group}\EdSharp Tutorial"; Filename: "{app}\Tutorial.htm"
Name: "{group}\EdSharp 5.0 beta Announcement"; Filename: "{app}\Announce.htm"
Name: "{group}\Uninstall EdSharp"; Filename: "{uninstallexe}"
Name: "{autodesktop}\EdSharp";    Filename: "{app}\EdSharp.exe"; WorkingDir: "{app}"

[Run]
; Install EdSharps JAWS scripts (Finish-page option, like DbDo). Delegates to
; EdSharp.exe --install-jaws-settings, whose C# implementation copies the
; settings family into every installed JAWS version and compiles them there.
Filename: "{app}\EdSharp.exe"; Parameters: "--install-jaws-settings"; WorkingDir: "{app}"; Description: "Install JAWS scripts for EdSharp (recommended if you use JAWS)"; Flags: postinstall skipifsilent
; Install the NVDA add-on by shell-executing the .nvda-addon file (NVDA
; registers itself as the handler). Unchecked by default; checking it opens
; NVDA's add-on install dialog. NVDA must be running, and be restarted after.
Filename: "{app}\EdSharp.nvda-addon"; Description: "Install NVDA add-on (NVDA must be running; restart NVDA afterward)"; Flags: postinstall shellexec skipifdoesntexist unchecked
; Pre-generate native images for faster startup (64-bit ngen).
Filename: "{code:NgenExe}"; Parameters: "uninstall EdSharp /nologo /silent"; Flags: runhidden; Check: HasNgen
Filename: "{code:NgenExe}"; Parameters: "install ""{app}\EdSharp.exe"" /AppBase:""{app}"" /nologo /silent"; Flags: runhidden; Check: HasNgen

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
  // 64-bit ngen ships with the .NET Framework runtime.
  result := ExpandConstant('{win}\Microsoft.NET\Framework64\v4.0.30319\ngen.exe');
end;

function HasNgen(): boolean;
begin
  result := FileExists(ExpandConstant('{code:NgenExe}'));
end;


