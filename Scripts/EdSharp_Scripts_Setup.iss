; Inno Setup source file for EdSharp Scripts
[Setup]
AppName=EdSharp Scripts
AppVerName=EdSharp Scripts 2.0
DefaultDirName={pf}\EdSharp\scripts
; DefaultGroupName=EdSharp Scripts
CreateUninstallRegKey=no
Uninstallable=no
ArchitecturesInstallIn64BitMode=x64 ia64
OutputBaseFilename=EdSharp_Scripts_setup
OutputDir=C:\EdSharp\scripts
Compression=lzma
SolidCompression=yes
DisableProgramGroupPage=yes
DisableReadyPage=yes
DisableDirPage=yes
DisableFinishedPage=no
PrivilegesRequired=None

[Files]
Source: "EdSharp.j*"; DestDir: "{code:GetScriptDir}"
Source: "homer.j*"; DestDir: "{code:GetScriptDir}"
Source: "MSAA.j*"; DestDir: "{code:GetScriptDir}"

[Icons]
Name: "{group}\Uninstall EdSharp Scripts"; Filename: "{uninstallexe}"

[Messages]
WizardSelectDir=Select Additional files folder
SelectDirDesc=
SelectDirLabel3=Additional files such as help and other files will be installed in the folder:
SelectDirBrowseLabel=
LicenseLabel3=Please read the following license agreement. Use the arrow keys to navigate the license agreement.
LicenseAccepted=I accept
LicenseNotAccepted=I do not accept
InfoBeforeClickLabel=After reading the following information click next.
InfoAfterClickLabel=After reading the following information click next.
FinishedHeadingLabel=
FinishedLabel=
ClickFinish=
ConfirmUninstall=Do you want to remove %1?
UninstalledAll=%1 has been removed from your computer.

[Run]
Filename: "{code:GetSCompilePath}"; Description: "Recompile script file EdSharp.jss (recommended)"; Parameters: """EdSharp.jss"""; WorkingDir: "{code:GetScriptDir}"; Flags: postinstall
Filename: "{code:GetSCompilePath}"; Description: "Recompile script file homer.jss (recommended)"; Parameters: """homer.jss"""; WorkingDir: "{code:GetScriptDir}"; Flags: postinstall

[Code]
var
  JAWSVersionsPage: TInputOptionWizardPage;
LanguagePage: TInputOptionWizardPage;
DataDirPage: TInputDirWizardPage;
JAWSVersions, Languages, Array2: array of String;
sVersion, sRootDir, sUserRootDir, sLanguage: String;
sRandomBaseName, sRunningVersion: String;
VersionCount, LanguageCount:Integer;
sDefaultUserPath, sDefaultSharedPath, sCNIPath: String;
InnoJCFExists, InnoJSBExists, InnoJSDExists, InnoJSSExists: Boolean;

Const
GW_Child = 5;
GW_NEXT = 2;
CNIFile = 'ConfigNames.ini';

function Show(sText: string): integer;
begin
result := SuppressibleMsgBox(sText, mbInformation, MB_OK, MB_OK);
end; // Show function

function Confirm(sText: string): integer;
begin
result := SuppressibleMsgBox(sText, mbConfirmation, MB_YESNO, MB_DEFBUTTON1);
end; // Confirm function

Function GetWindow(hwnd: longint; ucmd: cardinal): longint;
external 'GetWindow@user32.dll stdcall';

function GetWindowText(hWnd: HWND; lpString: String; nMaxCount: Integer): Integer;
external 'GetWindowTextA@user32.dll stdcall';

Function Pad(s: String): String;
begin
if (Length(s) - Pos('.', s) = 1) then
s := s + '0';
if (Pos('.', s)=2) then
s := '0' + s;
Result := s;
end;

function GetSCompilePath(s: String): String;
Var
sCompile: String;
sVer: String;
begin
if (s = '1') then
sVer := sRunningVersion
Else
sVer := sVersion;
RegQueryStringValue(HKLM, 'Software\Freedom Scientific\JAWS\'+sVer, 'Target', sCompile);
If (Length(sCompile) < 3) Then
begin
if (Pad(sVer) >= '06.00') then
sCompile := ExpandConstant('{pf}\Freedom Scientific\JAWS\' + sVer)
Else
sCompile := ExpandConstant('{sd}\JAWS' + Copy(sVer, 1, 1) + Copy(sVer, 3, 2));
end;
sCompile := sCompile + '\scompile.exe'
Result := sCompile;
end;

Procedure GetSubdirs(sPath: String; Mode: Integer);
Var
sTemp: String;
FindRec: TFindRec;
begin
if FindFirst(sPath + '*', FindRec) then begin
try
repeat
if (FindRec.Attributes and FILE_ATTRIBUTE_DIRECTORY <> 0) and (Pos('.', FindRec.Name) <> 1) then
begin
If (Mode = 0) Then
begin
Languages[LanguageCount] := FindRec.name;
LanguageCount := LanguageCount + 1;
End
Else
If (Mode = 1) Then
begin
if (LowerCase(FindRec.Name) = 'enu') then
LanguagePage.Add(FindRec.Name + ' (English US)')
Else
LanguagePage.Add(FindRec.Name);
Languages[LanguageCount] := FindRec.name;
LanguageCount := LanguageCount+1;
End
else if (Mode = 2) then
begin
if (Pos('JAWS', FindRec.Name)>0) then
sTemp := Copy(FindRec.Name, 5, 1) + '.' + Copy(FindRec.Name, 6,2)
else
sTemp := FindRec.Name;
if (Pad(sTemp) >= '01.00') and (Pad(sTemp) <= '99.90') then
begin
JAWSVersions[VersionCount] := sTemp;
VersionCount := VersionCount+1;
end
end;
end;
until not FindNext(FindRec);
finally
FindClose(FindRec);
end;
end;
end;

Procedure InstallInnoScripts;
Var
sDir: String;
ResultCode: Integer;
begin
sDir := GetCurrentDir();
if (Pad(sRunningVersion) >= '06.00') then
sDefaultUserPath := ExpandConstant('{userappdata}') + '\Freedom Scientific\JAWS\' + sRunningVersion + '\Settings\'
Else
sDefaultUserPath := ExpandConstant('{sd}\JAWS' + Copy(sRunningVersion, 1, 1) + Copy(sRunningVersion, 3, 2)+'\Settings\');
sDefaultSharedPath := ExpandConstant('{commonappdata}') + '\Freedom Scientific\JAWS\' + sRunningVersion + '\Settings\';
GetSubdirs(sDefaultUserPath, 0);
sDefaultUserPath := sDefaultUserPath + Languages[0] + '\';
sDefaultSharedPath := sDefaultSharedPath + Languages[0] + '\';
InnoJCFExists := FileExists (sDefaultUserPath + 'InnoSetup.jcf');
InnoJSBExists := FileExists (sDefaultUserPath + 'InnoSetup.jsb');
InnoJSDExists := FileExists (sDefaultUserPath + 'InnoSetup.jsd');
InnoJSSExists := FileExists (sDefaultUserPath + 'InnoSetup.jss');
(*
ExtractTemporaryFile('InnoSetup.jss');
ExtractTemporaryFile('InnoSetup.jcf');
ExtractTemporaryFile('InnoSetup.jsd');
FileCopy(ExpandConstant('{tmp}\InnoSetup.jss'), sDefaultUserPath+'InnoSetup.jss', False);
FileCopy(ExpandConstant('{tmp}\InnoSetup.jcf'), sDefaultUserPath+'InnoSetup.jcf', False);
FileCopy(ExpandConstant('{tmp}\InnoSetup.jsd'), sDefaultUserPath+'InnoSetup.jsd', False);
If Not FileExists(sDefaultUserPath + CNIFile) Then
FileCopy(sDefaultSharedPath + CNIFile, sDefaultUserPath+CNIFile, True);
sCNIPath := sDefaultUserPath + CNIFile;
SetIniString('ConfigNames', sRandomBaseName, 'InnoSetup', sCNIPath);
Exec(GetSCompilePath('1'), 'InnoSetup.jss', sDefaultUserPath, SW_HIDE,ewWaitUntilTerminated, ResultCode)
*)
SetCurrentDir(sDir);
end;

Procedure DeinitializeSetup;
Var
sDir: String;
begin
sDir := GetCurrentDir;
if (Pad(sRunningVersion) >= '05.00') then
DeleteIniEntry('ConfigNames', sRandomBaseName, sCNIPath);
if not InnoJCFExists then
DeleteFile(sDefaultUserPath+'InnoSetup.jcf');
if not InnoJSBExists then
DeleteFile(sDefaultUserPath+'InnoSetup.jsb');
if not InnoJSDExists then
DeleteFile(sDefaultUserPath+'InnoSetup.jsd');
if not InnoJSSExists then
DeleteFile(sDefaultUserPath+'InnoSetup.jss');
SetCurrentDir(sDir);
end;

function InitializeSetup(): Boolean;
Var
I, V2, Ret, Position: Integer;
H: Hwnd;
sTemp: String;
begin
Sleep(300);
SetArrayLength(JAWSVersions, 25);
SetArrayLength(Languages, 25);
SetLength(sRunningVersion, 201);
sRandomBaseName := Copy(ParamStr(0), Length(ParamStr(0))-11, 8);
H := FindWindowByWindowName('JAWS');
H := GetWindow(H, GW_CHILD);
H := GetWindow(H, GW_CHILD);
While (H<> 0) and (Pos('Version', sRunningVersion) = 0) do
begin
H := GetWindow(H, GW_Next);
 ret := GetWindowText(H, sRunningVersion, 200);
end;
Position := Pos('.', sRunningVersion);
sTemp := Copy(sRunningVersion, Position+1, 10);
sTemp := Copy(sTemp, 1, Pos('.', sTemp)-1);
if (sTemp = '00') then
sTemp := '0';
sRunningVersion := Copy(sRunningVersion, 14, Position-13) + sTemp;
if (RegGetSubkeyNames(HKLM, 'Software\Freedom Scientific\JAWS', Array2) and (GetArrayLength(Array2) > 0)) then
begin
if (Pos('.', Array2[0]) > 0) then
begin
V2 := GetArrayLength(Array2);
while (I < v2) do
begin
if (Pos('.', Array2[I])>0) and (Pad(Array2[I]) >= '01.00') and (Pad(Array2[I]) <= '99.90') then
begin
JAWSVersions[VersionCount] := Array2[i];
VersionCount := VersionCount+1;
end;
I := I + 1;
end;
End
Else
begin
SetArrayLength(JAWSVersions, 25);
VersionCount := 0;
GetSubDirs(ExpandConstant('{sd}\JAWS'), 2);
GetSubDirs(ExpandConstant('{pf}\Freedom Scientific\JAWS\'), 2);
End
if (Pad(sRunningVersion) >= '05.00') then
; InstallInnoScripts;
If (VersionCount > 0) Then
Result := True
Else
MsgBox('JAWS for Windows is not installed on this computer', MBError, MB_OK);
end
else
MsgBox('JAWS for Windows is not installed on this computer', MBError, MB_OK);
end;

procedure InitializeWizard;
Var
I: Integer;
begin
JAWSVersionsPage := CreateInputOptionPage(wpInfoBefore,
'Select JAWS Version', '',
'In which version of JAWS do you want to install the scripts.',
True, False);
For I :=0 to VersionCount-1 do
begin
JAWSVersionsPage.Add('JAWS ' + JAWSVersions[i]);
end;
JAWSVersionsPage.SelectedValueIndex := VersionCount-1;

LanguagePage := CreateInputOptionPage(JAWSVersionsPage.ID,
'Select Language folder', '',
'Select the language folder where you want to install the scripts',
True, False);

DataDirPage := CreateInputDirPage(LanguagePage.ID,
'Confirm script location folder', '',
'The scripts will be installed in the folder.',
False, '');
DataDirPage.Add('');
end;

function ShouldSkipPage(PageID: Integer): Boolean;
Var
SkipDataDir, SkipInstallScripts: Integer;
begin
if (PageID = JAWSVersionsPage.ID) and (SkipInstallScripts = 1) then
Result := True
Else
If (PageID = LanguagePage.ID) And ((SkipInstallScripts = 1) or (LanguageCount < 2)) Then
Result := True
Else
If (PageID = DataDirPage.ID) And ((SkipInstallScripts = 1) or (SkipDataDir = 1)) Then
Result := True
Else
Result := False;
end;

Function GetLanguageDir: String;
begin
LanguagePage.CheckListBox.Items.Clear
LanguageCount := 0;
GetSubdirs(sRootDir, 1);
if (LanguageCount > 0) then
LanguagePage.SelectedValueIndex := 0;
if (LanguagePage.SelectedValueIndex = -1) then
MsgBox('No valid language folder exists for the JAWS version you selected. You must select another JAWS version to proceed.', MBError, MB_OK);
Result := Languages[LanguagePage.SelectedValueIndex]
end;

function ValidateLanguagePath: Boolean;
var
L: Integer;
sPath: String;
begin
sPath := RemoveBackslash(LowerCase(DataDirPage.Values[0]));
L := Length(sPath);
If (Not DirExists(sPath)) Then
MsgBox('The selected path and folder cannot be found. Please reenter the path'#13#13'or click browse and select the appropriate path and folder.', MBError, MB_OK)
Else
begin
if (Copy(sPath,L-3,1) <> '\') or (Pos('\',Copy(sPath, L-2, 3) ) > 0) then
MsgBox('The selected folder is not a valid JAWS language folder.'#13#13'Examples of valid folder names are enu (English US)'#13#13'deu (German) and esp (Spanish).', MBError, MB_OK)
Else
begin
if (Copy(sPath,L-12,10)<>'\settings\') or (Pos('jaws',sPath)=0) or (L<9) then
MsgBox('The selected path and folder is not a valid JAWS settings folder.'#13#13'Click browse and select the appropriate folder.',MBError,MB_OK)
Else
begin
Result := True;
if (Pos('all users',sPath) > 0) then
if (MsgBox('You are attempting to install the scripts into the JAWS shared settings folder.'#13#13'Doing this may permanently overwrite or cause conflicts with the default Freedom Scientific scripts. Are you sure you want to proceed?',MBError,MB_YESNO or MB_DEFBUTTON2) = IDNO) then
Result := False;
end;
end;
end;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
If (CurPageID = JAWSVersionsPage.ID) Then
begin
sVersion := JAWSVersions[JAWSVersionsPage.SelectedValueIndex];
if (Pad(sVersion) < '06.00') then
begin
RegQueryStringValue(HKLM, 'Software\Freedom Scientific\JAWS\'+sVersion, 'Target', sRootDir);
sRootDir := sRootDir + '\Settings\';
End
Else
sRootDir := ExpandConstant('{userappdata}') + '\Freedom Scientific\JAWS\' + sVersion + '\Settings\';
sUserRootDir := sRootDir;
sLanguage := GetLanguageDir;
DataDirPage.Values[0] := sRootDir + sLanguage;
Result := True;
End
Else
If (CurPageID = LanguagePage.ID) Then
begin
sLanguage := Languages[LanguagePage.SelectedValueIndex];
DataDirPage.Values[0] := sRootDir + sLanguage;
Result := True;
End
Else
If (CurPageID = DataDirPage.ID) Then
begin
Result := ValidateLanguagePath;
End
Else
Result := True
end;

function GetScriptDir(Param: String): String;
begin
  Result := DataDirPage.Values[0];
end;

Function GetJSIPath(Param: String): String;
begin
if (Pad(sVersion) >= '06.00') then
Result := GetScriptDir('') + '\PersonalizedSettings'
else
Result := GetScriptDir('');
end;

function GetCNIPath(Param: String): String;
begin
  Result := sCNIPath;
end;
