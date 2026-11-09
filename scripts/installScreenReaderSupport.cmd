@echo off
rem installScreenReaderSupport.cmd -- put this app's JAWS scripts and NVDA
rem add-on where each screen reader looks for them.
rem
rem SHIPPED WITH EVERY HOMER APP THAT HAS SCRIPTS, and its installer checkbox is
rem CHECKED BY DEFAULT. A blind user installing a Homer tool wants the scripts;
rem making them notice and tick a box is the friction this suite exists to
rem remove.
rem
rem What it looks for, beside this script:
rem   <App>_JAWS.zip    the JAWS script files, unpacked into the user's own
rem                     JAWS settings folder for every JAWS version found
rem   <App>.nvda-addon  the NVDA add-on, handed to NVDA to install
rem
rem Either may be absent; what is here is installed and what is not is skipped
rem without complaint.
rem
rem AI NOTE FOR CUSTOMIZING: nothing here is app-specific. The app name comes
rem from the folder. If an app's scripts need compiling rather than copying, add
rem that step after the unpack and log its exit code.
rem
rem NOTHING PAUSES. LOG: %LOCALAPPDATA%\<App>\logs\<App>-<script>-<yyyymmdd-hhmmss>.log.
setlocal EnableExtensions EnableDelayedExpansion

rem The app name comes from the folder this script is installed into, climbing
rem one level from exec or scripts. INSTALLERS PUT THIS SCRIPT IN SCRIPTS
rem (1.43.20): with only exec climbed, the app was named "scripts", so FileDir's
rem finish page looked for scripts_JAWS.zip, found nothing, logged to
rem %LOCALAPPDATA%\scripts\logs, and installed no JAWS scripts at all.
for %%d in ("%~dp0.") do set "sApp=%%~nxd"
if /i "%sApp%"=="exec" for %%d in ("%~dp0..") do set "sApp=%%~nxd"
if /i "%sApp%"=="scripts" for %%d in ("%~dp0..") do set "sApp=%%~nxd"
set "sLogDir=%LOCALAPPDATA%\%sApp%\logs"
rem ONE LOG PER RUN (kit 1.79.2): <App>-<script>-<yyyymmdd-hhmmss>.log, as installCommon.cmd and
rem help\Logging.md have it, rather than one <App>-<script>-<yyyymmdd-hhmmss>.log that every run appended to.
for /f %%T in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd-HHmmss" 2^>nul') do set "sNow=%%T"
if not defined sNow set "sNow=unknown-time"
set "sLog=%sLogDir%\%sApp%-installScreenReaderSupport-%sNow%.log"
rem What the installer's Results box reports about the screen readers, one line
rem each; the installer clears it when Finish is pressed, before this runs.
set "sResult=%sLogDir%\%sApp%_screenReaders.txt"
if not exist "%sLogDir%" mkdir "%sLogDir%" >nul 2>&1

call :logLine "installScreenReaderSupport started %DATE% %TIME%"
call :logLine "Script: %~f0"
call :logLine "App: %sApp%"

rem STATE, FOR THE FINISH PAGE (1.43.43). "state jaws <file>" or "state nvda
rem <file>" writes one word to <file>: none (that reader is not on this
rem computer, or this app ships nothing for it), install (nothing of ours is
rem there), update (what is there differs from what ships), or reinstall (what
rem ships is what is there). The installer's box is worded from it, as every
rem other component's is. BOTH READERS THE SAME WAY (kit 1.98.0): the shipped
rem sources' fingerprint, every file's name and bytes, is compared with each
rem installed copy's, in every place the reader looks; a JAWS copy also needs
rem its compiled .jsb. The PowerShell part below explains the four steps.
if /i "%~1"=="state" (
  set "HS_MODE=state"
  set "HS_READER=%~2"
  set "HS_STATE=%~3"
  set "HS_APP=%sApp%"
  set "HS_ZIP=%~dp0%sApp%_JAWS.zip"
  set "HS_ADDON=%~dp0%sApp%.nvda-addon"
  set "HS_LOG=%sLog%"
  set "HS_RESULT=%sResult%"
  powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=[IO.File]::ReadAllText('%~f0'); iex $s.Substring($s.LastIndexOf('#<jawsScripts>') + 14)"
  exit /b 0
)
set "HS_MODE=install"

set "iInstalled=0"

rem ONE READER AT A TIME, WHEN ASKED (1.43.20). An installer gives the JAWS
rem scripts and the NVDA add-on a box each, so each box runs this with "jaws"
rem or "nvda"; with neither, both are installed, as before.
set "bJaws=1"
set "bNvda=1"
for %%a in (%*) do (
  if /i "%%~a"=="jaws" set "bNvda=0"
  if /i "%%~a"=="nvda" set "bJaws=0"
)
call :logLine "JAWS: %bJaws%  NVDA: %bNvda%"
if "%bJaws%"=="0" goto :nvda

rem ---- JAWS ----------------------------------------------------------------
rem COMPILED HERE OR NOT INSTALLED AT ALL (HomerDev 1.43.39). For each JAWS
rem version installed or in the user's settings, and each of its language
rem folders (Settings\enu, Settings\deu and so on; enu is made when a version
rem has none yet, kit 1.98.0), the sources in <App>_JAWS.zip are unpacked there
rem and compiled by THAT version's scompile.exe. JAWS need not be running.
rem No compiled .jsb is shipped: one built by another version may not suit
rem this one. If a version has no compiler, or any source fails to compile,
rem every file this run put in that folder is removed, any file it replaced is
rem put back as it was, and the failure is written for the installer's Results
rem box. The work is in PowerShell, at the end of this file after the marker
rem line, so it needs no second file beside this one.
if not exist "%~dp0%sApp%_JAWS.zip" goto :nvda
call :logLine "Found %sApp%_JAWS.zip"
echo Installing the JAWS scripts, compiled for each JAWS version found.
set "HS_APP=%sApp%"
set "HS_ZIP=%~dp0%sApp%_JAWS.zip"
set "HS_LOG=%sLog%"
set "HS_RESULT=%sResult%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=[IO.File]::ReadAllText('%~f0'); iex $s.Substring($s.LastIndexOf('#<jawsScripts>') + 14)"
call :logLine "JAWS scripts step exit code: %ERRORLEVEL%"
set /a iInstalled=iInstalled+1
:nvda
rem ---- NVDA ----------------------------------------------------------------
if "%bNvda%"=="0" goto :done
if not exist "%~dp0%sApp%.nvda-addon" goto :done
call :logLine "Found %sApp%.nvda-addon"
rem INSTALLED WITHOUT STARTING NVDA (HomerDev 1.43.46). Starting nvda.exe
rem --install-add-on brought a second screen reader up talking over JAWS, and
rem for EdSharp on 29 September stopped on an invalid command line parameter.
rem NVDA's own installer unpacks the add-on into %APPDATA%\nvda\addons and, at
rem its next start, moves it into place under its own name; an add-on already
rem there under its name is simply loaded. The PowerShell part puts it there
rem directly, moving any old copy aside to <name>.delete, which NVDA skips --
rem in every place NVDA looks (kit 1.98.0): %APPDATA%\nvda, made when NVDA is
rem installed but not yet run, and the userConfig of any running portable NVDA.
echo Installing the NVDA add-on; NVDA loads it when it next starts.
set "HS_MODE=nvda"
set "HS_APP=%sApp%"
set "HS_ADDON=%~dp0%sApp%.nvda-addon"
set "HS_LOG=%sLog%"
set "HS_RESULT=%sResult%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=[IO.File]::ReadAllText('%~f0'); iex $s.Substring($s.LastIndexOf('#<jawsScripts>') + 14)"
call :logLine "NVDA add-on step exit code: %ERRORLEVEL%"
set /a iInstalled=iInstalled+1

:done
if "%iInstalled%"=="0" (
  echo No screen reader support was installed.
) else (
  echo Screen reader support installed.
)
call :logLine "Screen reader support installed: %iInstalled% item(s)"
call :logLine "installScreenReaderSupport finished %DATE% %TIME%"
endlocal
exit /b 0

:logLine
echo %~1>> "%sLog%"
goto :eof

rem The PowerShell below is read and run by the JAWS step above; cmd never
rem reaches it, since every path through this file ends before it.
#<jawsScripts>
$sApp = $env:HS_APP; $sZip = $env:HS_ZIP; $sAddon = $env:HS_ADDON; $sLog = $env:HS_LOG; $sResult = $env:HS_RESULT
function logLine([string] $sText) { Add-Content -LiteralPath $sLog -Value $sText -Encoding UTF8 }
function resultLine([string] $sText) { Add-Content -LiteralPath $sResult -Value $sText -Encoding UTF8 }
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ONE DESIGN FOR BOTH READERS (kit 1.98.0, 2 November 2026). JAWS and NVDA are
# handled in the same four steps, so a fix to one is a fix to both:
#   1. DISCOVER every place the reader looks for scripts, including a reader that
#      is installed but has never been started, without the reader running.
#      JAWS: every installed version (its scompile.exe under Program Files) and
#      every language folder of its user settings, Settings\enu, Settings\deu
#      and so on; enu is made when a version has none yet. NVDA: its user
#      settings, %APPDATA%\nvda, made when NVDA is installed but not yet run,
#      and the userConfig folder of any portable NVDA that is running.
#   2. FINGERPRINT the shipped sources and each installed copy the same way:
#      every file's name and bytes, in name order, leaving out what the reader
#      makes for itself (a compiled .jsb, Python caches). The state for the
#      installer's box comes from comparing them: install (none there), update
#      (any differs), or reinstall (every copy current). Nothing is recorded
#      on the side: what is installed is what is compared.
#   3. INSTALL into every place found, never through the reader: JAWS sources
#      are unpacked and compiled by that version's own compiler, with every file
#      put back as it was if any source fails to compile; the NVDA add-on is
#      unpacked into the add-ons folder, the old copy moved aside first.
#   4. OPERATE: JAWS uses new scripts the next time the app comes to the front;
#      NVDA loads the add-on the next time it starts. Each place gets one result
#      line, worded the same way for both.

# ---- 2. Fingerprints ----------------------------------------------------------
function isMadeByReader([string] $sName) {
    return ($sName -like "*.jsb" -or $sName -like "*.pyc" -or $sName -match '(^|/)__pycache__/')
}
function printOfFiles($dFiles) {
    $oBuffer = New-Object IO.MemoryStream
    $lsNames = [string[]]@($dFiles.Keys)
    [Array]::Sort($lsNames, [StringComparer]::Ordinal)
    foreach ($sName in $lsNames) {
        $aName = [Text.Encoding]::UTF8.GetBytes($sName)
        $oBuffer.Write($aName, 0, $aName.Length)
        $aBytes = & $dFiles[$sName]
        # A VERSION ALONE IS NOT A CHANGE (kit 1.172.0, 9 November 2026). An app's build may stamp its own version
        # into the add-on's manifest.ini -- EdSharp's does -- so every new release read as Update though no script
        # had changed. The fingerprint now leaves out manifest.ini's version line: Update means the scripts changed,
        # Reinstall that they are the same.
        if ($sName -like "*manifest.ini") {
            $sText = [Text.Encoding]::UTF8.GetString($aBytes)
            $sText = [regex]::Replace($sText, '(?m)^[ \t]*version[ \t]*=.*(\r?\n)?', '')
            $aBytes = [Text.Encoding]::UTF8.GetBytes($sText)
        }
        $oBuffer.Write($aBytes, 0, $aBytes.Length)
    }
    $oHash = [Security.Cryptography.SHA256]::Create()
    return (($oHash.ComputeHash($oBuffer.ToArray()) | ForEach-Object { $_.ToString("x2") }) -join "")
}
# The shipped package: <App>_JAWS.zip or <App>.nvda-addon.
function archivePrint([string] $sPath) {
    $oArchive = [IO.Compression.ZipFile]::OpenRead($sPath)
    try {
        $dFiles = @{}
        foreach ($oEntry in @($oArchive.Entries | Where-Object { $_.Name })) {
            $sName = $oEntry.FullName.Replace("\", "/").ToLowerInvariant()
            if (isMadeByReader $sName) { continue }
            $oThis = $oEntry
            $dFiles[$sName] = { $oMemory = New-Object IO.MemoryStream; $oStream = $oThis.Open(); $oStream.CopyTo($oMemory); $oStream.Dispose(); ,$oMemory.ToArray() }.GetNewClosure()
        }
        return (printOfFiles $dFiles)
    } finally { $oArchive.Dispose() }
}
function archiveNames([string] $sPath) {
    $oArchive = [IO.Compression.ZipFile]::OpenRead($sPath)
    try { return @($oArchive.Entries | Where-Object { $_.Name } | ForEach-Object { $_.FullName.Replace("\", "/") } | Where-Object { -not (isMadeByReader $_.ToLowerInvariant()) }) }
    finally { $oArchive.Dispose() }
}
# An installed copy: the named files in a JAWS settings folder, which holds many
# other files too, or a whole NVDA add-on folder. "" when a file is missing.
function installedPrint([string] $sFolder, $lsNames) {
    $dFiles = @{}
    foreach ($sName in $lsNames) {
        $sFull = Join-Path $sFolder ($sName.Replace("/", "\"))
        if (-not (Test-Path -LiteralPath $sFull)) { return "" }
        $sPath = $sFull
        $dFiles[$sName.ToLowerInvariant()] = { ,[IO.File]::ReadAllBytes($sPath) }.GetNewClosure()
    }
    return (printOfFiles $dFiles)
}
function folderNames([string] $sFolder) {
    $sRoot = (Resolve-Path -LiteralPath $sFolder).Path.TrimEnd("\") + "\"
    return @(Get-ChildItem -LiteralPath $sFolder -Recurse -File -Force | ForEach-Object { $_.FullName.Substring($sRoot.Length).Replace("\", "/") } | Where-Object { -not (isMadeByReader $_.ToLowerInvariant()) })
}

# ---- 1. Discovery --------------------------------------------------------------
function jawsFolders() {
    $lsFound = @()
    $dCompilers = @{}
    foreach ($sProgram in @((Join-Path $env:ProgramFiles "Freedom Scientific\JAWS"), (Join-Path ${env:ProgramFiles(x86)} "Freedom Scientific\JAWS"))) {
        if (-not (Test-Path -LiteralPath $sProgram)) { continue }
        foreach ($oDir in @(Get-ChildItem -LiteralPath $sProgram -Directory)) {
            $sCompiler = Join-Path $oDir.FullName "scompile.exe"
            if ((Test-Path -LiteralPath $sCompiler) -and -not $dCompilers.ContainsKey($oDir.Name)) { $dCompilers[$oDir.Name] = $sCompiler }
        }
    }
    $sUserRoot = Join-Path $env:APPDATA "Freedom Scientific\JAWS"
    $lsVersions = @($dCompilers.Keys)
    if (Test-Path -LiteralPath $sUserRoot) { $lsVersions += @(Get-ChildItem -LiteralPath $sUserRoot -Directory | ForEach-Object { $_.Name }) }
    foreach ($sVersion in @($lsVersions | Sort-Object -Unique)) {
        $sSettingsRoot = Join-Path $sUserRoot "$sVersion\Settings"
        $lsLanguages = @()
        if (Test-Path -LiteralPath $sSettingsRoot) {
            $lsLanguages = @(Get-ChildItem -LiteralPath $sSettingsRoot -Directory | Where-Object { $_.Name -match '^[A-Za-z]{3}$' } | ForEach-Object { $_.Name })
        }
        if ($lsLanguages.Count -eq 0) {
            if (-not $dCompilers.ContainsKey($sVersion)) { continue }
            $lsLanguages = @("enu")
        }
        foreach ($sLanguage in $lsLanguages) {
            $lsFound += [pscustomobject]@{ Version = $sVersion; Language = $sLanguage; Settings = (Join-Path $sSettingsRoot $sLanguage); Compiler = [string]$dCompilers[$sVersion] }
        }
    }
    return $lsFound
}
function nvdaFolders() {
    $lsFound = @()
    $bInstalled = (Test-Path -LiteralPath (Join-Path $env:ProgramFiles "NVDA\nvda.exe")) -or (Test-Path -LiteralPath (Join-Path ${env:ProgramFiles(x86)} "NVDA\nvda.exe"))
    $sUser = Join-Path $env:APPDATA "nvda"
    if ($bInstalled -or (Test-Path -LiteralPath $sUser)) { $lsFound += $sUser }
    foreach ($oProcess in @(Get-Process -Name nvda -ErrorAction SilentlyContinue)) {
        $sExe = ""
        try { $sExe = [string]$oProcess.Path } catch { }
        if (-not $sExe) { continue }
        $sPortable = Join-Path (Split-Path -Parent $sExe) "userConfig"
        if ((Test-Path -LiteralPath $sPortable) -and ($lsFound -notcontains $sPortable)) { $lsFound += $sPortable }
    }
    return $lsFound
}

# An add-on's manifest, as name and value pairs.
function readManifest([string] $sText) {
    $dValues = @{}
    foreach ($sLine in ($sText -split "`r?`n")) {
        if ($sLine -match '^\s*(\w+)\s*=\s*"?(.*?)"?\s*$') { $dValues[$Matches[1]] = $Matches[2] }
    }
    return $dValues
}
function addonManifest([string] $sPath) {
    $oArchive = [IO.Compression.ZipFile]::OpenRead($sPath)
    try {
        $oEntry = $oArchive.GetEntry("manifest.ini")
        if (-not $oEntry) { return @{} }
        $oReader = New-Object IO.StreamReader($oEntry.Open())
        $sText = $oReader.ReadToEnd(); $oReader.Dispose()
        return (readManifest $sText)
    } finally { $oArchive.Dispose() }
}
# Where an add-on is installed in one NVDA settings folder; NVDA keeps one it
# has just accepted as <name>.pendingInstall until it restarts.
function addonFolder([string] $sConfig, [string] $sName) {
    foreach ($sLeaf in @($sName, "$sName.pendingInstall")) {
        $sFolder = Join-Path $sConfig "addons\$sLeaf"
        if (Test-Path -LiteralPath (Join-Path $sFolder "manifest.ini")) { return $sFolder }
    }
    return ""
}
# What an earlier installer kept on the side, before the installed copies were compared.
function removeOldRecords([string] $sSettings) {
    $sMarker = Join-Path $sSettings "$sApp.scripts.fingerprint"
    if (Test-Path -LiteralPath $sMarker) { try { Remove-Item -LiteralPath $sMarker -Force } catch { } }
    $sPrints = Join-Path $env:LOCALAPPDATA "$sApp\jawsScripts.inix"
    if (Test-Path -LiteralPath $sPrints) { try { Remove-Item -LiteralPath $sPrints -Force; logLine "Removed ${sPrints}: the installed scripts themselves are compared now" } catch { } }
}

# ---- 2. State, for the installer's box -------------------------------------------
if ($env:HS_MODE -eq "state") {
    $sState = "none"
    try {
        $iPlaces = 0; $iAny = 0; $iSame = 0
        if ($env:HS_READER -eq "jaws" -and (Test-Path -LiteralPath $sZip)) {
            $sShipped = archivePrint $sZip
            $lsNames = archiveNames $sZip
            foreach ($oFolder in @(jawsFolders | Where-Object { $_.Compiler })) {
                $iPlaces += 1
                $sHave = installedPrint $oFolder.Settings $lsNames
                $bCompiled = $true
                foreach ($sName in @($lsNames | Where-Object { $_ -like "*.jss" })) {
                    if (-not (Test-Path -LiteralPath (Join-Path $oFolder.Settings ([IO.Path]::ChangeExtension($sName.Replace("/", "\"), ".jsb"))))) { $bCompiled = $false }
                }
                $sWhat = if (-not $sHave) { "not installed" } elseif ($sHave -eq $sShipped -and $bCompiled) { "current" } else { "different" }
                if ($sHave) { $iAny += 1 }
                if ($sWhat -eq "current") { $iSame += 1 }
                logLine "state jaws: JAWS $($oFolder.Version) $($oFolder.Language): $sWhat"
            }
        } elseif ($env:HS_READER -eq "nvda" -and (Test-Path -LiteralPath $sAddon)) {
            $dShipped = addonManifest $sAddon
            $sShipped = archivePrint $sAddon
            foreach ($sConfig in @(nvdaFolders)) {
                $iPlaces += 1
                $sFolder = if ($dShipped["name"]) { addonFolder $sConfig $dShipped["name"] } else { "" }
                $sWhat = "not installed"
                if ($sFolder) {
                    $iAny += 1
                    $sHave = installedPrint $sFolder (folderNames $sFolder)
                    $sVersion = (readManifest ([IO.File]::ReadAllText((Join-Path $sFolder "manifest.ini"))))["version"]
                    $sWhat = if ($sHave -eq $sShipped) { "current" } else { "different" }
                    if ($sWhat -eq "current") { $iSame += 1 }
                    $sWhat += " (installed version $sVersion, shipped $($dShipped['version']))"
                }
                logLine "state nvda: $sConfig`: $sWhat"
            }
        }
        if ($iPlaces -gt 0) {
            $sState = if ($iAny -eq 0) { "install" } elseif ($iSame -eq $iPlaces) { "reinstall" } else { "update" }
        }
        logLine "state $($env:HS_READER): $iPlaces place(s), $iAny with ours, $iSame current; $sState"
    } catch {
        logLine ("state $($env:HS_READER): ERROR " + $_.Exception.Message + "; offered as install")
        $sState = "install"
    }
    Set-Content -LiteralPath $env:HS_STATE -Value $sState -Encoding ASCII
    exit 0
}

# ---- 3. Install: NVDA ------------------------------------------------------------
if ($env:HS_MODE -eq "nvda") {
    try {
        $dManifest = addonManifest $sAddon
        $sName = $dManifest["name"]
        if (-not $sName) { throw "the add-on's manifest.ini names no add-on" }
        logLine "NVDA add-on: name=$sName version=$($dManifest['version']) minimumNVDAVersion=$($dManifest['minimumNVDAVersion']) lastTestedNVDAVersion=$($dManifest['lastTestedNVDAVersion'])"
        $oArchive = [IO.Compression.ZipFile]::OpenRead($sAddon)
        try { if ($oArchive.GetEntry("installTasks.py")) { logLine "WARNING: the add-on has installTasks.py, whose onInstall step does not run in a direct install" } } finally { $oArchive.Dispose() }
        $bRunning = [bool](Get-Process -Name nvda -ErrorAction SilentlyContinue)
        logLine "NVDA running: $bRunning"
        $lsConfigs = @(nvdaFolders)
        if ($lsConfigs.Count -eq 0) {
            logLine "NVDA add-on: no NVDA found, installed or running"
            resultLine "NVDA add-on: not installed, because NVDA was not found"
        }
        foreach ($sConfig in $lsConfigs) {
            try {
                $sAddonsDir = Join-Path $sConfig "addons"
                New-Item -ItemType Directory -Force -Path $sAddonsDir | Out-Null
                $sTarget = Join-Path $sAddonsDir $sName
                $sOld = Join-Path $sAddonsDir "$sName.delete"
                $sUnpack = Join-Path $env:TEMP ("$sApp-nvda-" + (Get-Date -Format "yyyyMMddHHmmssfff"))
                [IO.Compression.ZipFile]::ExtractToDirectory($sAddon, $sUnpack)
                if (Test-Path -LiteralPath $sOld) { Remove-Item -LiteralPath $sOld -Recurse -Force -ErrorAction SilentlyContinue }
                $bMovedOld = $false
                if (Test-Path -LiteralPath $sTarget) { Move-Item -LiteralPath $sTarget -Destination $sOld; $bMovedOld = $true }
                try { Move-Item -LiteralPath $sUnpack -Destination $sTarget }
                catch { if ($bMovedOld) { Move-Item -LiteralPath $sOld -Destination $sTarget }; throw }
                if ($bMovedOld) { Remove-Item -LiteralPath $sOld -Recurse -Force -ErrorAction SilentlyContinue }
                logLine "NVDA add-on: installed at $sTarget"
                resultLine $(if ($bRunning) { "NVDA add-on ($sConfig): installed. NVDA uses it the next time it starts, so restart NVDA" } else { "NVDA add-on ($sConfig): installed. NVDA uses it the next time it starts" })
            } catch {
                logLine ("NVDA add-on ($sConfig): ERROR " + $_.Exception.Message)
                resultLine "NVDA add-on ($sConfig): NOT installed -- $($_.Exception.Message)"
            }
        }
    } catch {
        logLine ("NVDA add-on: ERROR " + $_.Exception.Message)
        resultLine "NVDA add-on: NOT installed -- $($_.Exception.Message)"
    }
    exit 0
}

# ---- 3. Install: JAWS ------------------------------------------------------------
$lsFolders = @(jawsFolders)
if ($lsFolders.Count -eq 0) {
    logLine "JAWS: no version found, installed or in user settings"
    resultLine "JAWS scripts: not installed, because JAWS was not found"
    exit 0
}
$oZip = [IO.Compression.ZipFile]::OpenRead($sZip)
try {
    foreach ($oFolder in $lsFolders) {
        $sVersion = $oFolder.Version; $sSettings = $oFolder.Settings; $sCompiler = $oFolder.Compiler
        $sPlace = "JAWS $sVersion $($oFolder.Language)"
        if (-not $sCompiler) {
            logLine "${sPlace}: no scompile.exe in its program folder; nothing installed here"
            resultLine "$sPlace scripts: not installed, because its script compiler was not found"
            continue
        }
        New-Item -ItemType Directory -Force -Path $sSettings | Out-Null
        removeOldRecords $sSettings
        # Every file this run will write, including the .jsb compiled from each
        # .jss, so that a failure can undo exactly what was done.
        $lsTargets = @()
        foreach ($oEntry in $oZip.Entries) {
            if (-not $oEntry.Name) { continue }
            $lsTargets += (Join-Path $sSettings $oEntry.FullName)
            if ($oEntry.Name -like "*.jss") { $lsTargets += (Join-Path $sSettings ([IO.Path]::ChangeExtension($oEntry.FullName, ".jsb"))) }
        }
        $sBackup = Join-Path $env:TEMP ("$sApp-jaws-$sVersion-$($oFolder.Language)-" + (Get-Date -Format "yyyyMMddHHmmssfff"))
        $dReplaced = @{}
        foreach ($sTarget in $lsTargets) {
            if (Test-Path -LiteralPath $sTarget) {
                $sCopy = Join-Path $sBackup ([IO.Path]::GetFileName($sTarget) + "." + $dReplaced.Count)
                New-Item -ItemType Directory -Force -Path $sBackup | Out-Null
                Copy-Item -LiteralPath $sTarget -Destination $sCopy -Force
                $dReplaced[$sTarget] = $sCopy
            }
        }
        $bFailed = $false
        $lsFailed = @()
        try {
            foreach ($oEntry in $oZip.Entries) {
                if (-not $oEntry.Name) { continue }
                $sTarget = Join-Path $sSettings $oEntry.FullName
                New-Item -ItemType Directory -Force -Path (Split-Path -Parent $sTarget) | Out-Null
                [IO.Compression.ZipFileExtensions]::ExtractToFile($oEntry, $sTarget, $true)
            }
            foreach ($oEntry in $oZip.Entries) {
                if ($oEntry.Name -notlike "*.jss") { continue }
                $sSource = Join-Path $sSettings $oEntry.FullName
                $sBinary = [IO.Path]::ChangeExtension($sSource, ".jsb")
                if (Test-Path -LiteralPath $sBinary) { Remove-Item -LiteralPath $sBinary -Force }
                $sOutput = & $sCompiler $sSource 2>&1 | Out-String
                $iExit = $LASTEXITCODE
                logLine ("${sPlace}: run exit=$iExit cmd=""$sCompiler $sSource""")
                if ($sOutput.Trim()) { logLine ("| " + ($sOutput.Trim() -replace "`r?`n", "`r`n| ")) }
                if ($iExit -ne 0 -or -not (Test-Path -LiteralPath $sBinary)) { $bFailed = $true; $lsFailed += $oEntry.Name }
            }
        } catch {
            $bFailed = $true
            logLine ("${sPlace}: ERROR " + $_.Exception.Message)
        }
        if ($bFailed) {
            foreach ($sTarget in $lsTargets) {
                if (Test-Path -LiteralPath $sTarget) { Remove-Item -LiteralPath $sTarget -Force }
                if ($dReplaced.ContainsKey($sTarget)) { Copy-Item -LiteralPath $dReplaced[$sTarget] -Destination $sTarget -Force }
            }
            logLine "${sPlace}: ERROR the scripts did not compile; every file this run placed was removed and $($dReplaced.Count) earlier file(s) put back"
            $sWhich = if ($lsFailed.Count) { " (" + ($lsFailed -join ", ") + ")" } else { "" }
            resultLine "$sPlace scripts: NOT installed -- they did not compile$sWhich; nothing was left behind, and the setup log has the compiler's words"
        } else {
            logLine "${sPlace}: scripts installed and compiled"
            resultLine "$sPlace scripts: installed. JAWS uses them the next time $sApp comes to the front"
        }
        if (Test-Path -LiteralPath $sBackup) { Remove-Item -LiteralPath $sBackup -Recurse -Force }
    }
} finally {
    $oZip.Dispose()
}
exit 0
