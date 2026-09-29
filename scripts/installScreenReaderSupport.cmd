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
rem NOTHING PAUSES. LOG: %LOCALAPPDATA%\<App>\logs\<App>_setup.log.
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
set "sLog=%sLogDir%\%sApp%_setup.log"
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
rem other component's is. JAWS: a fingerprint of the script sources in
rem <App>_JAWS.zip, compared with the one written into each JAWS version's
rem settings folder when its scripts last compiled there. NVDA: the add-on's
rem version in <App>.nvda-addon, compared with the installed add-on's.
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
rem version in the user's settings, the sources in <App>_JAWS.zip are unpacked
rem into its Settings\enu and compiled there by THAT version's scompile.exe.
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
rem directly, moving any old copy aside to <name>.delete, which NVDA skips.
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
$sApp = $env:HS_APP; $sZip = $env:HS_ZIP; $sLog = $env:HS_LOG; $sResult = $env:HS_RESULT
function logLine([string] $sText) { Add-Content -LiteralPath $sLog -Value $sText -Encoding UTF8 }
function resultLine([string] $sText) { Add-Content -LiteralPath $sResult -Value $sText -Encoding UTF8 }
Add-Type -AssemblyName System.IO.Compression.FileSystem

# The JAWS sources' fingerprint: every entry's name and bytes, in name order,
# compiled files left out. A zip's own bytes change with every build; its
# contents change only when a script does.
function zipFingerprint([string] $sPath) {
    $oArchive = [IO.Compression.ZipFile]::OpenRead($sPath)
    try {
        $oBuffer = New-Object IO.MemoryStream
        foreach ($oEntry in @($oArchive.Entries | Where-Object { $_.Name -and $_.Name -notlike "*.jsb" } | Sort-Object FullName)) {
            $aName = [Text.Encoding]::UTF8.GetBytes($oEntry.FullName.ToLowerInvariant())
            $oBuffer.Write($aName, 0, $aName.Length)
            $oStream = $oEntry.Open(); $oStream.CopyTo($oBuffer); $oStream.Dispose()
        }
        $oHash = [Security.Cryptography.SHA256]::Create()
        return (($oHash.ComputeHash($oBuffer.ToArray()) | ForEach-Object { $_.ToString("x2") }) -join "")
    } finally { $oArchive.Dispose() }
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

$sMarkerName = "$sApp.scripts.fingerprint"

if ($env:HS_MODE -eq "nvda") {
    $sAddon = $env:HS_ADDON
    try {
        $sNvdaConfig = Join-Path $env:APPDATA "nvda"
        if (-not (Test-Path -LiteralPath $sNvdaConfig)) { throw "NVDA's settings folder $sNvdaConfig was not found" }
        $dManifest = addonManifest $sAddon
        $sName = $dManifest["name"]
        if (-not $sName) { throw "the add-on's manifest.ini names no add-on" }
        logLine "NVDA add-on: name=$sName version=$($dManifest['version']) minimumNVDAVersion=$($dManifest['minimumNVDAVersion']) lastTestedNVDAVersion=$($dManifest['lastTestedNVDAVersion'])"
        $oArchive = [IO.Compression.ZipFile]::OpenRead($sAddon)
        try { if ($oArchive.GetEntry("installTasks.py")) { logLine "WARNING: the add-on has installTasks.py, whose onInstall step does not run in a direct install" } } finally { $oArchive.Dispose() }
        $bRunning = [bool](Get-Process -Name nvda -ErrorAction SilentlyContinue)
        logLine "NVDA running: $bRunning"
        $sAddonsDir = Join-Path $sNvdaConfig "addons"
        New-Item -ItemType Directory -Force -Path $sAddonsDir | Out-Null
        $sTarget = Join-Path $sAddonsDir $sName
        $sOld = Join-Path $sAddonsDir "$sName.delete"
        $sUnpack = Join-Path $env:TEMP ("$sApp-nvda-" + (Get-Date -Format "yyyyMMddHHmmss"))
        [IO.Compression.ZipFile]::ExtractToDirectory($sAddon, $sUnpack)
        if (Test-Path -LiteralPath $sOld) { Remove-Item -LiteralPath $sOld -Recurse -Force -ErrorAction SilentlyContinue }
        $bMovedOld = $false
        if (Test-Path -LiteralPath $sTarget) { Move-Item -LiteralPath $sTarget -Destination $sOld; $bMovedOld = $true }
        try { Move-Item -LiteralPath $sUnpack -Destination $sTarget }
        catch { if ($bMovedOld) { Move-Item -LiteralPath $sOld -Destination $sTarget }; throw }
        if ($bMovedOld) { Remove-Item -LiteralPath $sOld -Recurse -Force -ErrorAction SilentlyContinue }
        logLine "NVDA add-on: installed at $sTarget"
        resultLine $(if ($bRunning) { "NVDA add-on: installed. Restart NVDA to use it" } else { "NVDA add-on: installed. NVDA loads it when it next starts" })
    } catch {
        logLine ("NVDA add-on: ERROR " + $_.Exception.Message)
        resultLine "NVDA add-on: NOT installed -- $($_.Exception.Message)"
    }
    exit 0
}

if ($env:HS_MODE -eq "state") {
    $sState = "none"
    try {
        if ($env:HS_READER -eq "jaws") {
            $sRoot = Join-Path $env:APPDATA "Freedom Scientific\JAWS"
            $lsSettings = @()
            if ((Test-Path -LiteralPath $sZip) -and (Test-Path -LiteralPath $sRoot)) {
                $lsSettings = @(Get-ChildItem -LiteralPath $sRoot -Directory | ForEach-Object { Join-Path $_.FullName "Settings\enu" } |
                                Where-Object { Test-Path -LiteralPath $_ })
            }
            if ($lsSettings.Count -gt 0) {
                $sPrint = zipFingerprint $sZip
                $iSame = 0; $iAny = 0
                foreach ($sSettings in $lsSettings) {
                    $sMarker = Join-Path $sSettings $sMarkerName
                    if (Test-Path -LiteralPath $sMarker) {
                        $iAny += 1
                        if (([IO.File]::ReadAllText($sMarker)).Trim() -eq $sPrint) { $iSame += 1 }
                    } elseif (Test-Path -LiteralPath (Join-Path $sSettings "$sApp.jsb")) {
                        # Our scripts from before fingerprints were kept: there,
                        # but not known to be current.
                        $iAny += 1
                    }
                }
                if ($iAny -eq 0) { $sState = "install" }
                elseif ($iSame -eq $lsSettings.Count) { $sState = "reinstall" }
                else { $sState = "update" }
                logLine "state jaws: $($lsSettings.Count) JAWS version(s), $iAny with our scripts, $iSame current; $sState"
            } else { logLine "state jaws: none (no JAWS settings folder, or no $sApp`_JAWS.zip)" }
        } elseif ($env:HS_READER -eq "nvda") {
            $sAddon = $env:HS_ADDON
            $bNvda = (Test-Path -LiteralPath (Join-Path $env:APPDATA "nvda")) -or
                     (Test-Path -LiteralPath (Join-Path ${env:ProgramFiles(x86)} "NVDA\nvda.exe")) -or
                     (Test-Path -LiteralPath (Join-Path $env:ProgramFiles "NVDA\nvda.exe"))
            if ($bNvda -and (Test-Path -LiteralPath $sAddon)) {
                $dShipped = addonManifest $sAddon
                $sInstalled = Join-Path $env:APPDATA ("nvda\addons\" + $dShipped["name"] + "\manifest.ini")
                # NVDA keeps an add-on it has just accepted as <name>.pendingInstall
                # until it restarts; that counts as installed.
                $sPending = Join-Path $env:APPDATA ("nvda\addons\" + $dShipped["name"] + ".pendingInstall\manifest.ini")
                if (-not (Test-Path -LiteralPath $sInstalled) -and (Test-Path -LiteralPath $sPending)) { $sInstalled = $sPending }
                if (-not $dShipped["name"] -or -not (Test-Path -LiteralPath $sInstalled)) { $sState = "install" }
                else {
                    $dHave = readManifest ([IO.File]::ReadAllText($sInstalled))
                    $sState = if ($dHave["version"] -eq $dShipped["version"]) { "reinstall" } else { "update" }
                }
                $sHave = if (Test-Path -LiteralPath $sInstalled) { (readManifest ([IO.File]::ReadAllText($sInstalled)))["version"] } else { "none" }
                logLine "state nvda: add-on name=$($dShipped['name']) shipped version=$($dShipped['version']); installed version=$sHave; $sState"
                # NVDA's own record: an installed NVDA logs to %TEMP%\nvda.log and
                # keeps the previous session's as nvda-old.log; the lines about
                # add-ons, from both, go into this log (read with sharing, since
                # NVDA holds its log open).
                foreach ($sLogName in @("nvda-old.log", "nvda.log")) {
                    $sNvdaLog = Join-Path $env:TEMP $sLogName
                    if (-not (Test-Path -LiteralPath $sNvdaLog)) { logLine "  ${sLogName}: not found in $env:TEMP"; continue }
                    try {
                        $oStream = New-Object IO.FileStream($sNvdaLog, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
                        $oReader = New-Object IO.StreamReader($oStream)
                        $lsHits = @(($oReader.ReadToEnd() -split "`r?`n") | Where-Object { $_ -match [regex]::Escape($dShipped["name"]) -or $_ -match "addonHandler" })
                        $oReader.Dispose()
                        logLine "  ${sLogName}: $($lsHits.Count) line(s) about add-ons"
                        foreach ($sHit in @($lsHits | Select-Object -Last 15)) { logLine "  | $sHit" }
                    } catch { logLine "  ${sLogName}: could not be read: $($_.Exception.Message)" }
                }
            } else { logLine "state nvda: none (NVDA not found, or no $sApp.nvda-addon)" }
        }
    } catch {
        logLine ("state $($env:HS_READER): ERROR " + $_.Exception.Message + "; offered as install")
        $sState = "install"
    }
    Set-Content -LiteralPath $env:HS_STATE -Value $sState -Encoding ASCII
    exit 0
}
$sRoot = Join-Path $env:APPDATA "Freedom Scientific\JAWS"
if (-not (Test-Path -LiteralPath $sRoot)) {
    logLine "No JAWS settings folder at $sRoot"
    resultLine "JAWS scripts: not installed, because JAWS was not found"
    exit 0
}
$oZip = [IO.Compression.ZipFile]::OpenRead($sZip)
$sPrint = zipFingerprint $sZip
try {
    foreach ($oVersion in @(Get-ChildItem -LiteralPath $sRoot -Directory | Sort-Object Name)) {
        $sVersion = $oVersion.Name
        $sSettings = Join-Path $oVersion.FullName "Settings\enu"
        if (-not (Test-Path -LiteralPath $sSettings)) { continue }
        $sCompiler = @((Join-Path $env:ProgramFiles "Freedom Scientific\JAWS\$sVersion\scompile.exe"),
                       (Join-Path ${env:ProgramFiles(x86)} "Freedom Scientific\JAWS\$sVersion\scompile.exe")) |
                     Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
        if (-not $sCompiler) {
            logLine "JAWS ${sVersion}: no scompile.exe in its program folder; nothing installed for this version"
            resultLine "JAWS $sVersion scripts: not installed, because its script compiler was not found"
            continue
        }
        # Every file this run will write, including the .jsb compiled from each
        # .jss, so that a failure can undo exactly what was done.
        $lsTargets = @((Join-Path $sSettings $sMarkerName))
        foreach ($oEntry in $oZip.Entries) {
            if (-not $oEntry.Name) { continue }
            $lsTargets += (Join-Path $sSettings $oEntry.FullName)
            if ($oEntry.Name -like "*.jss") { $lsTargets += (Join-Path $sSettings ([IO.Path]::ChangeExtension($oEntry.FullName, ".jsb"))) }
        }
        $sBackup = Join-Path $env:TEMP ("$sApp-jaws-$sVersion-" + (Get-Date -Format "yyyyMMddHHmmss"))
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
                logLine ("JAWS ${sVersion}: run exit=$iExit cmd=""$sCompiler $sSource""")
                if ($sOutput.Trim()) { logLine ("| " + ($sOutput.Trim() -replace "`r?`n", "`r`n| ")) }
                if ($iExit -ne 0 -or -not (Test-Path -LiteralPath $sBinary)) { $bFailed = $true; $lsFailed += $oEntry.Name }
            }
        } catch {
            $bFailed = $true
            logLine ("JAWS ${sVersion}: ERROR " + $_.Exception.Message)
        }
        if ($bFailed) {
            foreach ($sTarget in $lsTargets) {
                if (Test-Path -LiteralPath $sTarget) { Remove-Item -LiteralPath $sTarget -Force }
                if ($dReplaced.ContainsKey($sTarget)) { Copy-Item -LiteralPath $dReplaced[$sTarget] -Destination $sTarget -Force }
            }
            logLine "JAWS ${sVersion}: ERROR the scripts did not compile; every file this run placed was removed and $($dReplaced.Count) earlier file(s) put back"
            $sWhich = if ($lsFailed.Count) { " (" + ($lsFailed -join ", ") + ")" } else { "" }
            resultLine "JAWS $sVersion scripts: NOT installed -- they did not compile$sWhich; nothing was left behind, and the setup log has the compiler's words"
        } else {
            # What makes the next installer's box say Reinstall rather than
            # Update: the fingerprint of the sources compiled here.
            Set-Content -LiteralPath (Join-Path $sSettings $sMarkerName) -Value $sPrint -Encoding ASCII
            logLine "JAWS ${sVersion}: scripts installed and compiled"
            resultLine "JAWS $sVersion scripts: installed and compiled"
        }
        if (Test-Path -LiteralPath $sBackup) { Remove-Item -LiteralPath $sBackup -Recurse -Force }
    }
} finally {
    $oZip.Dispose()
}
exit 0
