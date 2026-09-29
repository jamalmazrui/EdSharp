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
where nvda >nul 2>&1
if errorlevel 1 (
  if exist "%ProgramFiles(x86)%\NVDA\nvda.exe" (
    set "sNvda=%ProgramFiles(x86)%\NVDA\nvda.exe"
  ) else if exist "%ProgramFiles%\NVDA\nvda.exe" (
    set "sNvda=%ProgramFiles%\NVDA\nvda.exe"
  )
) else (
  set "sNvda=nvda"
)
if not defined sNvda (
  echo NVDA was not found on this computer, so its add-on was skipped.
  call :logLine "NVDA not found."
  >> "%sResult%" echo NVDA add-on: not installed, because NVDA was not found
  goto :done
)
echo Installing the NVDA add-on. NVDA will ask you to confirm.
call :logLine "Handing the add-on to NVDA: !sNvda!"
start "" "!sNvda!" --install-add-on="%~dp0%sApp%.nvda-addon"
>> "%sResult%" echo NVDA add-on: handed to NVDA to install, which asks you to confirm
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
$sRoot = Join-Path $env:APPDATA "Freedom Scientific\JAWS"
if (-not (Test-Path -LiteralPath $sRoot)) {
    logLine "No JAWS settings folder at $sRoot"
    resultLine "JAWS scripts: not installed, because JAWS was not found"
    exit 0
}
Add-Type -AssemblyName System.IO.Compression.FileSystem
$oZip = [IO.Compression.ZipFile]::OpenRead($sZip)
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
        $lsTargets = @()
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
            logLine "JAWS ${sVersion}: scripts installed and compiled"
            resultLine "JAWS $sVersion scripts: installed and compiled"
        }
        if (Test-Path -LiteralPath $sBackup) { Remove-Item -LiteralPath $sBackup -Recurse -Force }
    }
} finally {
    $oZip.Dispose()
}
exit 0
