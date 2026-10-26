@echo off
rem installCommon.cmd -- the common half of every Homer install<Component>.cmd.
rem
rem Called, not run: an install script does `call "%~dp0installCommon.cmd" setup`
rem and then uses what this sets. It exists because the same forty lines had
rem been copied into a dozen scripts across four apps and had drifted in every
rem one -- EdSharp logging to its own folder, HomerScribe logging beside itself
rem inside Program Files, two logging nowhere at all.
rem
rem WHAT IT SETS
rem   sApp      the application name, taken from the folder the caller sits in
rem   sLogDir   %LOCALAPPDATA%\<app>\logs
rem   log       a full path, <App>-<script>-yyyyMMdd-HHmmss.log
rem   sQuiet    1 when the installer is driving and nothing may wait for a key
rem
rem WHAT IT DOES NOT DO
rem   It does not hide the console. The console says briefly what was found and
rem   what was done; this file's log holds the detail, so a failure can be
rem   diagnosed from a file rather than from a window that has closed.

if /i "%~1"=="setup" goto :setup
if /i "%~1"=="log" goto :writeLog
if /i "%~1"=="say" goto :saySo
if /i "%~1"=="wingetHas" goto :wingetHas
if /i "%~1"=="wingetUpgrade" goto :wingetUpgrade
if /i "%~1"=="runLong" goto :runLong
exit /b 0

:setup
rem The app is the folder the CALLING script lives in, or its parent when that
rem folder is "scripts" -- so one file serves an app whatever it is called.
rem The caller's folder, with its trailing backslash removed so that ~nx
rem gives the folder's own name. If that name is scripts\ or exec\, the app is
rem one level up: resolve with ~f FIRST (which is what turns ".." into a real
rem path) and only then take the name. Taking ~nx of a path that still ends in
rem ".." would give ".." itself, and the log would land in %LOCALAPPDATA%\..\logs
rem -- a folder nobody would think to open.
if not defined sCallerDir set "sCallerDir=%~dp1"
set "sHere=%sCallerDir%"
if "%sHere:~-1%"=="\" set "sHere=%sHere:~0,-1%"
for %%d in ("%sHere%") do set "sApp=%%~nxd"
for %%d in ("%sHere%\..") do set "sParent=%%~fd"
for %%d in ("%sParent%") do set "sParentName=%%~nxd"
if /i "%sApp%"=="scripts" set "sApp=%sParentName%"
if /i "%sApp%"=="exec" set "sApp=%sParentName%"
if not defined sApp set "sApp=Homer"

set "sLogDir=%LOCALAPPDATA%\%sApp%\logs"
if not exist "%sLogDir%" mkdir "%sLogDir%" >nul 2>&1

rem WMIC IS GONE from Windows 11, so this used to fall through to zeros and
rem every log took the same name, overwriting the run before -- the very thing
rem the timestamp exists to prevent. PowerShell's Get-Date is always there.
for /f %%T in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd-HHmmss" 2^>nul') do set "sNow=%%T"
if not defined sNow set "sNow=unknown-time"
set "log=%sLogDir%\%sApp%-%sScript%-%sNow%.log"

rem QUIET means: nothing may wait for a key. It is set when the installer
rem passes "noPause" (Inno runs these with no console to read from, so a pause
rem would return at once or hang), or when HOMER_QUIET is in the environment.
rem Every script tests %noPause% before it pauses, and this is the one place
rem it is worked out.
set "sQuiet="
if defined HOMER_QUIET set "sQuiet=1"
for %%A in (%*) do if /i "%%A"=="noPause" set "sQuiet=1"
if defined sQuiet set "noPause=1"

call "%~f0" log "%sScript% started %DATE% %TIME%"
call "%~f0" log "Script: %sCallerDir%%sScript%.cmd"
call "%~f0" log "App: %sApp%"
call "%~f0" log "Log: %log%"
call "%~f0" log "Quiet: %sQuiet%"
exit /b 0

:writeLog
rem Always appends. A single > here would erase the header written above, and
rem the log would arrive missing exactly the part that says what machine it
rem ran on.
echo %~2>> "%log%" 2>nul
exit /b 0

:saySo
rem One line on the console, and the same line in the log, so the two agree.
echo %~2
echo %~2>> "%log%" 2>nul
exit /b 0

:wingetHas
rem wingetHas <id>: exit 0 when winget lists the package as installed, 1 when not (13 October 2026). It reads winget's
rem answer, not its exit code: winget list exits 0 on some machines while printing "No installed package found", and
rem trusting that made Node.js look installed under a package that was not there, so its update never happened.
winget list --id %~2 --exact --accept-source-agreements --disable-interactivity > "%TEMP%\homerWingetList.txt" 2>&1
findstr /i /c:"%~2" "%TEMP%\homerWingetList.txt" >nul 2>&1
set "iHomerHas=%ERRORLEVEL%"
del /q "%TEMP%\homerWingetList.txt" >nul 2>&1
if not "%iHomerHas%"=="0" exit /b 1
exit /b 0

:wingetUpgrade
rem wingetUpgrade <id> <name>: updates an installed package, says so, and tells the truth about the result. Setup has
rem already decided a newer version exists, so this updates rather than checking. Only winget's own "no newer version"
rem (-1978335189) counts as already current; any other non-zero result is a failure, reported as one, where every
rem installer used to say "already the newest winget offers" and succeed.
echo Updating %~3 to the newest release. Nothing is asked of you while it runs.
set "sHomerRunLong=winget upgrade --id %~2 --exact --silent --accept-source-agreements --accept-package-agreements --disable-interactivity"
call :runLongAs "%~3"
set "iHomerUpgrade=%ERRORLEVEL%"
>> "%log%" echo %date% %time% winget upgrade %~2 exit code %iHomerUpgrade%
if "%iHomerUpgrade%"=="0" (
  echo %~3 was updated.
  exit /b 0
)
if "%iHomerUpgrade%"=="-1978335189" (
  echo %~3 is already the newest release.
  exit /b 0
)
rem INSTALLED WITHOUT WINGET, UPDATED BY INSTALLING (15 October 2026). Setup now reads the version of a program it
rem finds by its file, so it can offer Update for a copy winget never installed, such as an mpv unpacked by hand.
rem winget upgrade answers such a copy with "no installed package found" (-1978335212), so the newest release is
rem installed in its default machine-wide folder instead, and the result is reported like any other.
rem Labels rather than a parenthesized block, so the exit code is read when it is set whether or not the calling
rem script turned on delayed expansion.
if not "%iHomerUpgrade%"=="-1978335212" exit /b 1
>> "%log%" echo %date% %time% winget knows no installed %~2, so the newest release is installed instead
set "sHomerRunLong=winget install --id %~2 --exact --silent --accept-source-agreements --accept-package-agreements --disable-interactivity"
call :runLongAs "%~3"
set "iHomerUpgrade=%ERRORLEVEL%"
>> "%log%" echo %date% %time% winget install %~2 exit code %iHomerUpgrade%
if not "%iHomerUpgrade%"=="0" exit /b 1
echo %~3 was updated.
exit /b 0

:runLong
rem runLong <label>: runs the command in sHomerRunLong, which the caller sets, for a step that takes a while.
call :runLongAs "%~2"
exit /b %ERRORLEVEL%

:runLongAs
rem A LONG STEP IS NEVER SILENT (HomerDev 1.93.7, the author's report of 25 October 2026: during long stretches the
rem installer evidently worked with nothing on the console). Install scripts sent each long step's whole output -- a
rem winget install, an ollama pull of several gigabytes, a pip install -- to the log alone, so after one line saying
rem what was starting, the console said nothing for minutes. runLong keeps the log's detail and gives the console a
rem short line now and then: "Pandoc: still working, 2 min, 45%" -- the percentage when the tool reports one -- first
rem after 20 seconds, then every 30, while the log gets the full output and a progress line each time. The PowerShell
rem that does it is at the end of this file, after the last exit, where cmd never runs it; it is read from here, so
rem every app that copies installCommon.cmd has it, with no change to any app's build.
set "sHomerRunLabel=%~1"
set "sHomerRunLog=%log%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s = [IO.File]::ReadAllText('%~f0'); $i = $s.LastIndexOf('#HOMER-RUNLONG'); Invoke-Expression $s.Substring($i)"
exit /b %ERRORLEVEL%

rem Nothing below this line is batch.
#HOMER-RUNLONG
$sCommand = $env:sHomerRunLong; $sLabel = $env:sHomerRunLabel; $sLog = $env:sHomerRunLog
if (-not $sLabel) { $sLabel = 'Working' }
function logLine($sText) { if ($sLog) { Add-Content -LiteralPath $sLog -Value ((Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + ' ' + $sText) -Encoding UTF8 } }
function spoken($nSeconds) { if ($nSeconds -lt 60) { return ([int]$nSeconds).ToString() + ' seconds' }; return ([int][Math]::Round($nSeconds / 60)).ToString() + ' min' }
function cleaned($sText) { $sText = [regex]::Replace($sText, "\x1B\[[0-9;?]*[A-Za-z]", ''); $l = @(); foreach ($sLine in ($sText -split "`n")) { $a = $sLine -split "`r"; $t = ($a | Where-Object { $_.Trim() -ne '' } | Select-Object -Last 1); if ($t) { $l += $t.TrimEnd() } }; return $l }
function readNew($sPath, [ref]$iAt) { try { $f = [IO.File]::Open($sPath, 'Open', 'Read', 'ReadWrite'); $f.Position = $iAt.Value; $r = New-Object IO.StreamReader($f); $t = $r.ReadToEnd(); $iAt.Value = $f.Position; $r.Close(); return $t } catch { return '' } }
$sOut = [IO.Path]::GetTempFileName(); $sErr = [IO.Path]::GetTempFileName()
logLine ('run: ' + $sCommand)
$dtStart = Get-Date
$p = Start-Process -FilePath $env:ComSpec -ArgumentList ('/d /s /c "' + $sCommand + '"') -RedirectStandardOutput $sOut -RedirectStandardError $sErr -NoNewWindow -PassThru
$null = $p.Handle
$iOut = 0; $iErr = 0; $nNext = 20; $sLast = ''
while (-not $p.HasExited) {
  Start-Sleep -Milliseconds 500
  $nSeconds = ((Get-Date) - $dtStart).TotalSeconds
  if ($nSeconds -ge $nNext) {
    $sNew = (readNew $sOut ([ref]$iOut)) + (readNew $sErr ([ref]$iErr))
    $lNew = @(cleaned $sNew); if ($lNew.Count -gt 0) { $sLast = $lNew[-1] }
    $oPercent = [regex]::Matches($sNew, '(\d{1,3}(?:\.\d)?)\s?%') | Select-Object -Last 1
    $sPercent = ''; if ($oPercent) { $sPercent = ', ' + [int][double]$oPercent.Groups[1].Value + '%' }
    Write-Host ($sLabel + ': still working, ' + (spoken $nSeconds) + $sPercent)
    logLine ('progress after ' + [int]$nSeconds + ' s' + $sPercent + ': ' + $sLast)
    $nNext += 30
  }
}
$p.WaitForExit()
$iCode = $p.ExitCode; if ($null -eq $iCode) { $iCode = 1 }
$sAll = (readNew $sOut ([ref]$iOut)) + (readNew $sErr ([ref]$iErr))
foreach ($sLine in @(cleaned ((Get-Content -Raw -LiteralPath $sOut -ErrorAction SilentlyContinue) + "`n" + (Get-Content -Raw -LiteralPath $sErr -ErrorAction SilentlyContinue)))) { logLine ('| ' + $sLine) }
logLine ('exit code ' + $iCode + ' after ' + [int]((Get-Date) - $dtStart).TotalSeconds + ' s: ' + $sCommand)
Remove-Item -LiteralPath $sOut, $sErr -ErrorAction SilentlyContinue
exit $iCode
