@echo off
rem finish.cmd -- the last thing a Homer installer does: show ONE Results
rem box saying what this install actually did, and start the program only after
rem that box has been dismissed.
rem
rem WHY THIS EXISTS. Inno runs the final-page checkboxes in the order they are
rem listed, and the launch is listed last. If the launch simply started the
rem program, the program's own window would arrive on top of whatever the other
rem checkboxes were still saying, and the user would never read the outcome. So
rem the launch entry runs this instead. It reads the setup log, reports what
rem happened, waits, and then starts the program.
rem
rem   finish.cmd <App>.exe [arguments for the app]
rem
rem The Results box is a PowerShell message box, which a screen reader reads as
rem an ordinary dialog. It reports only what was done in this session: an action
rem that did not run is not mentioned, and a count always matches its noun.
rem
rem LOG: %LOCALAPPDATA%\<App>\logs\<App>_setup.log, the same log the install
rem scripts write to, since a program under Program Files cannot write beside
rem its own .exe. The app name comes from the folder this script is installed
rem into, so nothing here is hardcoded.
setlocal EnableExtensions EnableDelayedExpansion

rem The app name comes from the folder this script is installed into. That is
rem exec in an installed copy, so climb one level when it is.
for %%d in ("%~dp0.") do set "sApp=%%~nxd"
if /i "%sApp%"=="exec" for %%d in ("%~dp0..") do set "sApp=%%~nxd"
set "sExe=%~1"
if "%sExe%"=="" set "sExe=%sApp%.exe"
shift
set "sArgs=%1 %2 %3 %4 %5 %6 %7 %8 %9"

set "sLogDir=%LOCALAPPDATA%\%sApp%\logs"
set "sLog=%sLogDir%\%sApp%_setup.log"
if not exist "%sLogDir%" mkdir "%sLogDir%" >nul 2>&1

echo finish started %DATE% %TIME%>> "%sLog%"
echo Script: %~f0>> "%sLog%"
echo App: %sApp%, executable: %sExe%>> "%sLog%"

echo Preparing the summary.

rem The summary is built in PowerShell because it has a message box and batch
rem does not. Everything it reports comes from the log the install scripts just
rem wrote, so it describes what happened rather than what was offered.
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Continue';" ^
  "$sLog='%sLog%'; $sApp='%sApp%';" ^
  "$lsLines=@(); if (Test-Path -LiteralPath $sLog) { $lsLines = Get-Content -LiteralPath $sLog -Tail 400 }" ^
  "$lsDone=@();" ^
  "if ($lsLines -match 'Model .* is present') { $lsDone += 'The local AI model is installed and ready.' }" ^
  "elseif ($lsLines -match 'installOllama started') { $lsDone += 'Ollama was set up. See the log if the model is not ready.' }" ^
  "if ($lsLines -match 'Screen reader support installed') { $lsDone += 'The JAWS scripts and the NVDA add-on are installed.' }" ^
  "elseif ($lsLines -match 'installScreenReaderSupport started') { $lsDone += 'Screen reader support was set up. See the log for detail.' }" ^
  "$sBody = \"$sApp is installed.`r`n`r`n\";" ^
  "if ($lsDone.Count -eq 0) { $sBody += 'No optional component was installed in this session.' }" ^
  "else { $sBody += ($lsDone -join \"`r`n\") }" ^
  "$sBody += \"`r`n`r`nThe full record is in:`r`n$sLog`r`n`r`nChoose OK to start $sApp.\";" ^
  "Add-Type -AssemblyName System.Windows.Forms | Out-Null;" ^
  "[void][System.Windows.Forms.MessageBox]::Show($sBody, \"$sApp setup results\", 'OK', 'Information')"

echo finish: results box dismissed %DATE% %TIME%>> "%sLog%"

rem Only now does the program start, so its window cannot cover the summary.
if exist "%~dp0%sExe%" (
  echo Starting %sApp%.
  echo Launching %sExe% %sArgs%>> "%sLog%"
  start "" "%~dp0%sExe%" %sArgs%
) else (
  echo %sExe% was not found, so nothing was started.
  echo ERROR: %~dp0%sExe% not found.>> "%sLog%"
)

endlocal
exit /b 0
