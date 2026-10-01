@echo off
rem installOllama.cmd -- install Ollama, the local AI service a Homer app uses.
rem
rem Run from the installer's final page, or on its own at any time.
rem
rem Two things learnt from a tester's first install:
rem
rem 1. Without --silent, winget hands over to Ollama's own installer, which puts
rem    up windows and may open a browser. They have to be closed by hand and it
rem    is not clear whether anything is still running.
rem
rem 2. After installing, ollama is NOT on the PATH of any console that was
rem    already open, including this one. Anything that looks for it with "where"
rem    will not find it, and will wrongly report that Ollama is not installed.
rem    So this script finds the program by looking where it is put.

setlocal enabledelayedexpansion

rem ---- the log -------------------------------------------------------------
rem The common half lives in the kit, so a fix reaches every install script in
rem every Homer app rather than the one being edited. It sets sApp, sLogDir,
rem log and sQuiet, and writes the environment header.
set "sScript=%~n0"
set "sCallerDir=%~dp0"
rem IF THE SHARED HALF IS MISSING, SAY SO. It was left out of the installer
rem once, and every script that calls it died at this line -- no message, no
rem log folder, nothing to diagnose from. A missing file must announce itself.
if not exist "%~dp0installCommon.cmd" (
  echo(
  echo installCommon.cmd is missing from %~dp0
  echo That file is part of this program. Reinstall, or copy it from the
  echo program's zip into this folder, and run this again.
  echo(
  if not defined noPause pause
  exit /b 1
)
call "%~dp0installCommon.cmd" setup "%~f0" %*




:afterLogSetup
rem THE FINISH PAGE HAS ALREADY DECIDED (1.43.51). It compared versions and
rem offered Install, Update or Reinstall; this script does what the box said,
rem and says that, rather than announcing a check of its own.
set "sVerb=update"
for %%a in (%*) do if /i "%%~a"=="reinstall" set "sVerb=reinstall"
for %%a in (%*) do if /i "%%~a"=="update" set "sVerb=update"
call "%~dp0installCommon.cmd" log "asked to: %sVerb% (when Ollama is already installed)"

call :findOllama
if defined ollamaExe goto :already

echo Downloading and installing Ollama. This is about 1 GB and takes a few minutes.
echo Nothing is asked of you while it runs.
echo(
winget install --id Ollama.Ollama --exact --silent --accept-source-agreements --accept-package-agreements
if errorlevel 1 goto :failed

echo(
echo Waiting for Ollama to start.
call :findOllama
if not defined ollamaExe goto :installedButLost

rem The service takes a few seconds to answer after the program appears.
for /l %%N in (1,1,30) do (
  "!ollamaExe!" list >nul 2>&1
  if not errorlevel 1 goto :ready
  timeout /t 2 /nobreak >nul
)

:ready
call :closeOllamaWindow
echo(
echo Ollama is installed and running. It starts with Windows from now on.
echo(
echo Installing the vision model next.
echo(
if exist "%~dp0installModels.cmd" call "%~dp0installModels.cmd" noPause
echo(
if not defined noPause pause
endlocal
exit /b 0

:already
rem IT IS HERE, SO THIS IS AN UPDATE OR A REINSTALL, NOT A NO-OP. The finish
rem page offered "Update Ollama from X to Y" and this script answered "already
rem installed" and exited in a second -- which is exactly what he noticed. A
rem winget upgrade is what an Update box means; when nothing is newer, winget
rem says so and returns at once, and that is the honest outcome.
echo(
if /i "%sVerb%"=="reinstall" goto :reinstall
echo Updating Ollama to the newest version: about 1 GB, and a few minutes.
echo Nothing is asked of you while it runs.
winget upgrade --id Ollama.Ollama --exact --silent --accept-source-agreements --accept-package-agreements >> "%log%" 2>&1
set "iUp=%ERRORLEVEL%"
call "%~dp0installCommon.cmd" log "winget upgrade exit code %iUp%"
if "%iUp%"=="0" echo Ollama was updated.
if not "%iUp%"=="0" echo Ollama was NOT updated; the log has winget's answer.
goto :afterChange

:reinstall
echo Reinstalling Ollama: about 1 GB, and a few minutes. Nothing is asked of you
echo while it runs.
winget install --id Ollama.Ollama --exact --force --silent --accept-source-agreements --accept-package-agreements >> "%log%" 2>&1
set "iUp=%ERRORLEVEL%"
call "%~dp0installCommon.cmd" log "winget install --force exit code %iUp%"
if "%iUp%"=="0" echo Ollama was reinstalled.
if not "%iUp%"=="0" echo Ollama was NOT reinstalled; the log has winget's answer.

:afterChange
call :closeOllamaWindow
call :findOllama
if not defined ollamaExe goto :installedButLost
echo Ollama is at !ollamaExe!
"!ollamaExe!" --version
echo(
echo Next, install the vision model with installModels.cmd.
echo(
if not defined noPause pause
endlocal
exit /b 0

:installedButLost
echo(
echo Ollama was installed, but this window cannot see it yet, which is normal:
echo a console keeps the PATH it started with. Open a NEW command window, or
echo run installModels.cmd from the Start menu folder, to install the model.
echo(
if not defined noPause pause
endlocal
exit /b 0

:failed
echo(
echo Ollama could not be installed automatically.
echo Download it instead from https://ollama.com/download
echo(
if not defined noPause pause
endlocal
exit /b 1

rem NO OLLAMA WINDOW LEFT OPEN (1.43.52). Ollama's own installer starts its
rem desktop app when it finishes, and the app opens a chat window. Homer apps
rem use Ollama behind the scenes and never its window, which only confuses
rem someone who did not ask for it. The window is closed politely -- the same
rem as pressing Alt+F4 -- so the app and the Ollama service keep running, as
rem every Homer app needs. Up to 20 seconds is allowed for it to appear.
:closeOllamaWindow
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$iClosed = 0; for ($i = 0; $i -lt 20; $i++) { foreach ($p in @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -like 'Ollama*' })) { if ($p.CloseMainWindow()) { $iClosed++; 'closed the Ollama window of ' + $p.ProcessName + ' (process ' + $p.Id + ')' } }; if ($iClosed -gt 0) { break }; Start-Sleep -Seconds 1 }; if ($iClosed -eq 0) { 'no Ollama window appeared' }" >> "%log%" 2>&1
goto :eof

:findOllama
rem The parenthesis in the variable name ProgramFiles(x86) must not appear
rem inside a parenthesised block, so it is copied out first.
set "progFiles86=%ProgramFiles(x86)%"
set "ollamaExe="
where ollama >nul 2>&1
if not errorlevel 1 set "ollamaExe=ollama"
if not defined ollamaExe if exist "%LOCALAPPDATA%\Programs\Ollama\ollama.exe" set "ollamaExe=%LOCALAPPDATA%\Programs\Ollama\ollama.exe"
if not defined ollamaExe if exist "%ProgramFiles%\Ollama\ollama.exe" set "ollamaExe=%ProgramFiles%\Ollama\ollama.exe"
if not defined ollamaExe if exist "!progFiles86!\Ollama\ollama.exe" set "ollamaExe=!progFiles86!\Ollama\ollama.exe"
if not defined ollamaExe if exist "%USERPROFILE%\AppData\Local\Programs\Ollama\ollama.exe" set "ollamaExe=%USERPROFILE%\AppData\Local\Programs\Ollama\ollama.exe"
goto :eof
