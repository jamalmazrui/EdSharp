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
if not exist "%~dp0homerInstall.cmd" (
  echo(
  echo homerInstall.cmd is missing from %~dp0
  echo That file is part of this program. Reinstall, or copy it from the
  echo program's zip into this folder, and run this again.
  echo(
  if not defined noPause pause
  exit /b 1
)
call "%~dp0homerInstall.cmd" setup "%~f0" %*




:afterLogSetup

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
echo Ollama is installed. Checking for a newer version. If there is one, it is
echo downloaded and installed now: about 1 GB, and a few minutes. Nothing is
echo asked of you while it runs.
winget upgrade --id Ollama.Ollama --exact --silent --accept-source-agreements --accept-package-agreements >> "%log%" 2>&1
set "iUp=%ERRORLEVEL%"
call "%~dp0homerInstall.cmd" log "winget upgrade exit code %iUp%"
if "%iUp%"=="0" echo Ollama was updated.
if not "%iUp%"=="0" echo Ollama is already the newest winget offers.
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
