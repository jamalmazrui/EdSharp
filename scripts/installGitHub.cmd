@echo off
setlocal enabledelayedexpansion
set "sScript=%~n0"
set "sCallerDir=%~dp0"
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
rem installGitHub.cmd -- part of EdSharp setup, Homer Tools pattern: probe first,
rem update when present, install when absent, pause on failure so the
rem reason is logged. NOTHING PAUSES: a console waiting for a keypress
rem interrupts the installation, and the summary shown at the very end --
rem after every checkbox has run -- is where the outcome is reported. The console says only what is happening in a few plain
rem words; every detail goes to the consolidated log.
rem 64-bit by rule: every winget call asks for the x64 build, and where a
rem package offers a machine-wide install it is taken, so components land in
rem the default Windows places -- Program Files, and the PATH every program
rem inherits -- rather than in a per-user corner EdSharp would have to hunt
rem for.
echo.

where git >nul 2>&1
if errorlevel 1 goto install_git
echo Updating Git
echo [installGitHub.cmd] winget upgrade Git.Git >> "%log%"
winget upgrade --id Git.Git -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements >> "%log%" 2>&1
echo [installGitHub.cmd] winget upgrade Git.Git exit %errorlevel% >> "%log%"
if errorlevel 1 (echo Already current.) else (echo Updated.)
goto after_git
:install_git
echo Installing Git
echo [installGitHub.cmd] winget install Git.Git >> "%log%"
winget install --id Git.Git -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements >> "%log%" 2>&1
echo [installGitHub.cmd] winget install Git.Git exit %errorlevel% >> "%log%"
if errorlevel 1 goto fail_git
goto after_git
:fail_git
echo The Git install did not finish. The log is:
echo %log%
echo [installGitHub.cmd] FAILED: Git.Git >> "%log%"
exit /b 3
:after_git

where gh >nul 2>&1
if errorlevel 1 goto install_gh
echo Updating the GitHub command line
echo [installGitHub.cmd] winget upgrade GitHub.cli >> "%log%"
winget upgrade --id GitHub.cli -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements >> "%log%" 2>&1
echo [installGitHub.cmd] winget upgrade GitHub.cli exit %errorlevel% >> "%log%"
if errorlevel 1 (echo Already current.) else (echo Updated.)
goto after_gh
:install_gh
echo Installing the GitHub command line
echo [installGitHub.cmd] winget install GitHub.cli >> "%log%"
winget install --id GitHub.cli -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements >> "%log%" 2>&1
echo [installGitHub.cmd] winget install GitHub.cli exit %errorlevel% >> "%log%"
if errorlevel 1 goto fail_gh
goto after_gh
:fail_gh
echo The The GitHub command line install did not finish. The log is:
echo %log%
echo [installGitHub.cmd] FAILED: GitHub.cli >> "%log%"
exit /b 3
:after_gh

echo Done.
echo [installGitHub.cmd] done >> "%log%"
exit /b 0
