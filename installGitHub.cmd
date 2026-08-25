@echo off
rem installGitHub.cmd -- part of EdSharp setup, Homer Tools pattern: probe first,
rem update when present, install when absent, pause on failure so the
rem message can be read. Tool output stays IN THIS WINDOW so progress
rem is readable with a screen reader; the consolidated log records
rem milestones and exit codes.
rem 64-bit by rule: every winget call asks for the x64 build, and where a
rem package offers a machine-wide install it is taken, so components land in
rem the default Windows places -- Program Files, and the PATH every program
rem inherits -- rather than in a per-user corner EdSharp would have to hunt
rem for.
setlocal
set "logFile=%LOCALAPPDATA%\EdSharp\logs\EdSharp_setup.log"
if not exist "%LOCALAPPDATA%\EdSharp\logs" mkdir "%LOCALAPPDATA%\EdSharp\logs" >nul 2>&1
echo [installGitHub.cmd] started %date% %time% >> "%logFile%"
echo If Windows asks permission during an install or update, a User Account
echo Control prompt appears on a separate screen; press Alt+Y to allow it.
echo A large download can also run quietly for several minutes.
echo.

where git >nul 2>&1
if errorlevel 1 goto install_git
echo Git is already installed; checking for an update.
echo [installGitHub.cmd] winget upgrade Git.Git >> "%logFile%"
winget upgrade --id Git.Git -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements
echo [installGitHub.cmd] winget upgrade Git.Git exit %errorlevel% >> "%logFile%"
if errorlevel 1 (echo Git is already current.) else (echo Git updated.)
goto after_git
:install_git
echo Installing Git with winget; this can take a few minutes.
echo [installGitHub.cmd] winget install Git.Git >> "%logFile%"
winget install --id Git.Git -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements
echo [installGitHub.cmd] winget install Git.Git exit %errorlevel% >> "%logFile%"
if errorlevel 1 goto fail_git
goto after_git
:fail_git
echo The Git install did not finish. The log is:
echo %logFile%
echo [installGitHub.cmd] FAILED: Git.Git >> "%logFile%"
pause
exit /b 3
:after_git

where gh >nul 2>&1
if errorlevel 1 goto install_gh
echo The GitHub command line is already installed; checking for an update.
echo [installGitHub.cmd] winget upgrade GitHub.cli >> "%logFile%"
winget upgrade --id GitHub.cli -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements
echo [installGitHub.cmd] winget upgrade GitHub.cli exit %errorlevel% >> "%logFile%"
if errorlevel 1 (echo The GitHub command line is already current.) else (echo The GitHub command line updated.)
goto after_gh
:install_gh
echo Installing The GitHub command line with winget; this can take a few minutes.
echo [installGitHub.cmd] winget install GitHub.cli >> "%logFile%"
winget install --id GitHub.cli -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements
echo [installGitHub.cmd] winget install GitHub.cli exit %errorlevel% >> "%logFile%"
if errorlevel 1 goto fail_gh
goto after_gh
:fail_gh
echo The The GitHub command line install did not finish. The log is:
echo %logFile%
echo [installGitHub.cmd] FAILED: GitHub.cli >> "%logFile%"
pause
exit /b 3
:after_gh

echo Done. Git and the GitHub command line are ready.
echo [installGitHub.cmd] done >> "%logFile%"
exit /b 0
