@echo off
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
setlocal
set "logFile=%LOCALAPPDATA%\EdSharp\logs\EdSharp_setup.log"
if not exist "%LOCALAPPDATA%\EdSharp\logs" mkdir "%LOCALAPPDATA%\EdSharp\logs" >nul 2>&1
echo [installGitHub.cmd] started %date% %time% >> "%logFile%"
echo.

where git >nul 2>&1
if errorlevel 1 goto install_git
echo Updating Git ...
echo [installGitHub.cmd] winget upgrade Git.Git >> "%logFile%"
winget upgrade --id Git.Git -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements >> "%logFile%" 2>&1
echo [installGitHub.cmd] winget upgrade Git.Git exit %errorlevel% >> "%logFile%"
if errorlevel 1 (echo Already current.) else (echo Updated.)
goto after_git
:install_git
echo Installing Git ...
echo [installGitHub.cmd] winget install Git.Git >> "%logFile%"
winget install --id Git.Git -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements >> "%logFile%" 2>&1
echo [installGitHub.cmd] winget install Git.Git exit %errorlevel% >> "%logFile%"
if errorlevel 1 goto fail_git
goto after_git
:fail_git
echo The Git install did not finish. The log is:
echo %logFile%
echo [installGitHub.cmd] FAILED: Git.Git >> "%logFile%"
exit /b 3
:after_git

where gh >nul 2>&1
if errorlevel 1 goto install_gh
echo Updating the GitHub command line ...
echo [installGitHub.cmd] winget upgrade GitHub.cli >> "%logFile%"
winget upgrade --id GitHub.cli -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements >> "%logFile%" 2>&1
echo [installGitHub.cmd] winget upgrade GitHub.cli exit %errorlevel% >> "%logFile%"
if errorlevel 1 (echo Already current.) else (echo Updated.)
goto after_gh
:install_gh
echo Installing the GitHub command line ...
echo [installGitHub.cmd] winget install GitHub.cli >> "%logFile%"
winget install --id GitHub.cli -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements >> "%logFile%" 2>&1
echo [installGitHub.cmd] winget install GitHub.cli exit %errorlevel% >> "%logFile%"
if errorlevel 1 goto fail_gh
goto after_gh
:fail_gh
echo The The GitHub command line install did not finish. The log is:
echo %logFile%
echo [installGitHub.cmd] FAILED: GitHub.cli >> "%logFile%"
exit /b 3
:after_gh

echo Done.
echo [installGitHub.cmd] done >> "%logFile%"
exit /b 0
