@echo off
rem installNode.cmd -- part of EdSharp setup, Homer Tools pattern: probe first,
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
echo [installNode.cmd] started %date% %time% >> "%logFile%"
echo.

where node >nul 2>&1
if errorlevel 1 goto install_node
echo Updating Node.js
echo [installNode.cmd] winget upgrade OpenJS.NodeJS.LTS >> "%logFile%"
winget upgrade --id OpenJS.NodeJS.LTS -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements >> "%logFile%" 2>&1
echo [installNode.cmd] winget upgrade OpenJS.NodeJS.LTS exit %errorlevel% >> "%logFile%"
if errorlevel 1 (echo Already current.) else (echo Updated.)
goto after_node
:install_node
echo Installing Node.js
echo [installNode.cmd] winget install OpenJS.NodeJS.LTS >> "%logFile%"
winget install --id OpenJS.NodeJS.LTS -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements >> "%logFile%" 2>&1
echo [installNode.cmd] winget install OpenJS.NodeJS.LTS exit %errorlevel% >> "%logFile%"
if errorlevel 1 goto fail_node
goto after_node
:fail_node
echo The Node.js LTS install did not finish. The log is:
echo %logFile%
echo [installNode.cmd] FAILED: OpenJS.NodeJS.LTS >> "%logFile%"
exit /b 3
:after_node

echo Done.
echo [installNode.cmd] done >> "%logFile%"
exit /b 0
