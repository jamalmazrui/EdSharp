@echo off
rem installPython.cmd -- part of EdSharp setup, Homer Tools pattern: probe first,
rem update when present, install when absent, pause on failure so the
rem message can be read. Tool output stays IN THIS WINDOW so progress
rem is readable with a screen reader; the consolidated log records
rem milestones and exit codes.
rem 64-bit by rule: every winget call asks for the x64 build, and where a
rem package offers a machine-wide install it is taken, so components land in
rem the default Windows places -- Program Files, and the PATH every program
rem inherits -- rather than in a per-user corner EdSharp would have to hunt
rem for. The python.org installer defaults to a per-user install; --scope
rem machine asks for the all-users install under Program Files instead, so
rem python.exe sits on the machine PATH where Compile finds it.
setlocal
set "logFile=%LOCALAPPDATA%\EdSharp\logs\EdSharp_setup.log"
if not exist "%LOCALAPPDATA%\EdSharp\logs" mkdir "%LOCALAPPDATA%\EdSharp\logs" >nul 2>&1
echo [installPython.cmd] started %date% %time% >> "%logFile%"
echo If Windows asks permission during an install or update, a User Account
echo Control prompt appears on a separate screen; press Alt+Y to allow it.
echo A large download can also run quietly for several minutes.
echo.

where python >nul 2>&1
if errorlevel 1 goto install_python
echo Python 3 is already installed; checking for an update.
echo [installPython.cmd] winget upgrade Python.Python.3.13 >> "%logFile%"
winget upgrade --id Python.Python.3.13 -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements
echo [installPython.cmd] winget upgrade Python.Python.3.13 exit %errorlevel% >> "%logFile%"
if errorlevel 1 (echo Python 3 is already current.) else (echo Python 3 updated.)
goto after_python
:install_python
echo Installing Python 3 with winget; this can take a few minutes.
echo [installPython.cmd] winget install Python.Python.3.13 >> "%logFile%"
winget install --id Python.Python.3.13 -e --architecture x64 --scope machine --silent --disable-interactivity --accept-package-agreements --accept-source-agreements
echo [installPython.cmd] winget install Python.Python.3.13 exit %errorlevel% >> "%logFile%"
if errorlevel 1 goto fail_python
goto after_python
:fail_python
echo The Python 3 install did not finish. The log is:
echo %logFile%
echo [installPython.cmd] FAILED: Python.Python.3.13 >> "%logFile%"
pause
exit /b 3
:after_python

echo Done. Python is ready.
echo [installPython.cmd] done >> "%logFile%"
exit /b 0
