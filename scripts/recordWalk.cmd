@echo off
rem recordWalk.cmd -- runs recordWalk.py, passing every argument through.
setlocal
set "sScript=%~dp0recordWalk.py"
where py >nul 2>nul
if not errorlevel 1 goto usePy
where python >nul 2>nul
if not errorlevel 1 goto usePython
echo Python was not found. Installing it with winget.
winget install -e --id Python.Python.3.13 --accept-package-agreements --accept-source-agreements
if errorlevel 1 echo winget could not install Python.& exit /b 2
:usePy
py -3 "%sScript%" %*
exit /b %errorlevel%
:usePython
python "%sScript%" %*
exit /b %errorlevel%
