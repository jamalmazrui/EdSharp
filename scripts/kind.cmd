@echo off
rem kind.cmd -- say what kind of Homer resource this folder is: an app, a
rem collection, a kit or a page, and the fact that decided it. Every kit
rem script asks the same question first; kind.py says how it is answered.
rem
rem     kind                      the project in this folder
rem     kind C:\BlindVibeCoding   another folder
rem     kind --word               just the word, for a script to read
setlocal
where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python was not found on the PATH. Install it with: winget install Python.Python.3.12
    endlocal
    exit /b 1
)
python "%~dp0kind.py" %*
set exitCode=%errorlevel%
endlocal & exit /b %exitCode%
