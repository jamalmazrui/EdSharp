@echo off
rem checkHomerApp.cmd -- gather evidence about a Homer app, without looking at it.
rem
rem A sighted developer checks by looking. A blind developer checks by
rem instrumenting, and in AI-assisted development that is not an accommodation:
rem nobody can eyeball generated code fast enough to trust it.
rem
rem This runs every check that can answer yes or no, records the command and the
rem exit code for each, and writes an evidence report saying what was verified,
rem what was not checked, and what remains uncertain. The last list is the
rem honest part.
rem
rem     checkHomerApp                    check the app in this folder
rem     checkHomerApp --path C:\JobDo    check another one
rem     checkHomerApp --build            build it first, and count that as evidence
rem     checkHomerApp --quiet            the report only
rem
rem Acceptance criteria belong to the app: put them in accept.inix beside the
rem source and this script runs them. templates\accept.inix in the kit is the
rem starting point.
rem
rem Writes evidence-<yyyymmdd-hhmmss>.md and checkHomerApp.log beside this
rem script. Exit code 0 when nothing failed, 1 when something did, so a build
rem script or a scheduled task can act on it.
setlocal
where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python was not found on the PATH. Install it from python.org
    echo        or with: winget install Python.Python.3.12
    endlocal
    exit /b 1
)
python "%~dp0checkHomerApp.py" %*
set exitCode=%errorlevel%
endlocal & exit /b %exitCode%
