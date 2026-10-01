@echo off
rem uiCheck.cmd -- run a program, press the keys, and check what actually appeared.
rem
rem     uiCheck                    every uiTest.inix beside this script
rem     uiCheck --path C:\JobDo    an app's own tests
rem     uiCheck --keep             leave the program running at the end
rem
rem It drives a real program through Windows UI Automation, which is the same
rem interface a screen reader uses to find out what is on screen. A control with
rem no accessible name is invisible to both, so a check that finds nothing is a
rem real finding rather than a limitation.
rem
rem It does NOT hear speech. UI Automation reports that a control exists and what
rem it is called, not what JAWS said. Every report says so.
rem
rem pywinauto does the driving and is installed on first run. Nothing else is
rem needed.
rem
rem Writes evidence-ui-<yyyymmdd-hhmmss>.md and uiCheck.log beside this script.
rem Exit code 0 when nothing failed, 1 when something did.
setlocal
where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python was not found on the PATH. Install it from python.org
    echo        or with: winget install Python.Python.3.12
    endlocal
    exit /b 1
)
python "%~dp0uiCheck.py" %*
set exitCode=%errorlevel%
endlocal & exit /b %exitCode%
