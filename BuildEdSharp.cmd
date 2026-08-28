@echo off
rem BuildEdSharp.cmd -- wrapper so BuildEdSharp.ps1 runs without typing
rem execution-policy parameters. Arguments are passed straight through.
rem The detailed log is BuildEdSharp.log beside these scripts, rewritten
rem fresh on every run.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0BuildEdSharp.ps1" %*
set iExit=%errorlevel%
if not exist "%~dp0BuildEdSharp.log" echo BuildEdSharp.cmd: PowerShell failed before any log could be written. Exit code %iExit%.
if not %iExit%==0 echo BuildEdSharp.cmd: build FAILED with exit code %iExit%. See BuildEdSharp.log.
if %iExit%==0 echo BuildEdSharp.cmd: build succeeded. See BuildEdSharp.log for details.
exit /b %iExit%
