@echo off
rem ============================================================
rem  tagRelease.cmd  -  Launcher for tagRelease.ps1.
rem
rem  Generic: drop tagRelease.cmd and tagRelease.ps1 into any repo
rem  root (next to its single ProgramName_setup.exe and matching
rem  ProgramName_setup.iss) and run this. The PowerShell script
rem  discovers the program, version, and GitHub owner/repo itself.
rem
rem  Invokes PowerShell with execution policy bypass for this single
rem  invocation only (does not change the system policy), forwarding
rem  any arguments. The PowerShell script writes a fresh
rem  .\tagRelease.log on every run via Start-Transcript.
rem
rem  Usage:
rem    tagRelease.cmd                publishes; warns on a dirty tree
rem    tagRelease.cmd -StrictTree    bails if the working tree is not clean
rem ============================================================

setlocal

set "sScriptDir=%~dp0"
set "sPsScript=%sScriptDir%tagRelease.ps1"

if not exist "%sPsScript%" (
    echo ERROR: Cannot find tagRelease.ps1 next to tagRelease.cmd
    echo Expected at: %sPsScript%
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%sPsScript%" %*
exit /b %ERRORLEVEL%
