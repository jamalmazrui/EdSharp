@echo off
rem ============================================================
rem  tagRelease.cmd  -  Launcher for tagRelease.ps1.
rem
rem  Generic: drop tagRelease.cmd and tagRelease.ps1 into any repo
rem  root (e.g. C:\EdSharp, C:\FileDir, C:\DbDo) and run this. The
rem  PowerShell script takes the app name from the directory name,
rem  finds the matching <App>_setup.iss, handles the version number,
rem  and publishes the GitHub release.
rem
rem  Invokes PowerShell with execution policy bypass for this single
rem  invocation only (does not change the system policy), forwarding
rem  any arguments. A fresh .\tagRelease.log is written on every run.
rem
rem  Usage:
rem    tagRelease.cmd                  the normal command; no flags needed.
rem                                    Bumps the version only if the current one
rem                                    was already released, so running it again
rem                                    after a rebuild publishes that same version.
rem    tagRelease.cmd -Version 5.1     set an explicit version
rem    tagRelease.cmd -NoBump          never bump, even if already released
rem    tagRelease.cmd -PrepareOnly     update the version files only, no release
rem    tagRelease.cmd -SkipStaleCheck  publish even if the installer looks stale
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
