@echo off
rem homerInstall.cmd -- the common half of every Homer install<Component>.cmd.
rem
rem Called, not run: an install script does `call "%~dp0homerInstall.cmd" setup`
rem and then uses what this sets. It exists because the same forty lines had
rem been copied into a dozen scripts across four apps and had drifted in every
rem one -- EdSharp logging to its own folder, HomerScribe logging beside itself
rem inside Program Files, two logging nowhere at all.
rem
rem WHAT IT SETS
rem   sApp      the application name, taken from the folder the caller sits in
rem   sLogDir   %LOCALAPPDATA%\<app>\logs
rem   log       a full path, <App>-<script>-yyyyMMdd-HHmmss.log
rem   sQuiet    1 when the installer is driving and nothing may wait for a key
rem
rem WHAT IT DOES NOT DO
rem   It does not hide the console. The console says briefly what was found and
rem   what was done; this file's log holds the detail, so a failure can be
rem   diagnosed from a file rather than from a window that has closed.

if /i "%~1"=="setup" goto :setup
if /i "%~1"=="log" goto :writeLog
if /i "%~1"=="say" goto :saySo
exit /b 0

:setup
rem The app is the folder the CALLING script lives in, or its parent when that
rem folder is "scripts" -- so one file serves an app whatever it is called.
rem The caller's folder, with its trailing backslash removed so that ~nx
rem gives the folder's own name. If that name is scripts\ or exec\, the app is
rem one level up: resolve with ~f FIRST (which is what turns ".." into a real
rem path) and only then take the name. Taking ~nx of a path that still ends in
rem ".." would give ".." itself, and the log would land in %LOCALAPPDATA%\..\logs
rem -- a folder nobody would think to open.
if not defined sCallerDir set "sCallerDir=%~dp1"
set "sHere=%sCallerDir%"
if "%sHere:~-1%"=="\" set "sHere=%sHere:~0,-1%"
for %%d in ("%sHere%") do set "sApp=%%~nxd"
for %%d in ("%sHere%\..") do set "sParent=%%~fd"
for %%d in ("%sParent%") do set "sParentName=%%~nxd"
if /i "%sApp%"=="scripts" set "sApp=%sParentName%"
if /i "%sApp%"=="exec" set "sApp=%sParentName%"
if not defined sApp set "sApp=Homer"

set "sLogDir=%LOCALAPPDATA%\%sApp%\logs"
if not exist "%sLogDir%" mkdir "%sLogDir%" >nul 2>&1

rem WMIC IS GONE from Windows 11, so this used to fall through to zeros and
rem every log took the same name, overwriting the run before -- the very thing
rem the timestamp exists to prevent. PowerShell's Get-Date is always there.
for /f %%T in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd-HHmmss" 2^>nul') do set "sNow=%%T"
if not defined sNow set "sNow=unknown-time"
set "log=%sLogDir%\%sApp%-%sScript%-%sNow%.log"

rem QUIET means: nothing may wait for a key. It is set when the installer
rem passes "noPause" (Inno runs these with no console to read from, so a pause
rem would return at once or hang), or when HOMER_QUIET is in the environment.
rem Every script tests %noPause% before it pauses, and this is the one place
rem it is worked out.
set "sQuiet="
if defined HOMER_QUIET set "sQuiet=1"
for %%A in (%*) do if /i "%%A"=="noPause" set "sQuiet=1"
if defined sQuiet set "noPause=1"

call "%~f0" log "%sScript% started %DATE% %TIME%"
call "%~f0" log "Script: %sCallerDir%%sScript%.cmd"
call "%~f0" log "App: %sApp%"
call "%~f0" log "Log: %log%"
call "%~f0" log "Quiet: %sQuiet%"
exit /b 0

:writeLog
rem Always appends. A single > here would erase the header written above, and
rem the log would arrive missing exactly the part that says what machine it
rem ran on.
echo %~2>> "%log%" 2>nul
exit /b 0

:saySo
rem One line on the console, and the same line in the log, so the two agree.
echo %~2
echo %~2>> "%log%" 2>nul
exit /b 0
